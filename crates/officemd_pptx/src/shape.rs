use std::{
    collections::BTreeMap,
    io::{Cursor, Read, Write},
};

use officemd_core::opc::OpcPackage;
use quick_xml::{
    Reader, Writer,
    events::{BytesStart, BytesText, Event},
};

use crate::{PptxError, extract::resolve_slide_parts};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxShapeInfo {
    pub slide_number: u32,
    pub shape_id: u32,
    pub text: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxCanvasOverflow {
    pub slide_number: u32,
    pub shape_id: u32,
}

/// Find top-level shapes whose declared bounds extend beyond the slide canvas.
/// Grouped shapes are omitted because their coordinates require group transforms.
///
/// # Errors
///
/// Returns an error when the presentation size or slide XML cannot be parsed.
pub fn inspect_canvas_overflows(content: &[u8]) -> Result<Vec<PptxCanvasOverflow>, PptxError> {
    let paths = resolve_slide_parts(content)?;
    let mut package = OpcPackage::from_bytes(content).map_err(PptxError::from)?;
    let presentation = package
        .read_part_bytes("ppt/presentation.xml")
        .map_err(PptxError::from)?
        .ok_or_else(|| PptxError::MissingPart("ppt/presentation.xml".to_string()))?;
    let (canvas_width, canvas_height) = parse_slide_size(&presentation)?;
    let mut overflows = Vec::new();
    for (number, path) in paths {
        let xml = package
            .read_part_bytes(&path)
            .map_err(PptxError::from)?
            .ok_or_else(|| PptxError::MissingPart(path.clone()))?;
        overflows.extend(parse_canvas_overflows(
            &xml,
            u32::try_from(number).unwrap_or(u32::MAX),
            canvas_width,
            canvas_height,
        )?);
    }
    Ok(overflows)
}

fn parse_slide_size(xml: &[u8]) -> Result<(i64, i64), PptxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref event) | Event::Empty(ref event))
                if local_name(event.name().as_ref()) == b"sldSz" =>
            {
                let width = attr_i64(event, b"cx");
                let height = attr_i64(event, b"cy");
                return width.zip(height).ok_or_else(|| {
                    PptxError::Xml("presentation slide size is incomplete".to_string())
                });
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(PptxError::Xml(error.to_string())),
        }
        buffer.clear();
    }
    Err(PptxError::Xml(
        "presentation slide size is missing".to_string(),
    ))
}

fn parse_canvas_overflows(
    xml: &[u8],
    slide_number: u32,
    canvas_width: i64,
    canvas_height: i64,
) -> Result<Vec<PptxCanvasOverflow>, PptxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut group_depth = 0usize;
    let mut in_shape = false;
    let mut in_transform = false;
    let mut shape_id = None;
    let mut x = None;
    let mut y = None;
    let mut width = None;
    let mut height = None;
    let mut overflows = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref event)) if local_name(event.name().as_ref()) == b"grpSp" => {
                group_depth += 1;
            }
            Ok(Event::End(ref event)) if local_name(event.name().as_ref()) == b"grpSp" => {
                group_depth = group_depth.saturating_sub(1);
            }
            Ok(Event::Start(ref event))
                if group_depth == 0 && local_name(event.name().as_ref()) == b"sp" =>
            {
                in_shape = true;
                in_transform = false;
                shape_id = None;
                x = None;
                y = None;
                width = None;
                height = None;
            }
            Ok(Event::Start(ref event) | Event::Empty(ref event))
                if in_shape && local_name(event.name().as_ref()) == b"cNvPr" =>
            {
                shape_id = attr_u32(event, b"id");
            }
            Ok(Event::Start(ref event))
                if in_shape && local_name(event.name().as_ref()) == b"xfrm" =>
            {
                in_transform = true;
            }
            Ok(Event::Start(ref event) | Event::Empty(ref event))
                if in_transform && local_name(event.name().as_ref()) == b"off" =>
            {
                x = attr_i64(event, b"x");
                y = attr_i64(event, b"y");
            }
            Ok(Event::Start(ref event) | Event::Empty(ref event))
                if in_transform && local_name(event.name().as_ref()) == b"ext" =>
            {
                width = attr_i64(event, b"cx");
                height = attr_i64(event, b"cy");
            }
            Ok(Event::End(ref event)) if local_name(event.name().as_ref()) == b"xfrm" => {
                in_transform = false;
            }
            Ok(Event::End(ref event)) if in_shape && local_name(event.name().as_ref()) == b"sp" => {
                if let (Some(shape_id), Some(x), Some(y), Some(width), Some(height)) =
                    (shape_id, x, y, width, height)
                    && (x < 0
                        || y < 0
                        || x.saturating_add(width) > canvas_width
                        || y.saturating_add(height) > canvas_height)
                {
                    overflows.push(PptxCanvasOverflow {
                        slide_number,
                        shape_id,
                    });
                }
                in_shape = false;
                in_transform = false;
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(PptxError::Xml(error.to_string())),
        }
        buffer.clear();
    }
    Ok(overflows)
}

/// Inspect text-bearing shapes using stable OOXML non-visual shape IDs.
///
/// # Errors
///
/// Returns an error when the package or selected slide XML is invalid.
pub fn inspect_shapes(
    content: &[u8],
    start: u32,
    end: u32,
) -> Result<Vec<PptxShapeInfo>, PptxError> {
    if start == 0 || end < start {
        return Err(PptxError::Xml(
            "slide range must be 1-based and ordered".to_string(),
        ));
    }
    let paths = resolve_slide_parts(content)?;
    let mut package = OpcPackage::from_bytes(content).map_err(PptxError::from)?;
    let mut shapes = Vec::new();
    for (number, path) in paths {
        let number = u32::try_from(number).unwrap_or(u32::MAX);
        if number < start || number > end {
            continue;
        }
        let xml = package
            .read_part_bytes(&path)
            .map_err(PptxError::from)?
            .ok_or_else(|| PptxError::MissingPart(path.clone()))?;
        shapes.extend(parse_shapes(&xml, number)?);
    }
    Ok(shapes)
}

/// Replace the complete text of one shape after checking its exact expected text.
///
/// # Errors
///
/// Returns an error when the locator is absent, ambiguous, or has unexpected text.
pub fn replace_shape_text(
    content: &[u8],
    slide_number: u32,
    shape_id: u32,
    expected_text: &str,
    replacement: &str,
) -> Result<Vec<u8>, PptxError> {
    let target_path = resolve_slide_parts(content)?
        .into_iter()
        .find(|(number, _)| u32::try_from(*number).ok() == Some(slide_number))
        .map(|(_, path)| path)
        .ok_or_else(|| PptxError::Xml(format!("slide not found: {slide_number}")))?;
    let mut parts = read_zip_parts(content)?;
    let xml = parts
        .get(&target_path)
        .ok_or_else(|| PptxError::MissingPart(target_path.clone()))?;
    let matches = parse_shapes(xml, slide_number)?
        .into_iter()
        .filter(|shape| shape.shape_id == shape_id)
        .collect::<Vec<_>>();
    if matches.len() != 1 {
        return Err(PptxError::Xml(format!(
            "shape locator resolved to {} shapes on slide {slide_number}: {shape_id}",
            matches.len()
        )));
    }
    if matches[0].text != expected_text {
        return Err(PptxError::Xml(format!(
            "shape text precondition failed on slide {slide_number}, shape {shape_id}"
        )));
    }
    let rewritten = rewrite_shape(xml, shape_id, replacement)?;
    parts.insert(target_path, rewritten);
    write_zip_parts(parts)
}

fn parse_shapes(xml: &[u8], slide_number: u32) -> Result<Vec<PptxShapeInfo>, PptxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut in_shape = false;
    let mut in_text = false;
    let mut shape_id = None;
    let mut text = String::new();
    let mut shapes = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref event)) if local_name(event.name().as_ref()) == b"sp" => {
                in_shape = true;
                shape_id = None;
                text.clear();
            }
            Ok(Event::Start(ref event) | Event::Empty(ref event))
                if in_shape && local_name(event.name().as_ref()) == b"cNvPr" =>
            {
                shape_id = attr_u32(event, b"id");
            }
            Ok(Event::Start(ref event))
                if in_shape && local_name(event.name().as_ref()) == b"t" =>
            {
                in_text = true;
            }
            Ok(Event::Text(event)) if in_text => {
                text.push_str(
                    &event
                        .unescape()
                        .map_err(|error| PptxError::Xml(error.to_string()))?,
                );
            }
            Ok(Event::CData(event)) if in_text => {
                text.push_str(&String::from_utf8_lossy(event.as_ref()));
            }
            Ok(Event::End(ref event)) if local_name(event.name().as_ref()) == b"t" => {
                in_text = false;
            }
            Ok(Event::End(ref event)) if local_name(event.name().as_ref()) == b"sp" => {
                if let Some(shape_id) = shape_id.take()
                    && !text.is_empty()
                {
                    shapes.push(PptxShapeInfo {
                        slide_number,
                        shape_id,
                        text: std::mem::take(&mut text),
                    });
                }
                in_shape = false;
                in_text = false;
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(PptxError::Xml(error.to_string())),
        }
        buffer.clear();
    }
    Ok(shapes)
}

fn rewrite_shape(
    xml: &[u8],
    target_shape_id: u32,
    replacement: &str,
) -> Result<Vec<u8>, PptxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::with_capacity(
        xml.len() + replacement.len(),
    )));
    let mut buffer = Vec::new();
    let mut in_shape = false;
    let mut target_shape = false;
    let mut in_text = false;
    let mut replacement_written = false;
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| PptxError::Xml(error.to_string()))?;
        match event {
            Event::Start(ref start) if local_name(start.name().as_ref()) == b"sp" => {
                in_shape = true;
                target_shape = false;
                writer.write_event(event.into_owned())?;
            }
            Event::Start(ref start) | Event::Empty(ref start)
                if in_shape && local_name(start.name().as_ref()) == b"cNvPr" =>
            {
                target_shape = attr_u32(start, b"id") == Some(target_shape_id);
                writer.write_event(event.into_owned())?;
            }
            Event::Start(ref start)
                if target_shape && local_name(start.name().as_ref()) == b"t" =>
            {
                in_text = true;
                writer.write_event(event.into_owned())?;
            }
            Event::Text(_) | Event::CData(_) if in_text && target_shape => {
                if !replacement_written {
                    writer.write_event(Event::Text(BytesText::new(replacement)))?;
                    replacement_written = true;
                }
            }
            Event::End(ref end) if local_name(end.name().as_ref()) == b"t" => {
                in_text = false;
                writer.write_event(event.into_owned())?;
            }
            Event::End(ref end) if local_name(end.name().as_ref()) == b"sp" => {
                in_shape = false;
                target_shape = false;
                writer.write_event(event.into_owned())?;
            }
            Event::Eof => break,
            _ => writer.write_event(event.into_owned())?,
        }
        buffer.clear();
    }
    if !replacement_written {
        return Err(PptxError::Xml(format!(
            "shape has no replaceable text node: {target_shape_id}"
        )));
    }
    Ok(writer.into_inner().into_inner())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.iter()
        .rposition(|byte| matches!(*byte, b':' | b'}'))
        .map_or(name, |index| &name[index + 1..])
}

fn attr_u32(element: &BytesStart<'_>, key: &[u8]) -> Option<u32> {
    element
        .attributes()
        .with_checks(false)
        .flatten()
        .find_map(|attribute| {
            (local_name(attribute.key.as_ref()) == key)
                .then(|| String::from_utf8_lossy(&attribute.value).parse().ok())
                .flatten()
        })
}

fn attr_i64(element: &BytesStart<'_>, key: &[u8]) -> Option<i64> {
    element
        .attributes()
        .with_checks(false)
        .flatten()
        .find_map(|attribute| {
            (local_name(attribute.key.as_ref()) == key)
                .then(|| String::from_utf8_lossy(&attribute.value).parse().ok())
                .flatten()
        })
}

fn read_zip_parts(content: &[u8]) -> Result<BTreeMap<String, Vec<u8>>, PptxError> {
    let mut archive = zip::ZipArchive::new(Cursor::new(content)).map_err(PptxError::from)?;
    let mut parts = BTreeMap::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index(index).map_err(PptxError::from)?;
        if file.is_dir() {
            continue;
        }
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes).map_err(PptxError::from)?;
        parts.insert(file.name().to_string(), bytes);
    }
    Ok(parts)
}

fn write_zip_parts(parts: BTreeMap<String, Vec<u8>>) -> Result<Vec<u8>, PptxError> {
    let mut output = Vec::new();
    {
        let mut archive = zip::ZipWriter::new(Cursor::new(&mut output));
        let options: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default();
        for (name, bytes) in parts {
            archive.start_file(name, options).map_err(PptxError::from)?;
            archive.write_all(&bytes).map_err(PptxError::from)?;
        }
        archive.finish().map_err(PptxError::from)?;
    }
    Ok(output)
}
