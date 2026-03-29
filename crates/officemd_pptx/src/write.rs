//! Generate PPTX files from the officemd IR.
//!
//! Converts an [`OoxmlDocument`] with `kind: Pptx` into a valid `.pptx` ZIP
//! archive that opens in Microsoft PowerPoint and LibreOffice Impress.

use std::fmt::Write as _;

use officemd_core::ir::{Block, Inline, OoxmlDocument, Paragraph, Slide, Table};
use officemd_core::opc::writer::{OpcWriter, RelEntry, xml_escape_attr, xml_escape_text};

const REL_TYPE_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

use crate::error::PptxError;

// --- OOXML constants ---

const CT_PRESENTATION: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const CT_SLIDE: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const CT_SLIDE_LAYOUT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const CT_SLIDE_MASTER: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";

const REL_TYPE_OFFICE_DOC: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_TYPE_SLIDE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_TYPE_SLIDE_LAYOUT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const REL_TYPE_SLIDE_MASTER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const REL_TYPE_NOTES_SLIDE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const CT_NOTES_SLIDE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";

const NS_P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Generate a `.pptx` file from an officemd IR document.
///
/// Each `Slide` in the IR becomes a slide in the presentation.
/// A minimal slide master and layout are included for compatibility.
///
/// # Errors
///
/// Returns an error if ZIP assembly fails.
pub fn generate_pptx(doc: &OoxmlDocument) -> Result<Vec<u8>, PptxError> {
    let mut w = OpcWriter::new();
    w.register_content_type_default(
        "rels",
        "application/vnd.openxmlformats-package.relationships+xml",
    );
    w.register_content_type_default("xml", "application/xml");
    w.register_content_type_override("/ppt/presentation.xml", CT_PRESENTATION);

    let mut pres_rels: Vec<RelEntry> = Vec::new();
    let mut rel_counter: usize = 1;

    // Add minimal slide master and layout
    w.add_xml_part("ppt/slideMasters/slideMaster1.xml", MINIMAL_SLIDE_MASTER)?;
    w.register_content_type_override("/ppt/slideMasters/slideMaster1.xml", CT_SLIDE_MASTER);

    w.add_xml_part("ppt/slideLayouts/slideLayout1.xml", MINIMAL_SLIDE_LAYOUT)?;
    w.register_content_type_override("/ppt/slideLayouts/slideLayout1.xml", CT_SLIDE_LAYOUT);

    // Slide master rels: link to layout
    w.add_part_rels(
        "ppt/slideMasters/slideMaster1.xml",
        &[RelEntry {
            id: "rId1".into(),
            rel_type: REL_TYPE_SLIDE_LAYOUT.to_string(),
            target: "../slideLayouts/slideLayout1.xml".to_string(),
            target_mode: None,
        }],
    )?;

    // Layout rels: link back to master
    w.add_part_rels(
        "ppt/slideLayouts/slideLayout1.xml",
        &[RelEntry {
            id: "rId1".into(),
            rel_type: REL_TYPE_SLIDE_MASTER.to_string(),
            target: "../slideMasters/slideMaster1.xml".to_string(),
            target_mode: None,
        }],
    )?;

    // Presentation rels: master
    let master_rid = format!("rId{rel_counter}");
    pres_rels.push(RelEntry {
        id: master_rid,
        rel_type: REL_TYPE_SLIDE_MASTER.to_string(),
        target: "slideMasters/slideMaster1.xml".to_string(),
        target_mode: None,
    });
    rel_counter += 1;

    // Build slides
    let mut slide_ids: Vec<(u32, String)> = Vec::new(); // (id, rId)

    for (i, slide) in doc.slides.iter().enumerate() {
        let slide_num = i + 1;
        let slide_path = format!("ppt/slides/slide{slide_num}.xml");
        let (slide_xml, hlink_rels) = build_slide_xml(slide);

        w.register_content_type_override(&format!("/{slide_path}"), CT_SLIDE);
        w.add_xml_part(&slide_path, &slide_xml)?;

        // Slide relationships: layout + optional notes + hyperlinks
        let mut slide_rels = vec![RelEntry {
            id: "rId1".into(),
            rel_type: REL_TYPE_SLIDE_LAYOUT.to_string(),
            target: "../slideLayouts/slideLayout1.xml".to_string(),
            target_mode: None,
        }];

        // Notes slide
        if let Some(notes) = &slide.notes
            && !notes.is_empty()
        {
            let notes_path = format!("ppt/notesSlides/notesSlide{slide_num}.xml");
            let notes_xml = build_notes_slide_xml(notes);
            w.register_content_type_override(&format!("/{notes_path}"), CT_NOTES_SLIDE);
            w.add_xml_part(&notes_path, &notes_xml)?;
            slide_rels.push(RelEntry {
                id: "rId2".into(),
                rel_type: REL_TYPE_NOTES_SLIDE.to_string(),
                target: format!("../notesSlides/notesSlide{slide_num}.xml"),
                target_mode: None,
            });
        }

        // Hyperlink relationships collected during slide XML generation
        slide_rels.extend(hlink_rels);

        w.add_part_rels(&slide_path, &slide_rels)?;

        let rid = format!("rId{rel_counter}");
        pres_rels.push(RelEntry {
            id: rid.clone(),
            rel_type: REL_TYPE_SLIDE.to_string(),
            target: format!("slides/slide{slide_num}.xml"),
            target_mode: None,
        });

        // Slide IDs start at 256 (PowerPoint convention)
        slide_ids.push((255 + slide_num as u32, rid));
        rel_counter += 1;
    }

    // Build presentation.xml
    let pres_xml = build_presentation_xml(&slide_ids);
    w.add_xml_part("ppt/presentation.xml", &pres_xml)?;
    w.add_part_rels("ppt/presentation.xml", &pres_rels)?;

    // Root relationship
    w.add_root_relationship(RelEntry {
        id: "rId1".to_string(),
        rel_type: REL_TYPE_OFFICE_DOC.to_string(),
        target: "ppt/presentation.xml".to_string(),
        target_mode: None,
    });

    Ok(w.finish()?)
}

// --- XML builders ---

fn build_presentation_xml(slide_ids: &[(u32, String)]) -> String {
    let mut sld_id_lst = String::new();
    for (id, rid) in slide_ids {
        let _ = write!(sld_id_lst, "<p:sldId id=\"{id}\" r:id=\"{rid}\"/>");
    }
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <p:presentation xmlns:p=\"{NS_P}\" xmlns:a=\"{NS_A}\" xmlns:r=\"{NS_R}\">\
         <p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>\
         <p:sldIdLst>{sld_id_lst}</p:sldIdLst>\
         <p:sldSz cx=\"9144000\" cy=\"5143500\" type=\"custom\"/>\
         <p:notesSz cx=\"6858000\" cy=\"9144000\"/>\
         </p:presentation>"
    )
}

/// Build slide XML and return any hyperlink relationships needed.
fn build_slide_xml(slide: &Slide) -> (String, Vec<RelEntry>) {
    let mut shapes = String::new();
    let mut shape_id: u32 = 2; // 1 is reserved for the shape tree
    let mut hlink_rels: Vec<RelEntry> = Vec::new();
    // rId1 is the slide layout relationship; hyperlinks start at rId2
    // (rId2 may be notes if present, but notes rels are added separately)
    let mut rel_counter: usize = 2;

    // Title shape
    if let Some(title) = &slide.title {
        let escaped = xml_escape_text(title);
        let _ = write!(
            shapes,
            "<p:sp>\
             <p:nvSpPr>\
             <p:cNvPr id=\"{shape_id}\" name=\"Title {shape_id}\"/>\
             <p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>\
             <p:nvPr><p:ph type=\"title\"/></p:nvPr>\
             </p:nvSpPr>\
             <p:spPr>\
             <a:xfrm><a:off x=\"457200\" y=\"274638\"/><a:ext cx=\"8229600\" cy=\"1143000\"/></a:xfrm>\
             </p:spPr>\
             <p:txBody>\
             <a:bodyPr/><a:lstStyle/>\
             <a:p><a:r><a:rPr lang=\"en-US\" dirty=\"0\"/>\
             <a:t>{escaped}</a:t></a:r></a:p>\
             </p:txBody>\
             </p:sp>"
        );
        shape_id += 1;
    }

    // Content shape with blocks (skip first paragraph if it duplicates the title)
    let content_blocks: Vec<&Block> = slide
        .blocks
        .iter()
        .enumerate()
        .filter(|(i, b)| {
            // Skip the first block if it's a paragraph whose text matches the title
            if *i == 0
                && slide.title.is_some()
                && let Block::Paragraph(p) = b
            {
                let text: String = p
                    .inlines
                    .iter()
                    .map(|il| match il {
                        Inline::Text(t) => t.as_str(),
                        Inline::Link(l) => l.display.as_str(),
                    })
                    .collect();
                return text.trim() != slide.title.as_deref().unwrap_or("").trim();
            }
            true
        })
        .map(|(_, b)| b)
        .collect();

    if !content_blocks.is_empty() {
        let mut body_paras = String::new();
        for block in &content_blocks {
            write_block_as_drawingml(&mut body_paras, block, &mut hlink_rels, &mut rel_counter);
        }

        let _ = write!(
            shapes,
            "<p:sp>\
             <p:nvSpPr>\
             <p:cNvPr id=\"{shape_id}\" name=\"Content {shape_id}\"/>\
             <p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>\
             <p:nvPr><p:ph idx=\"1\"/></p:nvPr>\
             </p:nvSpPr>\
             <p:spPr>\
             <a:xfrm><a:off x=\"457200\" y=\"1600200\"/><a:ext cx=\"8229600\" cy=\"4525963\"/></a:xfrm>\
             </p:spPr>\
             <p:txBody>\
             <a:bodyPr/><a:lstStyle/>\
             {body_paras}\
             </p:txBody>\
             </p:sp>"
        );
    }

    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <p:sld xmlns:p=\"{NS_P}\" xmlns:a=\"{NS_A}\" xmlns:r=\"{NS_R}\">\
         <p:cSld>\
         <p:spTree>\
         <p:nvGrpSpPr>\
         <p:cNvPr id=\"1\" name=\"\"/>\
         <p:cNvGrpSpPr/><p:nvPr/>\
         </p:nvGrpSpPr>\
         <p:grpSpPr/>\
         {shapes}\
         </p:spTree>\
         </p:cSld>\
         </p:sld>"
    );
    (xml, hlink_rels)
}

fn build_notes_slide_xml(notes: &[Paragraph]) -> String {
    let mut body = String::new();
    let mut unused_rels = Vec::new();
    let mut unused_counter = 1usize;
    for para in notes {
        write_drawingml_paragraph(&mut body, para, &mut unused_rels, &mut unused_counter);
    }
    if body.is_empty() {
        body.push_str("<a:p><a:endParaRPr lang=\"en-US\"/></a:p>");
    }
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <p:notes xmlns:p=\"{NS_P}\" xmlns:a=\"{NS_A}\" xmlns:r=\"{NS_R}\">\
         <p:cSld>\
         <p:spTree>\
         <p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>\
         <p:grpSpPr/>\
         <p:sp>\
         <p:nvSpPr>\
         <p:cNvPr id=\"2\" name=\"Notes Placeholder\"/>\
         <p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>\
         <p:nvPr><p:ph type=\"body\" idx=\"1\"/></p:nvPr>\
         </p:nvSpPr>\
         <p:spPr/>\
         <p:txBody><a:bodyPr/><a:lstStyle/>{body}</p:txBody>\
         </p:sp>\
         </p:spTree>\
         </p:cSld>\
         </p:notes>"
    )
}

fn write_block_as_drawingml(
    out: &mut String,
    block: &Block,
    rels: &mut Vec<RelEntry>,
    rel_counter: &mut usize,
) {
    match block {
        Block::Paragraph(para) => write_drawingml_paragraph(out, para, rels, rel_counter),
        Block::Table(table) => write_drawingml_table(out, table),
        Block::Separator => {
            // Render separator as empty paragraph with a line
            out.push_str("<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>---</a:t></a:r></a:p>");
        }
    }
}

fn write_drawingml_paragraph(
    out: &mut String,
    para: &Paragraph,
    rels: &mut Vec<RelEntry>,
    rel_counter: &mut usize,
) {
    out.push_str("<a:p>");
    for inline in &para.inlines {
        match inline {
            Inline::Text(text) => {
                let escaped = xml_escape_text(text);
                let _ = write!(
                    out,
                    "<a:r><a:rPr lang=\"en-US\" dirty=\"0\"/><a:t>{escaped}</a:t></a:r>"
                );
            }
            Inline::Link(link) => {
                let display = xml_escape_text(&link.display);
                let rid = format!("rId{}", *rel_counter);
                *rel_counter += 1;
                rels.push(RelEntry {
                    id: rid.clone(),
                    rel_type: REL_TYPE_HYPERLINK.to_string(),
                    target: link.target.clone(),
                    target_mode: Some("External".to_string()),
                });
                let escaped_rid = xml_escape_attr(&rid);
                let _ = write!(
                    out,
                    "<a:r><a:rPr lang=\"en-US\" dirty=\"0\" u=\"sng\">\
                     <a:hlinkClick r:id=\"{escaped_rid}\"/>\
                     </a:rPr><a:t>{display}</a:t></a:r>"
                );
            }
        }
    }
    out.push_str("</a:p>");
}

fn write_drawingml_table(out: &mut String, table: &Table) {
    // Render table as text paragraphs (PPTX tables require graphicFrame which
    // is complex). For initial implementation, format as text rows.
    if let Some(caption) = &table.caption {
        let escaped = xml_escape_text(caption);
        let _ = write!(
            out,
            "<a:p><a:r><a:rPr lang=\"en-US\" b=\"1\"/><a:t>{escaped}</a:t></a:r></a:p>"
        );
    }

    // Header row
    if !table.synthetic_headers && !table.headers.is_empty() {
        let header_text = table.headers.join(" | ");
        let escaped = xml_escape_text(&header_text);
        let _ = write!(
            out,
            "<a:p><a:r><a:rPr lang=\"en-US\" b=\"1\"/><a:t>{escaped}</a:t></a:r></a:p>"
        );
    }

    // Data rows
    for row in &table.rows {
        let row_text: String = row
            .iter()
            .map(|cell| {
                cell.content
                    .iter()
                    .map(|p| {
                        p.inlines
                            .iter()
                            .map(|i| match i {
                                Inline::Text(t) => t.as_str(),
                                Inline::Link(l) => l.display.as_str(),
                            })
                            .collect::<String>()
                    })
                    .collect::<Vec<_>>()
                    .join(" ")
            })
            .collect::<Vec<_>>()
            .join(" | ");
        let escaped = xml_escape_text(&row_text);
        let _ = write!(
            out,
            "<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>{escaped}</a:t></a:r></a:p>"
        );
    }
}

// --- Minimal static templates ---

const MINIMAL_SLIDE_MASTER: &str = "\
<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<p:sldMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill>\
<a:effectLst/></p:bgPr></p:bg>\
<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/>\
<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>\
<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" \
accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" \
accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>\
<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst>\
</p:sldMaster>";

const MINIMAL_SLIDE_LAYOUT: &str = "\
<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<p:sldLayout xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" \
type=\"blank\">\
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/>\
<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>\
</p:sldLayout>";

#[cfg(test)]
mod tests {
    use super::*;
    use officemd_core::ir::{
        Block, DocumentKind, Hyperlink, Inline, OoxmlDocument, Paragraph, Slide, Table, TableCell,
    };

    fn simple_para(text: &str) -> Paragraph {
        Paragraph {
            inlines: vec![Inline::Text(text.to_string())],
        }
    }

    fn simple_slide(title: &str, text: &str) -> Slide {
        Slide {
            number: 1,
            title: Some(title.to_string()),
            blocks: vec![Block::Paragraph(simple_para(text))],
            notes: None,
            comments: vec![],
        }
    }

    fn simple_pptx(slides: Vec<Slide>) -> OoxmlDocument {
        OoxmlDocument {
            kind: DocumentKind::Pptx,
            properties: None,
            sheets: vec![],
            slides,
            sections: vec![],
            pdf: None,
        }
    }

    #[test]
    fn generates_valid_pptx_single_slide() {
        let doc = simple_pptx(vec![simple_slide("Hello", "World")]);
        let bytes = generate_pptx(&doc).expect("generate");
        assert!(!bytes.is_empty());

        // Verify it's a valid PPTX with the right structure
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        assert!(pkg.has_part("ppt/presentation.xml"));
        assert!(pkg.has_part("ppt/slides/slide1.xml"));
        assert!(pkg.has_part("ppt/slideMasters/slideMaster1.xml"));
        assert!(pkg.has_part("ppt/slideLayouts/slideLayout1.xml"));
    }

    #[test]
    fn generates_pptx_with_multiple_slides() {
        let doc = simple_pptx(vec![
            simple_slide("Slide 1", "First"),
            simple_slide("Slide 2", "Second"),
            simple_slide("Slide 3", "Third"),
        ]);
        let bytes = generate_pptx(&doc).expect("generate");
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        assert!(pkg.has_part("ppt/slides/slide1.xml"));
        assert!(pkg.has_part("ppt/slides/slide2.xml"));
        assert!(pkg.has_part("ppt/slides/slide3.xml"));

        // Verify round-trip extraction
        let ir = crate::extract_ir(&bytes).expect("extract");
        assert_eq!(ir.slides.len(), 3);
    }

    #[test]
    fn slide_title_survives_round_trip() {
        let doc = simple_pptx(vec![simple_slide("My Title", "Body text")]);
        let bytes = generate_pptx(&doc).expect("generate");
        let ir = crate::extract_ir(&bytes).expect("extract");
        assert_eq!(ir.slides.len(), 1);
        assert_eq!(ir.slides[0].title.as_deref(), Some("My Title"));
    }

    #[test]
    fn generates_pptx_with_table_as_text() {
        let doc = simple_pptx(vec![Slide {
            number: 1,
            title: Some("Data".to_string()),
            blocks: vec![Block::Table(Table {
                caption: None,
                headers: vec!["Name".to_string(), "Value".to_string()],
                rows: vec![vec![
                    TableCell {
                        content: vec![simple_para("A")],
                    },
                    TableCell {
                        content: vec![simple_para("1")],
                    },
                ]],
                synthetic_headers: false,
            })],
            notes: None,
            comments: vec![],
        }]);
        let bytes = generate_pptx(&doc).expect("generate");
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        let slide_xml = pkg
            .read_part_string("ppt/slides/slide1.xml")
            .unwrap()
            .unwrap();
        assert!(slide_xml.contains("Name | Value"));
        assert!(slide_xml.contains("A | 1"));
    }

    #[test]
    fn generates_pptx_with_hyperlink() {
        let doc = simple_pptx(vec![Slide {
            number: 1,
            title: None,
            blocks: vec![Block::Paragraph(Paragraph {
                inlines: vec![Inline::Link(Hyperlink {
                    display: "Click here".to_string(),
                    target: "https://example.com".to_string(),
                    rel_id: None,
                })],
            })],
            notes: None,
            comments: vec![],
        }]);
        let bytes = generate_pptx(&doc).expect("generate");
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        let slide_xml = pkg
            .read_part_string("ppt/slides/slide1.xml")
            .unwrap()
            .unwrap();
        assert!(slide_xml.contains("Click here"));
        assert!(
            slide_xml.contains("a:hlinkClick"),
            "hyperlink should have hlinkClick element"
        );

        // Verify the relationship file contains the hyperlink target
        let rels_xml = pkg
            .read_part_string("ppt/slides/_rels/slide1.xml.rels")
            .unwrap()
            .unwrap();
        assert!(
            rels_xml.contains("https://example.com"),
            "hyperlink URL should be in slide rels"
        );
    }

    #[test]
    fn empty_pptx_is_valid() {
        let doc = simple_pptx(vec![]);
        let bytes = generate_pptx(&doc).expect("generate");
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        assert!(pkg.has_part("ppt/presentation.xml"));
    }

    #[test]
    fn xml_special_chars_in_slide_text() {
        let doc = simple_pptx(vec![simple_slide("A & B", "x < y > z")]);
        let bytes = generate_pptx(&doc).expect("generate");
        // Should not crash and produce valid PPTX
        let ir = crate::extract_ir(&bytes).expect("extract");
        assert_eq!(ir.slides.len(), 1);
    }

    #[test]
    fn generates_pptx_with_notes() {
        let doc = simple_pptx(vec![Slide {
            number: 1,
            title: Some("With Notes".to_string()),
            blocks: vec![Block::Paragraph(simple_para("Content"))],
            notes: Some(vec![simple_para("Speaker notes here")]),
            comments: vec![],
        }]);
        let bytes = generate_pptx(&doc).expect("generate");
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        assert!(pkg.has_part("ppt/notesSlides/notesSlide1.xml"));
        let notes_xml = pkg
            .read_part_string("ppt/notesSlides/notesSlide1.xml")
            .unwrap()
            .unwrap();
        assert!(notes_xml.contains("Speaker notes here"));
    }

    #[test]
    fn slide_has_explicit_shape_positions() {
        let doc = simple_pptx(vec![simple_slide("Title", "Body")]);
        let bytes = generate_pptx(&doc).expect("generate");
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        let xml = pkg
            .read_part_string("ppt/slides/slide1.xml")
            .unwrap()
            .unwrap();
        // Title and content shapes should have explicit positioning
        assert!(
            xml.contains("a:xfrm"),
            "shapes should have explicit transforms"
        );
        assert!(xml.contains("a:off"), "shapes should have explicit offsets");
    }

    #[test]
    fn slide_size_is_16x9() {
        let doc = simple_pptx(vec![simple_slide("Test", "Body")]);
        let bytes = generate_pptx(&doc).expect("generate");
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        let xml = pkg
            .read_part_string("ppt/presentation.xml")
            .unwrap()
            .unwrap();
        // 16:9 = 9144000 x 5143500 EMUs, type should be "custom"
        assert!(xml.contains("cx=\"9144000\""));
        assert!(xml.contains("cy=\"5143500\""));
        assert!(xml.contains("type=\"custom\""));
        assert!(!xml.contains("screen4x3"));
    }
}
