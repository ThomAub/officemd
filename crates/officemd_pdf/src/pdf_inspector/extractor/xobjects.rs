//! Form XObject and image XObject extraction.

use lopdf::{Document, Object, ObjectId, content::Content};

use crate::pdf_inspector::tounicode::FontCMaps;
use crate::pdf_inspector::types::{PdfRect, TextItem};

use super::fonts::CMapDecisionCache;
use super::interpreter::{MarkedContentMode, build_font_context, extract_content_stream};
use super::{get_number, multiply_matrices};

const MAX_FORM_XOBJECT_DEPTH: usize = 16;

pub(crate) enum XObjectType {
    Image,
    Form(ObjectId),
}

fn get_xobjects_from_resources(
    doc: &Document,
    resources: Option<&lopdf::Dictionary>,
) -> std::collections::HashMap<String, XObjectType> {
    let mut xobject_types = std::collections::HashMap::new();
    let Some(resources) = resources else {
        return xobject_types;
    };

    if let Ok(xobjects_ref) = resources.get(b"XObject") {
        let xobjects = if let Ok(obj_ref) = xobjects_ref.as_reference() {
            doc.get_dictionary(obj_ref).ok()
        } else {
            xobjects_ref.as_dict().ok()
        };

        if let Some(xobjects) = xobjects {
            for (name, value) in xobjects.iter() {
                let name_str = String::from_utf8_lossy(name).to_string();

                if let Ok(obj_ref) = value.as_reference() {
                    if let Ok(Object::Stream(stream)) = doc.get_object(obj_ref) {
                        if let Ok(subtype) = stream.dict.get(b"Subtype") {
                            if let Ok(subtype_name) = subtype.as_name() {
                                if subtype_name == b"Image" {
                                    xobject_types.insert(name_str, XObjectType::Image);
                                } else if subtype_name == b"Form" {
                                    xobject_types.insert(name_str, XObjectType::Form(obj_ref));
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    xobject_types
}

/// Get XObjects from page resources, categorized by type
pub(crate) fn get_page_xobjects(
    doc: &Document,
    page_id: ObjectId,
) -> std::collections::HashMap<String, XObjectType> {
    if let Ok(page_dict) = doc.get_dictionary(page_id) {
        let resources = if let Ok(res_ref) = page_dict.get(b"Resources") {
            if let Ok(obj_ref) = res_ref.as_reference() {
                doc.get_dictionary(obj_ref).ok()
            } else {
                res_ref.as_dict().ok()
            }
        } else {
            None
        };
        return get_xobjects_from_resources(doc, resources);
    }

    std::collections::HashMap::new()
}

fn get_form_xobjects(
    doc: &Document,
    form_dict: &lopdf::Dictionary,
) -> std::collections::HashMap<String, XObjectType> {
    let resources = if let Ok(res_ref) = form_dict.get(b"Resources") {
        if let Ok(obj_ref) = res_ref.as_reference() {
            doc.get_dictionary(obj_ref).ok()
        } else {
            res_ref.as_dict().ok()
        }
    } else {
        None
    };

    get_xobjects_from_resources(doc, resources)
}

/// Extract text items and rectangles from a Form XObject.
pub(crate) fn extract_form_xobject_content(
    doc: &Document,
    form_id: ObjectId,
    page_num: u32,
    font_cmaps: &FontCMaps,
    parent_ctm: &[f32; 6],
    cmap_decisions: &mut CMapDecisionCache,
) -> (Vec<TextItem>, Vec<PdfRect>) {
    extract_form_xobject_content_with_depth(
        doc,
        form_id,
        page_num,
        font_cmaps,
        parent_ctm,
        cmap_decisions,
        0,
    )
}

pub(crate) fn extract_form_xobject_content_with_depth(
    doc: &Document,
    form_id: ObjectId,
    page_num: u32,
    font_cmaps: &FontCMaps,
    parent_ctm: &[f32; 6],
    cmap_decisions: &mut CMapDecisionCache,
    depth: usize,
) -> (Vec<TextItem>, Vec<PdfRect>) {
    if depth >= MAX_FORM_XOBJECT_DEPTH {
        return (Vec::new(), Vec::new());
    }

    let Ok(Object::Stream(stream)) = doc.get_object(form_id) else {
        return (Vec::new(), Vec::new());
    };

    let Ok(content_data) = stream.decompressed_content() else {
        return (Vec::new(), Vec::new());
    };
    let Ok(content) = Content::decode(&content_data) else {
        return (Vec::new(), Vec::new());
    };

    let form_matrix = if let Ok(matrix_obj) = stream.dict.get(b"Matrix") {
        if let Ok(values) = matrix_obj.as_array() {
            let mut matrix = [1.0f32, 0.0, 0.0, 1.0, 0.0, 0.0];
            for (idx, value) in values.iter().take(6).enumerate() {
                matrix[idx] = get_number(value).unwrap_or(matrix[idx]);
            }
            matrix
        } else {
            [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
        }
    } else {
        [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
    };
    let initial_ctm = multiply_matrices(&form_matrix, parent_ctm);

    let form_fonts = get_form_fonts(doc, &stream.dict);
    let font_ctx = build_font_context(doc, &form_fonts);
    let form_xobjects = get_form_xobjects(doc, &stream.dict);

    let sink = extract_content_stream(
        doc,
        &content.operations,
        page_num,
        font_cmaps,
        &font_ctx,
        &form_xobjects,
        initial_ctm,
        MarkedContentMode::Disabled,
        cmap_decisions,
        |nested_form_id, nested_parent_ctm, decisions| {
            extract_form_xobject_content_with_depth(
                doc,
                nested_form_id,
                page_num,
                font_cmaps,
                nested_parent_ctm,
                decisions,
                depth + 1,
            )
        },
    );

    (sink.items, sink.rects)
}

/// Get fonts from a Form XObject's Resources
pub(crate) fn get_form_fonts<'a>(
    doc: &'a Document,
    form_dict: &lopdf::Dictionary,
) -> std::collections::BTreeMap<Vec<u8>, &'a lopdf::Dictionary> {
    let mut fonts = std::collections::BTreeMap::new();

    let resources = if let Ok(res_ref) = form_dict.get(b"Resources") {
        if let Ok(obj_ref) = res_ref.as_reference() {
            doc.get_dictionary(obj_ref).ok()
        } else {
            res_ref.as_dict().ok()
        }
    } else {
        return fonts;
    };

    let Some(resources) = resources else {
        return fonts;
    };

    let font_dict = if let Ok(font_ref) = resources.get(b"Font") {
        if let Ok(obj_ref) = font_ref.as_reference() {
            doc.get_dictionary(obj_ref).ok()
        } else {
            font_ref.as_dict().ok()
        }
    } else {
        return fonts;
    };

    let Some(font_dict) = font_dict else {
        return fonts;
    };

    for (name, value) in font_dict.iter() {
        if let Ok(obj_ref) = value.as_reference() {
            if let Ok(dict) = doc.get_dictionary(obj_ref) {
                fonts.insert(name.clone(), dict);
            }
        }
    }

    fonts
}
