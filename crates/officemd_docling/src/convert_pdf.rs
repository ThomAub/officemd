//! PDF-specific conversion from `OoxmlDocument` IR to Docling structures.

use officemd_core::ir::OoxmlDocument;

use crate::builder::DoclingDocumentBuilder;
use crate::model::{ContentLayer, DocItemLabel, PageItem, RefItem, Size, TextItem};

/// Convert PDF-specific IR content into Docling document structures.
pub fn convert_pdf(doc: &OoxmlDocument, builder: &mut DoclingDocumentBuilder) {
    let Some(pdf) = &doc.pdf else {
        return;
    };

    for page in &pdf.pages {
        // Page entry.
        builder.add_page(
            page.number,
            PageItem {
                size: Size {
                    width: 0.0,
                    height: 0.0,
                },
                page_no: page.number,
            },
        );

        // Page markdown content as a Text item under body.
        if !page.markdown.is_empty() {
            let text_ref = builder.add_text(TextItem {
                self_ref: String::new(),
                parent: Some(RefItem::new("#/body")),
                children: vec![],
                content_layer: ContentLayer::Body,
                label: DocItemLabel::Text,
                prov: vec![],
                orig: page.markdown.clone(),
                text: page.markdown.clone(),
                level: None,
                enumerated: None,
                marker: None,
            });
            builder.add_child_to_parent("#/body", &text_ref);
        }
    }
}
