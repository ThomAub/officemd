//! DOCX-specific conversion from `OoxmlDocument` IR to Docling structures.

use officemd_core::ir::OoxmlDocument;

use crate::builder::DoclingDocumentBuilder;
use crate::convert::{convert_blocks, convert_comments};
use crate::model::{ContentLayer, DocItemLabel, GroupItem, GroupLabel, RefItem, TextItem};

/// Convert DOCX sections from the shared IR into Docling groups and text items.
pub fn convert_docx(doc: &OoxmlDocument, builder: &mut DoclingDocumentBuilder) {
    for section in &doc.sections {
        if section.name == "body" {
            convert_blocks(&section.blocks, "#/body", ContentLayer::Body, builder);
            convert_comments(&section.comments, "#/body", ContentLayer::Body, builder);
        } else if section.name.starts_with("header") || section.name.starts_with("footer") {
            let label = if section.name.starts_with("header") {
                DocItemLabel::PageHeader
            } else {
                DocItemLabel::PageFooter
            };
            convert_furniture_section(&section.blocks, label, builder);
        } else {
            // Other sections (footnotes, endnotes, etc.) become groups under body.
            let group_ref = builder.add_group(GroupItem {
                self_ref: String::new(),
                parent: Some(RefItem::new("#/body")),
                children: vec![],
                content_layer: ContentLayer::Body,
                name: section.name.clone(),
                label: GroupLabel::Section,
            });
            builder.add_child_to_parent("#/body", &group_ref);

            convert_blocks(&section.blocks, &group_ref, ContentLayer::Body, builder);
            convert_comments(&section.comments, &group_ref, ContentLayer::Body, builder);
        }
    }
}

/// Convert blocks from a header/footer section into furniture.
/// Each paragraph gets a specific PageHeader/PageFooter label instead of Paragraph.
fn convert_furniture_section(
    blocks: &[officemd_core::ir::Block],
    label: DocItemLabel,
    builder: &mut DoclingDocumentBuilder,
) {
    for block in blocks {
        match block {
            officemd_core::ir::Block::Paragraph(p) => {
                let text = crate::convert::flatten_paragraph(p);
                let text_ref = builder.add_text(TextItem {
                    self_ref: String::new(),
                    parent: Some(RefItem::new("#/furniture")),
                    children: vec![],
                    content_layer: ContentLayer::Furniture,
                    label,
                    prov: vec![],
                    orig: text.clone(),
                    text,
                    level: None,
                    enumerated: None,
                    marker: None,
                });
                builder.add_child_to_parent("#/furniture", &text_ref);
            }
            officemd_core::ir::Block::Table(table) => {
                crate::convert::convert_ir_table(
                    table,
                    "#/furniture",
                    ContentLayer::Furniture,
                    builder,
                );
            }
            officemd_core::ir::Block::Separator => {}
        }
    }
}
