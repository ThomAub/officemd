//! XLSX-specific conversion from `OoxmlDocument` IR to Docling structures.

use officemd_core::ir::OoxmlDocument;

use crate::builder::DoclingDocumentBuilder;
use crate::convert::convert_ir_table;
use crate::model::{ContentLayer, DocItemLabel, GroupItem, GroupLabel, RefItem, TextItem};

/// Convert XLSX workbook sheets and tables from the shared IR into Docling groups and items.
pub fn convert_xlsx(doc: &OoxmlDocument, builder: &mut DoclingDocumentBuilder) {
    for sheet in &doc.sheets {
        // Each sheet becomes a Section group under body.
        let sheet_ref = builder.add_group(GroupItem {
            self_ref: String::new(),
            parent: Some(RefItem::new("#/body")),
            children: vec![],
            content_layer: ContentLayer::Body,
            name: sheet.name.clone(),
            label: GroupLabel::Section,
        });
        builder.add_child_to_parent("#/body", &sheet_ref);

        // Tables within the sheet.
        for table in &sheet.tables {
            convert_ir_table(table, &sheet_ref, ContentLayer::Body, builder);
        }

        // Formula notes become Text items under the sheet group.
        for note in &sheet.formulas {
            let text = format!("{} = {}", note.cell_ref, note.formula);
            let text_ref = builder.add_text(TextItem {
                self_ref: String::new(),
                parent: Some(RefItem::new(&sheet_ref)),
                children: vec![],
                content_layer: ContentLayer::Body,
                label: DocItemLabel::Text,
                prov: vec![],
                orig: text.clone(),
                text,
                level: None,
                enumerated: None,
                marker: None,
            });
            builder.add_child_to_parent(&sheet_ref, &text_ref);
        }

        // Hyperlinks as text items.
        for link in &sheet.hyperlinks {
            let text = if link.display.is_empty() {
                link.target.clone()
            } else {
                format!("{} ({})", link.display, link.target)
            };
            if text.is_empty() {
                continue;
            }
            let text_ref = builder.add_text(TextItem {
                self_ref: String::new(),
                parent: Some(RefItem::new(&sheet_ref)),
                children: vec![],
                content_layer: ContentLayer::Body,
                label: DocItemLabel::Text,
                prov: vec![],
                orig: text.clone(),
                text,
                level: None,
                enumerated: None,
                marker: None,
            });
            builder.add_child_to_parent(&sheet_ref, &text_ref);
        }
    }
}
