//! Shared conversion helpers used by format-specific modules.

use officemd_core::ir::{Block, CommentNote, Inline, Paragraph, Table};

use crate::builder::DoclingDocumentBuilder;
use crate::model::{
    ContentLayer, DocItemLabel, DoclingTableCell, GroupItem, GroupLabel, RefItem, TableData,
    TableItem, TextItem,
};

/// Flatten a paragraph's inlines into a single plain-text string.
#[must_use]
pub fn flatten_paragraph(p: &Paragraph) -> String {
    let estimated_len = p
        .inlines
        .iter()
        .map(|inline| match inline {
            Inline::Text(text) => text.len(),
            Inline::Link(link) => {
                if link.display.is_empty() {
                    link.target.len()
                } else {
                    link.display.len()
                }
            }
        })
        .sum();
    let mut buf = String::with_capacity(estimated_len);
    for inline in &p.inlines {
        match inline {
            Inline::Text(text) => buf.push_str(text),
            Inline::Link(link) => {
                if link.display.is_empty() {
                    buf.push_str(&link.target);
                } else {
                    buf.push_str(&link.display);
                }
            }
        }
    }
    buf
}

/// Convert a sequence of IR blocks into the given parent container.
pub fn convert_blocks(
    blocks: &[Block],
    parent_ref: &str,
    content_layer: ContentLayer,
    builder: &mut DoclingDocumentBuilder,
) {
    for block in blocks {
        match block {
            Block::Paragraph(p) => {
                let text = flatten_paragraph(p);
                let text_ref = builder.add_text(TextItem {
                    self_ref: String::new(),
                    parent: Some(RefItem::new(parent_ref)),
                    children: vec![],
                    content_layer,
                    label: DocItemLabel::Paragraph,
                    prov: vec![],
                    orig: text.clone(),
                    text,
                    level: None,
                    enumerated: None,
                    marker: None,
                });
                builder.add_child_to_parent(parent_ref, &text_ref);
            }
            Block::Table(table) => {
                convert_ir_table(table, parent_ref, content_layer, builder);
            }
            Block::Separator => {}
        }
    }
}

/// Convert an IR table (with optional caption) into the given parent container.
/// Returns the table's ref path.
pub fn convert_ir_table(
    table: &Table,
    parent_ref: &str,
    content_layer: ContentLayer,
    builder: &mut DoclingDocumentBuilder,
) -> String {
    // Caption becomes a sibling TextItem referenced from the table.
    let captions = if let Some(caption_text) = &table.caption {
        let cap_ref = builder.add_text(TextItem {
            self_ref: String::new(),
            parent: Some(RefItem::new(parent_ref)),
            children: vec![],
            content_layer,
            label: DocItemLabel::Caption,
            prov: vec![],
            orig: caption_text.clone(),
            text: caption_text.clone(),
            level: None,
            enumerated: None,
            marker: None,
        });
        builder.add_child_to_parent(parent_ref, &cap_ref);
        vec![RefItem::new(cap_ref)]
    } else {
        vec![]
    };

    let num_rows = table.rows.len();
    let num_cols = table
        .rows
        .first()
        .map_or(table.headers.len(), std::vec::Vec::len);

    let expected_cells: usize = table.rows.iter().map(std::vec::Vec::len).sum();
    let mut cells = Vec::with_capacity(expected_cells);
    for (row_idx, row) in table.rows.iter().enumerate() {
        for (col_idx, cell) in row.iter().enumerate() {
            let mut text = String::new();
            for (pidx, paragraph) in cell.content.iter().enumerate() {
                if pidx > 0 {
                    text.push('\n');
                }
                text.push_str(&flatten_paragraph(paragraph));
            }
            cells.push(DoclingTableCell {
                row_span: 1,
                col_span: 1,
                start_row_offset_idx: row_idx,
                end_row_offset_idx: row_idx + 1,
                start_col_offset_idx: col_idx,
                end_col_offset_idx: col_idx + 1,
                text,
                column_header: row_idx == 0,
                row_header: false,
                row_section: false,
            });
        }
    }

    let table_ref = builder.add_table(TableItem {
        self_ref: String::new(),
        parent: Some(RefItem::new(parent_ref)),
        children: vec![],
        content_layer,
        label: DocItemLabel::Table,
        prov: vec![],
        data: TableData {
            table_cells: cells,
            num_rows,
            num_cols,
        },
        captions,
    });
    builder.add_child_to_parent(parent_ref, &table_ref);
    table_ref
}

/// Convert comment notes into a "comments" group under the given parent.
pub fn convert_comments(
    comments: &[CommentNote],
    parent_ref: &str,
    content_layer: ContentLayer,
    builder: &mut DoclingDocumentBuilder,
) {
    if comments.is_empty() {
        return;
    }

    let group_ref = builder.add_group(GroupItem {
        self_ref: String::new(),
        parent: Some(RefItem::new(parent_ref)),
        children: vec![],
        content_layer,
        name: "comments".to_string(),
        label: GroupLabel::Section,
    });
    builder.add_child_to_parent(parent_ref, &group_ref);

    for comment in comments {
        let text = if comment.author.is_empty() {
            comment.text.clone()
        } else {
            format!("{}: {}", comment.author, comment.text)
        };
        let text_ref = builder.add_text(TextItem {
            self_ref: String::new(),
            parent: Some(RefItem::new(&group_ref)),
            children: vec![],
            content_layer,
            label: DocItemLabel::Footnote,
            prov: vec![],
            orig: text.clone(),
            text,
            level: None,
            enumerated: None,
            marker: None,
        });
        builder.add_child_to_parent(&group_ref, &text_ref);
    }
}
