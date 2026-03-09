//! Builder for assembling a `DoclingDocument` with automatic index allocation.

use std::collections::BTreeMap;

use crate::model::{
    ContentLayer, DoclingDocument, GroupItem, GroupLabel, PageItem, RefItem, TableItem, TextItem,
};

/// Accumulates items and manages index allocation for a Docling Document.
pub struct DoclingDocumentBuilder {
    name: String,
    body: GroupItem,
    furniture: GroupItem,
    groups: Vec<GroupItem>,
    texts: Vec<TextItem>,
    tables: Vec<TableItem>,
    pages: BTreeMap<String, PageItem>,
}

impl DoclingDocumentBuilder {
    #[must_use]
    pub fn new(name: &str) -> Self {
        Self {
            name: name.to_string(),
            body: GroupItem {
                self_ref: "#/body".to_string(),
                parent: None,
                children: vec![],
                content_layer: ContentLayer::Body,
                name: "_root_".to_string(),
                label: GroupLabel::Unspecified,
            },
            furniture: GroupItem {
                self_ref: "#/furniture".to_string(),
                parent: None,
                children: vec![],
                content_layer: ContentLayer::Furniture,
                name: "_root_".to_string(),
                label: GroupLabel::Unspecified,
            },
            groups: vec![],
            texts: vec![],
            tables: vec![],
            pages: BTreeMap::new(),
        }
    }

    /// Add a text item, assigning its `self_ref`. Returns the ref path.
    pub fn add_text(&mut self, mut item: TextItem) -> String {
        let idx = self.texts.len();
        let self_ref = format!("#/texts/{idx}");
        item.self_ref.clone_from(&self_ref);
        self.texts.push(item);
        self_ref
    }

    /// Add a table item, assigning its `self_ref`. Returns the ref path.
    pub fn add_table(&mut self, mut item: TableItem) -> String {
        let idx = self.tables.len();
        let self_ref = format!("#/tables/{idx}");
        item.self_ref.clone_from(&self_ref);
        self.tables.push(item);
        self_ref
    }

    /// Add a group item, assigning its `self_ref`. Returns the ref path.
    pub fn add_group(&mut self, mut group: GroupItem) -> String {
        let idx = self.groups.len();
        let self_ref = format!("#/groups/{idx}");
        group.self_ref.clone_from(&self_ref);
        self.groups.push(group);
        self_ref
    }

    /// Insert a page entry keyed by page number.
    pub fn add_page(&mut self, page_no: usize, page: PageItem) {
        self.pages.insert(page_no.to_string(), page);
    }

    /// Append a child reference to the given parent (body, furniture, or a group).
    pub fn add_child_to_parent(&mut self, parent_ref: &str, child_ref: &str) {
        let child = RefItem::new(child_ref);
        match parent_ref {
            "#/body" => self.body.children.push(child),
            "#/furniture" => self.furniture.children.push(child),
            _ if parent_ref.starts_with("#/groups/") => {
                if let Ok(idx) = parent_ref
                    .strip_prefix("#/groups/")
                    .unwrap_or("")
                    .parse::<usize>()
                    && let Some(g) = self.groups.get_mut(idx)
                {
                    g.children.push(child);
                }
            }
            _ => {}
        }
    }

    /// Consume the builder and produce the final document.
    #[must_use]
    pub fn build(self) -> DoclingDocument {
        DoclingDocument {
            schema_name: "DoclingDocument".to_string(),
            version: "1.9.0".to_string(),
            name: self.name,
            furniture: self.furniture,
            body: self.body,
            groups: self.groups,
            texts: self.texts,
            pictures: vec![],
            tables: self.tables,
            key_value_items: vec![],
            pages: self.pages,
        }
    }
}
