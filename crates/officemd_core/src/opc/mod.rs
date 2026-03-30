//! Shared Open Packaging Conventions (OPC) helpers for OOXML containers.

pub mod content_types;
pub mod package;
pub mod part;
pub mod properties;
pub mod relationships;
pub mod writer;

pub use content_types::ContentTypes;
pub use package::{OpcError, OpcPackage};
pub use part::OpcPart;
pub use properties::extract_properties;
pub use relationships::{
    load_relationships_for_part, part_base_dir, relationship_target_map, rels_path_for_part,
    resolve_relationship_target,
};
pub use writer::{OpcWriter, RelEntry, serialize_relationships, xml_escape_attr, xml_escape_text};
