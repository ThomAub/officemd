/// Metadata for an OPC part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpcPart {
    pub path: String,
    pub content_type: Option<String>,
}
