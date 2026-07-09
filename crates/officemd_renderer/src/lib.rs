//! Optional renderer adapters for OfficeMD.

pub mod discovery;
pub mod libreoffice;
pub mod poppler;
pub mod unavailable;

pub use unavailable::UnavailableRenderer;
