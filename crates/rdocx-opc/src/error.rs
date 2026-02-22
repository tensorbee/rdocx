//! Error types for OPC operations.

use thiserror::Error;

#[derive(Debug, Error)]
pub enum OpcError {
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("XML attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),

    #[error("part not found: {0}")]
    PartNotFound(String),

    #[error("invalid content types XML")]
    InvalidContentTypes,

    #[error("invalid relationship XML")]
    InvalidRelationship,

    #[error("UTF-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),
}

pub type Result<T> = std::result::Result<T, OpcError>;
