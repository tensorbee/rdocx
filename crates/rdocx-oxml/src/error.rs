//! Error types for OXML parsing.

use thiserror::Error;

#[derive(Debug, Error)]
pub enum OxmlError {
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("XML attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),

    #[error("unexpected element: {0}")]
    UnexpectedElement(String),

    #[error("missing required element: {0}")]
    MissingElement(String),

    #[error("invalid value: {0}")]
    InvalidValue(String),

    #[error("UTF-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),

    #[error("parse int error: {0}")]
    ParseInt(#[from] std::num::ParseIntError),
}

pub type Result<T> = std::result::Result<T, OxmlError>;
