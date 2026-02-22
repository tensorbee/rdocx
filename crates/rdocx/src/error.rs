//! Error types for the rdocx high-level API.

use thiserror::Error;

#[derive(Debug, Error)]
pub enum Error {
    #[error("OPC package error: {0}")]
    Opc(#[from] rdocx_opc::OpcError),

    #[error("OXML parsing error: {0}")]
    Oxml(#[from] rdocx_oxml::OxmlError),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("layout error: {0}")]
    Layout(#[from] rdocx_layout::LayoutError),

    #[error("document has no main document part")]
    NoDocumentPart,

    #[error("{0}")]
    Other(String),
}

pub type Result<T> = std::result::Result<T, Error>;
