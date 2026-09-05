//! Error types for the rdocx high-level API.

use thiserror::Error;

#[derive(Debug, Error)]
pub enum Error {
    #[error("OPC package error: {0}")]
    Opc(#[from] oxml_opc::OpcError),

    #[error("OXML parsing error: {0}")]
    Oxml(#[from] rdocx_oxml::OxmlError),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("layout error: {0}")]
    Layout(#[from] oxml_layout::LayoutError),

    #[error("PDF conformance error: {0}")]
    Pdf(#[from] oxml_pdf::PdfError),

    #[error("raster export error: {0}")]
    Raster(#[from] oxml_pdf::RasterError),

    #[error("document has no main document part")]
    NoDocumentPart,

    #[error("image dimensions are unavailable for {filename}")]
    UnavailableImageDimensions { filename: String },

    #[error("RTF parse error at byte {offset}: {message}")]
    Rtf { offset: usize, message: String },

    #[error("HTML import error at {location}: {message}")]
    Html { location: String, message: String },

    #[error("MHTML error in {part:?} at byte {offset}: {message}")]
    Mhtml {
        part: Option<String>,
        offset: u64,
        message: String,
    },

    #[error("ODT import error in {part:?} at byte {offset}: {message}")]
    Odt {
        part: Option<String>,
        offset: u64,
        message: String,
    },

    #[error("{operation} failed: {message}")]
    InvalidEmbeddedMutation {
        operation: &'static str,
        message: String,
    },

    #[error("{0}")]
    Other(String),
}

pub type Result<T> = std::result::Result<T, Error>;

#[cfg(test)]
mod tests {
    use super::Error;

    #[test]
    fn layout_error_wraps_the_shared_layout_error() {
        let shared = oxml_layout::LayoutError::Layout("shared failure".to_string());
        let error = Error::from(shared);
        assert!(matches!(
            error,
            Error::Layout(oxml_layout::LayoutError::Layout(message))
                if message == "shared failure"
        ));
    }
}
