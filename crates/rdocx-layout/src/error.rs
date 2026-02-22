//! Layout engine error types.

use thiserror::Error;

#[derive(Debug, Error)]
pub enum LayoutError {
    #[error("font not found: {0}")]
    FontNotFound(String),

    #[error("font parsing error: {0}")]
    FontParse(String),

    #[error("shaping error: {0}")]
    Shaping(String),

    #[error("layout error: {0}")]
    Layout(String),
}

pub type Result<T> = std::result::Result<T, LayoutError>;
