//! OOXML namespace constants.

/// WordprocessingML main namespace
pub const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// WordprocessingML namespace prefix
pub const W_PREFIX: &[u8] = b"w";
/// Transitional OfficeMath namespace.
pub const M_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
/// Canonical OfficeMath namespace prefix.
pub const M_PREFIX: &[u8] = b"m";

pub use oxml_core::xml::{MC_NS, R_NS, matches_local_name};
