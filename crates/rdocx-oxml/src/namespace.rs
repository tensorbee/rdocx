//! OOXML namespace constants.

/// WordprocessingML main namespace
pub const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// WordprocessingML namespace prefix
pub const W_PREFIX: &[u8] = b"w";

/// Relationships namespace
pub const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Markup Compatibility namespace
pub const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Check if a tag name matches an expected local name, accounting for namespace prefixes.
pub fn matches_local_name(tag: &[u8], local_name: &[u8]) -> bool {
    if tag == local_name {
        return true;
    }
    // Check for w:localName pattern
    if let Some(pos) = tag.iter().position(|&b| b == b':') {
        &tag[pos + 1..] == local_name
    } else {
        false
    }
}
