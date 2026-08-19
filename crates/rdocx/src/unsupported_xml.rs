//! Borrowed access to XML the high-level facade does not expose semantically.

use std::ops::Deref;

/// WordprocessingML namespace URI.
pub(crate) const WORD_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

enum UnsupportedXmlSource<'a> {
    Raw(&'a [u8]),
    Modeled {
        namespace_uri: &'static str,
        local_name: &'static str,
    },
}

/// A preserved XML subtree, or a modeled construct retained as an unsupported
/// reader fact for a compatibility iterator.
pub struct UnsupportedXmlRef<'a> {
    source: UnsupportedXmlSource<'a>,
}

impl<'a> UnsupportedXmlRef<'a> {
    pub(crate) fn new(raw: &'a [u8]) -> Self {
        Self {
            source: UnsupportedXmlSource::Raw(raw),
        }
    }

    pub(crate) fn modeled(namespace_uri: &'static str, local_name: &'static str) -> Self {
        Self {
            source: UnsupportedXmlSource::Modeled {
                namespace_uri,
                local_name,
            },
        }
    }

    /// Original subtree bytes.
    ///
    /// Modeled compatibility facts have no preserved raw subtree and return an
    /// empty slice. Consumers should use [`Self::local_name`] and
    /// [`Self::namespace_uri`] to classify those facts.
    pub fn bytes(&self) -> &'a [u8] {
        match self.source {
            UnsupportedXmlSource::Raw(raw) => raw,
            UnsupportedXmlSource::Modeled { .. } => &[],
        }
    }

    /// Qualified element name as it appears in preserved XML.
    ///
    /// Modeled compatibility facts return their local name.
    pub fn qualified_name(&self) -> &str {
        match self.source {
            UnsupportedXmlSource::Raw(raw) => first_element_name(raw),
            UnsupportedXmlSource::Modeled { local_name, .. } => local_name,
        }
    }

    /// Local element name.
    pub fn local_name(&self) -> &str {
        match self.source {
            UnsupportedXmlSource::Raw(raw) => first_element_name(raw)
                .rsplit_once(':')
                .map(|(_, local)| local)
                .unwrap_or_else(|| first_element_name(raw)),
            UnsupportedXmlSource::Modeled { local_name, .. } => local_name,
        }
    }

    /// Namespace URI declared on the preserved subtree's root element.
    ///
    /// A captured subtree cannot retain declarations inherited from its parent,
    /// so this returns `None` when that declaration is outside the subtree.
    /// Callers must treat that as unclassified content rather than assuming a
    /// namespace.
    pub fn namespace_uri(&self) -> Option<&str> {
        match self.source {
            UnsupportedXmlSource::Raw(raw) => {
                let name = first_element_name(raw);
                let prefix = name.split_once(':').map_or("", |(prefix, _)| prefix);
                namespace_declaration(raw, prefix)
            }
            UnsupportedXmlSource::Modeled { namespace_uri, .. } => Some(namespace_uri),
        }
    }

    /// Whether the preserved element has nested elements or non-whitespace text.
    pub fn has_child_content(&self) -> bool {
        match self.source {
            UnsupportedXmlSource::Raw(raw) => {
                let Some(start_end) = raw.iter().position(|byte| *byte == b'>') else {
                    return false;
                };
                let content = &raw[start_end + 1..];
                content.starts_with(b"<") && !content.starts_with(b"</")
                    || content
                        .iter()
                        .any(|byte| !byte.is_ascii_whitespace() && *byte != b'<')
            }
            UnsupportedXmlSource::Modeled { .. } => true,
        }
    }
}

impl AsRef<[u8]> for UnsupportedXmlRef<'_> {
    fn as_ref(&self) -> &[u8] {
        self.bytes()
    }
}

impl Deref for UnsupportedXmlRef<'_> {
    type Target = [u8];

    fn deref(&self) -> &Self::Target {
        self.bytes()
    }
}

fn first_element_name(raw: &[u8]) -> &str {
    let start = raw
        .iter()
        .position(|byte| *byte == b'<')
        .map_or(0, |index| index + 1);
    let end = raw[start..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(*byte, b'/' | b'>'))
        .map_or(raw.len(), |index| start + index);
    std::str::from_utf8(&raw[start..end]).unwrap_or("")
}

fn namespace_declaration<'a>(raw: &'a [u8], prefix: &str) -> Option<&'a str> {
    let start_end = raw.iter().position(|byte| *byte == b'>')?;
    let declaration = if prefix.is_empty() {
        "xmlns".to_owned()
    } else {
        format!("xmlns:{prefix}")
    };
    let bytes = declaration.as_bytes();
    let mut index = 0;
    while index + bytes.len() < start_end {
        let found = raw[index..start_end]
            .windows(bytes.len())
            .position(|candidate| candidate == bytes)?;
        let name_start = index + found;
        let name_end = name_start + bytes.len();
        let valid_before = name_start == 0 || raw[name_start - 1].is_ascii_whitespace();
        let valid_after =
            name_end < start_end && (raw[name_end].is_ascii_whitespace() || raw[name_end] == b'=');
        if valid_before && valid_after {
            let equals = raw[name_end..start_end]
                .iter()
                .position(|byte| *byte == b'=')
                .map(|offset| name_end + offset)?;
            let quote = *raw.get(equals + 1)?;
            if !matches!(quote, b'\'' | b'\"') {
                return None;
            }
            let value_start = equals + 2;
            let value_end = raw[value_start..start_end]
                .iter()
                .position(|byte| *byte == quote)
                .map(|offset| value_start + offset)?;
            return std::str::from_utf8(&raw[value_start..value_end]).ok();
        }
        index = name_end;
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn raw_fact_reports_its_declared_name_and_namespace() {
        let fact = UnsupportedXmlRef::new(br#"<x:item xmlns:x="urn:example"><x:child/></x:item>"#);

        assert_eq!(fact.qualified_name(), "x:item");
        assert_eq!(fact.local_name(), "item");
        assert_eq!(fact.namespace_uri(), Some("urn:example"));
        assert!(fact.has_child_content());
    }

    #[test]
    fn modeled_fact_has_identity_without_fabricating_xml() {
        let fact = UnsupportedXmlRef::modeled(WORD_NAMESPACE, "sdt");

        assert_eq!(fact.bytes(), b"");
        assert_eq!(fact.local_name(), "sdt");
        assert_eq!(fact.namespace_uri(), Some(WORD_NAMESPACE));
    }
}
