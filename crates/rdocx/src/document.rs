//! The main Document type — entry point for the rdocx API.

use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::path::Path;
use std::sync::{Arc, Mutex};

#[cfg(test)]
use std::cell::Cell;

use oxml_chart::{CT_ChartSpace, ChartData, ChartKind};
use oxml_core::app_properties::AppProperties;
use oxml_opc::content_types;
use oxml_opc::relationship::rel_types;
use oxml_opc::{OpcPackage, PackageReadLimits};
use oxml_sml::Workbook;
use quick_xml::Writer;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use rdocx_oxml::MathProperties;
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::{BodyContent, CT_Columns, CT_Document, CT_SectPr};
use rdocx_oxml::drawing::{CT_Anchor, CT_Drawing, CT_Inline, drawing_ns};
use rdocx_oxml::font_table::{EmbeddedFontReference, FontFaceKind, FontTable};
use rdocx_oxml::header_footer::{
    CT_HdrFtr, HdrFtrRef, HdrFtrType, VmlWatermark, replace_authored_watermark,
};
use rdocx_oxml::namespace::matches_local_name;
use rdocx_oxml::numbering::{
    CT_AbstractNum, CT_Lvl, CT_Num, CT_NumLvl, CT_Numbering, ST_LvlSuffix, ST_NumberFormat,
};
use rdocx_oxml::properties::{CT_PPr, CT_RPr};
use rdocx_oxml::settings::{
    CT_Settings, CharacterSpacingControl, CompatibilitySetting, DocumentProtection,
    ThemeFontLanguage,
};
use rdocx_oxml::shared::{ST_Jc, ST_PageOrientation, ST_SectionType};
use rdocx_oxml::styles::{CT_Styles, StyleType};
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::{CT_P, CT_R, RunContent};

use oxml_core::custom_properties::{CustomProperties, CustomProperty};
use rdocx_oxml::core_properties::CoreProperties;

use crate::Length;
use crate::content_control::ContentControlRef;
use crate::error::{Error, Result};
use crate::paragraph::{Paragraph, ParagraphRef};
use crate::revision::RevisionRef;
use crate::run::RunRef;
use crate::style::{self, Style, StyleBuilder};
use crate::table::{Table, TableRef};

/// Options that select a native document render projection.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct RenderOptions {
    /// The tracked-revision view to render.
    pub revision_view: rdocx_layout::RevisionView,
}

/// The package class declared by a Word main document part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WordPackageClass {
    /// A macro-free `.docx` document.
    Document,
    /// A macro-enabled `.docm` document.
    MacroEnabledDocument,
    /// A macro-free `.dotx` template.
    Template,
    /// A macro-enabled `.dotm` template.
    MacroEnabledTemplate,
}

/// Completeness and package class for a newly authored Word document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WordCreationProfile {
    /// The compact package graph historically emitted by `rdocx`.
    Minimal(WordPackageClass),
    /// A Word-compatible blank package with the standard owned support parts.
    WordCompatible(WordPackageClass),
}

/// One font-table record authored through the native facade.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct FontDefinition {
    pub name: String,
    pub alternate_name: Option<String>,
    pub family: Option<String>,
    pub pitch: Option<String>,
    pub embedded_fonts: Vec<EmbeddedFont>,
}

impl FontDefinition {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            alternate_name: None,
            family: None,
            pitch: None,
            embedded_fonts: Vec::new(),
        }
    }

    pub fn with_alternate_name(mut self, name: impl Into<String>) -> Self {
        self.alternate_name = Some(name.into());
        self
    }

    pub fn with_family(mut self, family: impl Into<String>) -> Self {
        self.family = Some(family.into());
        self
    }

    pub fn with_pitch(mut self, pitch: impl Into<String>) -> Self {
        self.pitch = Some(pitch.into());
        self
    }
}

/// The font-table slot occupied by an embedded font face.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum EmbeddedFontKind {
    Regular,
    Bold,
    Italic,
    BoldItalic,
}

/// The caller's explicit authorization and exact license identity.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct FontEmbeddingLicense {
    pub authorized: bool,
    pub identity: String,
}

impl FontEmbeddingLicense {
    pub fn new(authorized: bool, identity: impl Into<String>) -> Self {
        Self {
            authorized,
            identity: identity.into(),
        }
    }
}

/// Caller-owned bytes and metadata for one embedded font face.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EmbeddedFont {
    pub kind: EmbeddedFontKind,
    pub data: Vec<u8>,
    pub font_key: String,
    pub subsetted: bool,
    pub license: FontEmbeddingLicense,
}

impl EmbeddedFont {
    pub fn new(
        kind: EmbeddedFontKind,
        data: Vec<u8>,
        font_key: impl Into<String>,
        license: FontEmbeddingLicense,
    ) -> Self {
        Self {
            kind,
            data,
            font_key: font_key.into(),
            subsetted: false,
            license,
        }
    }

    pub fn with_subsetted(mut self, subsetted: bool) -> Self {
        self.subsetted = subsetted;
        self
    }
}

impl From<EmbeddedFontKind> for FontFaceKind {
    fn from(value: EmbeddedFontKind) -> Self {
        match value {
            EmbeddedFontKind::Regular => Self::Regular,
            EmbeddedFontKind::Bold => Self::Bold,
            EmbeddedFontKind::Italic => Self::Italic,
            EmbeddedFontKind::BoldItalic => Self::BoldItalic,
        }
    }
}

impl From<FontFaceKind> for EmbeddedFontKind {
    fn from(value: FontFaceKind) -> Self {
        match value {
            FontFaceKind::Regular => Self::Regular,
            FontFaceKind::Bold => Self::Bold,
            FontFaceKind::Italic => Self::Italic,
            FontFaceKind::BoldItalic => Self::BoldItalic,
        }
    }
}

impl WordPackageClass {
    pub(crate) fn content_type(self) -> &'static str {
        match self {
            Self::Document => content_types::WORD_DOCUMENT,
            Self::MacroEnabledDocument => content_types::WORD_DOCUMENT_MACRO_ENABLED,
            Self::Template => content_types::WORD_TEMPLATE,
            Self::MacroEnabledTemplate => content_types::WORD_TEMPLATE_MACRO_ENABLED,
        }
    }

    fn from_content_type(content_type: &str) -> Option<Self> {
        match content_type {
            content_types::WORD_DOCUMENT => Some(Self::Document),
            content_types::WORD_DOCUMENT_MACRO_ENABLED => Some(Self::MacroEnabledDocument),
            content_types::WORD_TEMPLATE => Some(Self::Template),
            content_types::WORD_TEMPLATE_MACRO_ENABLED => Some(Self::MacroEnabledTemplate),
            _ => None,
        }
    }
}

/// One direct child of a document body, in source order.
pub enum BodyItemRef<'a> {
    /// A body paragraph.
    Paragraph(ParagraphRef<'a>),
    /// A body table.
    Table(TableRef<'a>),
    /// A body-level content control.
    ContentControl(ContentControlRef<'a>),
    /// A preserved body child that rdocx does not model.
    UnsupportedXml(&'a [u8]),
}

const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";

enum UnsupportedXmlSource<'a> {
    Raw { raw: &'a [u8] },
    Modeled,
}

/// A borrowed fact for content the compatibility reader does not model.
pub struct UnsupportedXmlRef<'a> {
    source: UnsupportedXmlSource<'a>,
    qualified_name: Option<Cow<'a, str>>,
    local_name: Cow<'a, str>,
    namespace_uri: Option<Cow<'a, str>>,
    has_child_content: bool,
}

impl<'a> UnsupportedXmlRef<'a> {
    /// Inspect preserved reader-item bytes using the same unsupported XML
    /// projection as body compatibility items.
    pub fn from_bytes(raw: &'a [u8]) -> Self {
        Self::raw(raw, &[])
    }

    fn raw(raw: &'a [u8], inherited_namespaces: &'a [(String, String)]) -> Self {
        let names = raw_element_names(raw);
        let namespace_uri = names.as_ref().and_then(|(name, _)| {
            let prefix = name.split_once(':').map_or("", |(prefix, _)| prefix);
            raw_namespace_declaration(raw, prefix)
                .map(Cow::Owned)
                .or_else(|| inherited_namespace(inherited_namespaces, prefix).map(Cow::Borrowed))
        });
        let (qualified_name, local_name) = names.map_or_else(
            || (None, Cow::Borrowed("")),
            |(qualified_name, local_name)| {
                (Some(Cow::Owned(qualified_name)), Cow::Owned(local_name))
            },
        );
        Self {
            source: UnsupportedXmlSource::Raw { raw },
            qualified_name,
            local_name,
            namespace_uri,
            has_child_content: raw_has_child_content(raw),
        }
    }

    fn modeled(
        qualified_name: &'static str,
        namespace_uri: &'static str,
        local_name: &'static str,
        has_child_content: bool,
    ) -> Self {
        Self {
            source: UnsupportedXmlSource::Modeled,
            qualified_name: Some(Cow::Borrowed(qualified_name)),
            local_name: Cow::Borrowed(local_name),
            namespace_uri: Some(Cow::Borrowed(namespace_uri)),
            has_child_content,
        }
    }

    /// Original subtree bytes, or `None` for a modeled compatibility fact.
    pub fn raw_xml(&self) -> Option<&'a [u8]> {
        match self.source {
            UnsupportedXmlSource::Raw { raw } => Some(raw),
            UnsupportedXmlSource::Modeled => None,
        }
    }

    /// Qualified element name, when the retained XML has a valid root element.
    pub fn qualified_name(&self) -> Option<&str> {
        self.qualified_name.as_deref()
    }

    /// Local element name.
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Namespace URI resolved from local or inherited declarations.
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    /// Whether the retained element has a child element or visible text.
    pub fn has_child_content(&self) -> bool {
        self.has_child_content
    }
}

fn raw_element_names(raw: &[u8]) -> Option<(String, String)> {
    let mut reader = quick_xml::Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).ok()? {
            Event::Start(element) | Event::Empty(element) => {
                let qualified_name = std::str::from_utf8(element.name().as_ref())
                    .ok()?
                    .to_owned();
                let local_name = std::str::from_utf8(element.local_name().as_ref())
                    .ok()?
                    .to_owned();
                return Some((qualified_name, local_name));
            }
            Event::Eof => return None,
            _ => buffer.clear(),
        }
    }
}

fn raw_namespace_declaration(raw: &[u8], prefix: &str) -> Option<String> {
    let mut reader = quick_xml::Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let declaration = if prefix.is_empty() {
        "xmlns".to_owned()
    } else {
        format!("xmlns:{prefix}")
    };
    loop {
        match reader.read_event_into(&mut buffer).ok()? {
            Event::Start(element) | Event::Empty(element) => {
                let attribute = element
                    .attributes()
                    .flatten()
                    .find(|attribute| attribute.key.as_ref() == declaration.as_bytes())?;
                return attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                    .ok()
                    .map(Cow::into_owned);
            }
            Event::Eof => return None,
            _ => buffer.clear(),
        }
    }
}

fn inherited_namespace<'a>(namespaces: &'a [(String, String)], prefix: &str) -> Option<&'a str> {
    if prefix == "xml" {
        return Some(XML_NAMESPACE);
    }
    let declaration = if prefix.is_empty() {
        "xmlns".to_owned()
    } else {
        format!("xmlns:{prefix}")
    };
    namespaces
        .iter()
        .rev()
        .find(|(name, _)| name == &declaration)
        .map(|(_, uri)| uri.as_str())
}

fn raw_has_child_content(raw: &[u8]) -> bool {
    let mut reader = quick_xml::Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_)) => {
                if depth > 0 {
                    return true;
                }
                depth += 1;
            }
            Ok(Event::Empty(_)) if depth > 0 => return true,
            Ok(Event::Text(text)) if depth > 0 && !xml_text_is_whitespace(&text) => {
                return true;
            }
            Ok(Event::CData(text)) if depth > 0 && !xml_bytes_are_whitespace(text.as_ref()) => {
                return true;
            }
            Ok(Event::GeneralRef(reference)) if depth > 0 => match reference.resolve_char_ref() {
                Ok(Some(character)) if !character.is_ascii_whitespace() => return true,
                Ok(Some(_)) => {}
                Ok(None) => return true,
                Err(_) => return false,
            },
            Ok(Event::End(_)) => depth = depth.saturating_sub(1),
            Ok(Event::Eof) | Err(_) => return false,
            _ => {}
        }
        buffer.clear();
    }
}

fn xml_text_is_whitespace(text: &quick_xml::events::BytesText<'_>) -> bool {
    xml_bytes_are_whitespace(text.as_ref())
}

fn xml_bytes_are_whitespace(bytes: &[u8]) -> bool {
    bytes.iter().all(|byte| byte.is_ascii_whitespace())
}

fn typed_numbering_leaf_has_unmodeled(raw: &[u8], word_prefixes: &[String]) -> bool {
    if raw_has_child_content(raw) {
        return true;
    }
    let mut reader = quick_xml::Reader::from_reader(raw);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                let mut value_count = 0usize;
                for attribute in element.attributes() {
                    let Ok(attribute) = attribute else {
                        return true;
                    };
                    let name = attribute.key.as_ref();
                    if name == b"xmlns" || name.starts_with(b"xmlns:") {
                        continue;
                    }
                    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
                        return true;
                    };
                    let prefix = &name[..separator];
                    let local = &name[separator + 1..];
                    if local != b"val"
                        || !word_prefixes
                            .iter()
                            .any(|candidate| candidate.as_bytes() == prefix)
                    {
                        return true;
                    }
                    value_count += 1;
                }
                return value_count != 1;
            }
            Ok(Event::Eof) | Err(_) => return true,
            _ => buffer.clear(),
        }
    }
}

#[derive(Default)]
struct DocumentNamespaceScopes {
    root_declarations: Vec<(String, String)>,
    body_declarations: Vec<(String, String)>,
    body_bindings: Vec<(String, String)>,
}

fn document_namespace_scopes(xml: &[u8]) -> Result<DocumentNamespaceScopes> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut scopes: Vec<Vec<(String, String)>> = Vec::new();
    let mut root_declarations = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(rdocx_oxml::OxmlError::from)?;
        match event {
            Event::Start(ref element) => {
                let mut scope = scopes.last().cloned().unwrap_or_default();
                let declarations = namespace_declarations(element)?;
                apply_namespace_declarations(element, &mut scope)?;
                if scopes.is_empty() {
                    root_declarations = declarations.clone();
                }
                if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == WORD_NAMESPACE.as_bytes())
                    && matches_local_name(element.name().as_ref(), b"body")
                {
                    return Ok(DocumentNamespaceScopes {
                        root_declarations,
                        body_declarations: declarations,
                        body_bindings: scope,
                    });
                }
                scopes.push(scope);
            }
            Event::Empty(ref element) => {
                let mut scope = scopes.last().cloned().unwrap_or_default();
                let declarations = namespace_declarations(element)?;
                apply_namespace_declarations(element, &mut scope)?;
                if scopes.is_empty() {
                    root_declarations = declarations.clone();
                }
                if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == WORD_NAMESPACE.as_bytes())
                    && matches_local_name(element.name().as_ref(), b"body")
                {
                    return Ok(DocumentNamespaceScopes {
                        root_declarations,
                        body_declarations: declarations,
                        body_bindings: scope,
                    });
                }
            }
            Event::End(_) => {
                scopes.pop();
            }
            Event::Eof => {
                return Ok(DocumentNamespaceScopes {
                    root_declarations,
                    ..DocumentNamespaceScopes::default()
                });
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn namespace_declarations(element: &BytesStart<'_>) -> Result<Vec<(String, String)>> {
    let mut declarations = Vec::new();
    apply_namespace_declarations(element, &mut declarations)?;
    Ok(declarations)
}

fn apply_namespace_declarations(
    element: &BytesStart<'_>,
    scope: &mut Vec<(String, String)>,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(rdocx_oxml::OxmlError::from)?;
        let name = attribute.key.as_ref();
        if name != b"xmlns" && !name.starts_with(b"xmlns:") {
            continue;
        }
        let name = std::str::from_utf8(name)
            .map_err(rdocx_oxml::OxmlError::from)?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
            .map_err(rdocx_oxml::OxmlError::from)?
            .into_owned();
        scope.retain(|(candidate, _)| candidate != &name);
        scope.push((name, value));
    }
    Ok(())
}

fn unsafe_serializer_namespace_prefix(
    root_declarations: &[(String, String)],
    body_declarations: &[(String, String)],
) -> Option<String> {
    if let Some((name, _)) = body_declarations.first() {
        return Some(namespace_prefix(name));
    }

    root_declarations.iter().find_map(|(name, value)| {
        let prefix = namespace_prefix(name);
        let expected = match canonical_serializer_namespace(&prefix) {
            Some(expected) => expected,
            None if prefix.is_empty() => return Some("default".to_owned()),
            None => return None,
        };
        (value != expected).then_some(prefix)
    })
}

fn canonical_serializer_namespace(prefix: &str) -> Option<&'static str> {
    match prefix {
        "w" => Some(WORD_NAMESPACE),
        "r" => Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        "mc" => Some("http://schemas.openxmlformats.org/markup-compatibility/2006"),
        "wp" => Some("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        "a" => Some("http://schemas.openxmlformats.org/drawingml/2006/main"),
        "pic" => Some("http://schemas.openxmlformats.org/drawingml/2006/picture"),
        "c" => Some("http://schemas.openxmlformats.org/drawingml/2006/chart"),
        "w14" => Some("http://schemas.microsoft.com/office/word/2010/wordml"),
        "w15" => Some("http://schemas.microsoft.com/office/word/2012/wordml"),
        _ => None,
    }
}

fn namespace_prefix(declaration_name: &str) -> String {
    declaration_name
        .strip_prefix("xmlns:")
        .unwrap_or("")
        .to_owned()
}

#[derive(Debug)]
struct NestedNamespaceOwner {
    local_name: String,
    declarations: Vec<(String, String)>,
    snapshot: LogicalOwnerSnapshot,
    markers: Vec<NamespaceMarker>,
    ambiguous_without_namespace: bool,
    same_marker_structural_alternate: bool,
    same_namespace_structural_alternate: bool,
}

#[derive(Debug)]
struct ModeledOwnerSpan {
    local_name: String,
    start: usize,
    end: usize,
    declarations: Vec<(String, String)>,
    bindings: Vec<(String, String)>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct LogicalOwnerSnapshot {
    structure: Vec<String>,
    semantic: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct NamespaceMarker {
    raw: Vec<u8>,
    namespace_facts: Vec<(String, Option<String>)>,
}

fn is_modeled_owner(local_name: &[u8]) -> bool {
    matches!(
        local_name,
        b"p" | b"tbl" | b"tc" | b"sdt" | b"hyperlink" | b"r"
    )
}

fn is_modeled_word_child(parent: Option<&[u8]>, local_name: &[u8]) -> bool {
    match parent {
        Some(b"p") => matches!(
            local_name,
            b"pPr"
                | b"r"
                | b"fldSimple"
                | b"hyperlink"
                | b"sdt"
                | b"bookmarkStart"
                | b"bookmarkEnd"
                | b"commentRangeStart"
                | b"commentRangeEnd"
                | b"ins"
                | b"del"
                | b"moveFrom"
                | b"moveTo"
        ),
        Some(b"r") => matches!(
            local_name,
            b"rPr"
                | b"t"
                | b"delText"
                | b"tab"
                | b"br"
                | b"drawing"
                | b"fldChar"
                | b"instrText"
                | b"commentReference"
                | b"footnoteReference"
                | b"endnoteReference"
        ),
        Some(b"hyperlink") => {
            matches!(local_name, b"r" | b"ins" | b"del" | b"moveFrom" | b"moveTo")
        }
        Some(b"sdt") => matches!(local_name, b"sdtPr" | b"sdtEndPr" | b"sdtContent"),
        Some(b"sdtContent") => {
            matches!(local_name, b"p" | b"tbl" | b"tr" | b"tc" | b"r" | b"sdt")
        }
        Some(b"tbl") => matches!(local_name, b"tblPr" | b"tblGrid" | b"tr" | b"sdt"),
        Some(b"tblGrid") => local_name == b"gridCol",
        Some(b"tr") => matches!(local_name, b"trPr" | b"tc" | b"sdt"),
        Some(b"tc") => matches!(local_name, b"tcPr" | b"p" | b"tbl" | b"sdt"),
        _ => false,
    }
}

fn modeled_owner_spans(xml: &[u8]) -> Result<Vec<ModeledOwnerSpan>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut inside_body = false;
    let mut depth = 0usize;
    let mut owner_stack = Vec::new();
    let mut scope_stack: Vec<Vec<(String, String)>> = Vec::new();
    let mut spans: Vec<ModeledOwnerSpan> = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid nested namespace scope: {error}")))?;
        let is_word = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri)) if uri == WORD_NAMESPACE.as_bytes()
        );
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local_name = element.local_name();
                let mut bindings = scope_stack.last().cloned().unwrap_or_default();
                apply_namespace_declarations(&element, &mut bindings)?;
                if inside_body && is_word && is_modeled_owner(local_name.as_ref()) {
                    let index = spans.len();
                    spans.push(ModeledOwnerSpan {
                        local_name: std::str::from_utf8(local_name.as_ref())
                            .map_err(|error| {
                                Error::Other(format!("invalid nested element name: {error}"))
                            })?
                            .to_owned(),
                        start,
                        end: 0,
                        declarations: namespace_declarations(&element)?,
                        bindings: bindings.clone(),
                    });
                    owner_stack.push((depth, index));
                }
                if is_word && local_name.as_ref() == b"body" {
                    inside_body = true;
                }
                scope_stack.push(bindings);
                depth += 1;
            }
            Event::Empty(element) => {
                let local_name = element.local_name();
                let mut bindings = scope_stack.last().cloned().unwrap_or_default();
                apply_namespace_declarations(&element, &mut bindings)?;
                if inside_body && is_word && is_modeled_owner(local_name.as_ref()) {
                    spans.push(ModeledOwnerSpan {
                        local_name: std::str::from_utf8(local_name.as_ref())
                            .map_err(|error| {
                                Error::Other(format!("invalid nested element name: {error}"))
                            })?
                            .to_owned(),
                        start,
                        end,
                        declarations: namespace_declarations(&element)?,
                        bindings,
                    });
                }
            }
            Event::End(element) => {
                depth = depth.saturating_sub(1);
                scope_stack.pop();
                if owner_stack
                    .last()
                    .is_some_and(|(owner_depth, _)| *owner_depth == depth)
                {
                    let (_, index) = owner_stack.pop().expect("checked owner stack");
                    spans[index].end = end;
                }
                if inside_body && is_word && matches_local_name(element.name().as_ref(), b"body") {
                    inside_body = false;
                }
            }
            Event::Eof => {
                if spans.iter().any(|span| span.end == 0) {
                    return Err(Error::Other(
                        "unterminated modeled namespace owner".to_owned(),
                    ));
                }
                return Ok(spans);
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn element_declared_prefixes(
    element: &BytesStart<'_>,
    declarations: &[(String, String)],
    bindings: &[(String, String)],
    parent_local_name: Option<&[u8]>,
    shadowed_owner_prefixes: &HashSet<String>,
    allow_redundant_local_bindings: bool,
) -> Vec<String> {
    let prefixes = declarations
        .iter()
        .map(|(name, value)| (namespace_prefix(name), value.as_str()))
        .collect::<Vec<_>>();
    let local_declarations = local_namespace_prefixes(element);
    let mut used = Vec::new();
    let element_is_modeled_word =
        is_modeled_word_child(parent_local_name, element.local_name().as_ref());
    let qualified_name = element.name();
    let element_prefix = std::str::from_utf8(qualified_name.as_ref())
        .ok()
        .map(|name| name.split_once(':').map_or("", |(prefix, _)| prefix));
    if element_prefix.is_some_and(|prefix| {
        prefixes.iter().any(|(candidate, value)| {
            candidate == prefix
                && !shadowed_owner_prefixes.contains(prefix)
                && inherited_namespace(bindings, prefix) == Some(*value)
                && !(*value == WORD_NAMESPACE && element_is_modeled_word)
        }) && (!local_declarations.contains(element_prefix.unwrap_or_default())
            || allow_redundant_local_bindings)
    }) {
        used.push(element_prefix.unwrap_or_default().to_owned());
    }
    for attribute in element.attributes().flatten() {
        let Some(name) = std::str::from_utf8(attribute.key.as_ref()).ok() else {
            continue;
        };
        let Some((prefix, _)) = name.split_once(':') else {
            continue;
        };
        if prefix != "xmlns"
            && prefixes.iter().any(|(candidate, value)| {
                candidate == prefix
                    && !shadowed_owner_prefixes.contains(prefix)
                    && inherited_namespace(bindings, prefix) == Some(*value)
                    && !(*value == WORD_NAMESPACE && element_is_modeled_word)
            })
            && (!local_declarations.contains(prefix) || allow_redundant_local_bindings)
            && !used.iter().any(|candidate| candidate == prefix)
        {
            used.push(prefix.to_owned());
        }
    }
    used
}

fn local_namespace_prefixes(element: &BytesStart<'_>) -> HashSet<String> {
    element
        .attributes()
        .flatten()
        .filter_map(|attribute| {
            let name = std::str::from_utf8(attribute.key.as_ref()).ok()?;
            (name == "xmlns" || name.starts_with("xmlns:")).then(|| namespace_prefix(name))
        })
        .collect()
}

fn resolved_snapshot_name(
    qualified_name: &[u8],
    bindings: &[(String, String)],
    default_namespace_applies: bool,
) -> Result<String> {
    let qualified_name = std::str::from_utf8(qualified_name)
        .map_err(|error| Error::Other(format!("invalid logical owner name: {error}")))?;
    let (prefix, local_name) = qualified_name
        .split_once(':')
        .map_or(("", qualified_name), |(prefix, local)| (prefix, local));
    let namespace = if prefix == "xml" {
        Some(XML_NAMESPACE)
    } else if prefix.is_empty() && !default_namespace_applies {
        None
    } else {
        inherited_namespace(bindings, prefix)
    };
    let namespace = match namespace {
        Some(WORD_NAMESPACE) => WORD_NAMESPACE.to_owned(),
        Some(XML_NAMESPACE) => XML_NAMESPACE.to_owned(),
        _ if !prefix.is_empty() => format!("#prefix:{prefix}"),
        Some(namespace) => namespace.to_owned(),
        None => String::new(),
    };
    Ok(format!("{{{namespace}}}{local_name}"))
}

fn logical_owner_snapshot(
    owner_xml: &[u8],
    inherited_bindings: &[(String, String)],
    prospective_declarations: &[(String, String)],
) -> Result<LogicalOwnerSnapshot> {
    let mut reader = quick_xml::Reader::from_reader(owner_xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut owner_bindings = inherited_bindings.to_vec();
    owner_bindings.extend(prospective_declarations.iter().cloned());
    let mut scope_stack = vec![owner_bindings];
    let mut structure = Vec::new();
    let mut semantic = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid logical owner XML: {error}")))?;
        match event {
            event @ (Event::Start(_) | Event::Empty(_)) => {
                let is_start = matches!(&event, Event::Start(_));
                let element = match event {
                    Event::Start(element) | Event::Empty(element) => element,
                    _ => unreachable!(),
                };
                let mut bindings = scope_stack.last().cloned().unwrap_or_default();
                apply_namespace_declarations(&element, &mut bindings)?;
                let name = resolved_snapshot_name(element.name().as_ref(), &bindings, true)?;
                let mut attribute_names = Vec::new();
                let mut attribute_values = Vec::new();
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Other(format!("invalid logical owner attribute: {error}"))
                    })?;
                    if attribute.key.as_ref() == b"xmlns"
                        || attribute.key.as_ref().starts_with(b"xmlns:")
                    {
                        continue;
                    }
                    let attribute_name =
                        resolved_snapshot_name(attribute.key.as_ref(), &bindings, false)?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                        .map_err(|error| {
                            Error::Other(format!("invalid logical owner attribute: {error}"))
                        })?
                        .into_owned();
                    attribute_names.push(attribute_name.clone());
                    attribute_values.push(format!("{attribute_name}={value}"));
                }
                attribute_names.sort();
                attribute_values.sort();
                structure.push(format!("start:{name}:[{}]", attribute_names.join(",")));
                semantic.push(format!("start:{name}:[{}]", attribute_values.join(",")));
                if is_start {
                    scope_stack.push(bindings);
                } else {
                    structure.push(format!("end:{name}"));
                    semantic.push(format!("end:{name}"));
                }
            }
            Event::End(element) => {
                let bindings = scope_stack.last().cloned().unwrap_or_default();
                let name = resolved_snapshot_name(element.name().as_ref(), &bindings, true)?;
                structure.push(format!("end:{name}"));
                semantic.push(format!("end:{name}"));
                scope_stack.pop();
            }
            Event::Text(text) => {
                if xml_bytes_are_whitespace(text.as_ref()) {
                    buffer.clear();
                    continue;
                }
                structure.push("text".to_owned());
                semantic.push(format!("text:{}", String::from_utf8_lossy(text.as_ref())));
            }
            Event::CData(text) => {
                structure.push("cdata".to_owned());
                semantic.push(format!("cdata:{}", String::from_utf8_lossy(text.as_ref())));
            }
            Event::GeneralRef(reference) => {
                structure.push("reference".to_owned());
                semantic.push(format!(
                    "reference:{}",
                    String::from_utf8_lossy(reference.as_ref())
                ));
            }
            Event::Comment(comment) => {
                structure.push("comment".to_owned());
                semantic.push(format!(
                    "comment:{}",
                    String::from_utf8_lossy(comment.as_ref())
                ));
            }
            Event::PI(instruction) => {
                structure.push("processing-instruction".to_owned());
                semantic.push(format!(
                    "processing-instruction:{}",
                    String::from_utf8_lossy(instruction.as_ref())
                ));
            }
            Event::Eof => {
                return Ok(LogicalOwnerSnapshot {
                    structure,
                    semantic,
                });
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn marker_raw_without_owner_declarations(
    raw: &[u8],
    declarations: &[(String, String)],
) -> Result<Vec<u8>> {
    let mut reader = quick_xml::Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let element = match reader
        .read_event_into(&mut buffer)
        .map_err(|error| Error::Other(format!("invalid namespace marker XML: {error}")))?
    {
        Event::Start(element) | Event::Empty(element) => element,
        _ => return Ok(raw.to_vec()),
    };
    let removable = namespace_declarations(&element)?
        .into_iter()
        .filter(|candidate| declarations.iter().any(|expected| expected == candidate))
        .map(|(name, _)| name.into_bytes())
        .collect::<HashSet<_>>();
    if removable.is_empty() {
        return Ok(raw.to_vec());
    }

    let mut cursor = 1usize;
    while cursor < raw.len()
        && !raw[cursor].is_ascii_whitespace()
        && !matches!(raw[cursor], b'/' | b'>')
    {
        cursor += 1;
    }
    let mut removals = Vec::new();
    loop {
        let whitespace_start = cursor;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= raw.len() || matches!(raw[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < raw.len()
            && !raw[cursor].is_ascii_whitespace()
            && !matches!(raw[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= raw.len() || raw[cursor] != b'=' {
            return Ok(raw.to_vec());
        }
        cursor += 1;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= raw.len() || !matches!(raw[cursor], b'\'' | b'"') {
            return Ok(raw.to_vec());
        }
        let quote = raw[cursor];
        cursor += 1;
        while cursor < raw.len() && raw[cursor] != quote {
            cursor += 1;
        }
        if cursor >= raw.len() {
            return Ok(raw.to_vec());
        }
        cursor += 1;
        if removable.contains(&raw[name_start..name_end]) {
            removals.push((whitespace_start, cursor));
        }
    }
    if removals.is_empty() {
        return Ok(raw.to_vec());
    }
    let mut normalized = Vec::with_capacity(raw.len());
    let mut copied = 0usize;
    for (start, end) in removals {
        normalized.extend_from_slice(&raw[copied..start]);
        copied = end;
    }
    normalized.extend_from_slice(&raw[copied..]);
    Ok(normalized)
}

fn namespace_owner_markers(
    owner_xml: &[u8],
    declarations: &[(String, String)],
    inherited_bindings: &[(String, String)],
    allow_redundant_local_bindings: bool,
) -> Result<Vec<NamespaceMarker>> {
    let mut reader = quick_xml::Reader::from_reader(owner_xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut element_stack: Vec<Vec<u8>> = Vec::new();
    let mut empty_markers = Vec::new();
    let mut expanded_markers = Vec::new();
    // Candidate owners are inspected before their retained declarations have
    // been replayed. Seed those prospective bindings at the owner boundary so
    // descendant uses can still be matched, while declarations encountered
    // inside the raw subtree continue to shadow them normally.
    let mut owner_bindings = inherited_bindings.to_vec();
    owner_bindings.extend(declarations.iter().cloned());
    let mut scope_stack = vec![owner_bindings];
    let mut shadow_stack = vec![HashSet::<String>::new()];
    let mut depth = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid namespace owner marker: {error}")))?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let mut bindings = scope_stack.last().cloned().unwrap_or_default();
                apply_namespace_declarations(&element, &mut bindings)?;
                let mut shadowed = shadow_stack.last().cloned().unwrap_or_default();
                if depth > 0 {
                    shadowed.extend(local_namespace_prefixes(&element));
                    if allow_redundant_local_bindings {
                        for (name, value) in namespace_declarations(&element)? {
                            let prefix = namespace_prefix(&name);
                            if declarations.iter().any(|(candidate, owner_value)| {
                                namespace_prefix(candidate) == prefix && value == *owner_value
                            }) {
                                shadowed.remove(&prefix);
                            }
                        }
                    }
                }
                let used_prefixes = if depth > 0 {
                    element_declared_prefixes(
                        &element,
                        declarations,
                        &bindings,
                        element_stack.last().map(Vec::as_slice),
                        &shadowed,
                        allow_redundant_local_bindings,
                    )
                } else {
                    Vec::new()
                };
                let facts = used_prefixes
                    .into_iter()
                    .map(|prefix| {
                        let namespace = inherited_namespace(&bindings, &prefix).map(str::to_owned);
                        (prefix, namespace)
                    })
                    .collect::<Vec<_>>();
                stack.push((start, facts));
                element_stack.push(element.local_name().as_ref().to_vec());
                scope_stack.push(bindings);
                shadow_stack.push(shadowed);
                depth += 1;
            }
            Event::Empty(element) => {
                let mut bindings = scope_stack.last().cloned().unwrap_or_default();
                apply_namespace_declarations(&element, &mut bindings)?;
                let mut shadowed = shadow_stack.last().cloned().unwrap_or_default();
                if depth > 0 {
                    shadowed.extend(local_namespace_prefixes(&element));
                    if allow_redundant_local_bindings {
                        for (name, value) in namespace_declarations(&element)? {
                            let prefix = namespace_prefix(&name);
                            if declarations.iter().any(|(candidate, owner_value)| {
                                namespace_prefix(candidate) == prefix && value == *owner_value
                            }) {
                                shadowed.remove(&prefix);
                            }
                        }
                    }
                }
                let used_prefixes = if depth > 0 {
                    element_declared_prefixes(
                        &element,
                        declarations,
                        &bindings,
                        element_stack.last().map(Vec::as_slice),
                        &shadowed,
                        allow_redundant_local_bindings,
                    )
                } else {
                    Vec::new()
                };
                if !used_prefixes.is_empty() {
                    let namespace_facts = used_prefixes
                        .into_iter()
                        .map(|prefix| {
                            let namespace =
                                inherited_namespace(&bindings, &prefix).map(str::to_owned);
                            (prefix, namespace)
                        })
                        .collect();
                    empty_markers.push(NamespaceMarker {
                        raw: marker_raw_without_owner_declarations(
                            &owner_xml[start..end],
                            declarations,
                        )?,
                        namespace_facts,
                    });
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                scope_stack.pop();
                shadow_stack.pop();
                element_stack.pop();
                if let Some((marker_start, namespace_facts)) = stack.pop()
                    && !namespace_facts.is_empty()
                {
                    expanded_markers.push(NamespaceMarker {
                        raw: marker_raw_without_owner_declarations(
                            &owner_xml[marker_start..end],
                            declarations,
                        )?,
                        namespace_facts,
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    empty_markers.extend(expanded_markers);
    Ok(empty_markers)
}

fn candidate_namespace_owner_markers(
    owner_xml: &[u8],
    declarations: &[(String, String)],
    inherited_bindings: &[(String, String)],
    expected: &[NamespaceMarker],
) -> Result<Vec<NamespaceMarker>> {
    let candidates = namespace_owner_markers(owner_xml, declarations, inherited_bindings, true)?;
    let mut matched = vec![false; expected.len()];
    Ok(candidates
        .into_iter()
        .filter(|candidate| {
            let expected_index = expected
                .iter()
                .enumerate()
                .position(|(index, marker)| !matched[index] && marker.raw == candidate.raw);
            if let Some(index) = expected_index {
                matched[index] = true;
                true
            } else {
                false
            }
        })
        .collect())
}

fn marker_multiset_matches(
    expected: &[NamespaceMarker],
    actual: &[NamespaceMarker],
    include_namespace_facts: bool,
) -> bool {
    if expected.len() != actual.len() {
        return false;
    }
    let mut matched = vec![false; actual.len()];
    expected.iter().all(|marker| {
        let candidate = actual.iter().enumerate().position(|(index, candidate)| {
            !matched[index]
                && candidate.raw == marker.raw
                && (!include_namespace_facts || candidate.namespace_facts == marker.namespace_facts)
        });
        if let Some(index) = candidate {
            matched[index] = true;
            true
        } else {
            false
        }
    })
}

fn nested_modeled_namespace_owners(xml: &[u8]) -> Result<Vec<NestedNamespaceOwner>> {
    let spans = modeled_owner_spans(xml)?;
    let mut owners = Vec::new();
    for span in &spans {
        let mut declarations = span.declarations.clone();
        if declarations.is_empty() {
            continue;
        }
        let markers = namespace_owner_markers(
            &xml[span.start..span.end],
            &declarations,
            &span.bindings,
            false,
        )?;
        if markers.is_empty() {
            continue;
        }
        let used_prefixes = markers
            .iter()
            .flat_map(|marker| marker.namespace_facts.iter())
            .map(|(prefix, _)| prefix.as_str())
            .collect::<HashSet<_>>();
        declarations.retain(|(name, _)| used_prefixes.contains(namespace_prefix(name).as_str()));
        let snapshot =
            logical_owner_snapshot(&xml[span.start..span.end], &span.bindings, &declarations)?;
        let ambiguous_without_namespace = spans.iter().any(|candidate| {
            if candidate.start == span.start || candidate.local_name != span.local_name {
                return false;
            }
            let Ok(candidate_snapshot) = logical_owner_snapshot(
                &xml[candidate.start..candidate.end],
                &candidate.bindings,
                &declarations,
            ) else {
                return true;
            };
            if candidate_snapshot != snapshot {
                return false;
            }
            let Ok(candidate_markers) = candidate_namespace_owner_markers(
                &xml[candidate.start..candidate.end],
                &declarations,
                &candidate.bindings,
                &markers,
            ) else {
                return true;
            };
            marker_multiset_matches(&markers, &candidate_markers, false)
        });
        let same_marker_structural_alternate = spans.iter().any(|candidate| {
            if candidate.start == span.start || candidate.local_name != span.local_name {
                return false;
            }
            let Ok(candidate_snapshot) = logical_owner_snapshot(
                &xml[candidate.start..candidate.end],
                &candidate.bindings,
                &declarations,
            ) else {
                return true;
            };
            if candidate_snapshot.structure != snapshot.structure {
                return false;
            }
            let Ok(candidate_markers) = candidate_namespace_owner_markers(
                &xml[candidate.start..candidate.end],
                &declarations,
                &candidate.bindings,
                &markers,
            ) else {
                return true;
            };
            marker_multiset_matches(&markers, &candidate_markers, false)
        });
        let same_namespace_structural_alternate = spans.iter().any(|candidate| {
            if candidate.start == span.start || candidate.local_name != span.local_name {
                return false;
            }
            let Ok(candidate_snapshot) = logical_owner_snapshot(
                &xml[candidate.start..candidate.end],
                &candidate.bindings,
                &declarations,
            ) else {
                return true;
            };
            if candidate_snapshot.structure != snapshot.structure {
                return false;
            }
            let Ok(candidate_markers) = candidate_namespace_owner_markers(
                &xml[candidate.start..candidate.end],
                &declarations,
                &candidate.bindings,
                &markers,
            ) else {
                return true;
            };
            marker_multiset_matches(&markers, &candidate_markers, true)
        });
        owners.push(NestedNamespaceOwner {
            local_name: span.local_name.clone(),
            declarations,
            snapshot,
            markers,
            ambiguous_without_namespace,
            same_marker_structural_alternate,
            same_namespace_structural_alternate,
        });
    }
    Ok(owners)
}

fn unsafe_nested_namespace_prefix(owners: &[NestedNamespaceOwner]) -> Option<String> {
    owners.iter().find_map(|owner| {
        owner.declarations.iter().find_map(|(name, value)| {
            let prefix = namespace_prefix(name);
            if owner.local_name == "hyperlink" && prefix == "r" {
                return None;
            }
            canonical_serializer_namespace(&prefix)
                .is_some_and(|expected| value != expected)
                .then_some(prefix)
        })
    })
}

fn replay_nested_namespace_declarations(
    xml: &[u8],
    owners: &[NestedNamespaceOwner],
) -> Result<Vec<u8>> {
    if owners.is_empty() {
        return Ok(xml.to_vec());
    }
    let spans = modeled_owner_spans(xml)?;
    let mut targets: Vec<(usize, &NestedNamespaceOwner)> = Vec::new();
    for owner in owners {
        let mut candidates = Vec::new();
        for span in spans
            .iter()
            .filter(|span| span.local_name == owner.local_name)
        {
            let snapshot = logical_owner_snapshot(
                &xml[span.start..span.end],
                &span.bindings,
                &owner.declarations,
            )?;
            if snapshot.structure != owner.snapshot.structure {
                continue;
            }
            let markers = candidate_namespace_owner_markers(
                &xml[span.start..span.end],
                &owner.declarations,
                &span.bindings,
                &owner.markers,
            )?;
            if marker_multiset_matches(&owner.markers, &markers, false) {
                candidates.push((span, snapshot, markers));
            }
        }
        let namespace_candidates = candidates
            .iter()
            .filter(|(_, _, markers)| marker_multiset_matches(&owner.markers, markers, true))
            .collect::<Vec<_>>();
        let target = match namespace_candidates.as_slice() {
            [candidate] if !owner.same_namespace_structural_alternate => candidate.0,
            [candidate]
                if !owner.ambiguous_without_namespace
                    && candidate.1.semantic == owner.snapshot.semantic =>
            {
                candidate.0
            }
            [] if !owner.ambiguous_without_namespace => {
                let semantic_candidates = candidates
                    .iter()
                    .filter(|(_, snapshot, _)| snapshot.semantic == owner.snapshot.semantic)
                    .collect::<Vec<_>>();
                match semantic_candidates.as_slice() {
                    [candidate] => candidate.0,
                    [] if candidates.len() == 1 && !owner.same_marker_structural_alternate => {
                        candidates[0].0
                    }
                    _ => {
                        return Err(Error::Other(format!(
                            "cannot identify retained `{}` nested namespace owner after mutation",
                            owner.local_name
                        )));
                    }
                }
            }
            candidates => {
                if owner.ambiguous_without_namespace && owner.same_namespace_structural_alternate {
                    return Err(Error::Other(format!(
                        "cannot identify retained `{}` nested namespace owner after mutation",
                        owner.local_name
                    )));
                }
                let semantic_candidates = candidates
                    .iter()
                    .filter(|candidate| candidate.1.semantic == owner.snapshot.semantic)
                    .collect::<Vec<_>>();
                let [candidate] = semantic_candidates.as_slice() else {
                    return Err(Error::Other(format!(
                        "cannot identify retained `{}` nested namespace owner after mutation",
                        owner.local_name
                    )));
                };
                candidate.0
            }
        };
        if owner.declarations.iter().all(|declaration| {
            target
                .declarations
                .iter()
                .any(|candidate| candidate == declaration)
        }) {
            continue;
        }
        targets.push((target.start, owner));
    }
    if targets.is_empty() {
        return Ok(xml.to_vec());
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::new());
    let mut buffer = Vec::new();
    let mut replayed = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let (_namespace, event) =
            reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| {
                    Error::Other(format!("invalid serialized namespace scope: {error}"))
                })?;
        let starts_scope = matches!(&event, Event::Start(_));
        match event {
            Event::Start(mut element) | Event::Empty(mut element) => {
                if let Some((_, owner)) = targets.iter().find(|(target, _)| *target == start) {
                    let existing = namespace_declarations(&element)?;
                    for (name, value) in &owner.declarations {
                        match existing.iter().find(|(candidate, _)| candidate == name) {
                            Some((_, existing_value)) if existing_value == value => {}
                            Some(_) => {
                                return Err(Error::Other(format!(
                                    "cannot replay conflicting `{}` namespace after mutation",
                                    namespace_prefix(name),
                                )));
                            }
                            None => element.push_attribute((name.as_str(), value.as_str())),
                        }
                    }
                    replayed += 1;
                }
                writer.write_event(if starts_scope {
                    Event::Start(element)
                } else {
                    Event::Empty(element)
                })?;
            }
            Event::End(element) => writer.write_event(Event::End(element))?,
            Event::Eof => {
                if replayed < targets.len() {
                    return Err(Error::Other(
                        "cannot align retained nested namespace declarations after mutation"
                            .to_owned(),
                    ));
                }
                return Ok(writer.into_inner());
            }
            event => writer.write_event(event)?,
        }
        buffer.clear();
    }
}

/// One body child exposed through the compatibility reader facade.
#[non_exhaustive]
pub enum BodyContentRef<'a> {
    /// A body paragraph.
    Paragraph(ParagraphRef<'a>),
    /// A body table.
    Table(TableRef<'a>),
    /// Content that the compatibility facade does not model.
    UnsupportedXml(UnsupportedXmlRef<'a>),
}

/// Package and WordprocessingML identity state owned by the document facade.
///
/// Each set represents one OOXML scope. References are deliberately excluded
/// from definition sets, so a bookmark end, comment anchor, or numbering use
/// does not collide with the definition it names.
#[derive(Clone, Debug)]
pub(crate) struct DocumentIdentifiers {
    relationship_ids: HashMap<String, HashSet<String>>,
    bookmark_ids: HashSet<i32>,
    comment_ids: HashSet<i32>,
    drawing_ids: HashSet<u32>,
    abstract_numbering_ids: HashSet<u32>,
    numbering_instance_ids: HashSet<u32>,
    part_names: HashSet<String>,
    content_type_defaults: HashSet<String>,
    content_type_overrides: HashSet<String>,
    authored_content_type_defaults: HashSet<String>,
    preserved_relationship_ids: HashMap<String, HashSet<String>>,
    preserved_bookmark_ids: HashSet<i32>,
    preserved_comment_ids: HashSet<i32>,
    preserved_drawing_ids: HashSet<u32>,
    preserved_abstract_numbering_ids: HashSet<u32>,
    preserved_numbering_instance_ids: HashSet<u32>,
    preserved_part_names: HashSet<String>,
    authored_bookmark_ids: HashSet<i32>,
    authored_comment_ids: HashSet<i32>,
    authored_toc_bookmark_ids: HashSet<i32>,
    authored_bookmark_names: HashSet<String>,
    authored_story_parts: HashSet<String>,
    authored_story_relationship_ids: HashMap<String, HashSet<String>>,
    authored_bundle_relationship_ids: HashMap<String, HashSet<String>>,
    provisional_relationship_ids: HashMap<String, HashSet<String>>,
}

impl DocumentIdentifiers {
    fn scan(package: &OpcPackage) -> Result<Self> {
        let mut identifiers = Self {
            relationship_ids: HashMap::new(),
            bookmark_ids: HashSet::new(),
            comment_ids: HashSet::new(),
            drawing_ids: HashSet::new(),
            abstract_numbering_ids: HashSet::new(),
            numbering_instance_ids: HashSet::new(),
            part_names: package
                .parts
                .keys()
                .map(|part_name| part_name_identity(part_name))
                .collect(),
            content_type_defaults: package
                .content_types
                .defaults
                .keys()
                .map(|extension| extension.to_ascii_lowercase())
                .collect(),
            content_type_overrides: package
                .content_types
                .overrides
                .keys()
                .map(|part_name| part_name.to_ascii_lowercase())
                .collect(),
            authored_content_type_defaults: HashSet::new(),
            preserved_relationship_ids: HashMap::new(),
            preserved_bookmark_ids: HashSet::new(),
            preserved_comment_ids: HashSet::new(),
            preserved_drawing_ids: HashSet::new(),
            preserved_abstract_numbering_ids: HashSet::new(),
            preserved_numbering_instance_ids: HashSet::new(),
            preserved_part_names: HashSet::new(),
            authored_bookmark_ids: HashSet::new(),
            authored_comment_ids: HashSet::new(),
            authored_toc_bookmark_ids: HashSet::new(),
            authored_bookmark_names: HashSet::new(),
            authored_story_parts: HashSet::new(),
            authored_story_relationship_ids: HashMap::new(),
            authored_bundle_relationship_ids: HashMap::new(),
            provisional_relationship_ids: HashMap::new(),
        };
        identifiers.part_names.extend(
            package
                .part_rels
                .keys()
                .map(|name| part_name_identity(name)),
        );
        identifiers.part_names.extend(
            package
                .content_types
                .overrides
                .keys()
                .map(|name| part_name_identity(name)),
        );

        identifiers.scan_relationships("/", &package.package_rels.items)?;
        let mut relationship_owners = package.part_rels.iter().collect::<Vec<_>>();
        relationship_owners.sort_by_key(|(owner, _)| owner.as_str());
        for (owner, relationships) in relationship_owners {
            identifiers.scan_relationships(owner, &relationships.items)?;
        }
        let mut identity_parts: HashMap<String, HashSet<IdentifierXmlCategory>> = HashMap::new();
        if let Some(main_part) = package.main_document_part() {
            identity_parts
                .entry(part_name_identity(&main_part))
                .or_default()
                .extend([
                    IdentifierXmlCategory::Bookmark,
                    IdentifierXmlCategory::Drawing,
                ]);
            if let Some(relationships) = package.get_part_rels(&main_part) {
                for relationship in &relationships.items {
                    let categories: &[IdentifierXmlCategory] = match relationship.rel_type.as_str()
                    {
                        rel_types::HEADER
                        | rel_types::FOOTER
                        | rel_types::FOOTNOTES
                        | rel_types::ENDNOTES
                        | rel_types::GLOSSARY_DOCUMENT => &[
                            IdentifierXmlCategory::Bookmark,
                            IdentifierXmlCategory::Drawing,
                        ],
                        rel_types::COMMENTS => &[
                            IdentifierXmlCategory::Bookmark,
                            IdentifierXmlCategory::Comment,
                            IdentifierXmlCategory::Drawing,
                        ],
                        rel_types::NUMBERING => &[
                            IdentifierXmlCategory::AbstractNumbering,
                            IdentifierXmlCategory::NumberingInstance,
                        ],
                        _ => &[],
                    };
                    if !categories.is_empty() && relationship_is_internal(relationship) {
                        identity_parts
                            .entry(part_name_identity(&OpcPackage::resolve_rel_target(
                                &main_part,
                                &relationship.target,
                            )))
                            .or_default()
                            .extend(categories.iter().copied());
                    }
                }
            }
        }

        let mut parts = package.parts.iter().collect::<Vec<_>>();
        parts.sort_by_key(|(part_name, _)| part_name.as_str());
        for (part_name, bytes) in parts {
            if let Some(categories) = identity_parts.get(&part_name_identity(part_name))
                && identifier_xml_is_well_formed(bytes)
            {
                identifiers
                    .scan_xml_definitions(bytes, categories)
                    .map_err(|error| {
                        Error::Other(format!(
                            "cannot scan identifiers in XML part {part_name}: {error}"
                        ))
                    })?;
            }
        }
        identifiers.preserved_relationship_ids = identifiers.relationship_ids.clone();
        identifiers.preserved_bookmark_ids = identifiers.bookmark_ids.clone();
        identifiers.preserved_comment_ids = identifiers.comment_ids.clone();
        identifiers.preserved_drawing_ids = identifiers.drawing_ids.clone();
        identifiers.preserved_abstract_numbering_ids = identifiers.abstract_numbering_ids.clone();
        identifiers.preserved_numbering_instance_ids = identifiers.numbering_instance_ids.clone();
        identifiers.preserved_part_names = identifiers.part_names.clone();
        Ok(identifiers)
    }

    fn scan_relationships(
        &mut self,
        owner: &str,
        relationships: &[oxml_opc::relationship::Relationship],
    ) -> Result<()> {
        let owner_identity = relationship_owner_identity(owner);
        for relationship in relationships {
            if !self
                .relationship_ids
                .entry(owner_identity.clone())
                .or_default()
                .insert(relationship.id.clone())
            {
                return Err(Error::Other(format!(
                    "duplicate relationship id {} in owner {owner}",
                    relationship.id
                )));
            }
            if relationship_is_internal(relationship) {
                self.part_names
                    .insert(part_name_identity(&OpcPackage::resolve_rel_target(
                        owner,
                        &relationship.target,
                    )));
            }
        }
        Ok(())
    }

    fn scan_xml_definitions(
        &mut self,
        xml: &[u8],
        categories: &HashSet<IdentifierXmlCategory>,
    ) -> Result<()> {
        let mut reader = NsReader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = match reader.read_resolved_event_into(&mut buffer) {
                Ok(value) => value,
                Err(error) => return Err(Error::Other(error.to_string())),
            };
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    let local = element.local_name();
                    let is_drawing = matches!(
                        &namespace,
                        ResolveResult::Bound(namespace)
                            if namespace.as_ref() == drawing_ns::WP.as_bytes()
                    );
                    let is_word = matches!(
                        &namespace,
                        ResolveResult::Bound(namespace)
                            if namespace.as_ref() == WORD_NAMESPACE.as_bytes()
                    );
                    let category = if is_drawing && local.as_ref() == b"docPr" {
                        Some(IdentifierXmlCategory::Drawing)
                    } else if is_word && local.as_ref() == b"bookmarkStart" {
                        Some(IdentifierXmlCategory::Bookmark)
                    } else if is_word && local.as_ref() == b"comment" {
                        Some(IdentifierXmlCategory::Comment)
                    } else if is_word && local.as_ref() == b"abstractNum" {
                        Some(IdentifierXmlCategory::AbstractNumbering)
                    } else if is_word && local.as_ref() == b"num" {
                        Some(IdentifierXmlCategory::NumberingInstance)
                    } else {
                        None
                    };
                    if let Some(category) = category.filter(|value| categories.contains(value)) {
                        self.scan_xml_id(&reader, element, category)?;
                    }
                }
                Event::Eof => return Ok(()),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn scan_xml_id(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        category: IdentifierXmlCategory,
    ) -> Result<()> {
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            let accepted = match category {
                IdentifierXmlCategory::Drawing => {
                    matches!(namespace, ResolveResult::Unbound) && local.as_ref() == b"id"
                }
                IdentifierXmlCategory::Bookmark | IdentifierXmlCategory::Comment => {
                    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == WORD_NAMESPACE.as_bytes())
                        && local.as_ref() == b"id"
                }
                IdentifierXmlCategory::AbstractNumbering => {
                    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == WORD_NAMESPACE.as_bytes())
                        && local.as_ref() == b"abstractNumId"
                }
                IdentifierXmlCategory::NumberingInstance => {
                    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == WORD_NAMESPACE.as_bytes())
                        && local.as_ref() == b"numId"
                }
            };
            if !accepted {
                continue;
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| Error::Other(error.to_string()))?;
            match category {
                IdentifierXmlCategory::Drawing => {
                    let value = value
                        .parse::<u32>()
                        .map_err(|_| Error::Other(format!("invalid drawing id {value}")))?;
                    Self::insert_unique(&mut self.drawing_ids, value, "drawing")?;
                }
                IdentifierXmlCategory::Bookmark => {
                    let value = value
                        .parse::<i32>()
                        .map_err(|_| Error::Other(format!("invalid bookmark id {value}")))?;
                    Self::insert_unique(&mut self.bookmark_ids, value, "bookmark")?;
                }
                IdentifierXmlCategory::Comment => {
                    let value = value
                        .parse::<i32>()
                        .map_err(|_| Error::Other(format!("invalid comment id {value}")))?;
                    Self::insert_unique(&mut self.comment_ids, value, "comment")?;
                }
                IdentifierXmlCategory::AbstractNumbering => {
                    let value = value.parse::<u32>().map_err(|_| {
                        Error::Other(format!("invalid abstract numbering id {value}"))
                    })?;
                    Self::insert_unique(
                        &mut self.abstract_numbering_ids,
                        value,
                        "abstract numbering",
                    )?;
                }
                IdentifierXmlCategory::NumberingInstance => {
                    let value = value.parse::<u32>().map_err(|_| {
                        Error::Other(format!("invalid numbering instance id {value}"))
                    })?;
                    Self::insert_unique(
                        &mut self.numbering_instance_ids,
                        value,
                        "numbering instance",
                    )?;
                }
            }
            break;
        }
        Ok(())
    }

    fn insert_unique<T: std::hash::Hash + Eq + std::fmt::Display + Copy>(
        occupied: &mut HashSet<T>,
        value: T,
        category: &str,
    ) -> Result<()> {
        if occupied.insert(value) {
            Ok(())
        } else {
            Err(Error::Other(format!(
                "duplicate {category} id {value} in imported or preserved XML"
            )))
        }
    }

    pub(crate) fn reserve_bookmark_id(&mut self) -> Result<i32> {
        let id = reserve_lowest_i32(&mut self.bookmark_ids, 0, "bookmark")?;
        self.authored_bookmark_ids.insert(id);
        Ok(id)
    }

    pub(crate) fn reserve_preferred_bookmark_id(&mut self, preferred: i32) -> Result<i32> {
        if preferred >= 0 && self.bookmark_ids.insert(preferred) {
            Ok(preferred)
        } else {
            reserve_lowest_i32(&mut self.bookmark_ids, 0, "bookmark")
        }
    }

    pub(crate) fn restore_authored_bookmark(&mut self, id: i32, name: &str) {
        self.authored_toc_bookmark_ids.insert(id);
        self.authored_bookmark_names.insert(name.to_owned());
    }

    pub(crate) fn is_authored_bookmark(&self, id: i32) -> bool {
        self.authored_toc_bookmark_ids.contains(&id)
    }

    pub(crate) fn authored_bookmark_names(&self) -> &HashSet<String> {
        &self.authored_bookmark_names
    }

    pub(crate) fn reserve_comment_id(&mut self) -> Result<i32> {
        let id = reserve_i32(&mut self.comment_ids, 0, "comment")?;
        self.authored_comment_ids.insert(id);
        Ok(id)
    }

    pub(crate) fn reserve_drawing_id(&mut self) -> Result<u32> {
        reserve_u32(&mut self.drawing_ids, 1, "drawing")
    }

    fn reserve_numbering_ids(&mut self) -> Result<(u32, u32)> {
        let mut abstract_ids = self.abstract_numbering_ids.clone();
        let mut instance_ids = self.numbering_instance_ids.clone();
        let abstract_id = reserve_u32(&mut abstract_ids, 0, "abstract numbering")?;
        let instance_id = reserve_u32(&mut instance_ids, 1, "numbering instance")?;
        self.abstract_numbering_ids = abstract_ids;
        self.numbering_instance_ids = instance_ids;
        Ok((abstract_id, instance_id))
    }

    fn reserve_abstract_numbering_id(&mut self) -> Result<u32> {
        reserve_u32(&mut self.abstract_numbering_ids, 0, "abstract numbering")
    }

    fn reserve_numbering_instance_id(&mut self) -> Result<u32> {
        reserve_u32(&mut self.numbering_instance_ids, 1, "numbering instance")
    }

    pub(crate) fn reserve_relationship_id_checked(&mut self, owner: &str) -> Result<String> {
        self.ensure_relationship_capacity(owner, 1)?;
        let owner = relationship_owner_identity(owner);
        let occupied = self.relationship_ids.entry(owner).or_default();
        let maximum = occupied
            .iter()
            .filter_map(|id| id.strip_prefix("rId")?.parse::<u32>().ok())
            .max();
        let mut candidate = maximum.map_or(Ok(0), |value| {
            value
                .checked_add(1)
                .ok_or_else(|| Error::Other("relationship id range is exhausted".to_owned()))
        })?;
        loop {
            let id = format!("rId{candidate}");
            if occupied.insert(id.clone()) {
                return Ok(id);
            }
            candidate = candidate
                .checked_add(1)
                .ok_or_else(|| Error::Other("relationship id range is exhausted".to_owned()))?;
        }
    }

    fn reserve_bundle_relationship_id_checked(
        &mut self,
        owner: &str,
        rel_type: &str,
    ) -> Result<String> {
        let stem = match rel_type {
            rel_types::STYLES => "rdocxDeferredStyles",
            rel_types::NUMBERING => "rdocxDeferredNumbering",
            rel_types::SETTINGS => "rdocxDeferredSettings",
            rel_types::FOOTNOTES => "rdocxDeferredFootnotes",
            rel_types::COMMENTS => "rdocxDeferredComments",
            crate::comments::COMMENTS_EXTENDED_REL_TYPE => "rdocxDeferredCommentsExtended",
            _ => {
                return Err(Error::Other(format!(
                    "relationship type {rel_type} is not a document bundle"
                )));
            }
        };
        self.ensure_relationship_capacity(owner, 1)?;
        let owner = relationship_owner_identity(owner);
        let occupied = self.relationship_ids.entry(owner.clone()).or_default();
        let mut ordinal = 0u32;
        loop {
            let id = if ordinal == 0 {
                stem.to_owned()
            } else {
                format!("{stem}{ordinal}")
            };
            if occupied.insert(id.clone()) {
                self.authored_bundle_relationship_ids
                    .entry(owner.clone())
                    .or_default()
                    .insert(id.clone());
                self.provisional_relationship_ids
                    .entry(owner.clone())
                    .or_default()
                    .insert(id.clone());
                return Ok(id);
            }
            ordinal = ordinal.checked_add(1).ok_or_else(|| {
                Error::Other("bundle relationship id range is exhausted".to_owned())
            })?;
        }
    }

    pub(crate) fn reserve_requested_relationship_id_checked(
        &mut self,
        owner: &str,
        requested: &str,
    ) -> Result<String> {
        let owner = relationship_owner_identity(owner);
        let numeric_request = requested
            .strip_prefix("rId")
            .and_then(|value| value.parse::<u32>().ok());
        let is_available = !self
            .relationship_ids
            .get(&owner)
            .is_some_and(|occupied| occupied.contains(requested));
        if is_available {
            if numeric_request.is_some() {
                let provisional_count = self
                    .provisional_relationship_ids
                    .get(&owner)
                    .map_or(0, HashSet::len);
                let maximum = self
                    .relationship_ids
                    .get(&owner)
                    .into_iter()
                    .flatten()
                    .filter_map(|id| id.strip_prefix("rId")?.parse::<u32>().ok())
                    .chain(numeric_request)
                    .max();
                ensure_numeric_relationship_capacity(maximum, provisional_count)?;
            }
            self.relationship_ids
                .entry(owner.clone())
                .or_default()
                .insert(requested.to_owned());
            Ok(requested.to_owned())
        } else {
            self.reserve_relationship_id_checked(&owner)
        }
    }

    fn ensure_relationship_capacity(&self, owner: &str, additional: usize) -> Result<()> {
        let owner = relationship_owner_identity(owner);
        let maximum = self
            .relationship_ids
            .get(&owner)
            .into_iter()
            .flatten()
            .filter_map(|id| id.strip_prefix("rId")?.parse::<u32>().ok())
            .max();
        let provisional_count = self
            .provisional_relationship_ids
            .get(&owner)
            .map_or(0, HashSet::len)
            .checked_add(additional)
            .ok_or_else(|| Error::Other("relationship id range is exhausted".to_owned()))?;
        ensure_numeric_relationship_capacity(maximum, provisional_count)
    }

    fn reserve_part_name(
        &mut self,
        directory: &str,
        stem: &str,
        extension: &str,
    ) -> Result<String> {
        let mut namer = oxml_media::MediaNamer::scan(
            directory,
            stem,
            self.part_names.iter().map(String::as_str),
        );
        loop {
            let part_name = namer.next_part_name(extension);
            if self.part_names.insert(part_name_identity(&part_name)) {
                return Ok(part_name);
            }
        }
    }

    pub(crate) fn reserve_preferred_part_name(&mut self, preferred: &str) -> Result<String> {
        if !preferred.starts_with('/') {
            return Err(Error::Other(format!(
                "preferred part name {preferred} is not absolute"
            )));
        }
        let (directory, filename) = preferred.rsplit_once('/').ok_or_else(|| {
            Error::Other(format!("preferred part name {preferred} has no directory"))
        })?;
        let (name, extension) = filename.rsplit_once('.').ok_or_else(|| {
            Error::Other(format!("preferred part name {preferred} has no extension"))
        })?;
        if name.is_empty() || extension.is_empty() {
            return Err(Error::Other(format!(
                "preferred part name {preferred} has an empty stem or extension"
            )));
        }
        let digit_start = name
            .bytes()
            .rposition(|byte| !byte.is_ascii_digit())
            .map_or(0, |index| index + 1);
        let suffix = &name[digit_start..];
        let stem = suffix
            .parse::<usize>()
            .ok()
            .filter(|suffix| *suffix > 0)
            .map_or(name, |_| &name[..digit_start]);
        if stem.is_empty() {
            return Err(Error::Other(format!(
                "preferred part name {preferred} has no family stem"
            )));
        }
        if self.part_names.insert(part_name_identity(preferred)) {
            return Ok(preferred.to_owned());
        }
        self.reserve_part_name(directory, stem, extension)
    }

    fn register_content_type_default(&mut self, extension: &str) {
        let extension = extension.to_ascii_lowercase();
        self.content_type_defaults.insert(extension.clone());
        self.authored_content_type_defaults.insert(extension);
    }

    pub(crate) fn register_content_type_override(&mut self, part_name: &str) {
        self.content_type_overrides
            .insert(part_name.to_ascii_lowercase());
    }

    pub(crate) fn retire_authored_story_relationships(
        &mut self,
        owner: &str,
        relationship_ids: impl IntoIterator<Item = String>,
    ) {
        let owner = relationship_owner_identity(owner);
        let preserved = self
            .preserved_relationship_ids
            .get(&owner)
            .cloned()
            .unwrap_or_default();
        let removed = relationship_ids
            .into_iter()
            .filter(|id| !preserved.contains(id))
            .collect::<HashSet<_>>();
        if let Some(occupied) = self.relationship_ids.get_mut(&owner) {
            occupied.retain(|id| !removed.contains(id));
            if occupied.is_empty() {
                self.relationship_ids.remove(&owner);
            }
        }
        if let Some(authored) = self.authored_story_relationship_ids.get_mut(&owner) {
            authored.retain(|id| !removed.contains(id));
            if authored.is_empty() {
                self.authored_story_relationship_ids.remove(&owner);
            }
        }
        if let Some(authored) = self.authored_bundle_relationship_ids.get_mut(&owner) {
            authored.retain(|id| !removed.contains(id));
            if authored.is_empty() {
                self.authored_bundle_relationship_ids.remove(&owner);
            }
        }
        if let Some(provisional) = self.provisional_relationship_ids.get_mut(&owner) {
            provisional.retain(|id| !removed.contains(id));
            if provisional.is_empty() {
                self.provisional_relationship_ids.remove(&owner);
            }
        }
    }

    pub(crate) fn retire_authored_part(&mut self, part_name: &str) {
        let identity = part_name_identity(part_name);
        if !self.preserved_part_names.contains(&identity) {
            self.part_names.remove(&identity);
        }
        self.content_type_overrides
            .remove(&part_name.to_ascii_lowercase());
    }

    pub(crate) fn retire_authored_comment_ids(
        &mut self,
        comment_ids: impl IntoIterator<Item = i32>,
    ) {
        for id in comment_ids {
            if self.authored_comment_ids.remove(&id) && !self.preserved_comment_ids.contains(&id) {
                self.comment_ids.remove(&id);
            }
        }
    }

    pub(crate) fn relationship_is_preserved(&self, owner: &str, id: &str) -> bool {
        let owner = relationship_owner_identity(owner);
        self.preserved_relationship_ids
            .get(&owner)
            .is_some_and(|ids| ids.contains(id))
    }

    fn part_is_preserved(&self, part_name: &str) -> bool {
        self.preserved_part_names
            .contains(&part_name_identity(part_name))
    }

    pub(crate) fn reserve_fragment_part_name(&mut self, preferred: &str) -> Result<String> {
        if self.part_names.insert(part_name_identity(preferred)) {
            return Ok(preferred.to_owned());
        }
        let (stem, extension) = preferred
            .rsplit_once('.')
            .map_or((preferred, ""), |(stem, extension)| (stem, extension));
        let mut ordinal = 1u64;
        loop {
            let candidate = if extension.is_empty() {
                format!("{stem}-merge-{ordinal}")
            } else {
                format!("{stem}-merge-{ordinal}.{extension}")
            };
            if self.part_names.insert(part_name_identity(&candidate)) {
                return Ok(candidate);
            }
            ordinal = ordinal.checked_add(1).ok_or_else(|| {
                Error::Other("rich mail merge part-name range is exhausted".to_owned())
            })?;
        }
    }

    pub(crate) fn observe_package_graph(&mut self, package: &OpcPackage) -> Result<()> {
        for part_name in package
            .parts
            .keys()
            .chain(package.part_rels.keys())
            .chain(package.content_types.overrides.keys())
        {
            let identity = part_name_identity(part_name);
            if self.part_names.insert(identity.clone()) {
                self.preserved_part_names.insert(identity);
            }
        }
        self.content_type_defaults.extend(
            package
                .content_types
                .defaults
                .keys()
                .map(|extension| extension.to_ascii_lowercase()),
        );
        self.content_type_overrides.extend(
            package
                .content_types
                .overrides
                .keys()
                .map(|part_name| part_name.to_ascii_lowercase()),
        );

        let mut owners = std::iter::once(("/", &package.package_rels))
            .chain(
                package
                    .part_rels
                    .iter()
                    .map(|(owner, relationships)| (owner.as_str(), relationships)),
            )
            .collect::<Vec<_>>();
        owners.sort_by_key(|(owner, _)| *owner);
        for (owner, relationships) in owners {
            let owner_identity = relationship_owner_identity(owner);
            let mut seen = HashSet::new();
            for relationship in &relationships.items {
                if !seen.insert(relationship.id.clone()) {
                    return Err(Error::Other(format!(
                        "duplicate relationship id {} in owner {owner}",
                        relationship.id
                    )));
                }
                let occupied = self
                    .relationship_ids
                    .entry(owner_identity.clone())
                    .or_default();
                occupied.insert(relationship.id.clone());
                if relationship_is_internal(relationship) {
                    let part_name = part_name_identity(&OpcPackage::resolve_rel_target(
                        owner,
                        &relationship.target,
                    ));
                    if self.part_names.insert(part_name.clone()) {
                        self.preserved_part_names.insert(part_name);
                    }
                }
            }
        }
        Ok(())
    }

    fn reconcile_provenance(&mut self, source: &Self) {
        self.preserved_relationship_ids = intersect_relationship_registry(
            &source.preserved_relationship_ids,
            &self.relationship_ids,
        );
        self.authored_story_relationship_ids = intersect_relationship_registry(
            &source.authored_story_relationship_ids,
            &self.relationship_ids,
        );
        self.authored_bundle_relationship_ids = intersect_relationship_registry(
            &source.authored_bundle_relationship_ids,
            &self.relationship_ids,
        );
        self.provisional_relationship_ids = intersect_relationship_registry(
            &source.provisional_relationship_ids,
            &self.relationship_ids,
        );
        self.preserved_bookmark_ids = source
            .preserved_bookmark_ids
            .intersection(&self.bookmark_ids)
            .copied()
            .collect();
        self.preserved_comment_ids = source
            .preserved_comment_ids
            .intersection(&self.comment_ids)
            .copied()
            .collect();
        self.preserved_drawing_ids = source
            .preserved_drawing_ids
            .intersection(&self.drawing_ids)
            .copied()
            .collect();
        self.preserved_abstract_numbering_ids = source
            .preserved_abstract_numbering_ids
            .intersection(&self.abstract_numbering_ids)
            .copied()
            .collect();
        self.preserved_numbering_instance_ids = source
            .preserved_numbering_instance_ids
            .intersection(&self.numbering_instance_ids)
            .copied()
            .collect();
        self.preserved_part_names = source
            .preserved_part_names
            .intersection(&self.part_names)
            .cloned()
            .collect();
        self.authored_bookmark_ids = source
            .authored_bookmark_ids
            .intersection(&self.bookmark_ids)
            .copied()
            .collect();
        self.authored_comment_ids = source
            .authored_comment_ids
            .intersection(&self.comment_ids)
            .copied()
            .collect();
        self.authored_toc_bookmark_ids = source
            .authored_toc_bookmark_ids
            .intersection(&self.bookmark_ids)
            .copied()
            .collect();
        self.authored_bookmark_names = source.authored_bookmark_names.clone();
        self.authored_story_parts = source
            .authored_story_parts
            .iter()
            .filter(|part_name| self.part_names.contains(&part_name_identity(part_name)))
            .cloned()
            .collect();
        self.authored_content_type_defaults = source
            .authored_content_type_defaults
            .intersection(&self.content_type_defaults)
            .cloned()
            .collect();
    }
}

fn intersect_relationship_registry(
    source: &HashMap<String, HashSet<String>>,
    occupied: &HashMap<String, HashSet<String>>,
) -> HashMap<String, HashSet<String>> {
    source
        .iter()
        .filter_map(|(owner, ids)| {
            let owner = relationship_owner_identity(owner);
            let retained = ids
                .intersection(occupied.get(&owner)?)
                .cloned()
                .collect::<HashSet<_>>();
            (!retained.is_empty()).then_some((owner, retained))
        })
        .collect()
}

fn identifier_xml_is_well_formed(xml: &[u8]) -> bool {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_resolved_event_into(&mut buffer) {
            Ok((_, Event::Eof)) => return true,
            Ok(_) => buffer.clear(),
            Err(_) => return false,
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum IdentifierXmlCategory {
    Drawing,
    Bookmark,
    Comment,
    AbstractNumbering,
    NumberingInstance,
}

fn reserve_lowest_i32(occupied: &mut HashSet<i32>, minimum: i32, category: &str) -> Result<i32> {
    let candidate = (minimum..=i32::MAX)
        .find(|value| !occupied.contains(value))
        .ok_or_else(|| Error::Other(format!("{category} id range is exhausted")))?;
    if !occupied.insert(candidate) {
        return Err(Error::Other(format!(
            "pending {category} id collision for {candidate}"
        )));
    }
    Ok(candidate)
}

fn reserve_i32(occupied: &mut HashSet<i32>, minimum: i32, category: &str) -> Result<i32> {
    let candidate = occupied
        .iter()
        .copied()
        .max()
        .and_then(|value| value.checked_add(1))
        .filter(|value| *value >= minimum)
        .or_else(|| (minimum..=i32::MAX).find(|value| !occupied.contains(value)))
        .ok_or_else(|| Error::Other(format!("{category} id range is exhausted")))?;
    if !occupied.insert(candidate) {
        return Err(Error::Other(format!(
            "pending {category} id collision for {candidate}"
        )));
    }
    Ok(candidate)
}

fn reserve_u32(occupied: &mut HashSet<u32>, minimum: u32, category: &str) -> Result<u32> {
    let candidate = occupied
        .iter()
        .copied()
        .max()
        .and_then(|value| value.checked_add(1))
        .filter(|value| *value >= minimum)
        .or_else(|| occupied.is_empty().then_some(minimum))
        .ok_or_else(|| Error::Other(format!("{category} id range is exhausted")))?;
    if !occupied.insert(candidate) {
        return Err(Error::Other(format!(
            "pending {category} id collision for {candidate}"
        )));
    }
    Ok(candidate)
}

fn ensure_numeric_relationship_capacity(maximum: Option<u32>, required: usize) -> Result<()> {
    let available = maximum.map_or(u64::from(u32::MAX) + 1, |value| u64::from(u32::MAX - value));
    if u64::try_from(required).map_or(true, |required| required > available) {
        Err(Error::Other(
            "relationship id range is exhausted".to_owned(),
        ))
    } else {
        Ok(())
    }
}

fn reserve_relationship_from_cursor(
    occupied: &mut HashSet<String>,
    cursor: &mut u32,
) -> Result<String> {
    loop {
        let id = format!("rId{cursor}");
        if occupied.insert(id.clone()) {
            if let Some(next) = cursor.checked_add(1) {
                *cursor = next;
            }
            return Ok(id);
        }
        *cursor = cursor
            .checked_add(1)
            .ok_or_else(|| Error::Other("relationship id range is exhausted".to_owned()))?;
    }
}

fn reserve_part_from_set(
    occupied: &mut HashSet<String>,
    directory: &str,
    stem: &str,
    extension: &str,
) -> Result<String> {
    let mut namer =
        oxml_media::MediaNamer::scan(directory, stem, occupied.iter().map(String::as_str));
    loop {
        let part = namer.next_part_name(extension);
        if occupied.insert(part_name_identity(&part)) {
            return Ok(part);
        }
    }
}

fn part_name_identity(part_name: &str) -> String {
    part_name.to_ascii_lowercase()
}

fn relationship_owner_identity(owner: &str) -> String {
    if owner == "/" {
        "/".to_owned()
    } else {
        part_name_identity(owner)
    }
}

fn hdr_ftr_type_order(value: HdrFtrType) -> u8 {
    match value {
        HdrFtrType::Default => 0,
        HdrFtrType::First => 1,
        HdrFtrType::Even => 2,
    }
}

fn xml_relationship_ids_in_order(xml: &[u8]) -> Result<Vec<String>> {
    xml_relationship_ids_in_order_with_bindings(xml, &[])
}

fn xml_relationship_ids_in_order_with_bindings(
    xml: &[u8],
    inherited_bindings: &[(String, String)],
) -> Result<Vec<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut ids = Vec::new();
    loop {
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid story XML: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute
                        .map_err(|error| Error::Other(format!("invalid story XML: {error}")))?;
                    let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
                    let inherited_relationship_prefix =
                        matches!(namespace, ResolveResult::Unknown(_))
                            && attribute.key.as_ref().contains(&b':')
                            && attribute
                                .key
                                .as_ref()
                                .split(|byte| *byte == b':')
                                .next()
                                .and_then(|prefix| std::str::from_utf8(prefix).ok())
                                .is_some_and(|prefix| {
                                    let declaration = format!("xmlns:{prefix}");
                                    inherited_bindings.iter().any(|(name, value)| {
                                        name == &declaration && value == drawing_ns::R
                                    })
                                });
                    if matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == drawing_ns::R.as_bytes())
                        || inherited_relationship_prefix
                    {
                        let id = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|error| {
                                Error::Other(format!("invalid story relationship id: {error}"))
                            })?
                            .into_owned();
                        if !ids.contains(&id) {
                            ids.push(id);
                        }
                    }
                }
            }
            Event::Eof => return Ok(ids),
            _ => {}
        }
        buffer.clear();
    }
}

fn rewrite_authored_doc_pr_ids(xml: &[u8], occupied: &mut HashSet<u32>) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut edits = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid story XML: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let namespace = reader.resolver().resolve_element(element.name()).0;
                let is_doc_pr = matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == drawing_ns::WP.as_bytes())
                    && element.local_name().as_ref() == b"docPr";
                let mut replaced = false;
                for attribute in element.attributes() {
                    let attribute = attribute
                        .map_err(|error| Error::Other(format!("invalid story XML: {error}")))?;
                    let (attribute_namespace, local) =
                        reader.resolver().resolve_attribute(attribute.key);
                    if is_doc_pr
                        && matches!(attribute_namespace, ResolveResult::Unbound)
                        && local.as_ref() == b"id"
                    {
                        let id = reserve_u32(occupied, 1, "drawing")?;
                        let Some((start, end)) =
                            story_attribute_value_span(&xml[before..after], attribute.key.as_ref())
                        else {
                            return Err(Error::Other(
                                "story drawing id source was not found".to_owned(),
                            ));
                        };
                        edits.push((before + start, before + end, id.to_string().into_bytes()));
                        replaced = true;
                    }
                }
                if is_doc_pr && !replaced {
                    return Err(Error::Other(
                        "wp:docPr requires an unqualified id attribute".to_owned(),
                    ));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let mut updated = xml.to_vec();
    edits.sort_by_key(|(start, _, _)| *start);
    for (start, end, replacement) in edits.into_iter().rev() {
        updated.splice(start..end, replacement);
    }
    Ok(updated)
}

fn remap_xml_relationship_ids(xml: &[u8], remap: &HashMap<String, String>) -> Result<Vec<u8>> {
    if remap.is_empty() {
        return Ok(xml.to_vec());
    }
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut edits = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid story XML: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute
                        .map_err(|error| Error::Other(format!("invalid story XML: {error}")))?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::Other(format!("invalid story relationship id: {error}"))
                        })?;
                    if matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == drawing_ns::R.as_bytes())
                        && matches!(local.as_ref(), b"id" | b"embed" | b"link")
                        && let Some(updated) = remap.get(value.as_ref())
                    {
                        let Some((start, end)) =
                            story_attribute_value_span(&xml[before..after], attribute.key.as_ref())
                        else {
                            return Err(Error::Other(
                                "story relationship attribute source was not found".to_owned(),
                            ));
                        };
                        edits.push((before + start, before + end, updated.as_bytes().to_vec()));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let mut updated = xml.to_vec();
    edits.sort_by_key(|(start, _, _)| *start);
    for (start, end, replacement) in edits.into_iter().rev() {
        updated.splice(start..end, replacement);
    }
    Ok(updated)
}

fn story_attribute_value_span(element: &[u8], attribute_name: &[u8]) -> Option<(usize, usize)> {
    let mut index = 1usize;
    while index < element.len() && !element[index].is_ascii_whitespace() {
        index += 1;
    }
    while index < element.len() {
        while index < element.len() && element[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= element.len() || matches!(element[index], b'>' | b'/') {
            return None;
        }
        let name_start = index;
        while index < element.len()
            && !element[index].is_ascii_whitespace()
            && element[index] != b'='
        {
            index += 1;
        }
        let name_end = index;
        while index < element.len() && element[index].is_ascii_whitespace() {
            index += 1;
        }
        if element.get(index) != Some(&b'=') {
            return None;
        }
        index += 1;
        while index < element.len() && element[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *element.get(index)?;
        if !matches!(quote, b'\'' | b'"') {
            return None;
        }
        index += 1;
        let value_start = index;
        while index < element.len() && element[index] != quote {
            index += 1;
        }
        let value_end = index;
        if &element[name_start..name_end] == attribute_name {
            return Some((value_start, value_end));
        }
        index += 1;
    }
    None
}

enum DrawingMut<'a> {
    Inline(&'a mut CT_Inline),
    Anchor(&'a mut CT_Anchor),
}

impl DrawingMut<'_> {
    fn set_doc_pr_id(&mut self, id: u32) {
        match self {
            Self::Inline(drawing) => drawing.doc_pr_id = id,
            Self::Anchor(drawing) => drawing.doc_pr_id = id,
        }
    }

    fn remap_relationship(&mut self, remap: &HashMap<String, String>) {
        let (embed_id, chart_rel_id) = match self {
            Self::Inline(drawing) => (&mut drawing.embed_id, &mut drawing.chart_rel_id),
            Self::Anchor(drawing) => (&mut drawing.embed_id, &mut drawing.chart_rel_id),
        };
        if let Some(updated) = remap.get(embed_id) {
            embed_id.clone_from(updated);
        }
        if let Some(id) = chart_rel_id
            && let Some(updated) = remap.get(id)
        {
            id.clone_from(updated);
        }
    }
}

/// A Word document (.docx file).
///
/// This is the main entry point for reading, creating, and modifying
/// DOCX documents.
pub struct Document {
    pub(crate) package: OpcPackage,
    pub(crate) document: CT_Document,
    /// Namespace declarations on the original document root.
    root_namespace_declarations: Vec<(String, String)>,
    /// Namespace declarations local to the original document body.
    body_namespace_declarations: Vec<(String, String)>,
    /// Namespace declarations in scope on the original document body.
    body_namespace_bindings: Vec<(String, String)>,
    pub(crate) styles: CT_Styles,
    pub(crate) numbering: Option<CT_Numbering>,
    pub(crate) core_properties: Option<CoreProperties>,
    pub(crate) application_properties: Option<AppProperties>,
    pub(crate) custom_properties: Option<CustomProperties>,
    /// Package part containing the core properties, resolved from `_rels/.rels`.
    core_properties_part_name: Option<String>,
    application_properties_part_name: Option<String>,
    custom_properties_part_name: Option<String>,
    custom_properties_owned: bool,
    /// Part name for the main document
    pub(crate) doc_part_name: String,
    /// Part name the styles were loaded from, and where they are written back.
    /// Resolved through the relationship rather than assumed, so a document
    /// that keeps its styles somewhere other than `/word/styles.xml` is
    /// updated in place instead of gaining an orphaned second part.
    styles_part_name: Option<String>,
    /// Part name for numbering definitions, resolved the same way.
    numbering_part_name: Option<String>,
    /// Typed document settings loaded through the main document relationship.
    pub(crate) settings: Option<CT_Settings>,
    /// Existing settings relationship target. No conventional target is assumed.
    settings_part_name: Option<String>,
    /// Whether this facade instance created an optional settings graph.
    settings_owned: bool,
    /// Typed shared DrawingML theme loaded through the main-document relationship.
    theme: Option<oxml_drawing::theme::CT_OfficeStyleSheet>,
    /// Existing theme relationship target.
    theme_part_name: Option<String>,
    /// Whether the theme model must be serialized into the package.
    theme_dirty: bool,
    /// Typed font table loaded through the main-document relationship.
    font_table: Option<FontTable>,
    /// Existing font-table relationship target.
    font_table_part_name: Option<String>,
    /// Whether the font-table model must be serialized into the package.
    font_table_dirty: bool,
    /// Shared package and WordprocessingML identifier allocation state.
    pub(crate) identifiers: DocumentIdentifiers,
    /// Typed footnotes loaded through the main document relationship.
    pub(crate) footnotes: rdocx_oxml::footnotes::CT_Footnotes,
    /// Existing footnotes relationship target. No conventional target is assumed on read.
    pub(crate) footnotes_part_name: Option<String>,
    /// Whether a facade mutation requires complete typed footnote serialization.
    pub(crate) footnotes_dirty: bool,
    /// Typed comments loaded through the main document relationship.
    pub(crate) comments: Option<rdocx_oxml::comments::CT_Comments>,
    /// Existing comments relationship target. No target is invented on read.
    pub(crate) comments_part_name: Option<String>,
    /// Typed reply linkage and resolved state for comments.
    pub(crate) comments_extended: Option<rdocx_oxml::comments_extended::CT_CommentsEx>,
    /// Existing comments-extended relationship target.
    pub(crate) comments_extended_part_name: Option<String>,
    /// Whether this facade created the comments part and may remove it when empty.
    pub(crate) comments_owned: bool,
    /// Whether this facade created the comments-extended part and may remove it when empty.
    pub(crate) comments_extended_owned: bool,
    /// Embedded identities whose retained signature evidence is known invalid.
    pub(crate) embedded_invalidated_signatures: HashSet<(String, String)>,
    /// Whether retained package signature evidence is known invalid.
    pub(crate) package_signatures_invalidated: bool,
    /// Relationship-resolved glossary and building-block entries.
    pub(crate) glossary: Option<rdocx_oxml::glossary::CT_GlossaryDocument>,
    /// Existing glossary relationship target.
    pub(crate) glossary_part_name: Option<String>,
    /// Whether a bounded facade replacement changed the glossary model.
    pub(crate) glossary_dirty: bool,
    /// Normal layout, including system font discovery, computed on first use.
    layout_cache: Mutex<Option<Arc<rdocx_layout::WordLayoutResult>>>,
    /// Reusable normal-font engine retained across document edits.
    normal_layout_engine: Mutex<Option<rdocx_layout::engine::Engine>>,
    /// Bundled-font-only layout used by deterministic rendering.
    deterministic_layout_cache: Mutex<Option<Arc<rdocx_layout::WordLayoutResult>>>,
    /// Reusable bundled-font engine with caller faces at highest priority.
    bundled_fallback_layout_engine: Mutex<Option<rdocx_layout::engine::Engine>>,
}

enum ChartPackageSource<'a> {
    #[allow(dead_code)]
    Typed {
        chart: &'a CT_ChartSpace,
        workbook: &'a Workbook,
    },
    Authored {
        kind: ChartKind,
        data: &'a ChartData,
    },
}

/// Fallback part names used when a document does not already declare one.
const DEFAULT_STYLES_PART: &str = "/word/styles.xml";
const DEFAULT_NUMBERING_PART: &str = "/word/numbering.xml";
const DEFAULT_CORE_PROPERTIES_PART: &str = "/docProps/core.xml";
const DEFAULT_APP_PROPERTIES_PART: &str = "/docProps/app.xml";
const DEFAULT_FONT_TABLE_PART: &str = "/word/fontTable.xml";
const DEFAULT_THEME_PART: &str = "/word/theme/theme1.xml";
const STYLES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const NUMBERING_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const CORE_PROPERTIES_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const SETTINGS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
const DEFAULT_SETTINGS_PART: &str = "/word/settings.xml";
const FONT_TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
const EMBEDDED_FONT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.obfuscatedFont";
const DEFAULT_FONT_TABLE_XML: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
    r#"<w:font w:name="Calibri"/><w:font w:name="Times New Roman"/>"#,
    r#"</w:fonts>"#,
);

pub(crate) fn relationship_is_internal(
    relationship: &oxml_opc::relationship::Relationship,
) -> bool {
    matches!(relationship.target_mode.as_deref(), None | Some("Internal"))
}

fn bundle_relationship_order(rel_type: &str) -> Option<u8> {
    match rel_type {
        rel_types::STYLES => Some(0),
        rel_types::SETTINGS => Some(1),
        rel_types::NUMBERING => Some(2),
        rel_types::FOOTNOTES => Some(3),
        rel_types::COMMENTS => Some(4),
        crate::comments::COMMENTS_EXTENDED_REL_TYPE => Some(5),
        _ => None,
    }
}

fn bundle_relationship_precedes_story(rel_type: &str) -> bool {
    matches!(rel_type, rel_types::STYLES | rel_types::SETTINGS)
}

#[cfg(test)]
thread_local! {
    static LAYOUT_INVOCATIONS: Cell<usize> = const { Cell::new(0) };
    static FAIL_NEXT_HEADER_FOOTER_SERIALIZATION: Cell<bool> = const { Cell::new(false) };
}

#[cfg(test)]
fn record_layout_invocation() {
    LAYOUT_INVOCATIONS.set(LAYOUT_INVOCATIONS.get() + 1);
}

fn owned_font_aliases(font_aliases: &[(&str, &str)]) -> Vec<(String, String)> {
    font_aliases
        .iter()
        .map(|(requested, target)| ((*requested).to_owned(), (*target).to_owned()))
        .collect()
}

fn new_word_package(class: WordPackageClass) -> OpcPackage {
    let mut package = OpcPackage::with_main_part("word/document.xml", class.content_type());
    package
        .content_types
        .add_override(DEFAULT_STYLES_PART, STYLES_CONTENT_TYPE);
    package
}

fn new_word_compatible_package(
    class: WordPackageClass,
    document: &CT_Document,
    styles: &CT_Styles,
    settings: &CT_Settings,
    core_properties: &CoreProperties,
    application_properties: &AppProperties,
) -> OpcPackage {
    let mut package = new_word_package(class);
    package.set_part(
        "/word/document.xml",
        document
            .to_xml()
            .expect("fresh Word document model must serialize"),
    );
    package.set_part(
        DEFAULT_STYLES_PART,
        styles
            .to_xml()
            .expect("fresh Word styles model must serialize"),
    );
    package.set_part(
        DEFAULT_SETTINGS_PART,
        settings
            .to_xml()
            .expect("fresh Word settings model must serialize"),
    );
    package.set_part(
        DEFAULT_CORE_PROPERTIES_PART,
        core_properties
            .to_xml()
            .expect("fresh core properties model must serialize"),
    );
    package.set_part(
        DEFAULT_APP_PROPERTIES_PART,
        application_properties
            .to_xml()
            .expect("fresh application properties model must serialize"),
    );
    package.set_part(
        DEFAULT_THEME_PART,
        oxml_drawing::theme::OFFICE_DEFAULT_XML.as_bytes().to_vec(),
    );
    package.set_part(
        DEFAULT_FONT_TABLE_PART,
        DEFAULT_FONT_TABLE_XML.as_bytes().to_vec(),
    );

    package
        .content_types
        .add_override(DEFAULT_SETTINGS_PART, SETTINGS_CONTENT_TYPE);
    package
        .content_types
        .add_override(DEFAULT_FONT_TABLE_PART, FONT_TABLE_CONTENT_TYPE);
    package
        .content_types
        .add_override(DEFAULT_THEME_PART, content_types::THEME);
    package
        .content_types
        .add_override(DEFAULT_CORE_PROPERTIES_PART, content_types::CORE_PROPERTIES);
    package.content_types.add_override(
        DEFAULT_APP_PROPERTIES_PART,
        content_types::EXTENDED_PROPERTIES,
    );

    package
        .package_rels
        .add(rel_types::CORE_PROPERTIES, "docProps/core.xml");
    package
        .package_rels
        .add(rel_types::EXTENDED_PROPERTIES, "docProps/app.xml");
    let document_relationships = package.get_or_create_part_rels("/word/document.xml");
    document_relationships.add(rel_types::STYLES, "styles.xml");
    document_relationships.add_with_id("rdocxSettings", rel_types::SETTINGS, "settings.xml");
    document_relationships.add_with_id("rdocxTheme", rel_types::THEME, "theme/theme1.xml");
    document_relationships.add_with_id("rdocxFontTable", rel_types::FONT_TABLE, "fontTable.xml");
    package
}

fn default_application_properties() -> AppProperties {
    let mut properties = AppProperties::default();
    properties.application = Some("rdocx".to_owned());
    properties.application_version = Some(env!("CARGO_PKG_VERSION").to_owned());
    properties
}

fn validate_fresh_word_compatible_package(
    package: &OpcPackage,
    expected_class: WordPackageClass,
) -> Result<()> {
    let (_, actual_class) = validated_word_package_class(package)?;
    if actual_class != expected_class {
        return Err(Error::Other(
            "fresh Word package class does not match its profile".to_owned(),
        ));
    }
    let expected_parts = [
        ("/word/document.xml", expected_class.content_type()),
        (DEFAULT_STYLES_PART, STYLES_CONTENT_TYPE),
        (DEFAULT_SETTINGS_PART, SETTINGS_CONTENT_TYPE),
        (DEFAULT_THEME_PART, content_types::THEME),
        (DEFAULT_FONT_TABLE_PART, FONT_TABLE_CONTENT_TYPE),
        (DEFAULT_CORE_PROPERTIES_PART, content_types::CORE_PROPERTIES),
        (
            DEFAULT_APP_PROPERTIES_PART,
            content_types::EXTENDED_PROPERTIES,
        ),
    ];
    for (part_name, expected_content_type) in expected_parts {
        if !package.contains_part(part_name)
            || package
                .content_types
                .override_for(part_name)
                .is_none_or(|actual| actual != expected_content_type)
        {
            return Err(Error::Other(format!(
                "fresh Word package is missing required part {part_name}"
            )));
        }
    }
    if package.parts.len() != expected_parts.len()
        || package.content_types.overrides.len() != expected_parts.len()
    {
        return Err(Error::Other(
            "fresh Word package contains an unexpected owned part".to_owned(),
        ));
    }
    let expected_package_relationships = [
        (rel_types::DOCUMENT, "/word/document.xml"),
        (rel_types::CORE_PROPERTIES, DEFAULT_CORE_PROPERTIES_PART),
        (rel_types::EXTENDED_PROPERTIES, DEFAULT_APP_PROPERTIES_PART),
    ];
    let expected_document_relationships = [
        (rel_types::STYLES, DEFAULT_STYLES_PART),
        (rel_types::SETTINGS, DEFAULT_SETTINGS_PART),
        (rel_types::THEME, DEFAULT_THEME_PART),
        (rel_types::FONT_TABLE, DEFAULT_FONT_TABLE_PART),
    ];
    let document_relationships = package
        .get_part_rels("/word/document.xml")
        .ok_or_else(|| Error::Other("fresh Word package has no main relationships".to_owned()))?;
    if package.package_rels.items.len() != expected_package_relationships.len()
        || document_relationships.items.len() != expected_document_relationships.len()
        || package.part_rels.len() != 1
    {
        return Err(Error::Other(
            "fresh Word package relationship inventory is incomplete".to_owned(),
        ));
    }
    for (source, relationships, expected) in [
        (
            "/",
            &package.package_rels,
            expected_package_relationships.as_slice(),
        ),
        (
            "/word/document.xml",
            document_relationships,
            expected_document_relationships.as_slice(),
        ),
    ] {
        for (relationship_type, expected_target) in expected {
            let matching: Vec<_> = relationships
                .items
                .iter()
                .filter(|relationship| relationship.rel_type == *relationship_type)
                .collect();
            if matching.len() != 1
                || matching[0].target_mode.is_some()
                || OpcPackage::resolve_rel_target(source, &matching[0].target) != *expected_target
            {
                return Err(Error::Other(format!(
                    "fresh Word package relationship {relationship_type} is invalid"
                )));
            }
        }
    }
    for (source, relationships) in std::iter::once(("/", &package.package_rels)).chain(
        package
            .part_rels
            .iter()
            .map(|(source, rels)| (source.as_str(), rels)),
    ) {
        for relationship in &relationships.items {
            if !relationship_is_internal(relationship) {
                continue;
            }
            let target = OpcPackage::resolve_rel_target(source, &relationship.target);
            if !package.contains_part(&target) {
                return Err(Error::Other(format!(
                    "fresh Word relationship targets missing part {target}"
                )));
            }
        }
    }
    Ok(())
}

fn validated_word_package_class(package: &OpcPackage) -> Result<(String, WordPackageClass)> {
    let relationships: Vec<_> = package
        .package_rels
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == rel_types::DOCUMENT)
        .collect();
    let relationship = match relationships.as_slice() {
        [relationship] => *relationship,
        [] => return Err(Error::NoDocumentPart),
        _ => {
            return Err(Error::Other(
                "Word package has multiple officeDocument relationships".to_owned(),
            ));
        }
    };
    if !matches!(relationship.target_mode.as_deref(), None | Some("Internal")) {
        return Err(Error::Other(
            "Word package officeDocument relationship must be internal".to_owned(),
        ));
    }
    let target = relationship.target.as_str();
    if !crate::embedded::relationship_target_is_normalized_pack_uri(target)
        || target
            .split('/')
            .any(|segment| segment == "." || segment == "..")
    {
        return Err(Error::Other(
            "Word package officeDocument relationship has an unsafe target".to_owned(),
        ));
    }
    let doc_part_name = if target.starts_with('/') {
        target.to_owned()
    } else {
        format!("/{target}")
    };
    if !package.contains_part(&doc_part_name) {
        return Err(Error::NoDocumentPart);
    }
    let content_type = package
        .content_types
        .override_for(&doc_part_name)
        .ok_or_else(|| {
            Error::Other("Word main part requires an exact content-type override".to_owned())
        })?;
    let class = WordPackageClass::from_content_type(content_type).ok_or_else(|| {
        Error::Other(format!(
            "unsupported Word main-part content type: {content_type}"
        ))
    })?;
    Ok((doc_part_name, class))
}

fn take_paragraph<'a>(paragraph: &'a mut CT_P, index: &mut usize) -> Option<&'a mut CT_P> {
    if *index == 0 {
        Some(paragraph)
    } else {
        *index -= 1;
        None
    }
}

fn nth_paragraph_in_body<'a>(
    content: &'a mut [BodyContent],
    index: &mut usize,
) -> Option<&'a mut CT_P> {
    for child in content {
        let paragraph = match child {
            BodyContent::Paragraph(paragraph) => take_paragraph(paragraph, index),
            BodyContent::ContentControl(control) => nth_paragraph_in_control(control, index),
            BodyContent::Table(_) | BodyContent::RawXml(_) => None,
        };
        if paragraph.is_some() {
            return paragraph;
        }
    }
    None
}

fn nth_paragraph_in_control<'a>(
    control: &'a mut CT_Sdt,
    index: &mut usize,
) -> Option<&'a mut CT_P> {
    for child in &mut control.content {
        let paragraph = match child {
            SdtContent::Paragraph(paragraph) => take_paragraph(paragraph, index),
            SdtContent::Table(table) => nth_paragraph_in_table(table, index),
            SdtContent::Row(row) => nth_paragraph_in_row(row, index),
            SdtContent::Cell(cell) => nth_paragraph_in_cell(cell, index),
            SdtContent::ContentControl(control) => nth_paragraph_in_control(control, index),
            SdtContent::Run(_) | SdtContent::RawXml(_) => None,
        };
        if paragraph.is_some() {
            return paragraph;
        }
    }
    None
}

fn paragraph_count_in_control(control: &CT_Sdt) -> usize {
    control
        .content
        .iter()
        .map(|child| match child {
            SdtContent::Paragraph(_) => 1,
            SdtContent::Table(table) => paragraph_count_in_table(table),
            SdtContent::Row(row) => paragraph_count_in_row(row),
            SdtContent::Cell(cell) => paragraph_count_in_cell(cell),
            SdtContent::ContentControl(control) => paragraph_count_in_control(control),
            SdtContent::Run(_) | SdtContent::RawXml(_) => 0,
        })
        .sum()
}

fn paragraph_count_in_table(table: &CT_Tbl) -> usize {
    table
        .content_controls
        .iter()
        .map(|(_, _, control)| paragraph_count_in_control(control))
        .sum::<usize>()
        + table.rows.iter().map(paragraph_count_in_row).sum::<usize>()
}

fn paragraph_count_in_row(row: &CT_Row) -> usize {
    row.content_controls
        .iter()
        .map(|(_, _, control)| paragraph_count_in_control(control))
        .sum::<usize>()
        + row.cells.iter().map(paragraph_count_in_cell).sum::<usize>()
}

fn paragraph_count_in_cell(cell: &CT_Tc) -> usize {
    cell.content
        .iter()
        .map(|child| match child {
            CellContent::Paragraph(_) => 1,
            CellContent::ContentControl(control) => paragraph_count_in_control(control),
            CellContent::Table(_) => 0,
        })
        .sum()
}

fn nth_paragraph_in_table<'a>(table: &'a mut CT_Tbl, index: &mut usize) -> Option<&'a mut CT_P> {
    let CT_Tbl {
        rows,
        content_controls,
        ..
    } = table;
    let mut selected_control = None;
    let mut selected_row = None;
    for boundary in 0..=rows.len() {
        for (control_index, (_, _, control)) in content_controls
            .iter()
            .enumerate()
            .filter(|(_, (at, _, _))| *at == boundary)
        {
            let count = paragraph_count_in_control(control);
            if *index < count {
                selected_control = Some(control_index);
                break;
            }
            *index -= count;
        }
        if selected_control.is_some() {
            break;
        }
        if let Some(row) = rows.get(boundary) {
            let count = paragraph_count_in_row(row);
            if *index < count {
                selected_row = Some(boundary);
                break;
            }
            *index -= count;
        }
    }
    if let Some(control_index) = selected_control {
        nth_paragraph_in_control(&mut content_controls[control_index].2, index)
    } else if let Some(row_index) = selected_row {
        nth_paragraph_in_row(&mut rows[row_index], index)
    } else {
        None
    }
}

fn nth_paragraph_in_row<'a>(row: &'a mut CT_Row, index: &mut usize) -> Option<&'a mut CT_P> {
    let CT_Row {
        cells,
        content_controls,
        ..
    } = row;
    let mut selected_control = None;
    let mut selected_cell = None;
    for boundary in 0..=cells.len() {
        for (control_index, (_, _, control)) in content_controls
            .iter()
            .enumerate()
            .filter(|(_, (at, _, _))| *at == boundary)
        {
            let count = paragraph_count_in_control(control);
            if *index < count {
                selected_control = Some(control_index);
                break;
            }
            *index -= count;
        }
        if selected_control.is_some() {
            break;
        }
        if let Some(cell) = cells.get(boundary) {
            let count = paragraph_count_in_cell(cell);
            if *index < count {
                selected_cell = Some(boundary);
                break;
            }
            *index -= count;
        }
    }
    if let Some(control_index) = selected_control {
        nth_paragraph_in_control(&mut content_controls[control_index].2, index)
    } else if let Some(cell_index) = selected_cell {
        nth_paragraph_in_cell(&mut cells[cell_index], index)
    } else {
        None
    }
}

fn nth_paragraph_in_cell<'a>(cell: &'a mut CT_Tc, index: &mut usize) -> Option<&'a mut CT_P> {
    for child in &mut cell.content {
        let paragraph = match child {
            CellContent::Paragraph(paragraph) => take_paragraph(paragraph, index),
            CellContent::ContentControl(control) => nth_paragraph_in_control(control, index),
            CellContent::Table(_) => None,
        };
        if paragraph.is_some() {
            return paragraph;
        }
    }
    None
}

fn take_table<'a>(table: &'a mut CT_Tbl, index: &mut usize) -> Option<&'a mut CT_Tbl> {
    if *index == 0 {
        Some(table)
    } else {
        *index -= 1;
        None
    }
}

fn nth_table_in_body<'a>(
    content: &'a mut [BodyContent],
    index: &mut usize,
) -> Option<&'a mut CT_Tbl> {
    for child in content {
        let table = match child {
            BodyContent::Table(table) => take_table(table, index),
            BodyContent::ContentControl(control) => nth_table_in_control(control, index),
            BodyContent::Paragraph(_) | BodyContent::RawXml(_) => None,
        };
        if table.is_some() {
            return table;
        }
    }
    None
}

fn nth_table_in_control<'a>(control: &'a mut CT_Sdt, index: &mut usize) -> Option<&'a mut CT_Tbl> {
    for child in &mut control.content {
        let table = match child {
            SdtContent::Table(table) => take_table(table, index),
            SdtContent::Cell(cell) => nth_table_in_cell(cell, index),
            SdtContent::ContentControl(control) => nth_table_in_control(control, index),
            SdtContent::Paragraph(_)
            | SdtContent::Row(_)
            | SdtContent::Run(_)
            | SdtContent::RawXml(_) => None,
        };
        if table.is_some() {
            return table;
        }
    }
    None
}

fn nth_table_in_cell<'a>(cell: &'a mut CT_Tc, index: &mut usize) -> Option<&'a mut CT_Tbl> {
    for child in &mut cell.content {
        let table = match child {
            CellContent::Table(table) => take_table(table, index),
            CellContent::ContentControl(control) => nth_table_in_control(control, index),
            CellContent::Paragraph(_) => None,
        };
        if table.is_some() {
            return table;
        }
    }
    None
}

fn visit_body_paragraphs(content: &[BodyContent], visitor: &mut impl FnMut(&CT_P)) {
    for item in content {
        match item {
            BodyContent::Paragraph(paragraph) => visit_paragraph(paragraph, visitor),
            BodyContent::Table(table) => visit_table(table, visitor),
            BodyContent::ContentControl(control) => visit_sdt(control, visitor),
            BodyContent::RawXml(_) => {}
        }
    }
}

fn visit_paragraph(paragraph: &CT_P, visitor: &mut impl FnMut(&CT_P)) {
    visitor(paragraph);
    for (_, _, _, control) in &paragraph.content_controls {
        visit_sdt(control, visitor);
    }
}

fn visit_table(table: &CT_Tbl, visitor: &mut impl FnMut(&CT_P)) {
    for index in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == index)
        {
            visit_sdt(control, visitor);
        }
        if let Some(row) = table.rows.get(index) {
            visit_row(row, visitor);
        }
    }
}

fn visit_row(row: &CT_Row, visitor: &mut impl FnMut(&CT_P)) {
    for index in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == index)
        {
            visit_sdt(control, visitor);
        }
        if let Some(cell) = row.cells.get(index) {
            visit_cell(cell, visitor);
        }
    }
}

fn visit_cell(cell: &CT_Tc, visitor: &mut impl FnMut(&CT_P)) {
    for item in &cell.content {
        match item {
            CellContent::Paragraph(paragraph) => visit_paragraph(paragraph, visitor),
            CellContent::Table(table) => visit_table(table, visitor),
            CellContent::ContentControl(control) => visit_sdt(control, visitor),
        }
    }
}

fn visit_sdt(control: &CT_Sdt, visitor: &mut impl FnMut(&CT_P)) {
    for item in &control.content {
        match item {
            SdtContent::Paragraph(paragraph) => visit_paragraph(paragraph, visitor),
            SdtContent::Table(table) => visit_table(table, visitor),
            SdtContent::Row(row) => visit_row(row, visitor),
            SdtContent::Cell(cell) => visit_cell(cell, visitor),
            SdtContent::Run(_) | SdtContent::RawXml(_) => {}
            SdtContent::ContentControl(nested) => visit_sdt(nested, visitor),
        }
    }
}

fn visit_body_paragraphs_mut(content: &mut [BodyContent], visitor: &mut impl FnMut(&mut CT_P)) {
    for item in content {
        match item {
            BodyContent::Paragraph(paragraph) => visit_paragraph_mut(paragraph, visitor),
            BodyContent::Table(table) => visit_table_mut(table, visitor),
            BodyContent::ContentControl(control) => visit_sdt_mut(control, visitor),
            BodyContent::RawXml(_) => {}
        }
    }
}

fn visit_paragraph_mut(paragraph: &mut CT_P, visitor: &mut impl FnMut(&mut CT_P)) {
    visitor(paragraph);
    for (_, _, _, control) in &mut paragraph.content_controls {
        visit_sdt_mut(control, visitor);
    }
}

fn visit_table_mut(table: &mut CT_Tbl, visitor: &mut impl FnMut(&mut CT_P)) {
    for index in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter_mut()
            .filter(|(at, _, _)| *at == index)
        {
            visit_sdt_mut(control, visitor);
        }
        if let Some(row) = table.rows.get_mut(index) {
            visit_row_mut(row, visitor);
        }
    }
}

fn visit_row_mut(row: &mut CT_Row, visitor: &mut impl FnMut(&mut CT_P)) {
    for index in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter_mut()
            .filter(|(at, _, _)| *at == index)
        {
            visit_sdt_mut(control, visitor);
        }
        if let Some(cell) = row.cells.get_mut(index) {
            visit_cell_mut(cell, visitor);
        }
    }
}

fn visit_cell_mut(cell: &mut CT_Tc, visitor: &mut impl FnMut(&mut CT_P)) {
    for item in &mut cell.content {
        match item {
            CellContent::Paragraph(paragraph) => visit_paragraph_mut(paragraph, visitor),
            CellContent::Table(table) => visit_table_mut(table, visitor),
            CellContent::ContentControl(control) => visit_sdt_mut(control, visitor),
        }
    }
}

fn visit_sdt_mut(control: &mut CT_Sdt, visitor: &mut impl FnMut(&mut CT_P)) {
    for item in &mut control.content {
        match item {
            SdtContent::Paragraph(paragraph) => visit_paragraph_mut(paragraph, visitor),
            SdtContent::Table(table) => visit_table_mut(table, visitor),
            SdtContent::Row(row) => visit_row_mut(row, visitor),
            SdtContent::Cell(cell) => visit_cell_mut(cell, visitor),
            SdtContent::Run(_) | SdtContent::RawXml(_) => {}
            SdtContent::ContentControl(nested) => visit_sdt_mut(nested, visitor),
        }
    }
}

fn update_reciprocal_style_link(
    styles: &mut CT_Styles,
    style_id: &str,
    old_link: Option<&str>,
    new_link: Option<&str>,
) {
    if old_link != new_link
        && let Some(old_target) = old_link
            .and_then(|target_id| styles.styles.iter_mut().find(|s| s.style_id == target_id))
        && old_target.linked_style.as_deref() == Some(style_id)
    {
        old_target.linked_style = None;
    }
    if let Some(new_target) =
        new_link.and_then(|target_id| styles.styles.iter_mut().find(|s| s.style_id == target_id))
        && new_target
            .linked_style
            .as_deref()
            .is_none_or(|linked| linked == style_id)
    {
        new_target.linked_style = Some(style_id.to_owned());
    }
}

fn merge_style_update(
    existing: &rdocx_oxml::styles::CT_Style,
    authored: &mut rdocx_oxml::styles::CT_Style,
    cleared: u16,
) {
    authored.is_default = existing.is_default;
    if authored.based_on.is_none() && cleared & style::CLEAR_BASED_ON == 0 {
        authored.based_on = existing.based_on.clone();
    }
    if authored.next_style.is_none() && cleared & style::CLEAR_NEXT_STYLE == 0 {
        authored.next_style = existing.next_style.clone();
    }
    if authored.linked_style.is_none() && cleared & style::CLEAR_LINKED_STYLE == 0 {
        authored.linked_style = existing.linked_style.clone();
    }
    if authored.auto_redefine.is_none() && cleared & style::CLEAR_AUTO_REDEFINE == 0 {
        authored.auto_redefine = existing.auto_redefine;
    }
    if authored.hidden.is_none() && cleared & style::CLEAR_HIDDEN == 0 {
        authored.hidden = existing.hidden;
    }
    if authored.ui_priority.is_none() && cleared & style::CLEAR_PRIORITY == 0 {
        authored.ui_priority = existing.ui_priority;
    }
    if authored.semi_hidden.is_none() && cleared & style::CLEAR_SEMI_HIDDEN == 0 {
        authored.semi_hidden = existing.semi_hidden;
    }
    if authored.unhide_when_used.is_none() && cleared & style::CLEAR_UNHIDE_WHEN_USED == 0 {
        authored.unhide_when_used = existing.unhide_when_used;
    }
    if authored.quick_format.is_none() && cleared & style::CLEAR_QUICK_FORMAT == 0 {
        authored.quick_format = existing.quick_format;
    }
    if authored.locked.is_none() && cleared & style::CLEAR_LOCKED == 0 {
        authored.locked = existing.locked;
    }
    if let Some(properties) = &authored.ppr {
        let mut updated = existing.ppr.clone().unwrap_or_default();
        updated.merge_from(properties);
        authored.ppr = Some(updated);
    } else if cleared & style::CLEAR_PARAGRAPH_PROPERTIES == 0 {
        authored.ppr = existing.ppr.clone();
    }
    if let Some(properties) = &authored.rpr {
        let mut updated = existing.rpr.clone().unwrap_or_default();
        updated.merge_from(properties);
        authored.rpr = Some(updated);
    } else if cleared & style::CLEAR_RUN_PROPERTIES == 0 {
        authored.rpr = existing.rpr.clone();
    }
    authored.extra_xml = existing.extra_xml.clone();
    authored.extra_attributes = existing.extra_attributes.clone();
    authored.modeled_xml = existing.modeled_xml.clone();
    if let Some(properties) = &authored.table_properties {
        authored.table_properties = Some(merge_table_style_properties(
            existing.table_properties.as_ref(),
            properties,
        ));
    } else if cleared & style::CLEAR_TABLE_PROPERTIES == 0 {
        authored.table_properties = existing.table_properties.clone();
    }
    authored.table_properties_original = existing.table_properties_original.clone();
    authored.table_properties_xml = existing.table_properties_xml.clone();
    let authored_regions = std::mem::take(&mut authored.conditional_table_styles);
    let mut merged_regions = if cleared & style::CLEAR_CONDITIONAL_TABLE_STYLES == 0 {
        existing.conditional_table_styles.clone()
    } else {
        Vec::new()
    };
    let mut authored_region_names = std::collections::HashSet::new();
    for authored_region in authored_regions {
        if !authored_region_names.insert(authored_region.region.clone()) {
            merged_regions.push(authored_region);
            continue;
        }
        if let Some(index) = merged_regions
            .iter()
            .position(|current| current.region == authored_region.region)
        {
            merged_regions[index] =
                merge_conditional_table_style(&merged_regions[index], &authored_region);
        } else {
            merged_regions.push(authored_region);
        }
    }
    authored.conditional_table_styles = merged_regions;
}

fn merge_conditional_table_style(
    existing: &rdocx_oxml::styles::CT_TblStylePr,
    authored: &rdocx_oxml::styles::CT_TblStylePr,
) -> rdocx_oxml::styles::CT_TblStylePr {
    let mut updated = existing.clone();
    if let Some(properties) = &authored.paragraph_properties {
        let target = updated.paragraph_properties.get_or_insert_default();
        target.merge_from(properties);
    }
    if let Some(properties) = &authored.table_properties {
        updated.table_properties = Some(merge_table_style_properties(
            updated.table_properties.as_ref(),
            properties,
        ));
    }
    if let Some(properties) = &authored.cell_properties {
        updated.cell_properties = Some(merge_table_cell_style_properties(
            updated.cell_properties.as_ref(),
            properties,
        ));
    }
    updated
}

fn merge_table_borders(
    existing: Option<&rdocx_oxml::table::CT_TblBorders>,
    authored: &rdocx_oxml::table::CT_TblBorders,
) -> rdocx_oxml::table::CT_TblBorders {
    let mut updated = existing.cloned().unwrap_or_default();
    if authored.top.is_some() {
        updated.top.clone_from(&authored.top);
    }
    if authored.bottom.is_some() {
        updated.bottom.clone_from(&authored.bottom);
    }
    if authored.left.is_some() {
        updated.left.clone_from(&authored.left);
    }
    if authored.right.is_some() {
        updated.right.clone_from(&authored.right);
    }
    if authored.inside_h.is_some() {
        updated.inside_h.clone_from(&authored.inside_h);
    }
    if authored.inside_v.is_some() {
        updated.inside_v.clone_from(&authored.inside_v);
    }
    for raw in &authored.extra_xml {
        if !updated.extra_xml.contains(raw) {
            updated.extra_xml.push(raw.clone());
        }
    }
    updated
}

fn merge_table_cell_style_properties(
    existing: Option<&rdocx_oxml::table::CT_TcPr>,
    authored: &rdocx_oxml::table::CT_TcPr,
) -> rdocx_oxml::table::CT_TcPr {
    let mut updated = existing.cloned().unwrap_or_default();
    if authored.width.is_some() {
        updated.width.clone_from(&authored.width);
    }
    if authored.grid_span.is_some() {
        updated.grid_span = authored.grid_span;
    }
    if authored.h_merge.is_some() {
        updated.h_merge.clone_from(&authored.h_merge);
    }
    if authored.v_merge.is_some() {
        updated.v_merge = authored.v_merge;
    }
    if let Some(borders) = &authored.borders {
        updated.borders = Some(merge_table_borders(updated.borders.as_ref(), borders));
    }
    if authored.shading.is_some() {
        updated.shading.clone_from(&authored.shading);
    }
    if authored.v_align.is_some() {
        updated.v_align = authored.v_align;
    }
    if authored.no_wrap.is_some() {
        updated.no_wrap = authored.no_wrap;
    }
    if authored.text_direction.is_some() {
        updated.text_direction.clone_from(&authored.text_direction);
    }
    if authored.cnf_style.is_some() {
        updated.cnf_style.clone_from(&authored.cnf_style);
    }
    for raw in &authored.extra_xml {
        if !updated.extra_xml.contains(raw) {
            updated.extra_xml.push(raw.clone());
        }
    }
    updated
}

fn merge_table_style_properties(
    existing: Option<&rdocx_oxml::table::CT_TblPr>,
    authored: &rdocx_oxml::table::CT_TblPr,
) -> rdocx_oxml::table::CT_TblPr {
    let mut updated = existing.cloned().unwrap_or_default();
    if authored.style_id.is_some() {
        updated.style_id.clone_from(&authored.style_id);
    }
    if authored.width.is_some() {
        updated.width.clone_from(&authored.width);
    }
    if authored.jc.is_some() {
        updated.jc = authored.jc;
    }
    if let Some(borders) = &authored.borders {
        updated.borders = Some(merge_table_borders(updated.borders.as_ref(), borders));
    }
    if let Some(margins) = &authored.cell_margin {
        let target = updated.cell_margin.get_or_insert_default();
        if margins.top.is_some() {
            target.top = margins.top;
        }
        if margins.bottom.is_some() {
            target.bottom = margins.bottom;
        }
        if margins.left.is_some() {
            target.left = margins.left;
        }
        if margins.right.is_some() {
            target.right = margins.right;
        }
    }
    if authored.layout.is_some() {
        updated.layout.clone_from(&authored.layout);
    }
    if authored.indent.is_some() {
        updated.indent.clone_from(&authored.indent);
    }
    if authored.shading.is_some() {
        updated.shading.clone_from(&authored.shading);
    }
    if let Some(look) = &authored.look {
        let target = updated.look.get_or_insert_default();
        if look.val.is_some() {
            target.val.clone_from(&look.val);
        }
        if look.first_row.is_some() {
            target.first_row = look.first_row;
        }
        if look.last_row.is_some() {
            target.last_row = look.last_row;
        }
        if look.first_column.is_some() {
            target.first_column = look.first_column;
        }
        if look.last_column.is_some() {
            target.last_column = look.last_column;
        }
        if look.no_h_band.is_some() {
            target.no_h_band = look.no_h_band;
        }
        if look.no_v_band.is_some() {
            target.no_v_band = look.no_v_band;
        }
    }
    if authored.change.is_some() {
        updated.change.clone_from(&authored.change);
    }
    for raw in &authored.revision_xml {
        if !updated.revision_xml.contains(raw) {
            updated.revision_xml.push(raw.clone());
        }
    }
    for raw in &authored.extra_xml {
        if !updated.extra_xml.contains(raw) {
            updated.extra_xml.push(raw.clone());
        }
    }
    updated
}

fn body_references_style(content: &[BodyContent], style_id: &str) -> bool {
    content.iter().any(|item| match item {
        BodyContent::Paragraph(paragraph) => paragraph_references_style(paragraph, style_id),
        BodyContent::Table(table) => table_references_style(table, style_id),
        BodyContent::ContentControl(control) => control_references_style(control, style_id),
        BodyContent::RawXml(_) => false,
    })
}

fn paragraph_references_style(paragraph: &CT_P, style_id: &str) -> bool {
    paragraph
        .properties
        .as_ref()
        .and_then(|properties| properties.style_id.as_deref())
        == Some(style_id)
        || paragraph.runs.iter().any(|run| {
            run.properties
                .as_ref()
                .and_then(|properties| properties.style_id.as_deref())
                == Some(style_id)
        })
        || paragraph
            .content_controls
            .iter()
            .any(|(_, _, _, control)| control_references_style(control, style_id))
}

fn table_references_style(table: &CT_Tbl, style_id: &str) -> bool {
    table
        .properties
        .as_ref()
        .and_then(|properties| properties.style_id.as_deref())
        == Some(style_id)
        || table.rows.iter().any(|row| {
            row.cells.iter().any(|cell| {
                cell.content.iter().any(|item| match item {
                    CellContent::Paragraph(paragraph) => {
                        paragraph_references_style(paragraph, style_id)
                    }
                    CellContent::Table(table) => table_references_style(table, style_id),
                    CellContent::ContentControl(control) => {
                        control_references_style(control, style_id)
                    }
                })
            })
        })
        || table
            .content_controls
            .iter()
            .any(|(_, _, control)| control_references_style(control, style_id))
}

fn control_references_style(control: &CT_Sdt, style_id: &str) -> bool {
    control.content.iter().any(|item| match item {
        SdtContent::Paragraph(paragraph) => paragraph_references_style(paragraph, style_id),
        SdtContent::Table(table) => table_references_style(table, style_id),
        SdtContent::Row(row) => row.cells.iter().any(|cell| {
            cell.content.iter().any(|item| match item {
                CellContent::Paragraph(paragraph) => {
                    paragraph_references_style(paragraph, style_id)
                }
                CellContent::Table(table) => table_references_style(table, style_id),
                CellContent::ContentControl(control) => control_references_style(control, style_id),
            })
        }),
        SdtContent::Cell(cell) => cell.content.iter().any(|item| match item {
            CellContent::Paragraph(paragraph) => paragraph_references_style(paragraph, style_id),
            CellContent::Table(table) => table_references_style(table, style_id),
            CellContent::ContentControl(control) => control_references_style(control, style_id),
        }),
        SdtContent::Run(run) => {
            run.properties
                .as_ref()
                .and_then(|properties| properties.style_id.as_deref())
                == Some(style_id)
        }
        SdtContent::ContentControl(nested) => control_references_style(nested, style_id),
        SdtContent::RawXml(_) => false,
    })
}

fn xml_references_style(xml: &[u8], style_id: &str) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(error.to_string()))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == WORD_NAMESPACE.as_bytes())
                    && matches!(
                        element.local_name().as_ref(),
                        b"pStyle" | b"rStyle" | b"tblStyle"
                    ) =>
            {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == WORD_NAMESPACE.as_bytes())
                        && local.as_ref() == b"val"
                        && attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|error| Error::Other(error.to_string()))?
                            .as_ref()
                            == style_id
                    {
                        return Ok(true);
                    }
                }
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_relationship_ids(content: &[BodyContent], output: &mut Vec<String>) {
    for item in content {
        match item {
            BodyContent::Paragraph(paragraph) => {
                collect_paragraph_relationship_ids(paragraph, output)
            }
            BodyContent::Table(table) => collect_table_relationship_ids(table, output),
            BodyContent::ContentControl(control) => collect_sdt_relationship_ids(control, output),
            BodyContent::RawXml(_) => {}
        }
    }
}

fn collect_paragraph_relationship_ids(paragraph: &CT_P, output: &mut Vec<String>) {
    if let Some(section) = paragraph
        .properties
        .as_ref()
        .and_then(|properties| properties.sect_pr.as_ref())
    {
        for reference in section.header_refs.iter().chain(&section.footer_refs) {
            if !output.contains(&reference.rel_id) {
                output.push(reference.rel_id.clone());
            }
        }
    }
    for index in 0..=paragraph.runs.len() {
        for (_, _, _, control) in paragraph
            .content_controls
            .iter()
            .filter(|(at, _, _, _)| *at == index)
        {
            collect_sdt_relationship_ids(control, output);
        }
        for hyperlink in paragraph
            .hyperlinks
            .iter()
            .filter(|hyperlink| hyperlink.run_start == index)
        {
            if let Some(id) = &hyperlink.rel_id
                && !output.contains(id)
            {
                output.push(id.clone());
            }
        }
        let Some(run) = paragraph.runs.get(index) else {
            continue;
        };
        for content in &run.content {
            let RunContent::Drawing(drawing) = content else {
                continue;
            };
            let id = drawing
                .inline
                .as_ref()
                .filter(|drawing| drawing.raw_xml.is_none())
                .map(|drawing| drawing.chart_rel_id.as_ref().unwrap_or(&drawing.embed_id))
                .or_else(|| {
                    drawing
                        .anchor
                        .as_ref()
                        .filter(|drawing| drawing.raw_xml.is_none())
                        .map(|drawing| drawing.chart_rel_id.as_ref().unwrap_or(&drawing.embed_id))
                });
            if let Some(id) = id
                && !output.contains(id)
            {
                output.push(id.clone());
            }
        }
    }
}

fn collect_table_relationship_ids(table: &CT_Tbl, output: &mut Vec<String>) {
    for index in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == index)
        {
            collect_sdt_relationship_ids(control, output);
        }
        if let Some(row) = table.rows.get(index) {
            for cell_index in 0..=row.cells.len() {
                for (_, _, control) in row
                    .content_controls
                    .iter()
                    .filter(|(at, _, _)| *at == cell_index)
                {
                    collect_sdt_relationship_ids(control, output);
                }
                if let Some(cell) = row.cells.get(cell_index) {
                    for item in &cell.content {
                        match item {
                            CellContent::Paragraph(paragraph) => {
                                collect_paragraph_relationship_ids(paragraph, output)
                            }
                            CellContent::Table(table) => {
                                collect_table_relationship_ids(table, output)
                            }
                            CellContent::ContentControl(control) => {
                                collect_sdt_relationship_ids(control, output)
                            }
                        }
                    }
                }
            }
        }
    }
}

fn collect_sdt_relationship_ids(control: &CT_Sdt, output: &mut Vec<String>) {
    for item in &control.content {
        match item {
            SdtContent::Paragraph(paragraph) => {
                collect_paragraph_relationship_ids(paragraph, output)
            }
            SdtContent::Table(table) => collect_table_relationship_ids(table, output),
            SdtContent::Row(row) => {
                let table = CT_Tbl {
                    properties: None,
                    grid: None,
                    rows: vec![row.clone()],
                    extra_xml: Vec::new(),
                    content_controls: Vec::new(),
                };
                collect_table_relationship_ids(&table, output);
            }
            SdtContent::Cell(cell) => {
                for item in &cell.content {
                    match item {
                        CellContent::Paragraph(paragraph) => {
                            collect_paragraph_relationship_ids(paragraph, output)
                        }
                        CellContent::Table(table) => collect_table_relationship_ids(table, output),
                        CellContent::ContentControl(control) => {
                            collect_sdt_relationship_ids(control, output)
                        }
                    }
                }
            }
            SdtContent::ContentControl(nested) => collect_sdt_relationship_ids(nested, output),
            SdtContent::Run(run) => {
                for content in &run.content {
                    if let RunContent::Drawing(drawing) = content
                        && let Some(id) = drawing
                            .inline
                            .as_ref()
                            .filter(|drawing| drawing.raw_xml.is_none())
                            .map(|drawing| {
                                drawing.chart_rel_id.as_ref().unwrap_or(&drawing.embed_id)
                            })
                            .or_else(|| {
                                drawing
                                    .anchor
                                    .as_ref()
                                    .filter(|drawing| drawing.raw_xml.is_none())
                                    .map(|drawing| {
                                        drawing.chart_rel_id.as_ref().unwrap_or(&drawing.embed_id)
                                    })
                            })
                        && !output.contains(id)
                    {
                        output.push(id.clone());
                    }
                }
            }
            SdtContent::RawXml(_) => {}
        }
    }
}

fn visit_authored_drawings(content: &[BodyContent], visitor: &mut impl FnMut(&CT_Drawing)) {
    for item in content {
        match item {
            BodyContent::Paragraph(paragraph) => visit_paragraph_drawings(paragraph, visitor),
            BodyContent::Table(table) => visit_table_drawings(table, visitor),
            BodyContent::ContentControl(control) => visit_sdt_drawings(control, visitor),
            BodyContent::RawXml(_) => {}
        }
    }
}

fn visit_authored_drawings_mut(
    content: &mut [BodyContent],
    visitor: &mut impl FnMut(&mut CT_Drawing),
) {
    for item in content {
        match item {
            BodyContent::Paragraph(paragraph) => visit_paragraph_drawings_mut(paragraph, visitor),
            BodyContent::Table(table) => visit_table_drawings_mut(table, visitor),
            BodyContent::ContentControl(control) => visit_sdt_drawings_mut(control, visitor),
            BodyContent::RawXml(_) => {}
        }
    }
}

fn authored_drawing(drawing: &CT_Drawing) -> bool {
    drawing
        .inline
        .as_ref()
        .is_some_and(|value| value.raw_xml.is_none())
        || drawing
            .anchor
            .as_ref()
            .is_some_and(|value| value.raw_xml.is_none())
}

fn visit_run_drawings(run: &CT_R, visitor: &mut impl FnMut(&CT_Drawing)) {
    for content in &run.content {
        if let RunContent::Drawing(drawing) = content
            && authored_drawing(drawing)
        {
            visitor(drawing);
        }
    }
}

fn visit_run_drawings_mut(run: &mut CT_R, visitor: &mut impl FnMut(&mut CT_Drawing)) {
    for content in &mut run.content {
        if let RunContent::Drawing(drawing) = content
            && authored_drawing(drawing)
        {
            visitor(drawing);
        }
    }
}

fn visit_paragraph_drawings(paragraph: &CT_P, visitor: &mut impl FnMut(&CT_Drawing)) {
    for index in 0..=paragraph.runs.len() {
        for (_, _, _, control) in paragraph
            .content_controls
            .iter()
            .filter(|(at, _, _, _)| *at == index)
        {
            visit_sdt_drawings(control, visitor);
        }
        if let Some(run) = paragraph.runs.get(index) {
            visit_run_drawings(run, visitor);
        }
    }
}

fn visit_paragraph_drawings_mut(paragraph: &mut CT_P, visitor: &mut impl FnMut(&mut CT_Drawing)) {
    for index in 0..=paragraph.runs.len() {
        for (_, _, _, control) in paragraph
            .content_controls
            .iter_mut()
            .filter(|(at, _, _, _)| *at == index)
        {
            visit_sdt_drawings_mut(control, visitor);
        }
        if let Some(run) = paragraph.runs.get_mut(index) {
            visit_run_drawings_mut(run, visitor);
        }
    }
}

fn visit_table_drawings(table: &CT_Tbl, visitor: &mut impl FnMut(&CT_Drawing)) {
    for index in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == index)
        {
            visit_sdt_drawings(control, visitor);
        }
        if let Some(row) = table.rows.get(index) {
            visit_row_drawings(row, visitor);
        }
    }
}

fn visit_table_drawings_mut(table: &mut CT_Tbl, visitor: &mut impl FnMut(&mut CT_Drawing)) {
    for index in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter_mut()
            .filter(|(at, _, _)| *at == index)
        {
            visit_sdt_drawings_mut(control, visitor);
        }
        if let Some(row) = table.rows.get_mut(index) {
            visit_row_drawings_mut(row, visitor);
        }
    }
}

fn visit_row_drawings(row: &CT_Row, visitor: &mut impl FnMut(&CT_Drawing)) {
    for index in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == index)
        {
            visit_sdt_drawings(control, visitor);
        }
        if let Some(cell) = row.cells.get(index) {
            visit_cell_drawings(cell, visitor);
        }
    }
}

fn visit_row_drawings_mut(row: &mut CT_Row, visitor: &mut impl FnMut(&mut CT_Drawing)) {
    for index in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter_mut()
            .filter(|(at, _, _)| *at == index)
        {
            visit_sdt_drawings_mut(control, visitor);
        }
        if let Some(cell) = row.cells.get_mut(index) {
            visit_cell_drawings_mut(cell, visitor);
        }
    }
}

fn visit_cell_drawings(cell: &CT_Tc, visitor: &mut impl FnMut(&CT_Drawing)) {
    for item in &cell.content {
        match item {
            CellContent::Paragraph(paragraph) => visit_paragraph_drawings(paragraph, visitor),
            CellContent::Table(table) => visit_table_drawings(table, visitor),
            CellContent::ContentControl(control) => visit_sdt_drawings(control, visitor),
        }
    }
}

fn visit_cell_drawings_mut(cell: &mut CT_Tc, visitor: &mut impl FnMut(&mut CT_Drawing)) {
    for item in &mut cell.content {
        match item {
            CellContent::Paragraph(paragraph) => visit_paragraph_drawings_mut(paragraph, visitor),
            CellContent::Table(table) => visit_table_drawings_mut(table, visitor),
            CellContent::ContentControl(control) => visit_sdt_drawings_mut(control, visitor),
        }
    }
}

fn visit_sdt_drawings(control: &CT_Sdt, visitor: &mut impl FnMut(&CT_Drawing)) {
    for item in &control.content {
        match item {
            SdtContent::Paragraph(paragraph) => visit_paragraph_drawings(paragraph, visitor),
            SdtContent::Table(table) => visit_table_drawings(table, visitor),
            SdtContent::Row(row) => visit_row_drawings(row, visitor),
            SdtContent::Cell(cell) => visit_cell_drawings(cell, visitor),
            SdtContent::Run(run) => visit_run_drawings(run, visitor),
            SdtContent::ContentControl(nested) => visit_sdt_drawings(nested, visitor),
            SdtContent::RawXml(_) => {}
        }
    }
}

fn visit_sdt_drawings_mut(control: &mut CT_Sdt, visitor: &mut impl FnMut(&mut CT_Drawing)) {
    for item in &mut control.content {
        match item {
            SdtContent::Paragraph(paragraph) => visit_paragraph_drawings_mut(paragraph, visitor),
            SdtContent::Table(table) => visit_table_drawings_mut(table, visitor),
            SdtContent::Row(row) => visit_row_drawings_mut(row, visitor),
            SdtContent::Cell(cell) => visit_cell_drawings_mut(cell, visitor),
            SdtContent::Run(run) => visit_run_drawings_mut(run, visitor),
            SdtContent::ContentControl(nested) => visit_sdt_drawings_mut(nested, visitor),
            SdtContent::RawXml(_) => {}
        }
    }
}

impl Document {
    /// Create a new Word-compatible DOCX document with default page setup and styles.
    pub fn new() -> Self {
        Self::new_with_profile(WordCreationProfile::WordCompatible(
            WordPackageClass::Document,
        ))
    }

    /// Create a new document with explicit package completeness and class.
    pub fn new_with_profile(profile: WordCreationProfile) -> Self {
        let document = CT_Document::new();
        let styles = CT_Styles::new_default();
        let (class, compatible) = match profile {
            WordCreationProfile::Minimal(class) => (class, false),
            WordCreationProfile::WordCompatible(class) => (class, true),
        };
        let settings = compatible.then(CT_Settings::new);
        let core_properties = compatible.then(CoreProperties::default);
        let application_properties = compatible.then(default_application_properties);
        let mut package = if compatible {
            new_word_compatible_package(
                class,
                &document,
                &styles,
                settings.as_ref().expect("compatible settings exist"),
                core_properties
                    .as_ref()
                    .expect("compatible core properties exist"),
                application_properties
                    .as_ref()
                    .expect("compatible application properties exist"),
            )
        } else {
            new_word_package(class)
        };
        if compatible {
            validate_fresh_word_compatible_package(&package, class)
                .expect("fresh Word-compatible package graph must validate");
        }

        if !compatible {
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(rel_types::STYLES, "styles.xml");
        }
        let identifiers = DocumentIdentifiers::scan(&package)
            .expect("a freshly constructed package has valid identifiers");

        Document {
            package,
            document,
            root_namespace_declarations: Vec::new(),
            body_namespace_declarations: Vec::new(),
            body_namespace_bindings: Vec::new(),
            styles,
            numbering: None,
            core_properties,
            application_properties,
            custom_properties: None,
            core_properties_part_name: compatible.then(|| DEFAULT_CORE_PROPERTIES_PART.to_owned()),
            application_properties_part_name: compatible
                .then(|| DEFAULT_APP_PROPERTIES_PART.to_owned()),
            custom_properties_part_name: None,
            custom_properties_owned: false,
            doc_part_name: "/word/document.xml".to_string(),
            styles_part_name: Some(DEFAULT_STYLES_PART.to_owned()),
            numbering_part_name: None,
            settings,
            settings_part_name: compatible.then(|| DEFAULT_SETTINGS_PART.to_owned()),
            settings_owned: false,
            theme: compatible.then(oxml_drawing::theme::CT_OfficeStyleSheet::office_default),
            theme_part_name: compatible.then(|| DEFAULT_THEME_PART.to_owned()),
            theme_dirty: false,
            font_table: compatible.then(|| {
                FontTable::from_xml(DEFAULT_FONT_TABLE_XML.as_bytes())
                    .expect("fresh Word font table must parse")
            }),
            font_table_part_name: compatible.then(|| DEFAULT_FONT_TABLE_PART.to_owned()),
            font_table_dirty: false,
            identifiers,
            footnotes: rdocx_oxml::footnotes::CT_Footnotes::new(),
            footnotes_part_name: None,
            footnotes_dirty: false,
            comments: None,
            comments_part_name: None,
            comments_extended: None,
            comments_extended_part_name: None,
            comments_owned: false,
            comments_extended_owned: false,
            embedded_invalidated_signatures: HashSet::new(),
            package_signatures_invalidated: false,
            glossary: None,
            glossary_part_name: None,
            glossary_dirty: false,
            layout_cache: Mutex::new(None),
            normal_layout_engine: Mutex::new(None),
            deterministic_layout_cache: Mutex::new(None),
            bundled_fallback_layout_engine: Mutex::new(None),
        }
    }

    /// Clone all package and typed state while discarding derived layout caches.
    pub(crate) fn clone_for_staging(&self) -> Self {
        Self {
            package: self.package.clone(),
            document: self.document.clone(),
            root_namespace_declarations: self.root_namespace_declarations.clone(),
            body_namespace_declarations: self.body_namespace_declarations.clone(),
            body_namespace_bindings: self.body_namespace_bindings.clone(),
            styles: self.styles.clone(),
            numbering: self.numbering.clone(),
            core_properties: self.core_properties.clone(),
            application_properties: self.application_properties.clone(),
            custom_properties: self.custom_properties.clone(),
            core_properties_part_name: self.core_properties_part_name.clone(),
            application_properties_part_name: self.application_properties_part_name.clone(),
            custom_properties_part_name: self.custom_properties_part_name.clone(),
            custom_properties_owned: self.custom_properties_owned,
            doc_part_name: self.doc_part_name.clone(),
            styles_part_name: self.styles_part_name.clone(),
            numbering_part_name: self.numbering_part_name.clone(),
            settings: self.settings.clone(),
            settings_part_name: self.settings_part_name.clone(),
            settings_owned: self.settings_owned,
            theme: self.theme.clone(),
            theme_part_name: self.theme_part_name.clone(),
            theme_dirty: self.theme_dirty,
            font_table: self.font_table.clone(),
            font_table_part_name: self.font_table_part_name.clone(),
            font_table_dirty: self.font_table_dirty,
            identifiers: self.identifiers.clone(),
            footnotes: self.footnotes.clone(),
            footnotes_part_name: self.footnotes_part_name.clone(),
            footnotes_dirty: self.footnotes_dirty,
            comments: self.comments.clone(),
            comments_part_name: self.comments_part_name.clone(),
            comments_extended: self.comments_extended.clone(),
            comments_extended_part_name: self.comments_extended_part_name.clone(),
            comments_owned: self.comments_owned,
            comments_extended_owned: self.comments_extended_owned,
            embedded_invalidated_signatures: self.embedded_invalidated_signatures.clone(),
            package_signatures_invalidated: self.package_signatures_invalidated,
            glossary: self.glossary.clone(),
            glossary_part_name: self.glossary_part_name.clone(),
            glossary_dirty: self.glossary_dirty,
            layout_cache: Mutex::new(None),
            normal_layout_engine: Mutex::new(None),
            deterministic_layout_cache: Mutex::new(None),
            bundled_fallback_layout_engine: Mutex::new(None),
        }
    }

    fn canonicalize_authored_identifiers(&mut self) -> Result<()> {
        self.canonicalize_bookmark_ids()?;
        self.canonicalize_comment_ids()?;
        self.canonicalize_numbering_ids()?;
        self.canonicalize_drawing_ids()?;
        Ok(())
    }

    fn canonicalize_bookmark_ids(&mut self) -> Result<()> {
        let mut semantic = Vec::new();
        visit_body_paragraphs(&self.document.body.content, &mut |paragraph| {
            for marker in &paragraph.bookmark_markers {
                if marker.is_start()
                    && let Some(id) = marker.id()
                    && self.identifiers.authored_bookmark_ids.contains(&id)
                    && !semantic.contains(&id)
                {
                    semantic.push(id);
                }
            }
        });
        let mut occupied = self.identifiers.preserved_bookmark_ids.clone();
        let mut remap = HashMap::new();
        for old in semantic {
            remap.insert(old, reserve_lowest_i32(&mut occupied, 0, "bookmark")?);
        }
        if !remap.is_empty() {
            visit_body_paragraphs_mut(&mut self.document.body.content, &mut |paragraph| {
                let _ = paragraph.remap_authored_bookmark_ids(&remap);
            });
            for (old, new) in remap {
                self.identifiers.authored_bookmark_ids.remove(&old);
                self.identifiers.authored_bookmark_ids.insert(new);
            }
        }
        self.canonicalize_toc_bookmark_ids()?;
        Ok(())
    }

    fn canonicalize_toc_bookmark_ids(&mut self) -> Result<()> {
        let mut semantic = Vec::new();
        visit_body_paragraphs(&self.document.body.content, &mut |paragraph| {
            for marker in &paragraph.bookmark_markers {
                if marker.is_start()
                    && let Some(id) = marker.id()
                    && self.identifiers.authored_toc_bookmark_ids.contains(&id)
                    && !semantic.contains(&id)
                {
                    semantic.push(id);
                }
            }
        });
        if semantic.is_empty() {
            return Ok(());
        }
        let mut occupied = self.identifiers.bookmark_ids.clone();
        for id in &self.identifiers.authored_toc_bookmark_ids {
            occupied.remove(id);
        }
        let mut next = occupied.iter().copied().max().unwrap_or(0).checked_add(1);
        let mut remap = BTreeMap::new();
        for old in semantic {
            let new = next.ok_or_else(|| {
                Error::Other("table of contents exhausted the bookmark ID range".to_owned())
            })?;
            occupied.insert(new);
            next = new.checked_add(1);
            if old != new {
                remap.insert(old.to_string(), new.to_string());
            }
        }
        if remap.is_empty() {
            return Ok(());
        }
        let xml = self.document.to_xml()?;
        let updated = crate::field::patch_bookmark_ids(&xml, &remap)?;
        self.document = CT_Document::from_xml(&updated)?;
        for (old, new) in remap {
            let old = old.parse::<i32>().expect("bookmark remap key is numeric");
            let new = new.parse::<i32>().expect("bookmark remap value is numeric");
            self.identifiers.bookmark_ids.remove(&old);
            self.identifiers.bookmark_ids.insert(new);
            self.identifiers.preserved_bookmark_ids.remove(&old);
            self.identifiers.preserved_bookmark_ids.insert(new);
            self.identifiers.authored_toc_bookmark_ids.remove(&old);
            self.identifiers.authored_toc_bookmark_ids.insert(new);
        }
        Ok(())
    }

    fn canonicalize_comment_ids(&mut self) -> Result<()> {
        let mut semantic = Vec::new();
        visit_body_paragraphs(&self.document.body.content, &mut |paragraph| {
            for marker in &paragraph.comment_ranges {
                if let rdocx_oxml::text::CommentRangeMarker::Start { id, .. } = marker
                    && self.identifiers.comment_ids.contains(id)
                    && !self.identifiers.preserved_comment_ids.contains(id)
                    && !semantic.contains(id)
                {
                    semantic.push(*id);
                }
            }
        });
        if let Some(comments) = &self.comments {
            for comment in &comments.comments {
                if !self.identifiers.preserved_comment_ids.contains(&comment.id)
                    && self.identifiers.comment_ids.contains(&comment.id)
                    && !semantic.contains(&comment.id)
                {
                    semantic.push(comment.id);
                }
            }
        }
        let mut occupied = self.identifiers.preserved_comment_ids.clone();
        let mut remap = HashMap::new();
        for old in semantic {
            remap.insert(old, reserve_i32(&mut occupied, 0, "comment")?);
        }
        if remap.is_empty() {
            return Ok(());
        }
        visit_body_paragraphs_mut(&mut self.document.body.content, &mut |paragraph| {
            for marker in &mut paragraph.comment_ranges {
                match marker {
                    rdocx_oxml::text::CommentRangeMarker::Start { id, .. }
                    | rdocx_oxml::text::CommentRangeMarker::End { id, .. } => {
                        if let Some(updated) = remap.get(id) {
                            *id = *updated;
                        }
                    }
                }
            }
            for run in &mut paragraph.runs {
                for content in &mut run.content {
                    if let RunContent::CommentReference { id, .. } = content
                        && let Some(updated) = remap.get(id)
                    {
                        *id = *updated;
                    }
                }
            }
        });
        if let Some(comments) = &mut self.comments {
            for comment in &mut comments.comments {
                if let Some(updated) = remap.get(&comment.id) {
                    comment.id = *updated;
                }
            }
            if self.identifiers.preserved_comment_ids.is_empty() {
                comments.comments.sort_by_key(|comment| comment.id);
                let mut para_remap = HashMap::new();
                let mut next_para_id = 1u32;
                for comment in &mut comments.comments {
                    for para_id in comment.paragraph_ids.iter_mut().flatten() {
                        let updated = format!("{next_para_id:08X}");
                        next_para_id = next_para_id.checked_add(1).ok_or_else(|| {
                            Error::Other("comment paragraph id range is exhausted".to_owned())
                        })?;
                        para_remap.insert(para_id.clone(), updated.clone());
                        *para_id = updated;
                    }
                }
                if let Some(extended) = &mut self.comments_extended {
                    for entry in &mut extended.comments {
                        if let Some(updated) = para_remap.get(&entry.para_id) {
                            entry.para_id.clone_from(updated);
                        }
                        if let Some(parent) = &mut entry.para_id_parent
                            && let Some(updated) = para_remap.get(parent)
                        {
                            parent.clone_from(updated);
                        }
                    }
                    extended
                        .comments
                        .sort_by(|left, right| left.para_id.cmp(&right.para_id));
                }
            }
        }
        Ok(())
    }

    fn canonicalize_numbering_ids(&mut self) -> Result<()> {
        let Some(numbering) = self.numbering.as_ref() else {
            return Ok(());
        };
        let mut semantic_nums = Vec::new();
        visit_body_paragraphs(&self.document.body.content, &mut |paragraph| {
            if let Some(id) = paragraph
                .properties
                .as_ref()
                .and_then(|properties| properties.num_id)
                && !self
                    .identifiers
                    .preserved_numbering_instance_ids
                    .contains(&id)
                && self.identifiers.numbering_instance_ids.contains(&id)
                && !semantic_nums.contains(&id)
            {
                semantic_nums.push(id);
            }
        });
        for instance in &numbering.nums {
            if !self
                .identifiers
                .preserved_numbering_instance_ids
                .contains(&instance.num_id)
                && self
                    .identifiers
                    .numbering_instance_ids
                    .contains(&instance.num_id)
                && !semantic_nums.contains(&instance.num_id)
            {
                semantic_nums.push(instance.num_id);
            }
        }
        let abstract_by_num = numbering
            .nums
            .iter()
            .map(|instance| (instance.num_id, instance.abstract_num_id))
            .collect::<HashMap<_, _>>();
        let mut occupied_nums = self.identifiers.preserved_numbering_instance_ids.clone();
        let mut occupied_abstract = self.identifiers.preserved_abstract_numbering_ids.clone();
        let referenced_abstract = abstract_by_num.values().copied().collect::<HashSet<_>>();
        occupied_abstract.extend(
            numbering
                .abstract_nums
                .iter()
                .map(|definition| definition.abstract_num_id)
                .filter(|id| !referenced_abstract.contains(id)),
        );
        let mut num_remap = HashMap::new();
        let mut abstract_remap = HashMap::new();
        for old_num in semantic_nums {
            num_remap.insert(
                old_num,
                reserve_u32(&mut occupied_nums, 1, "numbering instance")?,
            );
            if let Some(old_abstract) = abstract_by_num.get(&old_num)
                && !self
                    .identifiers
                    .preserved_abstract_numbering_ids
                    .contains(old_abstract)
                && !abstract_remap.contains_key(old_abstract)
            {
                abstract_remap.insert(
                    *old_abstract,
                    reserve_u32(&mut occupied_abstract, 0, "abstract numbering")?,
                );
            }
        }
        visit_body_paragraphs_mut(&mut self.document.body.content, &mut |paragraph| {
            if let Some(id) = paragraph
                .properties
                .as_mut()
                .and_then(|properties| properties.num_id.as_mut())
                && let Some(updated) = num_remap.get(id)
            {
                *id = *updated;
            }
        });
        let numbering = self.numbering.as_mut().expect("checked above");
        for definition in &mut numbering.abstract_nums {
            if let Some(updated) = abstract_remap.get(&definition.abstract_num_id) {
                definition.abstract_num_id = *updated;
            }
        }
        for instance in &mut numbering.nums {
            if let Some(updated) = num_remap.get(&instance.num_id) {
                instance.num_id = *updated;
            }
            if let Some(updated) = abstract_remap.get(&instance.abstract_num_id) {
                instance.abstract_num_id = *updated;
            }
        }
        if self.identifiers.preserved_abstract_numbering_ids.is_empty()
            && self.identifiers.preserved_numbering_instance_ids.is_empty()
        {
            numbering
                .abstract_nums
                .sort_by_key(|definition| definition.abstract_num_id);
            numbering.nums.sort_by_key(|instance| instance.num_id);
        }
        Ok(())
    }

    fn opaque_relationship_ids_from_serialized_story(
        &self,
        modeled_relationships: &[String],
    ) -> Result<Vec<String>> {
        let nested_namespace_owners = self
            .package
            .get_part(&self.doc_part_name)
            .map(nested_modeled_namespace_owners)
            .transpose()?
            .unwrap_or_default();
        let serialize = |document: &CT_Document| -> Result<Vec<u8>> {
            let xml = document.to_xml()?;
            replay_nested_namespace_declarations(&xml, &nested_namespace_owners)
        };

        let current_xml = serialize(&self.document)?;
        let current_relationships = xml_relationship_ids_in_order(&current_xml)?;
        let mut masked = self.document.clone();
        let mut masked_values = HashSet::new();
        let mut remap = HashMap::new();
        for (index, relationship_id) in modeled_relationships.iter().enumerate() {
            let mut ordinal = index;
            let sentinel = loop {
                let candidate = format!("rdocx-modeled-relationship-{ordinal}");
                if !current_relationships.contains(&candidate)
                    && masked_values.insert(candidate.clone())
                {
                    break candidate;
                }
                ordinal = ordinal
                    .checked_add(modeled_relationships.len().max(1))
                    .ok_or_else(|| {
                        Error::Other("modeled relationship mask range is exhausted".to_owned())
                    })?;
            };
            remap.insert(relationship_id.clone(), sentinel);
        }

        visit_authored_drawings_mut(&mut masked.body.content, &mut |drawing| {
            if let Some(mut drawing) = drawing
                .inline
                .as_mut()
                .map(DrawingMut::Inline)
                .or_else(|| drawing.anchor.as_mut().map(DrawingMut::Anchor))
            {
                drawing.remap_relationship(&remap);
            }
        });
        visit_body_paragraphs_mut(&mut masked.body.content, &mut |paragraph| {
            for hyperlink in &mut paragraph.hyperlinks {
                if let Some(id) = &mut hyperlink.rel_id
                    && let Some(updated) = remap.get(id)
                {
                    id.clone_from(updated);
                }
            }
            if let Some(section) = paragraph
                .properties
                .as_mut()
                .and_then(|properties| properties.sect_pr.as_mut())
            {
                for reference in section
                    .header_refs
                    .iter_mut()
                    .chain(&mut section.footer_refs)
                {
                    if let Some(updated) = remap.get(&reference.rel_id) {
                        reference.rel_id.clone_from(updated);
                    }
                }
            }
        });
        if let Some(section) = &mut masked.body.sect_pr {
            for reference in section
                .header_refs
                .iter_mut()
                .chain(&mut section.footer_refs)
            {
                if let Some(updated) = remap.get(&reference.rel_id) {
                    reference.rel_id.clone_from(updated);
                }
            }
        }

        let masked_xml = serialize(&masked)?;
        Ok(xml_relationship_ids_in_order(&masked_xml)?
            .into_iter()
            .filter(|id| !masked_values.contains(id) && current_relationships.contains(id))
            .collect())
    }

    fn canonicalize_drawing_ids(&mut self) -> Result<()> {
        let owner = self.doc_part_name.clone();
        let owner_identity = relationship_owner_identity(&owner);
        if let Some(section) = &mut self.document.body.sect_pr {
            section
                .header_refs
                .sort_by_key(|reference| hdr_ftr_type_order(reference.hdr_ftr_type));
            section
                .footer_refs
                .sort_by_key(|reference| hdr_ftr_type_order(reference.hdr_ftr_type));
        }
        let preserved_relationships = self
            .identifiers
            .preserved_relationship_ids
            .get(&owner_identity)
            .cloned()
            .unwrap_or_default();
        let authored_bundle_relationships = self
            .identifiers
            .authored_bundle_relationship_ids
            .get(&owner_identity)
            .cloned()
            .unwrap_or_default();
        let mut story_relationships = Vec::new();
        collect_relationship_ids(&self.document.body.content, &mut story_relationships);
        if let Some(section) = &self.document.body.sect_pr {
            for reference in section.header_refs.iter().chain(&section.footer_refs) {
                if !story_relationships.contains(&reference.rel_id) {
                    story_relationships.push(reference.rel_id.clone());
                }
            }
        }
        let opaque_relationships =
            self.opaque_relationship_ids_from_serialized_story(&story_relationships)?;
        let relationships = self
            .package
            .get_part_rels(&owner)
            .map(|relationships| relationships.items.clone())
            .unwrap_or_default();
        story_relationships.retain(|id| {
            !preserved_relationships.contains(id)
                && relationships
                    .iter()
                    .any(|relationship| relationship.id == *id)
        });
        let mut bundle_relationships = relationships
            .iter()
            .filter_map(|relationship| {
                (!preserved_relationships.contains(&relationship.id)
                    && authored_bundle_relationships.contains(&relationship.id)
                    && relationship_is_internal(relationship))
                .then(|| {
                    bundle_relationship_order(&relationship.rel_type).map(|order| {
                        (
                            order,
                            relationship.target.as_str(),
                            relationship.id.as_str(),
                            bundle_relationship_precedes_story(&relationship.rel_type),
                        )
                    })
                })
                .flatten()
            })
            .collect::<Vec<_>>();
        bundle_relationships.sort_unstable();
        let mut semantic_relationships = bundle_relationships
            .iter()
            .filter(|(_, _, _, precedes_story)| *precedes_story)
            .map(|(_, _, id, _)| (*id).to_owned())
            .collect::<Vec<_>>();
        semantic_relationships.extend(story_relationships);
        if let Some(section) = &self.document.body.sect_pr {
            for reference in section.header_refs.iter().chain(&section.footer_refs) {
                if !preserved_relationships.contains(&reference.rel_id)
                    && !semantic_relationships.contains(&reference.rel_id)
                {
                    semantic_relationships.push(reference.rel_id.clone());
                }
            }
        }
        let mut unreferenced_relationships = relationships
            .iter()
            .filter(|relationship| {
                !preserved_relationships.contains(&relationship.id)
                    && !semantic_relationships.contains(&relationship.id)
                    && !authored_bundle_relationships.contains(&relationship.id)
                    && !opaque_relationships.contains(&relationship.id)
            })
            .map(|relationship| {
                let internal = relationship_is_internal(relationship);
                let target = if internal {
                    part_name_identity(&OpcPackage::resolve_rel_target(
                        &owner,
                        &relationship.target,
                    ))
                } else {
                    relationship.target.clone()
                };
                (
                    !internal,
                    relationship.rel_type.as_str(),
                    target,
                    relationship.target_mode.as_deref().unwrap_or(""),
                    relationship.id.as_str(),
                )
            })
            .collect::<Vec<_>>();
        unreferenced_relationships.sort_unstable();
        semantic_relationships.extend(
            unreferenced_relationships
                .into_iter()
                .map(|(_, _, _, _, id)| id.to_owned()),
        );
        semantic_relationships.extend(
            bundle_relationships
                .into_iter()
                .filter(|(_, _, _, precedes_story)| !precedes_story)
                .map(|(_, _, id, _)| id.to_owned()),
        );

        let mut occupied_relationships = preserved_relationships.clone();
        occupied_relationships.extend(opaque_relationships);
        occupied_relationships.extend(
            relationships
                .iter()
                .filter(|relationship| !semantic_relationships.contains(&relationship.id))
                .map(|relationship| relationship.id.clone()),
        );
        let mut relationship_cursor = if semantic_relationships.is_empty() {
            0
        } else {
            preserved_relationships
                .iter()
                .filter_map(|id| id.strip_prefix("rId")?.parse::<u32>().ok())
                .max()
                .map_or(Ok(0), |maximum| {
                    maximum.checked_add(1).ok_or_else(|| {
                        Error::Other("relationship id range is exhausted".to_owned())
                    })
                })?
        };
        let mut relationship_remap = HashMap::new();
        for old in &semantic_relationships {
            let updated = reserve_relationship_from_cursor(
                &mut occupied_relationships,
                &mut relationship_cursor,
            )?;
            relationship_remap.insert(old.clone(), updated);
        }
        if let Some(relationships) = self.package.get_part_rels_mut(&owner) {
            for relationship in &mut relationships.items {
                if let Some(updated) = relationship_remap.get(&relationship.id) {
                    relationship.id.clone_from(updated);
                }
            }
            let mut authored = Vec::new();
            relationships.items.retain(|relationship| {
                if preserved_relationships.contains(&relationship.id) {
                    true
                } else {
                    authored.push(relationship.clone());
                    false
                }
            });
            authored.sort_by_key(|relationship| {
                relationship
                    .id
                    .strip_prefix("rId")
                    .and_then(|value| value.parse::<u32>().ok())
                    .unwrap_or(u32::MAX)
            });
            relationships.items.extend(authored);
        }
        if !relationship_remap.is_empty() {
            if let Some(occupied) = self.identifiers.relationship_ids.get_mut(&owner_identity) {
                for old in relationship_remap.keys() {
                    occupied.remove(old);
                }
                occupied.extend(relationship_remap.values().cloned());
            }
            if let Some(authored) = self
                .identifiers
                .authored_story_relationship_ids
                .get_mut(&owner_identity)
            {
                let remapped = authored
                    .iter()
                    .map(|id| relationship_remap.get(id).unwrap_or(id).clone())
                    .collect();
                *authored = remapped;
            }
            if let Some(authored) = self
                .identifiers
                .authored_bundle_relationship_ids
                .get_mut(&owner_identity)
            {
                let remapped = authored
                    .iter()
                    .map(|id| relationship_remap.get(id).unwrap_or(id).clone())
                    .collect();
                *authored = remapped;
            }
            if let Some(provisional) = self
                .identifiers
                .provisional_relationship_ids
                .get_mut(&owner_identity)
            {
                provisional.retain(|id| !relationship_remap.contains_key(id));
                if provisional.is_empty() {
                    self.identifiers
                        .provisional_relationship_ids
                        .remove(&owner_identity);
                }
            }
        }
        let mut occupied_drawings = self.identifiers.preserved_drawing_ids.clone();
        let mut drawing_count = 0usize;
        visit_authored_drawings(&self.document.body.content, &mut |_| drawing_count += 1);
        let mut drawing_ids = Vec::with_capacity(drawing_count);
        for _ in 0..drawing_count {
            drawing_ids.push(reserve_u32(&mut occupied_drawings, 1, "drawing")?);
        }
        let mut drawing_ids = drawing_ids.into_iter();
        let mut drawing_error = None;
        visit_authored_drawings_mut(&mut self.document.body.content, &mut |drawing| {
            let Some(mut drawing) = drawing
                .inline
                .as_mut()
                .map(DrawingMut::Inline)
                .or_else(|| drawing.anchor.as_mut().map(DrawingMut::Anchor))
            else {
                drawing_error = Some(Error::Other(
                    "authored drawing has no inline or anchor payload".to_owned(),
                ));
                return;
            };
            let Some(id) = drawing_ids.next() else {
                drawing_error = Some(Error::Other(
                    "authored drawing identifier count changed during canonicalization".to_owned(),
                ));
                return;
            };
            drawing.set_doc_pr_id(id);
            drawing.remap_relationship(&relationship_remap);
        });
        if let Some(error) = drawing_error {
            return Err(error);
        }
        if drawing_ids.next().is_some() {
            return Err(Error::Other(
                "authored drawing identifier count changed during canonicalization".to_owned(),
            ));
        }
        visit_body_paragraphs_mut(&mut self.document.body.content, &mut |paragraph| {
            for hyperlink in &mut paragraph.hyperlinks {
                if let Some(id) = &mut hyperlink.rel_id
                    && let Some(updated) = relationship_remap.get(id)
                {
                    id.clone_from(updated);
                }
            }
            if let Some(section) = paragraph
                .properties
                .as_mut()
                .and_then(|properties| properties.sect_pr.as_mut())
            {
                for reference in section
                    .header_refs
                    .iter_mut()
                    .chain(&mut section.footer_refs)
                {
                    if let Some(updated) = relationship_remap.get(&reference.rel_id) {
                        reference.rel_id.clone_from(updated);
                    }
                }
            }
        });
        if let Some(section) = &mut self.document.body.sect_pr {
            for reference in section
                .header_refs
                .iter_mut()
                .chain(&mut section.footer_refs)
            {
                if let Some(updated) = relationship_remap.get(&reference.rel_id) {
                    reference.rel_id.clone_from(updated);
                }
            }
        }
        self.canonicalize_header_footer_relationships()?;
        self.canonicalize_image_parts()?;
        Ok(())
    }

    fn active_header_footer_parts(&self) -> Vec<String> {
        let mut parts = Vec::new();
        for (relationship_id, is_header) in self.header_footer_rel_ids() {
            if let Some(part) = self.header_footer_part_name(&relationship_id, is_header)
                && !parts.contains(&part)
            {
                parts.push(part);
            }
        }
        parts
    }

    fn main_story_relationship_ids(&self) -> Vec<String> {
        let mut ids = Vec::new();
        collect_relationship_ids(&self.document.body.content, &mut ids);
        if let Some(section) = &self.document.body.sect_pr {
            for reference in section.header_refs.iter().chain(&section.footer_refs) {
                if !ids.contains(&reference.rel_id) {
                    ids.push(reference.rel_id.clone());
                }
            }
        }
        ids
    }

    fn canonicalize_header_footer_relationships(&mut self) -> Result<()> {
        let parts = self.active_header_footer_parts();
        let mut occupied_drawings = self.identifiers.preserved_drawing_ids.clone();
        visit_authored_drawings(&self.document.body.content, &mut |drawing| {
            if let Some(inline) = &drawing.inline {
                occupied_drawings.insert(inline.doc_pr_id);
            }
            if let Some(anchor) = &drawing.anchor {
                occupied_drawings.insert(anchor.doc_pr_id);
            }
        });

        for owner in parts {
            let owner_identity = relationship_owner_identity(&owner);
            let fully_authored = self.identifiers.authored_story_parts.contains(&owner);
            let partially_authored = self
                .identifiers
                .authored_story_relationship_ids
                .get(&owner_identity)
                .cloned()
                .unwrap_or_default();
            if !fully_authored && partially_authored.is_empty() {
                continue;
            }
            let xml = self
                .package
                .get_part(&owner)
                .ok_or_else(|| Error::Other(format!("story part {owner} is missing")))?
                .to_vec();
            let referenced = xml_relationship_ids_in_order(&xml)?;
            let preserved = self
                .identifiers
                .preserved_relationship_ids
                .get(&owner_identity)
                .cloned()
                .unwrap_or_default();
            let relationships = self
                .package
                .get_part_rels(&owner)
                .map(|relationships| relationships.items.clone())
                .unwrap_or_default();
            let semantic = referenced
                .into_iter()
                .filter(|id| {
                    (fully_authored && !preserved.contains(id) || partially_authored.contains(id))
                        && relationships
                            .iter()
                            .any(|relationship| relationship.id == *id)
                })
                .collect::<Vec<_>>();
            let mut occupied = preserved.clone();
            occupied.extend(
                relationships
                    .iter()
                    .filter(|relationship| !semantic.contains(&relationship.id))
                    .map(|relationship| relationship.id.clone()),
            );
            let mut cursor = if semantic.is_empty() {
                0
            } else {
                preserved
                    .iter()
                    .filter_map(|id| id.strip_prefix("rId")?.parse::<u32>().ok())
                    .max()
                    .map_or(Ok(0), |maximum| {
                        maximum.checked_add(1).ok_or_else(|| {
                            Error::Other("relationship id range is exhausted".to_owned())
                        })
                    })?
            };
            let mut remap = HashMap::new();
            for old in semantic {
                remap.insert(
                    old,
                    reserve_relationship_from_cursor(&mut occupied, &mut cursor)?,
                );
            }
            if let Some(owner_relationships) = self.package.get_part_rels_mut(&owner) {
                for relationship in &mut owner_relationships.items {
                    if let Some(updated) = remap.get(&relationship.id) {
                        relationship.id.clone_from(updated);
                    }
                }
                if fully_authored {
                    let mut authored = owner_relationships
                        .items
                        .iter()
                        .filter(|relationship| !preserved.contains(&relationship.id))
                        .cloned()
                        .collect::<Vec<_>>();
                    authored.sort_by_key(|relationship| {
                        relationship
                            .id
                            .strip_prefix("rId")
                            .and_then(|value| value.parse::<u32>().ok())
                            .unwrap_or(u32::MAX)
                    });
                    let mut authored = authored.into_iter();
                    for relationship in &mut owner_relationships.items {
                        if !preserved.contains(&relationship.id) {
                            *relationship = authored
                                .next()
                                .expect("each authored relationship slot has a value");
                        }
                    }
                }
            }
            let xml = remap_xml_relationship_ids(&xml, &remap)?;
            if let Some(ids) = self
                .identifiers
                .authored_story_relationship_ids
                .get_mut(&owner_identity)
            {
                *ids = ids
                    .drain()
                    .map(|id| remap.get(&id).cloned().unwrap_or(id))
                    .collect();
            }
            let xml = if fully_authored {
                rewrite_authored_doc_pr_ids(&xml, &mut occupied_drawings)?
            } else {
                xml
            };
            self.package.set_part(&owner, xml);
        }
        Ok(())
    }

    fn canonicalize_image_parts(&mut self) -> Result<()> {
        let mut ordered = Vec::new();
        let mut owners = vec![(
            self.doc_part_name.clone(),
            self.main_story_relationship_ids(),
        )];
        for owner in self.active_header_footer_parts() {
            let owner_identity = relationship_owner_identity(&owner);
            if self.identifiers.authored_story_parts.contains(&owner)
                || self
                    .identifiers
                    .authored_story_relationship_ids
                    .contains_key(&owner_identity)
            {
                let ids = self
                    .package
                    .get_part(&owner)
                    .map(xml_relationship_ids_in_order)
                    .transpose()?
                    .unwrap_or_default();
                owners.push((owner, ids));
            }
        }
        for (owner, ids) in &owners {
            let Some(relationships) = self.package.get_part_rels(owner) else {
                continue;
            };
            for id in ids {
                if let Some(relationship) = relationships.get_by_id(id)
                    && relationship.rel_type == rel_types::IMAGE
                {
                    let part = OpcPackage::resolve_rel_target(owner, &relationship.target);
                    if !self.identifiers.part_is_preserved(&part) && !ordered.contains(&part) {
                        ordered.push(part);
                    }
                }
            }
        }
        let mut occupied = self.identifiers.preserved_part_names.clone();
        occupied.extend(
            self.package
                .parts
                .keys()
                .filter(|part| !ordered.contains(part))
                .map(|part| part_name_identity(part)),
        );
        let mut remap = HashMap::new();
        for old in ordered {
            let extension = old
                .rsplit_once('.')
                .map(|(_, extension)| extension)
                .unwrap_or("bin")
                .to_owned();
            remap.insert(
                old,
                reserve_part_from_set(&mut occupied, "/word/media", "image", &extension)?,
            );
        }
        for (owner, _) in &owners {
            if let Some(relationships) = self.package.get_part_rels_mut(owner) {
                for relationship in &mut relationships.items {
                    if relationship.rel_type == rel_types::IMAGE {
                        let old = OpcPackage::resolve_rel_target(owner, &relationship.target);
                        if let Some(updated) = remap.get(&old) {
                            relationship.target = relative_descendant_target(owner, updated);
                        }
                    }
                }
            }
        }
        let mut moved = Vec::new();
        for (old, updated) in remap {
            if old != updated
                && let Some(bytes) = self.package.remove_part(&old)
            {
                moved.push((updated.clone(), bytes));
                if let Some(content_type) = self.package.content_types.remove_override(&old) {
                    self.package
                        .content_types
                        .add_override(&updated, &content_type);
                }
            }
        }
        for (part_name, bytes) in moved {
            self.package.set_part(&part_name, bytes);
        }
        Ok(())
    }

    /// Commit staged package state without discarding reusable layout work.
    pub(crate) fn commit_staged_mutation(&mut self, mut candidate: Self) {
        candidate.package_signatures_invalidated |= candidate
            .package
            .package_rels
            .items
            .iter()
            .any(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE_ORIGIN)
            && (self.package_signatures_invalidated
                || crate::embedded::synchronized_package_mutation_invalidates_signature(
                    &self.package,
                    &candidate.package,
                ));
        std::mem::swap(
            &mut self.normal_layout_engine,
            &mut candidate.normal_layout_engine,
        );
        std::mem::swap(
            &mut self.bundled_fallback_layout_engine,
            &mut candidate.bundled_fallback_layout_engine,
        );
        *self = candidate;
    }

    /// Open a document from a file path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let package = OpcPackage::open(path)?;
        Self::from_package(package)
    }

    /// Open a password-protected agile-encrypted document from a file path.
    #[cfg(all(feature = "agile-encryption", not(target_arch = "wasm32")))]
    pub fn open_encrypted<P: AsRef<Path>>(path: P, password: &str) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let package = OpcPackage::from_encrypted_reader(file, password)?;
        Self::from_package(package)
    }

    /// Open a document from bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, PackageReadLimits::UNBOUNDED)
    }

    /// Open a password-protected agile-encrypted document from bytes.
    #[cfg(all(feature = "agile-encryption", not(target_arch = "wasm32")))]
    pub fn from_encrypted_bytes(bytes: &[u8], password: &str) -> Result<Self> {
        Self::from_encrypted_bytes_with_limits(bytes, password, PackageReadLimits::UNBOUNDED)
    }

    /// Open an agile-encrypted document while bounding OPC archive expansion.
    #[cfg(all(feature = "agile-encryption", not(target_arch = "wasm32")))]
    pub fn from_encrypted_bytes_with_limits(
        bytes: &[u8],
        password: &str,
        limits: PackageReadLimits,
    ) -> Result<Self> {
        let cursor = std::io::Cursor::new(bytes);
        let package = OpcPackage::from_encrypted_reader_with_limits(cursor, password, limits)?;
        Self::from_package(package)
    }

    /// Open a document from bytes while bounding OPC archive expansion.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: PackageReadLimits) -> Result<Self> {
        let cursor = std::io::Cursor::new(bytes);
        let package = OpcPackage::from_reader_with_limits(cursor, limits)?;
        Self::from_package(package)
    }

    /// Return the exact Word package class declared for the main document part.
    pub fn package_class(&self) -> Result<WordPackageClass> {
        validated_word_package_class(&self.package).map(|(_, class)| class)
    }

    /// Serialize a staged copy as the requested Word package class.
    ///
    /// This changes only the main part's content-type override. Macro projects
    /// and all other package payloads remain present.
    pub fn to_bytes_as(&self, class: WordPackageClass) -> Result<Vec<u8>> {
        let mut candidate = self.clone_for_staging();
        candidate.prepare_staged_package()?;
        candidate
            .package
            .content_types
            .add_override(&candidate.doc_part_name, class.content_type());
        candidate.package_signatures_invalidated |=
            crate::embedded::synchronized_package_mutation_invalidates_signature(
                &self.package,
                &candidate.package,
            );
        crate::embedded::persist_invalidated_package_signature(
            &mut candidate.package,
            candidate.package_signatures_invalidated,
        )?;
        let mut output = std::io::Cursor::new(Vec::new());
        candidate.package.write_to(&mut output)?;
        let bytes = output.into_inner();
        Self::from_bytes(&bytes)?;
        Ok(bytes)
    }

    /// Save a staged copy as the requested Word package class.
    pub fn save_as_package_class<P: AsRef<Path>>(
        &self,
        path: P,
        class: WordPackageClass,
    ) -> Result<()> {
        let bytes = self.to_bytes_as(class)?;
        write_atomic_file(
            path.as_ref(),
            &bytes,
            "invalid Word package file name",
            "could not allocate Word package save staging file",
        )?;
        Ok(())
    }

    /// Verify package signatures without asserting certificate-chain trust.
    #[cfg(feature = "digital-signatures")]
    pub fn verify_signatures(
        &self,
    ) -> std::result::Result<Vec<oxml_opc::SignatureReport>, oxml_opc::OpcError> {
        self.package.verify_signatures()
    }

    /// Sign the current typed document with the strict RSA-SHA256 profile.
    ///
    /// The private key must be PKCS#8 DER and the certificate must be X.509
    /// DER. Typed state is flushed and signed on a staged clone, so any
    /// serialization, key, certificate, or signature failure leaves this
    /// document unchanged. Certificate-chain trust remains caller policy.
    #[cfg(all(feature = "digital-signatures", not(target_arch = "wasm32")))]
    pub fn sign(
        &mut self,
        private_key_pkcs8_der: &[u8],
        certificate_der: &[u8],
    ) -> Result<oxml_opc::SignatureReport> {
        let mut candidate = self.clone_for_staging();
        candidate.prepare_staged_package()?;
        let report = candidate
            .package
            .sign(private_key_pkcs8_der, certificate_der)?;
        self.commit_staged_mutation(candidate);
        Ok(report)
    }

    pub(crate) fn from_package(package: OpcPackage) -> Result<Self> {
        let (doc_part_name, _) = validated_word_package_class(&package)?;

        let doc_xml = package
            .get_part(&doc_part_name)
            .ok_or(Error::NoDocumentPart)?;
        let namespace_scopes = document_namespace_scopes(doc_xml)?;
        let document = CT_Document::from_xml(doc_xml)?;

        // Resolve the part a relationship of the given type points at.
        let resolve_part = |rel_type: &str| -> Option<String> {
            let rels = package.get_part_rels(&doc_part_name)?;
            let rel = rels.items.iter().find(|relationship| {
                relationship.rel_type == rel_type && relationship_is_internal(relationship)
            })?;
            Some(OpcPackage::resolve_rel_target(&doc_part_name, &rel.target))
        };

        // Try to load styles, remembering where they came from.
        let styles_part_name = resolve_part(rel_types::STYLES);
        let styles = match styles_part_name
            .as_deref()
            .and_then(|p| package.get_part(p))
        {
            Some(styles_xml) => CT_Styles::from_xml(styles_xml)?,
            None => CT_Styles::new_default(),
        };

        // Try to load numbering definitions
        let numbering_part_name = resolve_part(rel_types::NUMBERING);
        let numbering = match numbering_part_name
            .as_deref()
            .and_then(|p| package.get_part(p))
        {
            Some(num_xml) => Some(CT_Numbering::from_xml(num_xml)?),
            None => None,
        };

        let settings_part_name = resolve_part(rel_types::SETTINGS);
        let settings = match settings_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
        {
            Some(xml) => Some(CT_Settings::from_xml(xml)?),
            None => None,
        };

        let theme_part_name = resolve_part(rel_types::THEME);
        let theme = match theme_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
        {
            Some(xml) => oxml_drawing::theme::CT_OfficeStyleSheet::from_xml(xml).ok(),
            None => None,
        };

        let font_table_part_name = resolve_part(rel_types::FONT_TABLE);
        let font_table = match font_table_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
        {
            Some(xml) => FontTable::from_xml(xml).ok(),
            None => None,
        };

        // Core properties are a package-level relationship, not a document part.
        let core_properties_part_name = package
            .package_rels
            .items
            .iter()
            .find(|relationship| {
                relationship.rel_type == CORE_PROPERTIES_REL_TYPE
                    && relationship_is_internal(relationship)
            })
            .map(|rel| OpcPackage::resolve_rel_target("/", &rel.target));
        let core_properties = core_properties_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
            .and_then(|xml| CoreProperties::from_xml(xml).ok());

        let application_properties_part_name = package
            .package_rels
            .items
            .iter()
            .find(|relationship| {
                relationship.rel_type == rel_types::EXTENDED_PROPERTIES
                    && relationship_is_internal(relationship)
            })
            .map(|rel| OpcPackage::resolve_rel_target("/", &rel.target));
        let application_properties = application_properties_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
            .and_then(|xml| AppProperties::from_xml(xml).ok());

        let custom_properties_part_name = package
            .package_rels
            .items
            .iter()
            .find(|relationship| {
                relationship.rel_type == rel_types::CUSTOM_PROPERTIES
                    && relationship_is_internal(relationship)
            })
            .map(|rel| OpcPackage::resolve_rel_target("/", &rel.target));
        let custom_properties = custom_properties_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
            .and_then(|xml| CustomProperties::from_xml(xml).ok());

        let footnotes_part_name = resolve_part(rel_types::FOOTNOTES);
        let footnotes = footnotes_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
            .and_then(|xml| rdocx_oxml::footnotes::CT_Footnotes::from_xml(xml).ok())
            .unwrap_or_default();

        let comments_part_name = resolve_part(rel_types::COMMENTS);
        let comments = match comments_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
        {
            Some(xml) => Some(rdocx_oxml::comments::CT_Comments::from_xml(xml)?),
            None => None,
        };
        let comments_extended_part_name = resolve_part(crate::comments::COMMENTS_EXTENDED_REL_TYPE);
        let comments_extended = match comments_extended_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
        {
            Some(xml) => Some(rdocx_oxml::comments_extended::CT_CommentsEx::from_xml(xml)?),
            None => None,
        };
        let glossary = crate::building_block::load_glossary(&package, &doc_part_name)?;
        let (glossary_part_name, glossary) = match glossary {
            Some((part_name, glossary)) => (Some(part_name), Some(glossary)),
            None => (None, None),
        };

        let package_signatures_invalidated =
            crate::embedded::known_invalid_package_signature_on_open(&package);
        let identifiers = DocumentIdentifiers::scan(&package)?;
        Ok(Document {
            package,
            document,
            root_namespace_declarations: namespace_scopes.root_declarations,
            body_namespace_declarations: namespace_scopes.body_declarations,
            body_namespace_bindings: namespace_scopes.body_bindings,
            styles,
            numbering,
            core_properties,
            application_properties,
            custom_properties,
            core_properties_part_name,
            application_properties_part_name,
            custom_properties_part_name,
            custom_properties_owned: false,
            doc_part_name,
            styles_part_name,
            numbering_part_name,
            settings,
            settings_part_name,
            settings_owned: false,
            theme,
            theme_part_name,
            theme_dirty: false,
            font_table,
            font_table_part_name,
            font_table_dirty: false,
            identifiers,
            footnotes,
            footnotes_part_name,
            footnotes_dirty: false,
            comments,
            comments_part_name,
            comments_extended,
            comments_extended_part_name,
            comments_owned: false,
            comments_extended_owned: false,
            embedded_invalidated_signatures: HashSet::new(),
            package_signatures_invalidated,
            glossary,
            glossary_part_name,
            glossary_dirty: false,
            layout_cache: Mutex::new(None),
            normal_layout_engine: Mutex::new(None),
            deterministic_layout_cache: Mutex::new(None),
            bundled_fallback_layout_engine: Mutex::new(None),
        })
    }

    /// Clear layouts derived from the current document state.
    pub(crate) fn invalidate_layout(&mut self) {
        self.layout_cache
            .get_mut()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .take();
        self.deterministic_layout_cache
            .get_mut()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .take();
    }

    #[cfg(test)]
    pub(crate) fn redaction_cache_identity(
        &self,
    ) -> (Option<usize>, Option<usize>, Option<usize>, Option<usize>) {
        let layout = self
            .layout_cache
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .as_ref()
            .map(|value| Arc::as_ptr(value) as usize);
        let normal_engine = self
            .normal_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .as_ref()
            .map(|value| value as *const _ as usize);
        let deterministic = self
            .deterministic_layout_cache
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .as_ref()
            .map(|value| Arc::as_ptr(value) as usize);
        let bundled_engine = self
            .bundled_fallback_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .as_ref()
            .map(|value| value as *const _ as usize);
        (layout, normal_engine, deterministic, bundled_engine)
    }

    /// Return the normal-font layout, computing it once after each mutation.
    fn cached_layout(&self) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
        let mut cache = self
            .layout_cache
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(layout) = cache.as_ref() {
            return Ok(Arc::clone(layout));
        }

        let input = self.build_layout_input();
        #[cfg(test)]
        record_layout_invocation();
        let mut engine = self
            .normal_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        let engine = engine.get_or_insert_with(rdocx_layout::engine::Engine::new);
        let layout = Arc::new(rdocx_layout::layout_document_with_reusable_engine(
            engine, &input,
        )?);
        *cache = Some(Arc::clone(&layout));
        Ok(layout)
    }

    /// Return the bundled-font-only layout, computing it once after mutation.
    fn cached_deterministic_layout(&self) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
        let mut cache = self
            .deterministic_layout_cache
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(layout) = cache.as_ref() {
            return Ok(Arc::clone(layout));
        }

        let input = self.build_layout_input();
        #[cfg(test)]
        record_layout_invocation();
        let layout = Arc::new(rdocx_layout::layout_document_deterministic_with_provenance(
            &input,
        )?);
        *cache = Some(Arc::clone(&layout));
        Ok(layout)
    }

    fn layout_for_options(
        &self,
        options: RenderOptions,
        deterministic: bool,
    ) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
        if options.revision_view == rdocx_layout::RevisionView::Accepted {
            return if deterministic {
                self.cached_deterministic_layout()
            } else {
                self.cached_layout()
            };
        }

        let mut input = self.build_layout_input();
        input.revision_view = options.revision_view;
        #[cfg(test)]
        record_layout_invocation();
        let layout = if deterministic {
            rdocx_layout::layout_document_deterministic_with_provenance(&input)?
        } else {
            let mut engine = self
                .normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner);
            let engine = engine.get_or_insert_with(rdocx_layout::engine::Engine::new);
            rdocx_layout::layout_document_with_reusable_engine(engine, &input)?
        };
        Ok(Arc::new(layout))
    }

    /// Save the document to a file path.
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.prepare_staged_package()?;
        crate::embedded::persist_invalidated_package_signature(
            &mut candidate.package,
            candidate.package_signatures_invalidated,
        )?;
        candidate.package.save(path)?;
        Ok(())
    }

    /// Save the document to a byte vector.
    pub fn to_bytes(&mut self) -> Result<Vec<u8>> {
        let mut candidate = self.clone_for_staging();
        candidate.prepare_staged_package()?;
        crate::embedded::persist_invalidated_package_signature(
            &mut candidate.package,
            candidate.package_signatures_invalidated,
        )?;
        let mut buf = std::io::Cursor::new(Vec::new());
        candidate.package.write_to(&mut buf)?;
        Ok(buf.into_inner())
    }

    /// Save a password-protected document using the fixed Agile write profile.
    #[cfg(all(feature = "agile-encryption", not(target_arch = "wasm32")))]
    pub fn save_encrypted<P: AsRef<Path>>(&self, path: P, password: &str) -> Result<()> {
        let bytes = self.to_encrypted_bytes(password)?;
        write_atomic_file(
            path.as_ref(),
            &bytes,
            "invalid file name",
            "could not allocate encrypted-save staging file",
        )?;
        Ok(())
    }

    /// Save a password-protected document to a byte vector.
    #[cfg(all(feature = "agile-encryption", not(target_arch = "wasm32")))]
    pub fn to_encrypted_bytes(&self, password: &str) -> Result<Vec<u8>> {
        let mut candidate = self.clone_for_staging();
        candidate.prepare_staged_package()?;
        crate::embedded::persist_invalidated_package_signature(
            &mut candidate.package,
            candidate.package_signatures_invalidated,
        )?;
        let mut encrypted = Vec::new();
        candidate
            .package
            .write_encrypted_to(&mut encrypted, password)?;
        Ok(encrypted)
    }

    pub(crate) fn prepare_staged_package(&mut self) -> Result<()> {
        self.preflight_flush_required_bundles()?;
        self.canonicalize_authored_identifiers()?;
        self.package_signatures_invalidated |=
            self.retained_package_signature_would_be_invalidated()?;
        self.flush_to_package()
    }

    pub(crate) fn prepare_and_reopen_staged(mut self) -> Result<Self> {
        self.prepare_staged_package()?;
        self.reopen_prepared_staged()
    }

    pub(crate) fn reopen_prepared_staged(self) -> Result<Self> {
        self.reopen_prepared_staged_with_limits(PackageReadLimits::UNBOUNDED)
    }

    pub(crate) fn reopen_prepared_staged_with_limits(
        self,
        limits: PackageReadLimits,
    ) -> Result<Self> {
        let provenance = self.identifiers.clone();
        let embedded_invalidated_signatures = self.embedded_invalidated_signatures.clone();
        let package_signatures_invalidated = self.package_signatures_invalidated;
        let mut output = std::io::Cursor::new(Vec::new());
        self.package.write_to(&mut output)?;
        let mut reopened = Self::from_bytes_with_limits(output.get_ref(), limits)?;
        reopened.identifiers.reconcile_provenance(&provenance);
        reopened.embedded_invalidated_signatures = embedded_invalidated_signatures;
        reopened.package_signatures_invalidated = package_signatures_invalidated;
        Ok(reopened)
    }

    fn preflight_flush_required_bundles(&mut self) -> Result<()> {
        self.reserve_styles_bundle()?;
        if self.numbering.is_some() {
            self.reserve_numbering_bundle()?;
        }
        if self.footnotes_dirty && !self.footnotes.footnotes.is_empty() {
            self.reserve_footnotes_bundle()?;
        }
        if let Some(part_name) = self.comments_part_name.clone()
            && self.comments.is_some()
        {
            self.ensure_part_relationship_checked(
                &part_name,
                rel_types::COMMENTS,
                crate::comments::COMMENTS_CONTENT_TYPE,
            )
            .map_err(|error| {
                Error::Other(format!("comments relationship allocation failed: {error}"))
            })?;
        }
        if let Some(part_name) = self.comments_extended_part_name.clone()
            && self.comments_extended.is_some()
        {
            self.ensure_part_relationship_checked(
                &part_name,
                crate::comments::COMMENTS_EXTENDED_REL_TYPE,
                crate::comments::COMMENTS_EXTENDED_CONTENT_TYPE,
            )
            .map_err(|error| {
                Error::Other(format!(
                    "comments-extended relationship allocation failed: {error}"
                ))
            })?;
        }
        if self.core_properties.is_some() {
            self.reserve_core_properties_bundle()?;
        }
        if self.application_properties.is_some() {
            self.reserve_application_properties_bundle()?;
        }
        if self.custom_properties.is_some() {
            self.reserve_custom_properties_bundle()?;
        }
        Ok(())
    }

    /// Write the in-memory document/styles back into the OPC package parts.
    pub(crate) fn flush_to_package(&mut self) -> Result<()> {
        // A read-only save with unsafe retained namespace shadows keeps the
        // producer's complete scopes and every retained raw subtree byte for
        // byte. A modified document fails closed when canonical serialization
        // could change the meaning or bytes of retained content.
        let existing_document_xml = self.package.get_part(&self.doc_part_name);
        let typed_document_is_unchanged = existing_document_xml
            .and_then(|xml| CT_Document::from_xml(xml).ok())
            .as_ref()
            == Some(&self.document);
        let nested_namespace_owners = existing_document_xml
            .map(nested_modeled_namespace_owners)
            .transpose()?
            .unwrap_or_default();
        let unsafe_prefix = unsafe_serializer_namespace_prefix(
            &self.root_namespace_declarations,
            &self.body_namespace_declarations,
        )
        .or_else(|| unsafe_nested_namespace_prefix(&nested_namespace_owners));
        let doc_xml = if typed_document_is_unchanged && unsafe_prefix.is_some() {
            existing_document_xml
                .expect("compared existing document XML")
                .to_vec()
        } else {
            if let Some(prefix) = unsafe_prefix {
                return Err(Error::Other(format!(
                    "cannot serialize a modified document with a shadowed `{prefix}` namespace"
                )));
            }
            let serialized = self.document.to_xml()?;
            replay_nested_namespace_declarations(&serialized, &nested_namespace_owners)?
        };
        self.package.set_part(&self.doc_part_name, doc_xml);

        // Serialize the styles part. A document opened without one still gets
        // rdocx's defaults written out, so make sure it is reachable: an
        // unreferenced, untyped part would simply be ignored by Word.
        let styles_xml = self.styles.to_xml()?;
        let styles_part_name = self.styles_part_name.clone();
        let styles_part = self
            .reserve_document_part_bundle(
                styles_part_name.as_deref(),
                DEFAULT_STYLES_PART,
                rel_types::STYLES,
                STYLES_CONTENT_TYPE,
            )
            .map_err(|error| {
                Error::Other(format!("styles relationship allocation failed: {error}"))
            })?;
        self.styles_part_name = Some(styles_part.clone());
        self.package.set_part(&styles_part, styles_xml);

        // Serialize numbering definitions if we have any
        if let Some(numbering_xml) = self
            .numbering
            .as_ref()
            .map(CT_Numbering::to_xml)
            .transpose()?
        {
            let numbering_part_name = self.numbering_part_name.clone();
            let numbering_part = self
                .reserve_document_part_bundle(
                    numbering_part_name.as_deref(),
                    DEFAULT_NUMBERING_PART,
                    rel_types::NUMBERING,
                    NUMBERING_CONTENT_TYPE,
                )
                .map_err(|error| {
                    Error::Other(format!("numbering relationship allocation failed: {error}"))
                })?;
            self.numbering_part_name = Some(numbering_part.clone());
            self.package.set_part(&numbering_part, numbering_xml);
        }

        // F-155 exposes settings as a read-only projection. Parsed settings
        // retain their complete producer bytes and are written back only to
        // the relationship-resolved part they came from.
        if let (Some(settings), Some(part_name)) = (&self.settings, &self.settings_part_name) {
            self.package.set_part(part_name, settings.to_xml()?);
        }

        if self.theme_dirty {
            let theme = self
                .theme
                .as_ref()
                .ok_or_else(|| Error::Other("dirty theme model is missing".to_owned()))?;
            let part_name = self
                .theme_part_name
                .as_ref()
                .ok_or_else(|| Error::Other("dirty theme part name is missing".to_owned()))?;
            self.package.set_part(
                part_name,
                theme.to_xml().map_err(|error| {
                    Error::Other(format!("theme serialization failed: {error}"))
                })?,
            );
        }

        if self.font_table_dirty {
            let table = self
                .font_table
                .as_ref()
                .ok_or_else(|| Error::Other("dirty font-table model is missing".to_owned()))?;
            let part_name = self
                .font_table_part_name
                .as_ref()
                .ok_or_else(|| Error::Other("dirty font-table part name is missing".to_owned()))?;
            self.package.set_part(part_name, table.to_xml()?);
        }

        // Preserve parsed footnote bytes until a facade mutation makes the typed view dirty.
        if self.footnotes_dirty && !self.footnotes.footnotes.is_empty() {
            let fx = self.footnotes.to_xml_footnotes()?;
            let footnotes_part_name = self.footnotes_part_name.clone();
            let footnotes_part = self
                .reserve_document_part_bundle(
                    footnotes_part_name.as_deref(),
                    "/word/footnotes.xml",
                    rel_types::FOOTNOTES,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
                )
                .map_err(|error| {
                    Error::Other(format!("footnotes relationship allocation failed: {error}"))
                })?;
            self.footnotes_part_name = Some(footnotes_part.clone());
            self.package.set_part(&footnotes_part, fx);
        }

        // An existing comments part is modelled and flushed to its resolved
        // relationship target. An absent part remains absent until the later
        // comment API creates one deliberately.
        if let (Some(comments), Some(part_name)) = (&self.comments, self.comments_part_name.clone())
        {
            let xml = comments.to_xml()?;
            self.package.set_part(&part_name, xml);
            self.ensure_part_relationship_checked(
                &part_name,
                rel_types::COMMENTS,
                crate::comments::COMMENTS_CONTENT_TYPE,
            )
            .map_err(|error| {
                Error::Other(format!("comments relationship allocation failed: {error}"))
            })?;
        }

        if let (Some(comments), Some(part_name)) = (
            &self.comments_extended,
            self.comments_extended_part_name.clone(),
        ) {
            let xml = comments.to_xml()?;
            self.package.set_part(&part_name, xml);
            self.ensure_part_relationship_checked(
                &part_name,
                crate::comments::COMMENTS_EXTENDED_REL_TYPE,
                crate::comments::COMMENTS_EXTENDED_CONTENT_TYPE,
            )
            .map_err(|error| {
                Error::Other(format!(
                    "comments-extended relationship allocation failed: {error}"
                ))
            })?;
        }

        if self.glossary_dirty {
            let glossary = self
                .glossary
                .as_ref()
                .ok_or_else(|| Error::Other("dirty glossary model is missing".to_owned()))?;
            let part_name = self
                .glossary_part_name
                .as_ref()
                .ok_or_else(|| Error::Other("dirty glossary part name is missing".to_owned()))?;
            self.package.set_part(part_name, glossary.to_xml()?);
        }

        // Serialize core properties to the package relationship's target.
        if let Some(core_xml) = self
            .core_properties
            .as_ref()
            .map(CoreProperties::to_xml)
            .transpose()?
        {
            let core_part = self.reserve_core_properties_bundle()?;
            self.package.set_part(&core_part, core_xml);
        }
        if let Some(application_xml) = self
            .application_properties
            .as_ref()
            .map(AppProperties::to_xml)
            .transpose()?
        {
            let application_part = self.reserve_application_properties_bundle()?;
            self.package.set_part(&application_part, application_xml);
        }
        if let Some(custom_xml) = self
            .custom_properties
            .as_ref()
            .map(CustomProperties::to_xml)
            .transpose()?
        {
            let custom_part = self.reserve_custom_properties_bundle()?;
            self.package.set_part(&custom_part, custom_xml);
        }

        Ok(())
    }

    /// Make sure `part_name` is reachable from the main document: it needs a
    /// relationship of `rel_type` and a content-type override.
    pub(crate) fn ensure_part_relationship_checked(
        &mut self,
        part_name: &str,
        rel_type: &str,
        content_type: &str,
    ) -> Result<()> {
        let doc_part_name = self.doc_part_name.clone();
        let already_linked =
            self.package
                .get_part_rels(&doc_part_name)
                .is_some_and(|relationships| {
                    relationships.items.iter().any(|relationship| {
                        relationship.rel_type == rel_type
                            && relationship_is_internal(relationship)
                            && OpcPackage::resolve_rel_target(&doc_part_name, &relationship.target)
                                == part_name
                    })
                });
        let pending_id = if already_linked {
            None
        } else {
            Some(
                self.identifiers
                    .reserve_relationship_id_checked(&doc_part_name)?,
            )
        };
        self.package
            .content_types
            .add_override(part_name, content_type);
        self.identifiers.register_content_type_override(part_name);
        if let Some(id) = pending_id {
            // Relationship targets are relative to the source part's directory.
            let target = relative_target(&doc_part_name, part_name);
            self.package
                .get_or_create_part_rels(&doc_part_name)
                .add_with_id(&id, rel_type, &target);
            if bundle_relationship_order(rel_type).is_some() {
                self.identifiers
                    .authored_bundle_relationship_ids
                    .entry(relationship_owner_identity(&doc_part_name))
                    .or_default()
                    .insert(id);
            }
        }
        Ok(())
    }

    fn reserve_document_part_bundle(
        &mut self,
        existing: Option<&str>,
        preferred: &str,
        rel_type: &str,
        content_type: &str,
    ) -> Result<String> {
        self.identifiers.observe_package_graph(&self.package)?;
        let part_name = match existing {
            Some(part_name) => part_name.to_owned(),
            None => self.identifiers.reserve_preferred_part_name(preferred)?,
        };
        let owner = self.doc_part_name.clone();
        let already_linked = self
            .package
            .get_part_rels(&owner)
            .is_some_and(|relationships| {
                relationships.items.iter().any(|relationship| {
                    relationship.rel_type == rel_type
                        && relationship_is_internal(relationship)
                        && OpcPackage::resolve_rel_target(&owner, &relationship.target) == part_name
                })
            });
        let pending_id = if already_linked {
            None
        } else {
            Some(
                self.identifiers
                    .reserve_bundle_relationship_id_checked(&owner, rel_type)?,
            )
        };
        self.package
            .content_types
            .add_override(&part_name, content_type);
        self.identifiers.register_content_type_override(&part_name);
        if let Some(id) = pending_id {
            let target = relative_target(&owner, &part_name);
            self.package
                .get_or_create_part_rels(&owner)
                .add_with_id(&id, rel_type, &target);
        }
        Ok(part_name)
    }

    fn reserve_core_properties_bundle(&mut self) -> Result<String> {
        let existing = self.core_properties_part_name.clone();
        let part_name = self.reserve_package_properties_bundle(
            existing.as_deref(),
            DEFAULT_CORE_PROPERTIES_PART,
            rel_types::CORE_PROPERTIES,
            content_types::CORE_PROPERTIES,
        )?;
        self.core_properties_part_name = Some(part_name.clone());
        Ok(part_name)
    }

    fn reserve_application_properties_bundle(&mut self) -> Result<String> {
        let existing = self.application_properties_part_name.clone();
        let part_name = self.reserve_package_properties_bundle(
            existing.as_deref(),
            DEFAULT_APP_PROPERTIES_PART,
            rel_types::EXTENDED_PROPERTIES,
            content_types::EXTENDED_PROPERTIES,
        )?;
        self.application_properties_part_name = Some(part_name.clone());
        Ok(part_name)
    }

    fn reserve_custom_properties_bundle(&mut self) -> Result<String> {
        let existing = self.custom_properties_part_name.clone();
        let part_name = self.reserve_package_properties_bundle(
            existing.as_deref(),
            "/docProps/custom.xml",
            rel_types::CUSTOM_PROPERTIES,
            content_types::CUSTOM_PROPERTIES,
        )?;
        self.custom_properties_part_name = Some(part_name.clone());
        Ok(part_name)
    }

    fn reserve_package_properties_bundle(
        &mut self,
        existing: Option<&str>,
        preferred: &str,
        rel_type: &str,
        content_type: &str,
    ) -> Result<String> {
        self.identifiers.observe_package_graph(&self.package)?;
        let part_name = match existing {
            Some(part_name) => part_name.to_owned(),
            None => self.identifiers.reserve_preferred_part_name(preferred)?,
        };
        let already_linked = self.package.package_rels.items.iter().any(|relationship| {
            relationship.rel_type == rel_type
                && relationship_is_internal(relationship)
                && OpcPackage::resolve_rel_target("/", &relationship.target) == part_name
        });
        let pending_id = if already_linked {
            None
        } else {
            Some(self.identifiers.reserve_relationship_id_checked("/")?)
        };
        self.package
            .content_types
            .add_override(&part_name, content_type);
        self.identifiers.register_content_type_override(&part_name);
        if let Some(id) = pending_id {
            let target = part_name.strip_prefix('/').unwrap_or(&part_name);
            self.package.package_rels.add_with_id(&id, rel_type, target);
        }
        Ok(part_name)
    }

    pub(crate) fn add_internal_relationship_checked(
        &mut self,
        owner: &str,
        rel_type: &str,
        target: &str,
    ) -> Result<String> {
        let id = self.identifiers.reserve_relationship_id_checked(owner)?;
        self.package
            .get_or_create_part_rels(owner)
            .add_with_id(&id, rel_type, target);
        Ok(id)
    }

    fn add_external_relationship_checked(
        &mut self,
        owner: &str,
        rel_type: &str,
        target: &str,
    ) -> Result<String> {
        let id = self.identifiers.reserve_relationship_id_checked(owner)?;
        self.package.get_or_create_part_rels(owner).items.push(
            oxml_opc::relationship::Relationship {
                id: id.clone(),
                rel_type: rel_type.to_owned(),
                target: target.to_owned(),
                target_mode: Some("External".to_owned()),
            },
        );
        Ok(id)
    }

    /// Stage a chart, its editable workbook, and the Word drawing that reaches them.
    fn add_chart_package(
        &mut self,
        source: ChartPackageSource<'_>,
        width: Length,
        height: Length,
    ) -> Result<()> {
        if width.to_emu() <= 0 || height.to_emu() <= 0 {
            return Err(Error::Other(
                "chart width and height must be positive".to_owned(),
            ));
        }

        let mut candidate = self.clone_for_staging();
        candidate
            .identifiers
            .observe_package_graph(&candidate.package)?;
        if matches!(&source, ChartPackageSource::Authored { .. }) {
            candidate.ensure_authored_chart_theme()?;
        }
        let chart_part = candidate
            .identifiers
            .reserve_part_name("/word/charts", "chart", "xml")?;
        let workbook_part =
            candidate
                .identifiers
                .reserve_part_name("/word/embeddings", "Workbook", "xlsx")?;
        let document_relationship_id = candidate
            .identifiers
            .reserve_relationship_id_checked(&candidate.doc_part_name)
            .map_err(|error| {
                Error::Other(format!("chart relationship allocation failed: {error}"))
            })?;
        candidate
            .package
            .get_or_create_part_rels(&candidate.doc_part_name)
            .add_with_id(
                &document_relationship_id,
                rel_types::CHART,
                &relative_target(&candidate.doc_part_name, &chart_part),
            );
        let workbook_relationship_id = candidate
            .identifiers
            .reserve_relationship_id_checked(&chart_part)
            .map_err(|error| {
                Error::Other(format!(
                    "chart workbook relationship allocation failed: {error}"
                ))
            })?;
        candidate
            .package
            .get_or_create_part_rels(&chart_part)
            .add_with_id(
                &workbook_relationship_id,
                rel_types::PACKAGE,
                &relative_target(&chart_part, &workbook_part),
            );
        let (chart_xml, workbook_bytes) = match source {
            ChartPackageSource::Typed { chart, workbook } => (
                chart_with_workbook_relationship(chart, &workbook_relationship_id)?,
                workbook
                    .to_xlsx_bytes()
                    .map_err(|error| Error::Other(format!("invalid chart workbook: {error}")))?,
            ),
            ChartPackageSource::Authored { kind, data } => {
                oxml_chart::authored_chart_parts(kind, data, &workbook_relationship_id)
                    .map_err(|error| Error::Other(format!("invalid chart data: {error}")))?
            }
        };

        let mut inline =
            CT_Inline::new_chart(&document_relationship_id, width.to_emu(), height.to_emu());
        inline.doc_pr_id = candidate.identifiers.reserve_drawing_id()?;
        let drawing = CT_Drawing::inline(inline);
        let run = CT_R {
            alt_drawings: Vec::new(),
            properties: None,
            content: vec![RunContent::Drawing(drawing)],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
        };
        let mut paragraph = CT_P::new();
        paragraph.runs.push(run);
        candidate
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));
        candidate.document.to_xml()?;

        candidate.package.set_part(&chart_part, chart_xml);
        candidate.package.set_part(&workbook_part, workbook_bytes);
        candidate
            .package
            .content_types
            .add_override(&chart_part, content_types::CHART);
        candidate
            .identifiers
            .register_content_type_override(&chart_part);
        candidate
            .package
            .content_types
            .add_override(&workbook_part, content_types::EMBEDDED_WORKBOOK);
        candidate
            .identifiers
            .register_content_type_override(&workbook_part);

        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    fn ensure_authored_chart_theme(&mut self) -> Result<()> {
        let document_part = self.doc_part_name.clone();
        let reusable = self
            .package
            .get_part_rels(&document_part)
            .and_then(|relationships| {
                relationships.items.iter().find(|relationship| {
                    relationship.rel_type == rel_types::THEME
                        && relationship_is_internal(relationship)
                })
            })
            .and_then(|relationship| {
                let part_name =
                    OpcPackage::resolve_rel_target(&document_part, &relationship.target);
                if self.package.content_types.content_type_for(&part_name)
                    != Some(content_types::THEME)
                {
                    return None;
                }
                let bytes = if self.theme_dirty
                    && self.theme_part_name.as_deref() == Some(part_name.as_str())
                {
                    self.theme.as_ref()?.to_xml().ok()?
                } else {
                    self.package.get_part(&part_name)?.to_vec()
                };
                let theme = oxml_drawing::theme::CT_OfficeStyleSheet::from_xml(&bytes).ok()?;
                Some((part_name, theme, bytes))
            });
        if let Some((part_name, theme, bytes)) = reusable {
            if self.theme_dirty && self.theme_part_name.as_deref() == Some(part_name.as_str()) {
                self.package.set_part(&part_name, bytes);
            }
            self.theme = Some(theme);
            self.theme_part_name = Some(part_name);
            return Ok(());
        }

        let theme = oxml_drawing::theme::CT_OfficeStyleSheet::office_default();
        let theme_xml = theme
            .to_xml()
            .map_err(|error| Error::Other(format!("invalid default Office theme: {error}")))?;
        self.identifiers.observe_package_graph(&self.package)?;
        let theme_part = self
            .identifiers
            .reserve_preferred_part_name(DEFAULT_THEME_PART)?;
        self.package.set_part(&theme_part, theme_xml);
        self.package
            .content_types
            .add_override(&theme_part, content_types::THEME);
        self.identifiers.register_content_type_override(&theme_part);

        let target = relative_target(&document_part, &theme_part);
        let relationships = self.package.get_or_create_part_rels(&document_part);
        if let Some(relationship) = relationships
            .items
            .iter_mut()
            .find(|relationship| relationship.rel_type == rel_types::THEME)
        {
            relationship.target = target;
            relationship.target_mode = None;
        } else {
            let id = self
                .identifiers
                .reserve_relationship_id_checked(&document_part)?;
            relationships.add_with_id(&id, rel_types::THEME, &target);
        }
        self.theme = Some(theme);
        self.theme_part_name = Some(theme_part);
        self.theme_dirty = false;
        Ok(())
    }

    // ---- Paragraph access ----

    /// Iterate over direct body items in source order.
    ///
    /// Unlike [`Self::paragraphs`] and [`Self::tables`], this retains the
    /// interleaving of paragraphs, tables, content controls, and preserved
    /// unmodelled XML.
    pub fn body_items(&self) -> impl Iterator<Item = BodyItemRef<'_>> {
        self.document.body.content.iter().map(|item| match item {
            BodyContent::Paragraph(paragraph) => {
                BodyItemRef::Paragraph(ParagraphRef { inner: paragraph })
            }
            BodyContent::Table(table) => BodyItemRef::Table(TableRef { inner: table }),
            BodyContent::ContentControl(control) => {
                BodyItemRef::ContentControl(ContentControlRef { inner: control })
            }
            BodyContent::RawXml(raw) => BodyItemRef::UnsupportedXml(raw),
        })
    }

    /// Iterate over paragraphs, tables, and unsupported content in body order.
    ///
    /// This compatibility view reports content controls as modeled unsupported
    /// facts without fabricating raw XML bytes.
    pub fn body_content(&self) -> impl Iterator<Item = BodyContentRef<'_>> {
        self.document.body.content.iter().map(|item| match item {
            BodyContent::Paragraph(paragraph) => {
                BodyContentRef::Paragraph(ParagraphRef { inner: paragraph })
            }
            BodyContent::Table(table) => BodyContentRef::Table(TableRef { inner: table }),
            BodyContent::ContentControl(control) => {
                BodyContentRef::UnsupportedXml(UnsupportedXmlRef::modeled(
                    "w:sdt",
                    WORD_NAMESPACE,
                    "sdt",
                    control.has_child_content(),
                ))
            }
            BodyContent::RawXml(raw) => BodyContentRef::UnsupportedXml(UnsupportedXmlRef::raw(
                raw,
                &self.body_namespace_bindings,
            )),
        })
    }

    /// Whether the document declares a page background that may affect its
    /// visible appearance.
    pub fn has_document_background(&self) -> bool {
        self.document.background_xml.is_some()
    }

    /// Whether any document section carries layout formatting.
    ///
    /// This includes the final body section and sections attached to paragraph
    /// properties. Callers that cannot preserve page layout can reject such
    /// documents without treating section XML as ordinary body content.
    pub fn has_section_layout_formatting(&self) -> bool {
        self.section_properties_in_order().any(section_has_layout)
    }

    /// Whether any document section retains XML that the semantic section
    /// reader does not model.
    pub fn has_unmodeled_section_properties(&self) -> bool {
        self.section_properties_in_order()
            .any(|properties| !properties.extra_xml.is_empty() || properties.change.is_some())
    }

    /// Get immutable references to all paragraphs.
    pub fn paragraphs(&self) -> Vec<ParagraphRef<'_>> {
        self.document
            .body
            .paragraphs()
            .map(|p| ParagraphRef { inner: p })
            .collect()
    }

    /// Get an immutable reference to a paragraph by index (among paragraphs only).
    pub fn paragraph(&self, index: usize) -> Option<ParagraphRef<'_>> {
        self.document
            .body
            .paragraphs()
            .nth(index)
            .map(|p| ParagraphRef { inner: p })
    }

    /// All footnotes as (id, plain text), in file order.
    ///
    /// Separator entries are excluded. They live in the same stream and are
    /// retained by the model so a round trip preserves them, but they are not
    /// notes and never were part of this listing.
    pub fn footnotes(&self) -> Vec<(i32, String)> {
        self.footnotes
            .footnotes
            .iter()
            .filter(|f| f.note_type == rdocx_oxml::footnotes::NoteType::Normal)
            .map(|f| {
                let text = f
                    .paragraphs
                    .iter()
                    .map(|p| p.text())
                    .collect::<Vec<_>>()
                    .join("\n");
                (f.id, text)
            })
            .collect()
    }

    /// Add a footnote with the given text; returns its id. Pair with
    /// `Paragraph::add_footnote_ref` to reference it from the body.
    pub fn add_footnote(&mut self, text: &str) -> i32 {
        let mut candidate = self.clone_for_staging();
        candidate
            .reserve_footnotes_bundle()
            .expect("an in-memory document can allocate a footnotes part");
        candidate.invalidate_layout();
        candidate.footnotes_dirty = true;
        use rdocx_oxml::footnotes::CT_Footnote;
        use rdocx_oxml::text::CT_P;
        let id = candidate
            .footnotes
            .footnotes
            .iter()
            .map(|f| f.id)
            .max()
            .unwrap_or(1)
            + 1;
        let mut p = CT_P::new();
        p.add_run(text);
        candidate.footnotes.footnotes.push(CT_Footnote {
            id,
            note_type: rdocx_oxml::footnotes::NoteType::Normal,
            paragraphs: vec![p],
        });
        self.commit_staged_mutation(candidate);
        id
    }

    /// Add a paragraph with the given text and return a mutable reference.
    pub fn add_paragraph(&mut self, text: &str) -> Paragraph<'_> {
        self.invalidate_layout();
        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        self.document.body.content.push(BodyContent::Paragraph(p));
        match self.document.body.content.last_mut().unwrap() {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Get the number of paragraphs.
    pub fn paragraph_count(&self) -> usize {
        self.document.body.paragraphs().count()
    }

    /// Return every valid modeled main-document revision in document order.
    pub fn revisions(&self) -> Vec<RevisionRef<'_>> {
        self.document
            .revisions()
            .into_iter()
            .map(|inner| RevisionRef { inner })
            .collect()
    }

    /// Return valid document-protection metadata recorded in the settings part.
    ///
    /// This reports author intent and password-verification metadata. It does
    /// not enforce an access-control boundary or verify a password.
    pub fn document_protection(&self) -> Option<&DocumentProtection> {
        self.settings.as_ref()?.document_protection()
    }

    /// Get the plain text of body paragraphs and table cells in document order.
    pub fn text(&self) -> String {
        let mut result = String::new();
        for content in &self.document.body.content {
            match content {
                BodyContent::Paragraph(paragraph) => {
                    result.push_str(&paragraph.text());
                    result.push('\n');
                }
                BodyContent::Table(table) => {
                    for row in &table.rows {
                        for cell in &row.cells {
                            for content in &cell.content {
                                if let CellContent::Paragraph(paragraph) = content {
                                    result.push_str(&paragraph.text());
                                    result.push('\t');
                                }
                            }
                        }
                        result.push('\n');
                    }
                }
                BodyContent::ContentControl(_) => {}
                BodyContent::RawXml(_) => {}
            }
        }
        result
    }

    /// Get a mutable reference to a paragraph by index (among paragraphs only).
    pub fn paragraph_mut(&mut self, index: usize) -> Option<Paragraph<'_>> {
        self.invalidate_layout();
        let mut remaining = index;
        nth_paragraph_in_body(&mut self.document.body.content, &mut remaining)
            .map(|inner| Paragraph { inner })
    }

    // ---- Table access ----

    /// Get immutable references to all tables.
    pub fn tables(&self) -> Vec<TableRef<'_>> {
        self.document
            .body
            .tables()
            .map(|t| TableRef { inner: t })
            .collect()
    }

    /// Get an immutable table by index among tables only.
    pub fn table(&self, index: usize) -> Option<TableRef<'_>> {
        self.document
            .body
            .tables()
            .nth(index)
            .map(|inner| TableRef { inner })
    }

    /// Get a mutable table by index among tables only.
    pub fn table_mut(&mut self, index: usize) -> Option<Table<'_>> {
        self.invalidate_layout();
        let mut remaining = index;
        nth_table_in_body(&mut self.document.body.content, &mut remaining)
            .map(|inner| Table { inner })
    }

    /// Add a table with the specified number of rows and columns.
    /// Returns a mutable reference for further configuration.
    pub fn add_table(&mut self, rows: usize, cols: usize) -> Table<'_> {
        self.invalidate_layout();
        use rdocx_oxml::table::{CT_Row, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc};
        use rdocx_oxml::units::Twips;

        // Default column width: divide 9360tw (6.5" printable at 1" margins) evenly.
        // A zero-column table has no grid to divide; clamp so this cannot divide by zero.
        let col_width = Twips(9360 / cols.max(1) as i32);

        let grid = CT_TblGrid {
            columns: (0..cols)
                .map(|_| CT_TblGridCol { width: col_width })
                .collect(),
            ..Default::default()
        };

        let mut tbl = CT_Tbl::new();
        tbl.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(col_width.0 * cols as i32)),
            ..Default::default()
        });
        tbl.grid = Some(grid);

        for _ in 0..rows {
            let mut row = CT_Row::new();
            for _ in 0..cols {
                row.cells.push(CT_Tc::new());
            }
            tbl.rows.push(row);
        }

        self.document.body.content.push(BodyContent::Table(tbl));
        match self.document.body.content.last_mut().unwrap() {
            BodyContent::Table(t) => Table { inner: t },
            _ => unreachable!(),
        }
    }

    /// Get the number of tables.
    pub fn table_count(&self) -> usize {
        self.document.body.tables().count()
    }

    // ---- Content insertion ----

    /// Get the number of body content elements (paragraphs + tables).
    pub fn content_count(&self) -> usize {
        self.document.body.content_count()
    }

    /// Insert a paragraph at the given body index.
    ///
    /// Returns a mutable `Paragraph` for further configuration.
    /// # Panics
    ///
    /// Panics if `index > content_count()`. (Unlike [`Self::insert_document`]
    /// and [`Self::insert_toc`], which clamp an out-of-range index to the end.)
    pub fn insert_paragraph(&mut self, index: usize, text: &str) -> Paragraph<'_> {
        self.invalidate_layout();
        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        self.document.body.insert_paragraph(index, p);
        match &mut self.document.body.content[index] {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Insert a table at the given body index.
    ///
    /// Returns a mutable `Table` for further configuration.
    /// A `cols` of 0 produces a table with no columns rather than panicking.
    ///
    /// # Panics
    ///
    /// Panics if `index > content_count()`. (Unlike [`Self::insert_document`]
    /// and [`Self::insert_toc`], which clamp an out-of-range index to the end.)
    pub fn insert_table(&mut self, index: usize, rows: usize, cols: usize) -> Table<'_> {
        self.invalidate_layout();
        use rdocx_oxml::table::{CT_Row, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc};
        use rdocx_oxml::units::Twips;

        let col_width = Twips(9360 / cols.max(1) as i32);
        let grid = CT_TblGrid {
            columns: (0..cols)
                .map(|_| CT_TblGridCol { width: col_width })
                .collect(),
            ..Default::default()
        };

        let mut tbl = CT_Tbl::new();
        tbl.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(col_width.0 * cols as i32)),
            ..Default::default()
        });
        tbl.grid = Some(grid);

        for _ in 0..rows {
            let mut row = CT_Row::new();
            for _ in 0..cols {
                row.cells.push(CT_Tc::new());
            }
            tbl.rows.push(row);
        }

        self.document.body.insert_table(index, tbl);
        match &mut self.document.body.content[index] {
            BodyContent::Table(t) => Table { inner: t },
            _ => unreachable!(),
        }
    }

    /// Find the body content index of the first paragraph containing the given text.
    pub fn find_content_index(&self, text: &str) -> Option<usize> {
        self.document.body.find_paragraph_index(text)
    }

    /// Remove the content at the given body index.
    ///
    /// Returns `true` if an element was removed, `false` if the index was out of bounds.
    pub fn remove_content(&mut self, index: usize) -> bool {
        self.invalidate_layout();
        self.document.body.remove(index).is_some()
    }

    // ---- Image support ----

    /// Add an inline image to the document.
    ///
    /// Embeds the image data (PNG, JPEG, etc.) into the package and adds a
    /// paragraph containing the image. Returns a mutable reference to the
    /// paragraph for further configuration.
    ///
    /// `width` and `height` specify the display size.
    pub fn add_picture(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
    ) -> Paragraph<'_> {
        self.try_add_picture(image_data, image_filename, width, height)
            .expect("an in-memory document cannot exhaust image identifiers");
        let Some(BodyContent::Paragraph(paragraph)) = self.document.body.content.last_mut() else {
            unreachable!("picture staging appends one paragraph")
        };
        Paragraph { inner: paragraph }
    }

    fn try_add_picture(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
    ) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        let (part_name, format) = candidate.reserve_image_part(image_data, image_filename)?;
        let owner = candidate.doc_part_name.clone();
        let rel_id = candidate
            .identifiers
            .reserve_relationship_id_checked(&owner)?;
        let drawing_id = candidate.identifiers.reserve_drawing_id()?;
        let rel_target = candidate.install_reserved_image_part(&part_name, image_data, format);
        candidate
            .package
            .get_or_create_part_rels(&owner)
            .add_with_id(&rel_id, rel_types::IMAGE, &rel_target);

        let mut inline = CT_Inline::new(&rel_id, width.to_emu(), height.to_emu());
        inline.doc_pr_id = drawing_id;

        let drawing = CT_Drawing::inline(inline);
        let run = CT_R {
            alt_drawings: Vec::new(),
            properties: None,
            content: vec![RunContent::Drawing(drawing)],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
        };

        let mut p = CT_P::new();
        p.runs.push(run);
        candidate
            .document
            .body
            .content
            .push(BodyContent::Paragraph(p));
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Add an editable inline chart to the document.
    ///
    /// The ChartML cache and embedded workbook are authored from the same
    /// validated data. Invalid dimensions or data leave the document unchanged.
    pub fn add_chart(
        &mut self,
        kind: ChartKind,
        width: Length,
        height: Length,
        data: &ChartData,
    ) -> Result<Paragraph<'_>> {
        self.add_chart_package(ChartPackageSource::Authored { kind, data }, width, height)?;
        let Some(BodyContent::Paragraph(paragraph)) = self.document.body.content.last_mut() else {
            unreachable!("chart package assembly appends one paragraph");
        };
        Ok(Paragraph { inner: paragraph })
    }

    /// Add an inline image at its native size using 72 DPI when none is declared.
    ///
    /// Returns an error without changing the document when the image dimensions
    /// cannot be determined.
    pub fn add_picture_auto(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
    ) -> Result<Paragraph<'_>> {
        let native_size = oxml_media::probe(image_data)
            .and_then(|info| info.native_size(72.0))
            .ok_or_else(|| Error::UnavailableImageDimensions {
                filename: image_filename.to_owned(),
            })?;

        self.try_add_picture(
            image_data,
            image_filename,
            Length::emu(native_size.width_emu),
            Length::emu(native_size.height_emu),
        )?;
        let Some(BodyContent::Paragraph(paragraph)) = self.document.body.content.last_mut() else {
            unreachable!("picture staging appends one paragraph")
        };
        Ok(Paragraph { inner: paragraph })
    }

    /// Add a full-page background image behind text.
    ///
    /// The image is placed at position (0,0) relative to the page with
    /// dimensions matching the page size from section properties.
    /// It is inserted at the beginning of the document body so it renders
    /// behind all other content.
    pub fn add_background_image(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
    ) -> Paragraph<'_> {
        let mut candidate = self.clone_for_staging();
        let rel_id = candidate.embed_image(image_data, image_filename);

        // Get page dimensions from section properties (default US Letter)
        let sect = candidate
            .document
            .body
            .sect_pr
            .as_ref()
            .cloned()
            .unwrap_or_else(CT_SectPr::default_letter);
        let page_width_emu = sect
            .page_width
            .unwrap_or(rdocx_oxml::units::Twips(12240))
            .to_emu()
            .0;
        let page_height_emu = sect
            .page_height
            .unwrap_or(rdocx_oxml::units::Twips(15840))
            .to_emu()
            .0;

        let mut anchor = CT_Anchor::background(&rel_id, page_width_emu, page_height_emu);
        anchor.doc_pr_id = candidate
            .identifiers
            .reserve_drawing_id()
            .expect("an in-memory document cannot exhaust every drawing identifier");
        let drawing = CT_Drawing::anchor(anchor);
        let run = CT_R {
            alt_drawings: Vec::new(),
            properties: None,
            content: vec![RunContent::Drawing(drawing)],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
        };

        let mut p = CT_P::new();
        p.runs.push(run);
        candidate.document.body.insert_paragraph(0, p);
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        match &mut self.document.body.content[0] {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Add an anchored (floating) image to the document.
    ///
    /// If `behind_text` is true, the image renders behind text content.
    /// The image is inserted at the beginning of the document body.
    pub fn add_anchored_image(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
        behind_text: bool,
    ) -> Paragraph<'_> {
        let mut candidate = self.clone_for_staging();
        let rel_id = candidate.embed_image(image_data, image_filename);

        let mut anchor = CT_Anchor::background(&rel_id, width.to_emu(), height.to_emu());
        anchor.doc_pr_id = candidate
            .identifiers
            .reserve_drawing_id()
            .expect("an in-memory document cannot exhaust every drawing identifier");
        anchor.behind_doc = behind_text;

        let drawing = CT_Drawing::anchor(anchor);
        let run = CT_R {
            alt_drawings: Vec::new(),
            properties: None,
            content: vec![RunContent::Drawing(drawing)],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
        };

        let mut p = CT_P::new();
        p.runs.push(run);
        candidate.document.body.insert_paragraph(0, p);
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        match &mut self.document.body.content[0] {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Store image bytes as a new media part and declare its content type.
    ///
    /// Returns the relationship target to use when referencing it, e.g.
    /// `media/image3.png`. No relationship is created here: an image referenced
    /// from a header or footer must be related to *that* part, not the
    /// document, so the caller decides where it is attached.
    fn store_image_part(&mut self, image_data: &[u8], filename: &str) -> String {
        let (part_name, format) = self
            .reserve_image_part(image_data, filename)
            .expect("an in-memory package cannot exhaust every media part suffix");
        self.install_reserved_image_part(&part_name, image_data, format)
    }

    fn reserve_image_part(
        &mut self,
        image_data: &[u8],
        filename: &str,
    ) -> Result<(String, oxml_media::ImageFormat)> {
        let format = oxml_media::resolve(image_data, filename);
        let part_name =
            self.identifiers
                .reserve_part_name("/word/media", "image", format.extension())?;
        Ok((part_name, format))
    }

    fn install_reserved_image_part(
        &mut self,
        part_name: &str,
        image_data: &[u8],
        format: oxml_media::ImageFormat,
    ) -> String {
        let extension = format.extension();

        self.package.set_part(part_name, image_data.to_vec());
        let content_type = format.content_type();
        match self.package.content_types.content_type_for(part_name) {
            Some(existing) if existing == content_type => {}
            Some(_) => self
                .package
                .content_types
                .add_override(part_name, content_type),
            None => {
                self.package
                    .content_types
                    .add_default(extension, content_type);
                self.identifiers.register_content_type_default(extension);
            }
        }
        if self.package.content_types.contains_override(part_name) {
            self.identifiers.register_content_type_override(part_name);
        }

        part_name
            .strip_prefix("/word/")
            .unwrap_or(part_name)
            .to_owned()
    }

    /// Embed an image into the OPC package and return the relationship ID.
    ///
    /// Public so callers can pre-embed an image and then pass the returned
    /// `rel_id` to [`crate::Cell::add_picture`] for inline cell images.
    pub fn embed_image(&mut self, image_data: &[u8], filename: &str) -> String {
        let mut candidate = self.clone_for_staging();
        let (part_name, format) = candidate
            .reserve_image_part(image_data, filename)
            .expect("an in-memory package cannot exhaust every media part suffix");
        let owner = candidate.doc_part_name.clone();
        let relationship_id = candidate
            .identifiers
            .reserve_relationship_id_checked(&owner)
            .expect("an in-memory package cannot exhaust every relationship identifier");
        let rel_target = candidate.install_reserved_image_part(&part_name, image_data, format);
        candidate
            .package
            .get_or_create_part_rels(&owner)
            .add_with_id(&relationship_id, rel_types::IMAGE, &rel_target);
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        relationship_id
    }

    /// Whether the given numbering definition renders as bullets (true)
    /// or numbers (false). None if the id is unknown.
    pub fn numbering_is_bullet(&self, num_id: u32) -> Option<bool> {
        let numbering = self.numbering.as_ref()?;
        let abstract_num = numbering.get_abstract_num_for(num_id)?;
        let fmt = abstract_num.levels.first()?.num_fmt.as_ref()?;
        match fmt {
            rdocx_oxml::numbering::ST_NumberFormat::Bullet => Some(true),
            rdocx_oxml::numbering::ST_NumberFormat::Other(_) => None,
            _ => Some(false),
        }
    }

    /// Resolve the public reader projection for one numbering level.
    ///
    /// `has_unmodeled_properties` reports retained XML or attributes attached
    /// to this numbering instance, definition, or level. Modeled producer
    /// metadata that this projection does not expose is not reported here.
    pub fn numbering_level(&self, num_id: u32, level: u32) -> Option<NumberingLevel<'_>> {
        if num_id == 0 {
            return None;
        }
        let numbering = self.numbering.as_ref()?;
        let instance = numbering.nums.iter().find(|item| item.num_id == num_id)?;
        let definition = numbering
            .abstract_nums
            .iter()
            .find(|item| item.abstract_num_id == instance.abstract_num_id)?;
        let level = definition.levels.iter().find(|item| item.ilvl == level)?;
        let format = level.num_fmt.as_ref();

        Some(NumberingLevel {
            level: level.ilvl,
            format: format
                .map(NumberingFormat::from_st)
                .unwrap_or(NumberingFormat::Decimal),
            format_name: format.map(ST_NumberFormat::to_str).unwrap_or("decimal"),
            start: level.start.unwrap_or(1),
            suffix: level
                .suffix
                .map(ListLevelSuffix::from_st)
                .unwrap_or(ListLevelSuffix::Tab),
            level_text: level.lvl_text.as_deref(),
            alignment: level.lvl_jc.map(list_level_alignment),
            paragraph_style: level.p_style.as_deref(),
            restart: level.restart.map(ListLevelRestart::from_st),
            legal_numbering: level.legal,
            indent_left: level
                .ppr
                .as_ref()
                .and_then(|properties| properties.ind_left),
            indent_hanging: level
                .ppr
                .as_ref()
                .and_then(|properties| properties.ind_hanging),
            indent_first_line: level
                .ppr
                .as_ref()
                .and_then(|properties| properties.ind_first_line),
            marker_properties: level.rpr.as_ref(),
            template_code: level.template_code.as_deref(),
            tentative: level.tentative,
            has_unmodeled_properties: !instance.extra_xml.is_empty()
                || !instance.extra_attributes.is_empty()
                || instance
                    .abstract_num_id_raw
                    .as_ref()
                    .is_some_and(|(_, raw, prefixes)| {
                        typed_numbering_leaf_has_unmodeled(raw, prefixes)
                    })
                || instance
                    .level_overrides
                    .iter()
                    .any(numbering_override_has_unmodeled)
                || !definition.extra_xml.is_empty()
                || !definition.extra_attributes.is_empty()
                || definition
                    .nsid_raw
                    .as_ref()
                    .is_some_and(|(_, raw, prefixes)| {
                        typed_numbering_leaf_has_unmodeled(raw, prefixes)
                    })
                || definition
                    .multi_level_type_raw
                    .as_ref()
                    .is_some_and(|(_, raw, prefixes)| {
                        typed_numbering_leaf_has_unmodeled(raw, prefixes)
                    })
                || definition
                    .tmpl_raw
                    .as_ref()
                    .is_some_and(|(_, raw, prefixes)| {
                        typed_numbering_leaf_has_unmodeled(raw, prefixes)
                    })
                || !level.extra_xml.is_empty()
                || !level.extra_attributes.is_empty()
                || level.start_raw.as_ref().is_some_and(|(_, raw, prefixes)| {
                    typed_numbering_leaf_has_unmodeled(raw, prefixes)
                })
                || level
                    .num_fmt_raw
                    .as_ref()
                    .is_some_and(|(_, raw, prefixes)| {
                        typed_numbering_leaf_has_unmodeled(raw, prefixes)
                    })
                || level
                    .p_style_raw
                    .as_ref()
                    .is_some_and(|(_, raw, prefixes)| {
                        typed_numbering_leaf_has_unmodeled(raw, prefixes)
                    })
                || level
                    .restart_raw
                    .as_ref()
                    .is_some_and(|(_, raw, prefixes)| {
                        typed_numbering_leaf_has_unmodeled(raw, prefixes)
                    })
                || level.legal_raw.as_ref().is_some_and(|(_, raw, prefixes)| {
                    typed_numbering_leaf_has_unmodeled(raw, prefixes)
                })
                || level.suffix_raw.as_ref().is_some_and(|(_, raw, prefixes)| {
                    typed_numbering_leaf_has_unmodeled(raw, prefixes)
                })
                || level
                    .lvl_text_raw
                    .as_ref()
                    .is_some_and(|(_, raw, prefixes)| {
                        typed_numbering_leaf_has_unmodeled(raw, prefixes)
                    })
                || level.lvl_jc_raw.as_ref().is_some_and(|(_, raw, prefixes)| {
                    typed_numbering_leaf_has_unmodeled(raw, prefixes)
                })
                || level.ppr_raw.is_some()
                || level.rpr_raw.is_some(),
            has_paragraph_presentation: level.ppr.as_ref().is_some_and(|properties| {
                Self::has_list_paragraph_presentation(level.ilvl, properties)
            }),
            has_marker_presentation: level
                .rpr
                .as_ref()
                .is_some_and(|properties| properties != &CT_RPr::default()),
        })
    }

    fn has_list_paragraph_presentation(level: u32, properties: &CT_PPr) -> bool {
        let Some(standard_indent) = (u64::from(level) + 1)
            .checked_mul(720)
            .and_then(|value| i32::try_from(value).ok())
        else {
            return true;
        };
        let standard = CT_PPr {
            ind_left: Some(rdocx_oxml::units::Twips(standard_indent)),
            ind_hanging: Some(rdocx_oxml::units::Twips(360)),
            ..CT_PPr::default()
        };
        properties != &standard
    }

    /// Append an external hyperlink to the last paragraph (creating one if
    /// the document is empty): adds the External relationship and wraps the
    /// new run in a hyperlink span.
    pub fn append_hyperlink(&mut self, text: &str, url: &str) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        let owner = candidate.doc_part_name.clone();
        let rel_id = candidate
            .add_external_relationship_checked(&owner, rel_types::HYPERLINK, url)
            .expect("an in-memory document can install a hyperlink relationship");

        if !matches!(
            candidate.document.body.content.last(),
            Some(BodyContent::Paragraph(_))
        ) {
            candidate
                .document
                .body
                .content
                .push(BodyContent::Paragraph(CT_P::new()));
        }
        let Some(BodyContent::Paragraph(p)) = candidate.document.body.content.last_mut() else {
            unreachable!();
        };
        crate::Paragraph { inner: p }.add_hyperlink(text, &rel_id);
        self.commit_staged_mutation(candidate);
    }

    /// Add an external hyperlink relationship and return its relationship ID.
    ///
    /// Use this with [`crate::Paragraph::add_hyperlink`] when the target
    /// paragraph is not the last body paragraph, such as a paragraph inside a
    /// table cell.
    pub fn add_hyperlink_relationship(&mut self, url: &str) -> String {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        let owner = candidate.doc_part_name.clone();
        let id = candidate
            .add_external_relationship_checked(&owner, rel_types::HYPERLINK, url)
            .expect("an in-memory document can install a hyperlink relationship");
        self.commit_staged_mutation(candidate);
        id
    }

    /// Get a builder for the last paragraph in the body, if any. Lets
    /// callers interleave plain runs with `append_hyperlink` calls.
    pub fn last_paragraph_mut(&mut self) -> Option<Paragraph<'_>> {
        self.invalidate_layout();
        match self.document.body.content.last_mut() {
            Some(BodyContent::Paragraph(p)) => Some(Paragraph { inner: p }),
            _ => None,
        }
    }

    /// Fetch the raw bytes of an embedded image by its relationship ID.
    pub fn image_data(&self, rel_id: &str) -> Option<Vec<u8>> {
        let rels = self.package.get_part_rels(&self.doc_part_name)?;
        let rel = rels.items.iter().find(|relationship| {
            relationship.id == rel_id
                && relationship.rel_type == rel_types::IMAGE
                && relationship_is_internal(relationship)
        })?;
        let target = OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
        self.package.get_part(&target).map(|b| b.to_vec())
    }

    /// Resolve a hyperlink relationship ID to its external URL.
    pub fn hyperlink_url(&self, rel_id: &str) -> Option<String> {
        use oxml_opc::relationship::rel_types;
        let rels = self.package.get_part_rels(&self.doc_part_name)?;
        rels.items
            .iter()
            .find(|relationship| {
                relationship.id == rel_id
                    && relationship.rel_type == rel_types::HYPERLINK
                    && relationship.target_mode.as_deref() == Some("External")
            })
            .map(|r| r.target.clone())
    }

    // ---- Header/Footer ----

    /// Set the default header text.
    ///
    /// Creates a header part with the given text and references it from
    /// the section properties.
    pub fn set_header(&mut self, text: &str) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        candidate
            .set_header_footer_part(text, true, HdrFtrType::Default)
            .expect("an in-memory document can install a header");
        self.commit_staged_mutation(candidate);
    }

    /// Set the default footer text.
    pub fn set_footer(&mut self, text: &str) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        candidate
            .set_header_footer_part(text, false, HdrFtrType::Default)
            .expect("an in-memory document can install a footer");
        self.commit_staged_mutation(candidate);
    }

    /// Set the first-page header text.
    pub fn set_first_page_header(&mut self, text: &str) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        candidate.set_different_first_page(true);
        candidate
            .set_header_footer_part(text, true, HdrFtrType::First)
            .expect("an in-memory document can install a first-page header");
        self.commit_staged_mutation(candidate);
    }

    /// Set the first-page footer text.
    pub fn set_first_page_footer(&mut self, text: &str) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        candidate.set_different_first_page(true);
        candidate
            .set_header_footer_part(text, false, HdrFtrType::First)
            .expect("an in-memory document can install a first-page footer");
        self.commit_staged_mutation(candidate);
    }

    /// Get the default header text, if set.
    pub fn header_text(&self) -> Option<String> {
        self.get_header_footer_text(true, HdrFtrType::Default)
    }

    /// Get the default footer text, if set.
    pub fn footer_text(&self) -> Option<String> {
        self.get_header_footer_text(false, HdrFtrType::Default)
    }

    /// Whether the document references any header or footer part.
    ///
    /// This is broader than [`Self::header_text`] and [`Self::footer_text`]:
    /// a referenced part can contain drawings, fields, tables, or other visible
    /// content without contributing literal text.
    pub fn has_header_footer_content(&self) -> bool {
        !self.header_footer_rel_ids().is_empty()
    }

    /// Set the default header to an inline image.
    ///
    /// Creates a header part with an image paragraph. The image is embedded
    /// in the header part's relationships.
    pub fn set_header_image(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
    ) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        candidate
            .set_header_footer_image_part(
                image_data,
                image_filename,
                width,
                height,
                true,
                HdrFtrType::Default,
            )
            .expect("an in-memory document can install a header image");
        self.commit_staged_mutation(candidate);
    }

    /// Set a Word-compatible text watermark in every active header variant.
    pub fn set_text_watermark(&mut self, text: &str) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.apply_watermark(|_, _, _| {
            Ok((
                VmlWatermark::Text {
                    text: text.to_owned(),
                    width_pt: 468.0,
                    height_pt: 117.0,
                    rotation_degrees: 315.0,
                    color: "D9D9D9".to_owned(),
                    font_family: Some("Calibri".to_owned()),
                    opacity: 0.5,
                },
                None,
            ))
        })?;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Set an image watermark in every active header variant.
    pub fn set_image_watermark(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
    ) -> Result<()> {
        if width.to_emu() <= 0 || height.to_emu() <= 0 {
            return Err(Error::Other(
                "watermark image width and height must be positive".to_owned(),
            ));
        }

        let mut candidate = self.clone_for_staging();
        let image_target = candidate.store_image_part(image_data, image_filename);
        let image_part_name =
            OpcPackage::resolve_rel_target(&candidate.doc_part_name, &image_target);
        candidate.apply_watermark(|package, identifiers, part_name| {
            let target = relative_target(part_name, &image_part_name);
            let relationship_id = identifiers
                .reserve_relationship_id_checked(part_name)
                .map_err(|error| {
                    Error::Other(format!(
                        "watermark relationship allocation failed for {part_name}: {error}"
                    ))
                })?;
            package.get_or_create_part_rels(part_name).add_with_id(
                &relationship_id,
                rel_types::IMAGE,
                &target,
            );
            Ok((
                VmlWatermark::Image {
                    relationship_id: relationship_id.clone(),
                    width_pt: width.to_pt(),
                    height_pt: height.to_pt(),
                    rotation_degrees: 0.0,
                    opacity: 0.5,
                },
                Some(relationship_id),
            ))
        })?;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    fn apply_watermark(
        &mut self,
        mut watermark_for_part: impl FnMut(
            &mut OpcPackage,
            &mut DocumentIdentifiers,
            &str,
        ) -> Result<(VmlWatermark, Option<String>)>,
    ) -> Result<()> {
        self.ensure_watermark_header_inheritance()?;
        let relationships = self.package.get_part_rels(&self.doc_part_name);
        let header_ids = self
            .header_footer_rel_ids()
            .into_iter()
            .filter_map(|(relationship_id, is_header)| {
                (is_header
                    && relationships.is_some_and(|relationships| {
                        relationships
                            .get_by_id(&relationship_id)
                            .is_some_and(|relationship| {
                                relationship.rel_type == rel_types::HEADER
                                    && relationship_is_internal(relationship)
                            })
                    }))
                .then_some(relationship_id)
            })
            .collect::<Vec<_>>();

        for relationship_id in header_ids {
            let target = self
                .package
                .get_part_rels(&self.doc_part_name)
                .and_then(|relationships| relationships.get_by_id(&relationship_id))
                .filter(|relationship| relationship_is_internal(relationship))
                .map(|relationship| relationship.target.clone())
                .ok_or_else(|| {
                    Error::Other(format!(
                        "active header relationship {relationship_id} is missing"
                    ))
                })?;
            let part_name = OpcPackage::resolve_rel_target(&self.doc_part_name, &target);
            let xml = self
                .package
                .get_part(&part_name)
                .map(<[u8]>::to_vec)
                .ok_or_else(|| {
                    Error::Other(format!("active header part {part_name} is missing"))
                })?;
            let fully_authored = self.identifiers.authored_story_parts.contains(&part_name);
            let (watermark, authored_relationship) =
                watermark_for_part(&mut self.package, &mut self.identifiers, &part_name)?;
            let updated = replace_authored_watermark(&xml, &watermark)?;
            let referenced = xml_relationship_ids_in_order(&updated)?
                .into_iter()
                .collect::<HashSet<_>>();
            self.prune_authored_story_relationships(&part_name, &referenced);
            self.package.set_part(&part_name, updated);
            if !fully_authored {
                let authored = self
                    .identifiers
                    .authored_story_relationship_ids
                    .entry(relationship_owner_identity(&part_name))
                    .or_default();
                authored.retain(|id| referenced.contains(id));
                if let Some(id) = authored_relationship {
                    authored.insert(id);
                }
                if authored.is_empty() {
                    self.identifiers
                        .authored_story_relationship_ids
                        .remove(&relationship_owner_identity(&part_name));
                }
            }
        }
        self.invalidate_layout();
        Ok(())
    }

    fn ensure_watermark_header_inheritance(&mut self) -> Result<()> {
        self.section_properties_mut();
        let even_enabled = self.even_headers_enabled();
        let internal_header_ids = self
            .package
            .get_part_rels(&self.doc_part_name)
            .map(|relationships| {
                relationships
                    .items
                    .iter()
                    .filter(|relationship| {
                        relationship.rel_type == rel_types::HEADER
                            && relationship_is_internal(relationship)
                    })
                    .map(|relationship| relationship.id.clone())
                    .collect::<HashSet<_>>()
            })
            .unwrap_or_default();
        let insertions = {
            let mut effective = [false; 3];
            let mut insertions = Vec::new();
            let mut inspect = |location: Option<usize>, section: &CT_SectPr| {
                for reference in &section.header_refs {
                    if internal_header_ids.contains(&reference.rel_id) {
                        effective[header_type_index(reference.hdr_ftr_type)] = true;
                    }
                }
                for hdr_type in [HdrFtrType::Default, HdrFtrType::First, HdrFtrType::Even] {
                    let active = match hdr_type {
                        HdrFtrType::Default => true,
                        HdrFtrType::First => section.title_pg.unwrap_or(false),
                        HdrFtrType::Even => even_enabled,
                    };
                    let index = header_type_index(hdr_type);
                    if active && !effective[index] {
                        insertions.push((location, hdr_type));
                        effective[index] = true;
                    }
                }
            };
            for (index, content) in self.document.body.content.iter().enumerate() {
                if let BodyContent::Paragraph(paragraph) = content
                    && let Some(section) = paragraph
                        .properties
                        .as_ref()
                        .and_then(|properties| properties.sect_pr.as_ref())
                {
                    inspect(Some(index), section);
                }
            }
            if let Some(section) = self.document.body.sect_pr.as_ref() {
                inspect(None, section);
            }
            insertions
        };

        for (location, hdr_type) in insertions {
            let relationship_id = self.create_watermark_header_relationship(hdr_type)?;
            let reference = HdrFtrRef {
                hdr_ftr_type: hdr_type,
                rel_id: relationship_id,
            };
            match location {
                Some(index) => {
                    let BodyContent::Paragraph(paragraph) = &mut self.document.body.content[index]
                    else {
                        unreachable!("recorded section owner changed")
                    };
                    paragraph
                        .properties
                        .as_mut()
                        .and_then(|properties| properties.sect_pr.as_mut())
                        .expect("recorded section disappeared")
                        .header_refs
                        .push(reference);
                }
                None => self
                    .document
                    .body
                    .sect_pr
                    .as_mut()
                    .expect("final section disappeared")
                    .header_refs
                    .push(reference),
            }
        }
        Ok(())
    }

    fn create_watermark_header_relationship(&mut self, hdr_type: HdrFtrType) -> Result<String> {
        let part_name = self.identifiers.reserve_part_name(
            "/word",
            &format!(
                "headerWatermark{}",
                match hdr_type {
                    HdrFtrType::Default => "Default",
                    HdrFtrType::First => "First",
                    HdrFtrType::Even => "Even",
                }
            ),
            "xml",
        )?;
        let empty_header = CT_HdrFtr::new().to_xml_header()?;
        self.package.set_part(&part_name, empty_header);
        self.package.content_types.add_override(
            &part_name,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        );
        self.identifiers.register_content_type_override(&part_name);
        self.identifiers
            .authored_story_parts
            .insert(part_name.clone());
        let target = relative_target(&self.doc_part_name, &part_name);
        let owner = self.doc_part_name.clone();
        let id = self
            .identifiers
            .reserve_relationship_id_checked(&owner)
            .map_err(|error| {
                Error::Other(format!(
                    "watermark header relationship allocation failed: {error}"
                ))
            })?;
        self.package
            .get_or_create_part_rels(&owner)
            .add_with_id(&id, rel_types::HEADER, &target);
        Ok(id)
    }

    /// Set the default footer to an inline image.
    pub fn set_footer_image(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
    ) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        candidate
            .set_header_footer_image_part(
                image_data,
                image_filename,
                width,
                height,
                false,
                HdrFtrType::Default,
            )
            .expect("an in-memory document can install a footer image");
        self.commit_staged_mutation(candidate);
    }

    /// Set a header from raw XML bytes with associated images.
    ///
    /// This is useful for copying complex headers from template documents
    /// that contain grouped shapes, VML, or other elements not easily
    /// recreated through the high-level API.
    ///
    /// Each entry in `images` is `(rel_id, image_data, image_filename)`:
    /// - `rel_id`: the relationship ID referenced in the header XML (e.g. "rId1")
    /// - `image_data`: the raw image bytes
    /// - `image_filename`: used to derive the part name and content type (e.g. "image5.png")
    pub fn set_raw_header_with_images(
        &mut self,
        header_xml: Vec<u8>,
        images: &[(&str, &[u8], &str)],
        hdr_type: HdrFtrType,
    ) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        candidate
            .set_raw_hdr_ftr_with_images(header_xml, images, true, hdr_type)
            .expect("raw header package preflight failed");
        self.commit_staged_mutation(candidate);
    }

    /// Set a footer from raw XML bytes with associated images.
    pub fn set_raw_footer_with_images(
        &mut self,
        footer_xml: Vec<u8>,
        images: &[(&str, &[u8], &str)],
        hdr_type: HdrFtrType,
    ) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        candidate
            .set_raw_hdr_ftr_with_images(footer_xml, images, false, hdr_type)
            .expect("raw footer package preflight failed");
        self.commit_staged_mutation(candidate);
    }

    /// Set the default header to an inline image with a colored background.
    ///
    /// Creates a header part where the paragraph has shading fill set to
    /// `bg_color` (hex string, e.g. "000000" for black) and contains the
    /// inline image.
    pub fn set_header_image_with_background(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
        bg_color: &str,
    ) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        candidate
            .set_header_footer_image_bg_part(
                image_data,
                image_filename,
                width,
                height,
                Some(bg_color),
                true,
                HdrFtrType::Default,
            )
            .expect("an in-memory document can install a header background image");
        self.commit_staged_mutation(candidate);
    }

    /// Set the first-page header to an inline image.
    pub fn set_first_page_header_image(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
    ) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        candidate.set_different_first_page(true);
        candidate
            .set_header_footer_image_part(
                image_data,
                image_filename,
                width,
                height,
                true,
                HdrFtrType::First,
            )
            .expect("an in-memory document can install a first-page header image");
        self.commit_staged_mutation(candidate);
    }

    /// Where a header/footer of this kind lives, and how to declare it.
    ///
    /// All four public entry points differ only in what goes *inside* the part;
    /// the surrounding bookkeeping — part name, content type, relationship,
    /// section reference — is identical, and lives here.
    ///
    /// Note the fixed `1` in the part name: rdocx manages one header and one
    /// footer per [`HdrFtrType`] for the document's single section. Setting a
    /// header of the same type again replaces the existing part.
    fn hdr_ftr_slot_metadata(
        is_header: bool,
        hdr_type: HdrFtrType,
    ) -> (String, &'static str, &'static str) {
        let type_suffix = match hdr_type {
            HdrFtrType::Default => "",
            HdrFtrType::First => "First",
            HdrFtrType::Even => "Even",
        };
        if is_header {
            (
                format!("/word/header{type_suffix}1.xml"),
                rel_types::HEADER,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            )
        } else {
            (
                format!("/word/footer{type_suffix}1.xml"),
                rel_types::FOOTER,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            )
        }
    }

    /// Install a header/footer part: store its bytes, declare the content type,
    /// relate it to the document, and point the section properties at it.
    ///
    /// Any previous reference of the same [`HdrFtrType`] is replaced.
    fn reserve_hdr_ftr_part_name(
        &mut self,
        is_header: bool,
        hdr_type: HdrFtrType,
    ) -> Result<String> {
        self.identifiers.observe_package_graph(&self.package)?;
        let (preferred, expected_rel_type, _) = Self::hdr_ftr_slot_metadata(is_header, hdr_type);
        let reference = self.document.body.sect_pr.as_ref().and_then(|section| {
            let references = if is_header {
                &section.header_refs
            } else {
                &section.footer_refs
            };
            references
                .iter()
                .find(|reference| reference.hdr_ftr_type == hdr_type)
        });
        if let Some(part_name) = reference
            .and_then(|reference| {
                self.package
                    .get_part_rels(&self.doc_part_name)
                    .and_then(|relationships| relationships.get_by_id(&reference.rel_id))
            })
            .filter(|relationship| {
                relationship.rel_type == expected_rel_type && relationship_is_internal(relationship)
            })
            .map(|relationship| {
                OpcPackage::resolve_rel_target(&self.doc_part_name, &relationship.target)
            })
        {
            return Ok(part_name);
        }
        self.identifiers.reserve_preferred_part_name(&preferred)
    }

    fn install_hdr_ftr_part(
        &mut self,
        part_name: String,
        xml: Vec<u8>,
        is_header: bool,
        hdr_type: HdrFtrType,
    ) -> Result<String> {
        let (_, rel_type, content_type) = Self::hdr_ftr_slot_metadata(is_header, hdr_type);

        // Setting the same header twice must not leave the first relationship
        // behind pointing at the same part.
        let rel_target = relative_target(&self.doc_part_name, &part_name);
        let existing = self
            .package
            .get_part_rels(&self.doc_part_name)
            .and_then(|rels| {
                rels.items
                    .iter()
                    .find(|relationship| {
                        relationship.rel_type == rel_type
                            && relationship_is_internal(relationship)
                            && OpcPackage::resolve_rel_target(
                                &self.doc_part_name,
                                &relationship.target,
                            ) == part_name
                    })
                    .map(|relationship| relationship.id.clone())
            });
        let pending_relationship = match existing {
            Some(existing) => (existing, false),
            None => (
                self.identifiers
                    .reserve_relationship_id_checked(&self.doc_part_name)
                    .map_err(|error| {
                        Error::Other(format!(
                            "header or footer relationship allocation failed: {error}"
                        ))
                    })?,
                true,
            ),
        };

        self.package.set_part(&part_name, xml);
        self.identifiers
            .authored_story_parts
            .insert(part_name.clone());
        self.package
            .content_types
            .add_override(&part_name, content_type);
        self.identifiers.register_content_type_override(&part_name);
        let (rel_id, install_relationship) = pending_relationship;
        if install_relationship {
            self.package
                .get_or_create_part_rels(&self.doc_part_name)
                .add_with_id(&rel_id, rel_type, &rel_target);
        }

        let sect = self.section_properties_mut();
        let refs = if is_header {
            &mut sect.header_refs
        } else {
            &mut sect.footer_refs
        };
        refs.retain(|r| r.hdr_ftr_type != hdr_type);
        refs.push(HdrFtrRef {
            hdr_ftr_type: hdr_type,
            rel_id,
        });

        Ok(part_name)
    }

    fn clear_authored_story_relationships(&mut self, part_name: &str) {
        self.prune_authored_story_relationships(part_name, &HashSet::new());
    }

    fn prune_authored_story_relationships(
        &mut self,
        part_name: &str,
        referenced: &HashSet<String>,
    ) {
        let preserved = self
            .identifiers
            .preserved_relationship_ids
            .get(&relationship_owner_identity(part_name))
            .cloned()
            .unwrap_or_default();
        let Some(relationships) = self.package.get_part_rels_mut(part_name) else {
            return;
        };
        let mut removed = Vec::new();
        relationships.items.retain(|relationship| {
            if preserved.contains(&relationship.id) || referenced.contains(&relationship.id) {
                true
            } else {
                removed.push(relationship.clone());
                false
            }
        });
        let relationships_are_empty = relationships.items.is_empty();
        if relationships_are_empty {
            self.package.remove_part_rels(part_name);
        }
        self.identifiers.retire_authored_story_relationships(
            part_name,
            removed.iter().map(|relationship| relationship.id.clone()),
        );
        let mut image_parts = removed
            .iter()
            .filter(|relationship| {
                relationship.rel_type == rel_types::IMAGE && relationship_is_internal(relationship)
            })
            .map(|relationship| OpcPackage::resolve_rel_target(part_name, &relationship.target))
            .collect::<HashSet<_>>();
        image_parts.retain(|candidate| {
            !self.package.part_rels.iter().any(|(owner, relationships)| {
                relationships.items.iter().any(|relationship| {
                    relationship_is_internal(relationship)
                        && OpcPackage::resolve_rel_target(owner, &relationship.target) == *candidate
                })
            })
        });
        for image_part in image_parts {
            if self.identifiers.part_is_preserved(&image_part) {
                continue;
            }
            self.package.remove_part(&image_part);
            self.package.remove_part_rels(&image_part);
            self.package.content_types.remove_override(&image_part);
            self.identifiers.retire_authored_part(&image_part);
            let Some((_, extension)) = image_part.rsplit_once('.') else {
                continue;
            };
            let default_is_used = self.package.parts.keys().any(|part_name| {
                !self.package.content_types.contains_override(part_name)
                    && part_name
                        .rsplit_once('.')
                        .is_some_and(|(_, candidate)| candidate == extension)
            });
            if !default_is_used
                && self
                    .identifiers
                    .authored_content_type_defaults
                    .remove(extension)
            {
                self.package.content_types.remove_default(extension);
                self.identifiers.content_type_defaults.remove(extension);
            }
        }
    }

    /// Serialize a header/footer body, choosing the right root element.
    fn serialize_hdr_ftr(hdr_ftr: &CT_HdrFtr, is_header: bool) -> Result<Vec<u8>> {
        let xml = if is_header {
            hdr_ftr.to_xml_header()
        } else {
            hdr_ftr.to_xml_footer()
        };
        Ok(xml?)
    }

    fn set_header_footer_part(
        &mut self,
        text: &str,
        is_header: bool,
        hdr_type: HdrFtrType,
    ) -> Result<()> {
        let mut hdr_ftr = CT_HdrFtr::new();
        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        hdr_ftr.paragraphs.push(p);

        let xml = Self::serialize_hdr_ftr(&hdr_ftr, is_header)?;
        let part_name = self.reserve_hdr_ftr_part_name(is_header, hdr_type)?;
        self.clear_authored_story_relationships(&part_name);
        self.install_hdr_ftr_part(part_name, xml, is_header, hdr_type)?;
        Ok(())
    }

    fn set_raw_hdr_ftr_with_images(
        &mut self,
        xml: Vec<u8>,
        images: &[(&str, &[u8], &str)],
        is_header: bool,
        hdr_type: HdrFtrType,
    ) -> Result<()> {
        let part_name = self.reserve_hdr_ftr_part_name(is_header, hdr_type)?;
        let mut requested = HashSet::new();
        for &(rel_id, _, _) in images {
            if !requested.insert(rel_id) {
                return Err(Error::Other(format!(
                    "duplicate supplied header or footer relationship id {rel_id}"
                )));
            }
        }
        self.clear_authored_story_relationships(&part_name);
        let mut relationship_remap = HashMap::new();
        for &(rel_id, _, _) in images {
            let allocated = self
                .identifiers
                .reserve_requested_relationship_id_checked(&part_name, rel_id)
                .map_err(|error| {
                    Error::Other(format!(
                        "header or footer image relationship allocation failed: {error}"
                    ))
                })?;
            if allocated != rel_id {
                relationship_remap.insert(rel_id.to_owned(), allocated);
            }
        }
        let xml = remap_xml_relationship_ids(&xml, &relationship_remap)?;
        let part_name = self.install_hdr_ftr_part(part_name, xml, is_header, hdr_type)?;

        for &(rel_id, image_data, image_filename) in images {
            let img_rel_target = self.store_image_part(image_data, image_filename);
            let allocated = relationship_remap
                .get(rel_id)
                .map_or(rel_id, String::as_str);
            self.package
                .get_or_create_part_rels(&part_name)
                .add_with_id(allocated, rel_types::IMAGE, &img_rel_target);
        }
        Ok(())
    }

    fn set_header_footer_image_part(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
        is_header: bool,
        hdr_type: HdrFtrType,
    ) -> Result<()> {
        self.set_header_footer_image_bg_part(
            image_data,
            image_filename,
            width,
            height,
            None,
            is_header,
            hdr_type,
        )
    }

    fn set_header_footer_image_bg_part(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
        bg_color: Option<&str>,
        is_header: bool,
        hdr_type: HdrFtrType,
    ) -> Result<()> {
        use rdocx_oxml::properties::CT_Shd;

        let part_name = self.reserve_hdr_ftr_part_name(is_header, hdr_type)?;
        self.clear_authored_story_relationships(&part_name);

        // The image relationship belongs to the header/footer part, not the
        // document, because that is where the drawing referencing it lives.
        let img_rel_target = self.store_image_part(image_data, image_filename);
        let img_rel_id = self
            .add_internal_relationship_checked(&part_name, rel_types::IMAGE, &img_rel_target)
            .map_err(|error| {
                Error::Other(format!(
                    "header or footer image relationship allocation failed: {error}"
                ))
            })?;

        let mut inline = CT_Inline::new(&img_rel_id, width.to_emu(), height.to_emu());
        inline.doc_pr_id = self.identifiers.reserve_drawing_id()?;
        let run = CT_R {
            alt_drawings: Vec::new(),
            properties: None,
            content: vec![RunContent::Drawing(CT_Drawing::inline(inline))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
        };

        let mut p = CT_P::new();
        p.runs.push(run);
        if let Some(color) = bg_color {
            p.properties = Some(CT_PPr {
                shading: Some(CT_Shd {
                    val: "clear".to_string(),
                    color: Some("auto".to_string()),
                    fill: Some(color.to_string()),
                }),
                ..Default::default()
            });
        }

        let mut hdr_ftr = CT_HdrFtr::new();
        hdr_ftr.paragraphs.push(p);

        let xml = Self::serialize_hdr_ftr(&hdr_ftr, is_header)?;
        self.install_hdr_ftr_part(part_name, xml, is_header, hdr_type)?;
        Ok(())
    }

    fn get_header_footer_text(&self, is_header: bool, hdr_type: HdrFtrType) -> Option<String> {
        let sect = self.document.body.sect_pr.as_ref()?;
        let refs = if is_header {
            &sect.header_refs
        } else {
            &sect.footer_refs
        };
        let rels = self.package.get_part_rels(&self.doc_part_name)?;
        let expected_type = if is_header {
            rel_types::HEADER
        } else {
            rel_types::FOOTER
        };
        let rel = refs
            .iter()
            .filter(|reference| reference.hdr_ftr_type == hdr_type)
            .find_map(|reference| {
                rels.get_by_id(&reference.rel_id).filter(|relationship| {
                    relationship.rel_type == expected_type && relationship_is_internal(relationship)
                })
            })?;
        let part_name = OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
        let xml = self.package.get_part(&part_name)?;
        let hdr_ftr = CT_HdrFtr::from_xml(xml).ok()?;
        Some(hdr_ftr.text())
    }

    // ---- Numbering/Lists ----

    /// Ensure a numbering part exists.
    ///
    /// The relationship and content-type override are added by
    /// [`Self::flush_to_package`], which knows the resolved part name and will
    /// not create a second numbering relationship if one already exists.
    fn ensure_numbering(&mut self) -> &mut CT_Numbering {
        self.numbering.get_or_insert_with(CT_Numbering::new)
    }

    fn reserve_numbering_bundle(&mut self) -> Result<()> {
        self.identifiers.observe_package_graph(&self.package)?;
        if let Some(numbering) = &self.numbering {
            self.identifiers.abstract_numbering_ids.extend(
                numbering
                    .abstract_nums
                    .iter()
                    .map(|definition| definition.abstract_num_id),
            );
            self.identifiers
                .numbering_instance_ids
                .extend(numbering.nums.iter().map(|instance| instance.num_id));
        }
        let part_name = match self.numbering_part_name.as_deref() {
            Some(part_name) => part_name.to_owned(),
            None => self
                .identifiers
                .reserve_preferred_part_name(DEFAULT_NUMBERING_PART)?,
        };
        let already_linked =
            self.package
                .get_part_rels(&self.doc_part_name)
                .is_some_and(|relationships| {
                    relationships.items.iter().any(|relationship| {
                        relationship.rel_type == rel_types::NUMBERING
                            && relationship_is_internal(relationship)
                            && OpcPackage::resolve_rel_target(
                                &self.doc_part_name,
                                &relationship.target,
                            ) == part_name
                    })
                });
        let pending_id = if already_linked {
            None
        } else {
            Some(
                self.identifiers
                    .reserve_bundle_relationship_id_checked(
                        &self.doc_part_name,
                        rel_types::NUMBERING,
                    )
                    .map_err(|error| {
                        Error::Other(format!("numbering relationship allocation failed: {error}"))
                    })?,
            )
        };
        self.package
            .content_types
            .add_override(&part_name, NUMBERING_CONTENT_TYPE);
        self.identifiers.register_content_type_override(&part_name);
        if let Some(id) = pending_id {
            let target = relative_target(&self.doc_part_name, &part_name);
            self.package
                .get_or_create_part_rels(&self.doc_part_name)
                .add_with_id(&id, rel_types::NUMBERING, &target);
        }
        self.numbering_part_name = Some(part_name);
        Ok(())
    }

    fn reserve_styles_bundle(&mut self) -> Result<()> {
        let existing = self.styles_part_name.clone();
        let part_name = self
            .reserve_document_part_bundle(
                existing.as_deref(),
                DEFAULT_STYLES_PART,
                rel_types::STYLES,
                STYLES_CONTENT_TYPE,
            )
            .map_err(|error| {
                Error::Other(format!("styles relationship allocation failed: {error}"))
            })?;
        self.styles_part_name = Some(part_name);
        Ok(())
    }

    fn reserve_footnotes_bundle(&mut self) -> Result<()> {
        let existing = self.footnotes_part_name.clone();
        let part_name = self
            .reserve_document_part_bundle(
                existing.as_deref(),
                "/word/footnotes.xml",
                rel_types::FOOTNOTES,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            )
            .map_err(|error| {
                Error::Other(format!("footnotes relationship allocation failed: {error}"))
            })?;
        self.footnotes_part_name = Some(part_name);
        Ok(())
    }

    /// Add a bullet list item at the given indentation level (0-based).
    ///
    /// If no bullet list definition exists yet, one is created automatically.
    /// Returns a mutable `Paragraph` for further configuration.
    pub fn add_bullet_list_item(&mut self, text: &str, level: u32) -> Paragraph<'_> {
        let mut candidate = self.clone_for_staging();
        candidate
            .reserve_numbering_bundle()
            .expect("an in-memory document can allocate a numbering part");
        candidate.invalidate_layout();
        // Find or create a bullet list numId
        let existing = candidate.numbering.as_ref().and_then(|numbering| {
            numbering.nums.iter().find(|n| {
                numbering
                    .get_abstract_num_for(n.num_id)
                    .map(|a| {
                        a.levels.first().and_then(|l| l.num_fmt.as_ref())
                            == Some(&rdocx_oxml::numbering::ST_NumberFormat::Bullet)
                    })
                    .unwrap_or(false)
            })
        });
        let num_id = if let Some(existing) = existing {
            existing.num_id
        } else {
            let (abstract_id, num_id) = candidate
                .identifiers
                .reserve_numbering_ids()
                .expect("an in-memory document cannot exhaust numbering identifiers");
            candidate.ensure_numbering().add_list_with_ids(
                &[(rdocx_oxml::numbering::ST_NumberFormat::Bullet, Some(1))],
                abstract_id,
                num_id,
            )
        };

        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        let ppr = CT_PPr {
            num_id: Some(num_id),
            num_ilvl: Some(level),
            ..Default::default()
        };
        p.properties = Some(ppr);

        candidate
            .document
            .body
            .content
            .push(BodyContent::Paragraph(p));
        self.commit_staged_mutation(candidate);
        match self.document.body.content.last_mut().unwrap() {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Add a numbered list item at the given indentation level (0-based).
    ///
    /// If no numbered list definition exists yet, one is created automatically.
    /// Returns a mutable `Paragraph` for further configuration.
    pub fn add_numbered_list_item(&mut self, text: &str, level: u32) -> Paragraph<'_> {
        let mut candidate = self.clone_for_staging();
        candidate
            .reserve_numbering_bundle()
            .expect("an in-memory document can allocate a numbering part");
        candidate.invalidate_layout();
        // Find or create a numbered list numId
        let existing = candidate.numbering.as_ref().and_then(|numbering| {
            numbering.nums.iter().find(|n| {
                numbering
                    .get_abstract_num_for(n.num_id)
                    .map(|a| {
                        a.levels.first().and_then(|l| l.num_fmt.as_ref())
                            == Some(&rdocx_oxml::numbering::ST_NumberFormat::Decimal)
                    })
                    .unwrap_or(false)
            })
        });
        let num_id = if let Some(existing) = existing {
            existing.num_id
        } else {
            let (abstract_id, num_id) = candidate
                .identifiers
                .reserve_numbering_ids()
                .expect("an in-memory document cannot exhaust numbering identifiers");
            candidate.ensure_numbering().add_list_with_ids(
                &[(rdocx_oxml::numbering::ST_NumberFormat::Decimal, Some(1))],
                abstract_id,
                num_id,
            )
        };

        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        let ppr = CT_PPr {
            num_id: Some(num_id),
            num_ilvl: Some(level),
            ..Default::default()
        };
        p.properties = Some(ppr);

        candidate
            .document
            .body
            .content
            .push(BodyContent::Paragraph(p));
        self.commit_staged_mutation(candidate);
        match self.document.body.content.last_mut().unwrap() {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Create a list definition with explicit per-level formats and return
    /// its numId.
    ///
    /// Unlike [`Self::add_bullet_list_item`] / [`Self::add_numbered_list_item`],
    /// which share one bullet and one numbered definition per document, every
    /// call creates a fresh definition — so separate lists restart their
    /// numbering, and one definition can mix formats across levels (e.g. a
    /// bullet list whose nested level is decimal). Attach paragraphs with
    /// [`crate::Paragraph::set_numbering`].
    ///
    /// `levels[i]` configures level `i`; deeper unspecified levels fall back
    /// to the standard template rotation for the last specified format's
    /// family. An empty slice produces the standard numbered template. Word
    /// supports nine levels, so entries after index eight are ignored.
    ///
    /// ```no_run
    /// use rdocx::{Document, ListLevel};
    ///
    /// let mut doc = Document::new();
    /// let num_id = doc.add_list_definition(&[
    ///     ListLevel::bullet(),
    ///     ListLevel::decimal().start(3),
    /// ]);
    /// doc.add_paragraph("first bullet").set_numbering(num_id, 0);
    /// doc.add_paragraph("third decimal").set_numbering(num_id, 1);
    /// ```
    pub fn add_list_definition(&mut self, levels: &[ListLevel]) -> u32 {
        if levels
            .iter()
            .take(9)
            .enumerate()
            .any(|(level, value)| validate_list_level(level as u32, value).is_err())
        {
            return 0;
        }
        let mut candidate = self.clone_for_staging();
        candidate
            .reserve_numbering_bundle()
            .expect("an in-memory document can allocate a numbering part");
        candidate.invalidate_layout();
        let formats: Vec<(ST_NumberFormat, Option<u32>)> = levels
            .iter()
            .take(9)
            .map(|level| (level.format.to_st(), level.start))
            .collect();
        let (abstract_id, num_id) = candidate
            .identifiers
            .reserve_numbering_ids()
            .expect("an in-memory document cannot exhaust numbering identifiers");
        let num_id = candidate
            .ensure_numbering()
            .add_list_with_ids(&formats, abstract_id, num_id);
        if let Some(definition) = candidate.numbering.as_mut().and_then(|numbering| {
            numbering
                .abstract_nums
                .iter_mut()
                .find(|definition| definition.abstract_num_id == abstract_id)
        }) {
            for (level, value) in levels.iter().take(9).enumerate() {
                let existing = &definition.levels[level];
                definition.levels[level] = merge_list_level(existing, value)
                    .expect("list level was validated before staged allocation");
            }
        }
        if candidate.validate_numbering_graph().is_err() {
            return 0;
        }
        self.commit_staged_mutation(candidate);
        num_id
    }

    /// Inspect all abstract numbering definitions in package order.
    pub fn numbering_definitions(&self) -> Vec<NumberingDefinition> {
        self.numbering
            .as_ref()
            .map(|numbering| {
                numbering
                    .abstract_nums
                    .iter()
                    .map(|definition| NumberingDefinition {
                        id: definition.abstract_num_id,
                        levels: definition
                            .levels
                            .iter()
                            .map(|level| NumberingDefinitionLevel {
                                level: level.ilvl,
                                properties: list_level_from_ct(level),
                            })
                            .collect(),
                        paragraph_style_links: definition
                            .levels
                            .iter()
                            .map(|level| level.p_style.clone())
                            .collect(),
                        has_unmodeled_properties: !definition.extra_xml.is_empty()
                            || !definition.extra_attributes.is_empty()
                            || definition
                                .nsid_raw
                                .as_ref()
                                .is_some_and(|(_, raw, prefixes)| {
                                    typed_numbering_leaf_has_unmodeled(raw, prefixes)
                                })
                            || definition.multi_level_type_raw.as_ref().is_some_and(
                                |(_, raw, prefixes)| {
                                    typed_numbering_leaf_has_unmodeled(raw, prefixes)
                                },
                            )
                            || definition
                                .tmpl_raw
                                .as_ref()
                                .is_some_and(|(_, raw, prefixes)| {
                                    typed_numbering_leaf_has_unmodeled(raw, prefixes)
                                })
                            || definition.levels.iter().any(|level| {
                                numbering_level_has_unmodeled(level)
                                    || matches!(
                                        level.num_fmt.as_ref(),
                                        Some(ST_NumberFormat::Other(_))
                                    )
                            }),
                    })
                    .collect()
            })
            .unwrap_or_default()
    }

    /// Inspect one abstract numbering definition by identifier.
    pub fn numbering_definition(&self, id: u32) -> Option<NumberingDefinition> {
        self.numbering_definitions()
            .into_iter()
            .find(|definition| definition.id == id)
    }

    /// Inspect all numbering instances in package order.
    pub fn numbering_instances(&self) -> Vec<NumberingInstance> {
        self.numbering
            .as_ref()
            .map(|numbering| {
                numbering
                    .nums
                    .iter()
                    .map(|instance| NumberingInstance {
                        id: instance.num_id,
                        definition_id: instance.abstract_num_id,
                        level_overrides: instance
                            .level_overrides
                            .iter()
                            .map(|value| NumberingLevelOverride {
                                level: value.ilvl,
                                start: value.start_override,
                                replacement: value.level.as_ref().map(list_level_from_ct),
                                paragraph_style_link: value
                                    .level
                                    .as_ref()
                                    .and_then(|level| level.p_style.clone()),
                                has_unmodeled_properties: numbering_override_has_unmodeled(value),
                            })
                            .collect(),
                        has_unmodeled_properties: !instance.extra_xml.is_empty()
                            || !instance.extra_attributes.is_empty()
                            || instance.abstract_num_id_raw.as_ref().is_some_and(
                                |(_, raw, prefixes)| {
                                    typed_numbering_leaf_has_unmodeled(raw, prefixes)
                                },
                            )
                            || instance
                                .level_overrides
                                .iter()
                                .any(numbering_override_has_unmodeled),
                    })
                    .collect()
            })
            .unwrap_or_default()
    }

    /// Inspect one numbering instance by identifier.
    pub fn numbering_instance(&self, id: u32) -> Option<NumberingInstance> {
        self.numbering_instances()
            .into_iter()
            .find(|instance| instance.id == id)
    }

    /// Create an abstract numbering definition and return its identifier.
    pub fn add_numbering_definition(&mut self, levels: &[ListLevel]) -> Result<u32> {
        validate_numbering_level_slice(levels)?;
        let authored = levels
            .iter()
            .enumerate()
            .map(|(level, value)| ct_level_from_list(level as u32, value))
            .collect::<Result<Vec<_>>>()?;
        let mut candidate = self.clone_for_staging();
        candidate.reserve_numbering_bundle()?;
        let id = candidate.identifiers.reserve_abstract_numbering_id()?;
        let mut definition = CT_AbstractNum::new(id);
        definition.multi_level_type = Some("hybridMultilevel".to_owned());
        definition.levels = authored;
        candidate.ensure_numbering().push_abstract_num(definition);
        candidate.validate_numbering_graph()?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(id)
    }

    /// Replace the modeled levels of an existing abstract definition.
    pub fn update_numbering_definition(
        &mut self,
        id: u32,
        levels: &[NumberingDefinitionLevel],
    ) -> Result<()> {
        validate_numbering_definition_levels(levels)?;
        let mut candidate = self.clone_for_staging();
        let numbering = candidate
            .numbering
            .as_mut()
            .ok_or_else(|| Error::Other(format!("numbering definition {id} does not exist")))?;
        let definition = numbering
            .abstract_nums
            .iter_mut()
            .find(|definition| definition.abstract_num_id == id)
            .ok_or_else(|| Error::Other(format!("numbering definition {id} does not exist")))?;
        if definition.levels.len() != levels.len() {
            return Err(Error::Other(format!(
                "numbering definition {id} updates must retain its {} levels",
                definition.levels.len()
            )));
        }
        definition.levels = definition
            .levels
            .iter()
            .map(|existing| {
                let value = levels
                    .iter()
                    .find(|value| value.level == existing.ilvl)
                    .ok_or_else(|| {
                        Error::Other(format!(
                            "numbering definition {id} update is missing level {}",
                            existing.ilvl
                        ))
                    })?;
                merge_list_level(existing, &value.properties)
            })
            .collect::<Result<Vec<_>>>()?;
        candidate.reserve_numbering_bundle()?;
        candidate.validate_numbering_graph()?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Remove an abstract definition that no instance references.
    pub fn remove_numbering_definition(&mut self, id: u32) -> Result<bool> {
        let Some(numbering) = self.numbering.as_ref() else {
            return Ok(false);
        };
        if !numbering
            .abstract_nums
            .iter()
            .any(|definition| definition.abstract_num_id == id)
        {
            return Ok(false);
        }
        if let Some(instance) = numbering
            .nums
            .iter()
            .find(|instance| instance.abstract_num_id == id)
        {
            return Err(Error::Other(format!(
                "numbering definition {id} is referenced by instance {}",
                instance.num_id
            )));
        }
        let mut candidate = self.clone_for_staging();
        candidate
            .numbering
            .as_mut()
            .expect("checked numbering part")
            .remove_abstract_num(id);
        candidate.validate_numbering_graph()?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(true)
    }

    /// Create a numbering instance for an existing definition.
    pub fn add_numbering_instance(
        &mut self,
        definition_id: u32,
        overrides: &[NumberingLevelOverride],
    ) -> Result<u32> {
        let mut candidate = self.clone_for_staging();
        candidate.reserve_numbering_bundle()?;
        let numbering = candidate.ensure_numbering();
        if !numbering
            .abstract_nums
            .iter()
            .any(|definition| definition.abstract_num_id == definition_id)
        {
            return Err(Error::Other(format!(
                "numbering definition {definition_id} does not exist"
            )));
        }
        let level_overrides = overrides
            .iter()
            .map(numbering_override_from_public)
            .collect::<Result<Vec<_>>>()?;
        let id = candidate.identifiers.reserve_numbering_instance_id()?;
        candidate.ensure_numbering().push_num(CT_Num {
            num_id: id,
            abstract_num_id: definition_id,
            abstract_num_id_raw: None,
            level_overrides,
            extra_xml: Vec::new(),
            extra_attributes: Vec::new(),
        });
        candidate.validate_numbering_graph()?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(id)
    }

    /// Replace the definition and overrides used by an existing instance.
    pub fn update_numbering_instance(
        &mut self,
        id: u32,
        definition_id: u32,
        overrides: &[NumberingLevelOverride],
    ) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        let numbering = candidate
            .numbering
            .as_mut()
            .ok_or_else(|| Error::Other(format!("numbering instance {id} does not exist")))?;
        if !numbering
            .abstract_nums
            .iter()
            .any(|definition| definition.abstract_num_id == definition_id)
        {
            return Err(Error::Other(format!(
                "numbering definition {definition_id} does not exist"
            )));
        }
        let instance = numbering
            .nums
            .iter_mut()
            .find(|instance| instance.num_id == id)
            .ok_or_else(|| Error::Other(format!("numbering instance {id} does not exist")))?;
        let mut updated = Vec::with_capacity(overrides.len());
        for value in overrides {
            let existing = instance
                .level_overrides
                .iter()
                .find(|existing| existing.ilvl == value.level);
            updated.push(numbering_override_for_update(existing, value)?);
        }
        instance.abstract_num_id = definition_id;
        instance.level_overrides = updated;
        candidate.reserve_numbering_bundle()?;
        candidate.validate_numbering_graph()?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Remove an unreferenced numbering instance.
    pub fn remove_numbering_instance(&mut self, id: u32) -> Result<bool> {
        let Some(numbering) = self.numbering.as_ref() else {
            return Ok(false);
        };
        if !numbering.nums.iter().any(|instance| instance.num_id == id) {
            return Ok(false);
        }
        if self.document_references_numbering_instance(id)? {
            return Err(Error::Other(format!(
                "numbering instance {id} is referenced by document content"
            )));
        }
        let mut candidate = self.clone_for_staging();
        candidate
            .numbering
            .as_mut()
            .expect("checked numbering part")
            .remove_num(id);
        candidate.validate_numbering_graph()?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(true)
    }

    /// Validate the complete typed numbering graph.
    pub fn validate_numbering_graph(&self) -> Result<()> {
        let numbering = self.numbering.as_ref();
        let Some(numbering) = numbering else {
            for style in &self.styles.styles {
                if style.style_type == StyleType::Paragraph
                    && style
                        .ppr
                        .as_ref()
                        .and_then(|properties| properties.num_id)
                        .is_some_and(|num_id| num_id != 0)
                {
                    return Err(Error::Other(format!(
                        "style '{}' references a missing numbering part",
                        style.style_id
                    )));
                }
            }
            let live_references = self.document_numbering_references()?;
            if let Some((num_id, level)) = live_references.first() {
                return Err(Error::Other(format!(
                    "document content references missing numbering instance {num_id} at level {level}"
                )));
            }
            return Ok(());
        };
        let mut definition_ids = HashSet::new();
        for definition in &numbering.abstract_nums {
            if !definition_ids.insert(definition.abstract_num_id) {
                return Err(Error::Other(format!(
                    "duplicate numbering definition id {}",
                    definition.abstract_num_id
                )));
            }
            let mut levels = HashSet::new();
            for level in &definition.levels {
                if level.ilvl > 8 || !levels.insert(level.ilvl) {
                    return Err(Error::Other(format!(
                        "numbering definition {} has invalid or duplicate level {}",
                        definition.abstract_num_id, level.ilvl
                    )));
                }
                validate_ct_numbering_level(level)?;
                if let Some(style_id) = level.p_style.as_deref()
                    && self.styles.get_by_id(style_id).is_none()
                {
                    return Err(Error::Other(format!(
                        "numbering definition {} references missing style '{style_id}'",
                        definition.abstract_num_id
                    )));
                }
            }
        }
        let mut instance_ids = HashSet::new();
        for instance in &numbering.nums {
            if instance.num_id == 0 || !instance_ids.insert(instance.num_id) {
                return Err(Error::Other(format!(
                    "invalid or duplicate numbering instance id {}",
                    instance.num_id
                )));
            }
            let definition = numbering
                .abstract_nums
                .iter()
                .find(|definition| definition.abstract_num_id == instance.abstract_num_id)
                .ok_or_else(|| {
                    Error::Other(format!(
                        "numbering instance {} references missing definition {}",
                        instance.num_id, instance.abstract_num_id
                    ))
                })?;
            let mut override_levels = HashSet::new();
            for level_override in &instance.level_overrides {
                if level_override.ilvl > 8 || !override_levels.insert(level_override.ilvl) {
                    return Err(Error::Other(format!(
                        "numbering instance {} has invalid or duplicate override level {}",
                        instance.num_id, level_override.ilvl
                    )));
                }
                if !definition
                    .levels
                    .iter()
                    .any(|level| level.ilvl == level_override.ilvl)
                {
                    return Err(Error::Other(format!(
                        "numbering instance {} override level {} is not owned by definition {}",
                        instance.num_id, level_override.ilvl, instance.abstract_num_id
                    )));
                }
                if let Some(level) = &level_override.level {
                    if level.ilvl != level_override.ilvl {
                        return Err(Error::Other(format!(
                            "numbering instance {} override level {} owns replacement level {}",
                            instance.num_id, level_override.ilvl, level.ilvl
                        )));
                    }
                    validate_ct_numbering_level(level)?;
                    if let Some(style_id) = level.p_style.as_deref()
                        && self.styles.get_by_id(style_id).is_none()
                    {
                        return Err(Error::Other(format!(
                            "numbering instance {} override references missing style '{style_id}'",
                            instance.num_id
                        )));
                    }
                }
            }
        }
        let mut style_links = HashMap::<&str, (u32, u32)>::new();
        for style in &self.styles.styles {
            if style.style_type != StyleType::Paragraph {
                continue;
            }
            let Some(properties) = style.ppr.as_ref() else {
                continue;
            };
            let Some(num_id) = properties.num_id else {
                if properties.num_ilvl.is_some() {
                    return Err(Error::Other(format!(
                        "style '{}' has a numbering level without an instance",
                        style.style_id
                    )));
                }
                continue;
            };
            if num_id == 0 {
                continue;
            }
            let level = properties.num_ilvl.unwrap_or(0);
            let instance = numbering
                .nums
                .iter()
                .find(|instance| instance.num_id == num_id)
                .ok_or_else(|| {
                    Error::Other(format!(
                        "style '{}' references missing numbering instance {num_id}",
                        style.style_id
                    ))
                })?;
            let definition = numbering
                .abstract_nums
                .iter()
                .find(|definition| definition.abstract_num_id == instance.abstract_num_id)
                .expect("numbering instances were validated above");
            let base = definition
                .levels
                .iter()
                .find(|candidate| candidate.ilvl == level)
                .ok_or_else(|| {
                    Error::Other(format!(
                        "style '{}' references numbering instance {num_id} at missing level {level}",
                        style.style_id
                    ))
                })?;
            let effective = instance
                .level_overrides
                .iter()
                .find(|value| value.ilvl == level)
                .and_then(|value| value.level.as_ref())
                .unwrap_or(base);
            if effective.p_style.as_deref() != Some(style.style_id.as_str()) {
                return Err(Error::Other(format!(
                    "style '{}' and numbering instance {num_id} level {level} are not reciprocally linked",
                    style.style_id
                )));
            }
            style_links.insert(style.style_id.as_str(), (num_id, level));
        }
        for definition in &numbering.abstract_nums {
            for level in &definition.levels {
                let Some(style_id) = level.p_style.as_deref() else {
                    continue;
                };
                let linked = style_links.get(style_id).is_some_and(|(num_id, ilvl)| {
                    *ilvl == level.ilvl
                        && numbering.nums.iter().any(|instance| {
                            instance.num_id == *num_id
                                && instance.abstract_num_id == definition.abstract_num_id
                        })
                });
                if !linked {
                    return Err(Error::Other(format!(
                        "numbering definition {} level {} and style '{style_id}' are not reciprocally linked",
                        definition.abstract_num_id, level.ilvl
                    )));
                }
            }
        }
        for instance in &numbering.nums {
            for level in instance
                .level_overrides
                .iter()
                .filter_map(|value| value.level.as_ref())
            {
                let Some(style_id) = level.p_style.as_deref() else {
                    continue;
                };
                if style_links.get(style_id) != Some(&(instance.num_id, level.ilvl)) {
                    return Err(Error::Other(format!(
                        "numbering instance {} replacement level {} and style '{style_id}' are not reciprocally linked",
                        instance.num_id, level.ilvl
                    )));
                }
            }
        }
        let live_references = self.document_numbering_references()?;
        for (num_id, level) in live_references {
            let instance = numbering
                .nums
                .iter()
                .find(|instance| instance.num_id == num_id)
                .ok_or_else(|| {
                    Error::Other(format!(
                        "document content references missing numbering instance {num_id} at level {level}"
                    ))
                })?;
            let definition = numbering
                .abstract_nums
                .iter()
                .find(|definition| definition.abstract_num_id == instance.abstract_num_id)
                .expect("numbering instances were validated above");
            if !definition
                .levels
                .iter()
                .any(|candidate| candidate.ilvl == level)
            {
                return Err(Error::Other(format!(
                    "document content references numbering instance {num_id} at missing level {level}"
                )));
            }
        }
        Ok(())
    }

    fn document_numbering_references(&self) -> Result<Vec<(u32, u32)>> {
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let mut references = Vec::new();
        for xml in candidate
            .word_story_part_names()
            .into_iter()
            .filter_map(|part_name| {
                candidate
                    .package
                    .get_part(&part_name)
                    .filter(|bytes| identifier_xml_is_well_formed(bytes))
            })
        {
            for direct in xml_paragraph_numbering_properties(xml)? {
                let style_id = direct.style_id.as_deref().or_else(|| {
                    candidate
                        .styles
                        .get_default(StyleType::Paragraph)
                        .map(|style| style.style_id.as_str())
                });
                let mut effective =
                    style::resolve_paragraph_properties(style_id, &candidate.styles);
                candidate.apply_effective_numbering(&mut effective, Some(&direct), style_id);
                if let Some(num_id) = effective.num_id
                    && num_id != 0
                {
                    references.push((num_id, effective.num_ilvl.unwrap_or(0)));
                }
            }
        }
        Ok(references)
    }

    fn word_story_part_names(&self) -> Vec<String> {
        let mut story_parts = vec![self.doc_part_name.clone()];
        if let Some(relationships) = self.package.get_part_rels(&self.doc_part_name) {
            story_parts.extend(
                relationships
                    .items
                    .iter()
                    .filter(|relationship| {
                        relationship_is_internal(relationship)
                            && matches!(
                                relationship.rel_type.as_str(),
                                rel_types::HEADER
                                    | rel_types::FOOTER
                                    | rel_types::FOOTNOTES
                                    | rel_types::ENDNOTES
                                    | rel_types::COMMENTS
                                    | rel_types::GLOSSARY_DOCUMENT
                            )
                    })
                    .map(|relationship| {
                        OpcPackage::resolve_rel_target(&self.doc_part_name, &relationship.target)
                    }),
            );
        }
        story_parts
    }

    fn document_references_numbering_instance(&self, id: u32) -> Result<bool> {
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let mut owners = candidate.word_story_part_names();
        owners.extend(candidate.styles_part_name.clone());
        for xml in owners.into_iter().filter_map(|part_name| {
            candidate
                .package
                .get_part(&part_name)
                .filter(|bytes| identifier_xml_is_well_formed(bytes))
        }) {
            if xml_references_numbering_instance(xml, id)? {
                return Ok(true);
            }
        }
        Ok(false)
    }

    /// Redefine one level (0–8) of an existing list definition, for callers
    /// that only learn a deeper level's format when content first reaches it.
    ///
    /// Returns `false` when `num_id` is unknown or `level` is out of range.
    pub fn set_list_level(&mut self, num_id: u32, level: u32, spec: ListLevel) -> bool {
        if validate_list_level(level, &spec).is_err() {
            return false;
        }
        let mut candidate = self.clone_for_staging();
        let updated = candidate.numbering.as_mut().is_some_and(|numbering| {
            let Some(abstract_id) = numbering
                .nums
                .iter()
                .find(|instance| instance.num_id == num_id)
                .map(|instance| instance.abstract_num_id)
            else {
                return false;
            };
            let Some(definition) = numbering
                .abstract_nums
                .iter_mut()
                .find(|definition| definition.abstract_num_id == abstract_id)
            else {
                return false;
            };
            let Some(index) = definition
                .levels
                .iter()
                .position(|existing| existing.ilvl == level)
            else {
                return false;
            };
            let Ok(updated) = merge_list_level(&definition.levels[index], &spec) else {
                return false;
            };
            definition.levels[index] = updated;
            true
        });
        if updated {
            candidate
                .reserve_numbering_bundle()
                .expect("an in-memory document can allocate a numbering part");
            if candidate.validate_numbering_graph().is_err() {
                return false;
            }
            candidate.invalidate_layout();
            self.commit_staged_mutation(candidate);
        }
        updated
    }

    /// Atomically associate a paragraph style with one concrete numbering level.
    ///
    /// Both the style's paragraph properties and the numbering level's
    /// paragraph-style link are published together. Existing links must either
    /// match this exact tuple or the operation fails without changing the
    /// document.
    pub fn link_style_to_numbering(
        &mut self,
        style_id: &str,
        num_id: u32,
        level: u32,
    ) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        style::validate_style_graph(&candidate.styles)?;
        candidate.validate_numbering_graph()?;
        candidate.reserve_styles_bundle()?;
        candidate.reserve_numbering_bundle()?;

        let style_index = candidate
            .styles
            .styles
            .iter()
            .position(|style| style.style_id == style_id)
            .ok_or_else(|| Error::Other(format!("style '{style_id}' does not exist")))?;
        if candidate.styles.styles[style_index].style_type != StyleType::Paragraph {
            return Err(Error::Other(format!(
                "style '{style_id}' is not a paragraph style"
            )));
        }

        let numbering = candidate
            .numbering
            .as_ref()
            .ok_or_else(|| Error::Other(format!("numbering instance {num_id} does not exist")))?;
        let instance_index = numbering
            .nums
            .iter()
            .position(|instance| instance.num_id == num_id)
            .ok_or_else(|| Error::Other(format!("numbering instance {num_id} does not exist")))?;
        let definition_id = numbering.nums[instance_index].abstract_num_id;
        let definition_index = numbering
            .abstract_nums
            .iter()
            .position(|definition| definition.abstract_num_id == definition_id)
            .expect("the numbering graph was validated");
        let override_index = numbering.nums[instance_index]
            .level_overrides
            .iter()
            .position(|value| value.ilvl == level && value.level.is_some());
        let definition_level_index = numbering.abstract_nums[definition_index]
            .levels
            .iter()
            .position(|value| value.ilvl == level)
            .ok_or_else(|| {
                Error::Other(format!("numbering instance {num_id} has no level {level}"))
            })?;

        let style_numbering = candidate.styles.styles[style_index]
            .ppr
            .as_ref()
            .map(|properties| (properties.num_id, properties.num_ilvl))
            .unwrap_or((None, None));
        if style_numbering != (None, None) && style_numbering != (Some(num_id), Some(level)) {
            return Err(Error::Other(format!(
                "style '{style_id}' already references a different numbering level"
            )));
        }

        for (candidate_definition_index, definition) in numbering.abstract_nums.iter().enumerate() {
            for (candidate_level_index, candidate_level) in definition.levels.iter().enumerate() {
                if candidate_level.p_style.as_deref() == Some(style_id)
                    && (candidate_definition_index, candidate_level_index)
                        != (definition_index, definition_level_index)
                {
                    return Err(Error::Other(format!(
                        "style '{style_id}' is already linked to another numbering level"
                    )));
                }
            }
        }
        for (candidate_instance_index, instance) in numbering.nums.iter().enumerate() {
            for (candidate_override_index, value) in instance.level_overrides.iter().enumerate() {
                if value
                    .level
                    .as_ref()
                    .and_then(|level| level.p_style.as_deref())
                    == Some(style_id)
                    && Some(candidate_override_index)
                        != override_index.filter(|_| candidate_instance_index == instance_index)
                {
                    return Err(Error::Other(format!(
                        "style '{style_id}' is already linked to another numbering level"
                    )));
                }
            }
        }

        let target_style = override_index
            .and_then(|index| {
                numbering.nums[instance_index].level_overrides[index]
                    .level
                    .as_ref()
            })
            .unwrap_or(&numbering.abstract_nums[definition_index].levels[definition_level_index])
            .p_style
            .as_deref();
        if target_style.is_some() && target_style != Some(style_id) {
            return Err(Error::Other(format!(
                "numbering instance {num_id} level {level} is already linked to another style"
            )));
        }

        let properties = candidate.styles.styles[style_index]
            .ppr
            .get_or_insert_with(CT_PPr::default);
        properties.num_id = Some(num_id);
        properties.num_ilvl = Some(level);
        let numbering = candidate
            .numbering
            .as_mut()
            .expect("the numbering part was resolved above");
        if let Some(index) = override_index {
            numbering.nums[instance_index].level_overrides[index]
                .level
                .as_mut()
                .expect("the override was selected only when it owns a level")
                .p_style = Some(style_id.to_owned());
        } else {
            numbering.abstract_nums[definition_index].levels[definition_level_index].p_style =
                Some(style_id.to_owned());
        }

        style::validate_style_graph(&candidate.styles)?;
        candidate.validate_numbering_graph()?;
        candidate.flush_to_package()?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Atomically remove one exact paragraph-style numbering association.
    pub fn unlink_style_from_numbering(
        &mut self,
        style_id: &str,
        num_id: u32,
        level: u32,
    ) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        style::validate_style_graph(&candidate.styles)?;
        candidate.validate_numbering_graph()?;

        let style_index = candidate
            .styles
            .styles
            .iter()
            .position(|style| style.style_id == style_id)
            .ok_or_else(|| Error::Other(format!("style '{style_id}' does not exist")))?;
        if candidate.styles.styles[style_index].style_type != StyleType::Paragraph {
            return Err(Error::Other(format!(
                "style '{style_id}' is not a paragraph style"
            )));
        }
        let numbering = candidate
            .numbering
            .as_ref()
            .ok_or_else(|| Error::Other(format!("numbering instance {num_id} does not exist")))?;
        let instance_index = numbering
            .nums
            .iter()
            .position(|instance| instance.num_id == num_id)
            .ok_or_else(|| Error::Other(format!("numbering instance {num_id} does not exist")))?;
        let definition_id = numbering.nums[instance_index].abstract_num_id;
        let definition_index = numbering
            .abstract_nums
            .iter()
            .position(|definition| definition.abstract_num_id == definition_id)
            .expect("the numbering graph was validated");
        let override_index = numbering.nums[instance_index]
            .level_overrides
            .iter()
            .position(|value| value.ilvl == level && value.level.is_some());
        let definition_level_index = numbering.abstract_nums[definition_index]
            .levels
            .iter()
            .position(|value| value.ilvl == level)
            .ok_or_else(|| {
                Error::Other(format!("numbering instance {num_id} has no level {level}"))
            })?;

        let style_numbering = candidate.styles.styles[style_index]
            .ppr
            .as_ref()
            .map(|properties| (properties.num_id, properties.num_ilvl))
            .unwrap_or((None, None));
        let target_style = override_index
            .and_then(|index| {
                numbering.nums[instance_index].level_overrides[index]
                    .level
                    .as_ref()
            })
            .unwrap_or(&numbering.abstract_nums[definition_index].levels[definition_level_index])
            .p_style
            .as_deref();
        if style_numbering == (None, None) && target_style.is_none() {
            return Ok(());
        }
        if style_numbering != (Some(num_id), Some(level)) || target_style != Some(style_id) {
            return Err(Error::Other(format!(
                "style '{style_id}' and numbering instance {num_id} level {level} are not linked"
            )));
        }

        candidate.reserve_styles_bundle()?;
        candidate.reserve_numbering_bundle()?;
        let properties = candidate.styles.styles[style_index]
            .ppr
            .as_mut()
            .expect("the exact style link was validated above");
        properties.num_id = None;
        properties.num_ilvl = None;
        let numbering = candidate
            .numbering
            .as_mut()
            .expect("the numbering part was resolved above");
        if let Some(index) = override_index {
            numbering.nums[instance_index].level_overrides[index]
                .level
                .as_mut()
                .expect("the override was selected only when it owns a level")
                .p_style = None;
        } else {
            numbering.abstract_nums[definition_index].levels[definition_level_index].p_style = None;
        }

        style::validate_style_graph(&candidate.styles)?;
        candidate.validate_numbering_graph()?;
        candidate.flush_to_package()?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    // ---- Style access ----

    /// Get all styles.
    pub fn styles(&self) -> Vec<Style<'_>> {
        self.styles
            .styles
            .iter()
            .map(|s| Style { inner: s })
            .collect()
    }

    /// Find a style by its ID.
    pub fn style(&self, style_id: &str) -> Option<Style<'_>> {
        self.styles.get_by_id(style_id).map(|s| Style { inner: s })
    }

    // ---- Style manipulation ----

    /// Add a custom style after validating the complete style graph.
    pub fn add_style(&mut self, builder: StyleBuilder) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.reserve_styles_bundle()?;
        let (authored, _) = builder.build();
        if candidate.styles.get_by_id(&authored.style_id).is_some() {
            return Err(Error::Other(format!(
                "style '{}' already exists",
                authored.style_id
            )));
        }
        if authored
            .ppr
            .as_ref()
            .is_some_and(|properties| properties.num_id.is_some() || properties.num_ilvl.is_some())
        {
            return Err(Error::Other(format!(
                "style '{}' numbering must be established with link_style_to_numbering",
                authored.style_id
            )));
        }
        let linked_style = authored.linked_style.clone();
        let style_id = authored.style_id.clone();
        candidate.styles.styles.push(authored);
        update_reciprocal_style_link(
            &mut candidate.styles,
            &style_id,
            None,
            linked_style.as_deref(),
        );
        style::validate_style_graph(&candidate.styles)?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Replace an existing style after validating the complete style graph.
    pub fn set_style(&mut self, builder: StyleBuilder) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.reserve_styles_bundle()?;
        let (mut authored, cleared) = builder.build();
        let index = candidate
            .styles
            .styles
            .iter()
            .position(|style| style.style_id == authored.style_id)
            .ok_or_else(|| Error::Other(format!("style '{}' does not exist", authored.style_id)))?;
        let existing = &candidate.styles.styles[index];
        if existing.style_type != authored.style_type {
            return Err(Error::Other(format!(
                "style '{}' has type '{}', not '{}'",
                authored.style_id,
                existing.style_type.to_str(),
                authored.style_type.to_str()
            )));
        }
        let old_link = existing.linked_style.clone();
        merge_style_update(existing, &mut authored, cleared);
        let existing_numbering = existing
            .ppr
            .as_ref()
            .map(|properties| (properties.num_id, properties.num_ilvl))
            .unwrap_or((None, None));
        let authored_numbering = authored
            .ppr
            .as_ref()
            .map(|properties| (properties.num_id, properties.num_ilvl))
            .unwrap_or((None, None));
        if authored_numbering != existing_numbering {
            return Err(Error::Other(format!(
                "style '{}' numbering must be changed with link_style_to_numbering or unlink_style_from_numbering",
                authored.style_id
            )));
        }
        let style_id = authored.style_id.clone();
        let new_link = authored.linked_style.clone();
        candidate.styles.styles[index] = authored;
        update_reciprocal_style_link(
            &mut candidate.styles,
            &style_id,
            old_link.as_deref(),
            new_link.as_deref(),
        );
        style::validate_style_graph(&candidate.styles)?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Select the sole default style for one style type.
    pub fn set_default_style(&mut self, style_type: StyleType, style_id: &str) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.reserve_styles_bundle()?;
        let target = candidate
            .styles
            .get_by_id(style_id)
            .ok_or_else(|| Error::Other(format!("style '{style_id}' does not exist")))?;
        if target.style_type != style_type {
            return Err(Error::Other(format!(
                "style '{style_id}' has type '{}', not '{}'",
                target.style_type.to_str(),
                style_type.to_str()
            )));
        }
        for style in &mut candidate.styles.styles {
            if style.style_type == style_type {
                style.is_default = style.style_id == style_id;
            }
        }
        style::validate_style_graph(&candidate.styles)?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Remove an unreferenced style, returning whether it existed.
    pub fn remove_style(&mut self, style_id: &str) -> Result<bool> {
        let Some(index) = self
            .styles
            .styles
            .iter()
            .position(|style| style.style_id == style_id)
        else {
            return Ok(false);
        };
        if let Some(owner) = self.styles.styles.iter().find(|style| {
            style.style_id != style_id
                && (style.based_on.as_deref() == Some(style_id)
                    || style.next_style.as_deref() == Some(style_id)
                    || style.linked_style.as_deref() == Some(style_id))
        }) {
            return Err(Error::Other(format!(
                "style '{style_id}' is referenced by style '{}'",
                owner.style_id
            )));
        }
        if self.document_references_style(style_id)? {
            return Err(Error::Other(format!(
                "style '{style_id}' is referenced by document content"
            )));
        }
        if self.numbering.as_ref().is_some_and(|numbering| {
            numbering.abstract_nums.iter().any(|definition| {
                definition
                    .levels
                    .iter()
                    .any(|level| level.p_style.as_deref() == Some(style_id))
            }) || numbering.nums.iter().any(|instance| {
                instance.level_overrides.iter().any(|value| {
                    value
                        .level
                        .as_ref()
                        .is_some_and(|level| level.p_style.as_deref() == Some(style_id))
                })
            })
        }) {
            return Err(Error::Other(format!(
                "style '{style_id}' is referenced by a numbering level"
            )));
        }

        let mut candidate = self.clone_for_staging();
        candidate.reserve_styles_bundle()?;
        candidate.styles.styles.remove(index);
        style::validate_style_graph(&candidate.styles)?;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(true)
    }

    /// Validate all style IDs, defaults, links, next styles, and inheritance chains.
    pub fn validate_style_graph(&self) -> Result<()> {
        style::validate_style_graph(&self.styles)
    }

    fn document_references_style(&self, style_id: &str) -> Result<bool> {
        if body_references_style(&self.document.body.content, style_id)
            || xml_references_style(&self.document.to_xml()?, style_id)?
            || self
                .footnotes
                .footnotes
                .iter()
                .flat_map(|note| &note.paragraphs)
                .any(|paragraph| paragraph_references_style(paragraph, style_id))
            || self.comments.as_ref().is_some_and(|comments| {
                comments
                    .comments
                    .iter()
                    .flat_map(|comment| &comment.paragraphs)
                    .any(|paragraph| paragraph_references_style(paragraph, style_id))
            })
            || self.glossary.as_ref().is_some_and(|glossary| {
                glossary
                    .doc_parts
                    .iter()
                    .any(|part| body_references_style(&part.body.content, style_id))
            })
        {
            return Ok(true);
        }

        let Some(relationships) = self.package.get_part_rels(&self.doc_part_name) else {
            return Ok(false);
        };
        for relationship in &relationships.items {
            if relationship_is_internal(relationship)
                && matches!(
                    relationship.rel_type.as_str(),
                    rel_types::HEADER
                        | rel_types::FOOTER
                        | rel_types::FOOTNOTES
                        | rel_types::ENDNOTES
                        | rel_types::COMMENTS
                        | rel_types::GLOSSARY_DOCUMENT
                )
            {
                let part_name =
                    OpcPackage::resolve_rel_target(&self.doc_part_name, &relationship.target);
                if let Some(xml) = self.package.get_part(&part_name)
                    && xml_references_style(xml, style_id)?
                {
                    return Ok(true);
                }
            }
        }
        Ok(false)
    }

    /// Resolve the effective paragraph properties for a given style ID,
    /// walking the full inheritance chain (docDefaults → basedOn → ...).
    pub fn resolve_paragraph_properties(&self, style_id: Option<&str>) -> CT_PPr {
        style::resolve_paragraph_properties(style_id, &self.styles)
    }

    /// Resolve inherited, numbering-level, and direct properties for a
    /// concrete paragraph.
    pub fn effective_paragraph_properties(&self, paragraph: &ParagraphRef<'_>) -> CT_PPr {
        let direct = paragraph.inner.properties.as_ref();
        let style_id = direct
            .and_then(|properties| properties.style_id.as_deref())
            .or_else(|| {
                self.styles
                    .get_default(StyleType::Paragraph)
                    .map(|style| style.style_id.as_str())
            });
        let mut effective = style::resolve_paragraph_properties(style_id, &self.styles);
        self.apply_effective_numbering(&mut effective, direct, style_id);

        if let Some((num_id, level)) = effective.num_id.zip(effective.num_ilvl)
            && let Some(definition) = self.resolved_numbering_level_definition(num_id, level)
            && let Some(properties) = &definition.ppr
        {
            effective.merge_from(properties);
        }
        if let Some(properties) = direct {
            effective.merge_from(properties);
        }
        effective
    }

    fn apply_effective_numbering(
        &self,
        effective: &mut CT_PPr,
        direct: Option<&CT_PPr>,
        style_id: Option<&str>,
    ) {
        effective.num_id = direct
            .and_then(|properties| properties.num_id)
            .or(effective.num_id);
        effective.num_ilvl = direct
            .and_then(|properties| properties.num_ilvl)
            .or(effective.num_ilvl);
        if let Some(num_id) = effective.num_id {
            effective.num_ilvl = effective
                .num_ilvl
                .or_else(|| self.numbering_level_for_style(num_id, style_id));
        }
    }

    /// Resolve the effective run properties for the given paragraph and character styles,
    /// walking the full inheritance chain.
    pub fn resolve_run_properties(
        &self,
        para_style_id: Option<&str>,
        run_style_id: Option<&str>,
    ) -> CT_RPr {
        style::resolve_run_properties(para_style_id, run_style_id, &self.styles)
    }

    /// Resolve inherited, paragraph-mark, and direct properties for one
    /// concrete body run.
    pub fn effective_run_properties(
        &self,
        paragraph: &ParagraphRef<'_>,
        run: &RunRef<'_>,
    ) -> CT_RPr {
        let paragraph_properties = paragraph.inner.properties.as_ref();
        let direct = run.inner.properties.as_ref();
        let mut effective = style::resolve_run_properties(
            paragraph_properties.and_then(|properties| properties.style_id.as_deref()),
            direct.and_then(|properties| properties.style_id.as_deref()),
            &self.styles,
        );
        if let Some(properties) =
            paragraph_properties.and_then(|properties| properties.rpr.as_ref())
        {
            effective.merge_from(properties);
        }
        if let Some(properties) = direct {
            effective.merge_from(properties);
        }
        effective
    }

    fn resolved_numbering_level_definition(
        &self,
        num_id: u32,
        level: u32,
    ) -> Option<&rdocx_oxml::numbering::CT_Lvl> {
        if num_id == 0 {
            return None;
        }
        self.numbering
            .as_ref()?
            .get_abstract_num_for(num_id)?
            .levels
            .iter()
            .find(|definition| definition.ilvl == level)
    }

    fn numbering_level_for_style(&self, num_id: u32, style_id: Option<&str>) -> Option<u32> {
        let definition = self.numbering.as_ref()?.get_abstract_num_for(num_id)?;
        let mut style_id = style_id?;
        for _ in 0..self.styles.styles.len() {
            if let Some(level) = definition
                .levels
                .iter()
                .find(|level| level.p_style.as_deref() == Some(style_id))
            {
                return Some(level.ilvl);
            }
            style_id = self.styles.get_by_id(style_id)?.based_on.as_deref()?;
        }
        None
    }

    fn section_properties_in_order(&self) -> impl Iterator<Item = &CT_SectPr> {
        self.document
            .body
            .content
            .iter()
            .filter_map(|content| match content {
                BodyContent::Paragraph(paragraph) => paragraph
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.sect_pr.as_ref()),
                BodyContent::Table(_) | BodyContent::ContentControl(_) | BodyContent::RawXml(_) => {
                    None
                }
            })
            .chain(self.document.body.sect_pr.iter())
    }

    // ---- Section/Page setup ----

    /// Get the section properties (page size, margins).
    pub fn section_properties(&self) -> Option<&CT_SectPr> {
        self.document.body.sect_pr.as_ref()
    }

    /// Get a mutable reference to section properties, creating defaults if needed.
    pub fn section_properties_mut(&mut self) -> &mut CT_SectPr {
        self.invalidate_layout();
        self.document
            .body
            .sect_pr
            .get_or_insert_with(CT_SectPr::default_letter)
    }

    /// Set page size.
    pub fn set_page_size(&mut self, width: Length, height: Length) {
        let sect = self.section_properties_mut();
        sect.page_width = Some(width.as_twips());
        sect.page_height = Some(height.as_twips());
    }

    /// Set page orientation to landscape (swaps width and height if needed).
    pub fn set_landscape(&mut self) {
        let sect = self.section_properties_mut();
        sect.orientation = Some(ST_PageOrientation::Landscape);
        // Swap width/height if portrait dimensions
        if let (Some(w), Some(h)) = (sect.page_width, sect.page_height)
            && w.0 < h.0
        {
            sect.page_width = Some(h);
            sect.page_height = Some(w);
        }
    }

    /// Set page orientation to portrait (swaps width and height if needed).
    pub fn set_portrait(&mut self) {
        let sect = self.section_properties_mut();
        sect.orientation = Some(ST_PageOrientation::Portrait);
        // Swap width/height if landscape dimensions
        if let (Some(w), Some(h)) = (sect.page_width, sect.page_height)
            && w.0 > h.0
        {
            sect.page_width = Some(h);
            sect.page_height = Some(w);
        }
    }

    /// Set all page margins.
    pub fn set_margins(&mut self, top: Length, right: Length, bottom: Length, left: Length) {
        let sect = self.section_properties_mut();
        sect.margin_top = Some(top.as_twips());
        sect.margin_right = Some(right.as_twips());
        sect.margin_bottom = Some(bottom.as_twips());
        sect.margin_left = Some(left.as_twips());
    }

    /// Set equal-width column layout.
    pub fn set_columns(&mut self, num: u32, spacing: Length) {
        let sect = self.section_properties_mut();
        sect.columns = Some(CT_Columns {
            num: Some(num),
            space: Some(spacing.as_twips()),
            equal_width: Some(true),
            sep: None,
            columns: Vec::new(),
        });
    }

    /// Set header and footer distances from page edges.
    pub fn set_header_footer_distance(&mut self, header: Length, footer: Length) {
        let sect = self.section_properties_mut();
        sect.header_distance = Some(header.as_twips());
        sect.footer_distance = Some(footer.as_twips());
    }

    /// Set the gutter margin.
    pub fn set_gutter(&mut self, gutter: Length) {
        self.section_properties_mut().gutter = Some(gutter.as_twips());
    }

    /// Enable or disable different first page header/footer.
    pub fn set_different_first_page(&mut self, val: bool) {
        self.section_properties_mut().title_pg = Some(val);
    }

    /// Enable or disable automatic document hyphenation.
    pub fn set_auto_hyphenation(&mut self, enabled: bool) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        let created = candidate.settings_part_name.is_none();
        candidate
            .identifiers
            .observe_package_graph(&candidate.package)?;
        let part_name = match &candidate.settings_part_name {
            Some(part_name) => part_name.clone(),
            None => candidate
                .identifiers
                .reserve_preferred_part_name(DEFAULT_SETTINGS_PART)?,
        };
        candidate
            .settings
            .get_or_insert_with(CT_Settings::new)
            .set_automatic_hyphenation(enabled)?;
        candidate.invalidate_layout();
        candidate.settings_part_name = Some(part_name.clone());
        candidate.settings_owned |= created;
        candidate
            .ensure_part_relationship_checked(
                &part_name,
                rel_types::SETTINGS,
                SETTINGS_CONTENT_TYPE,
            )
            .map_err(|error| {
                Error::Other(format!("settings relationship allocation failed: {error}"))
            })?;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Return document-wide OfficeMath defaults from the settings part.
    pub fn math_properties(&self) -> Option<&MathProperties> {
        self.settings.as_ref()?.math_properties()
    }

    /// Set document-wide OfficeMath defaults in the relationship-resolved settings part.
    pub fn set_math_properties(&mut self, properties: MathProperties) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        let created = candidate.settings_part_name.is_none();
        candidate
            .identifiers
            .observe_package_graph(&candidate.package)?;
        let part_name = match &candidate.settings_part_name {
            Some(part_name) => part_name.clone(),
            None => candidate
                .identifiers
                .reserve_preferred_part_name(DEFAULT_SETTINGS_PART)?,
        };
        candidate
            .settings
            .get_or_insert_with(CT_Settings::new)
            .set_math_properties(properties)?;
        candidate.invalidate_layout();
        candidate.settings_part_name = Some(part_name.clone());
        candidate.settings_owned |= created;
        candidate
            .ensure_part_relationship_checked(
                &part_name,
                rel_types::SETTINGS,
                SETTINGS_CONTENT_TYPE,
            )
            .map_err(|error| {
                Error::Other(format!("settings relationship allocation failed: {error}"))
            })?;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Return a document variable by name.
    pub fn document_variable(&self, name: &str) -> Option<&str> {
        self.settings.as_ref()?.document_variable(name)
    }

    /// Set a document variable in the relationship-resolved settings part.
    pub fn set_document_variable(&mut self, name: &str, value: &str) -> Result<()> {
        let mut candidate = self.settings_mutation_candidate()?;
        candidate
            .settings
            .get_or_insert_with(CT_Settings::new)
            .set_document_variable(name.to_owned(), value.to_owned())?;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Remove one document variable by name.
    pub fn remove_document_variable(&mut self, name: &str) -> Result<Option<String>> {
        let Some(settings) = &self.settings else {
            return Ok(None);
        };
        if settings.document_variable(name).is_none() {
            return Ok(None);
        }
        let mut candidate = self.settings_mutation_candidate()?;
        let removed = candidate
            .settings
            .as_mut()
            .expect("settings candidate retains its model")
            .remove_document_variable(name)?
            .map(|variable| variable.value);
        candidate.prune_empty_owned_settings();
        self.commit_staged_mutation(candidate);
        Ok(removed)
    }

    /// Return typed compatibility settings in package order.
    pub fn compatibility_settings(&self) -> &[CompatibilitySetting] {
        self.settings
            .as_ref()
            .map(CT_Settings::compatibility_settings)
            .unwrap_or_default()
    }

    /// Add or replace one compatibility setting selected by name and URI.
    pub fn set_compatibility_setting(&mut self, name: &str, uri: &str, value: &str) -> Result<()> {
        let mut candidate = self.settings_mutation_candidate()?;
        candidate
            .settings
            .get_or_insert_with(CT_Settings::new)
            .set_compatibility_setting(CompatibilitySetting {
                name: name.to_owned(),
                uri: uri.to_owned(),
                value: value.to_owned(),
            })?;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Remove one compatibility setting selected by name and URI.
    pub fn remove_compatibility_setting(
        &mut self,
        name: &str,
        uri: &str,
    ) -> Result<Option<CompatibilitySetting>> {
        let Some(settings) = &self.settings else {
            return Ok(None);
        };
        if !settings
            .compatibility_settings()
            .iter()
            .any(|setting| setting.name == name && setting.uri == uri)
        {
            return Ok(None);
        }
        let mut candidate = self.settings_mutation_candidate()?;
        let removed = candidate
            .settings
            .as_mut()
            .expect("settings candidate retains its model")
            .remove_compatibility_setting(name, uri)?;
        candidate.prune_empty_owned_settings();
        self.commit_staged_mutation(candidate);
        Ok(removed)
    }

    pub fn default_tab_stop(&self) -> Option<oxml_core::Twips> {
        self.settings.as_ref()?.default_tab_stop()
    }

    pub fn set_default_tab_stop(&mut self, value: oxml_core::Twips) -> Result<()> {
        let mut candidate = self.settings_mutation_candidate()?;
        candidate
            .settings
            .get_or_insert_with(CT_Settings::new)
            .set_default_tab_stop(value)?;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    pub fn remove_default_tab_stop(&mut self) -> Result<Option<oxml_core::Twips>> {
        if self.default_tab_stop().is_none() {
            return Ok(None);
        }
        let mut candidate = self.settings_mutation_candidate()?;
        let removed = candidate
            .settings
            .as_mut()
            .expect("settings candidate retains its model")
            .remove_default_tab_stop()?;
        candidate.prune_empty_owned_settings();
        self.commit_staged_mutation(candidate);
        Ok(removed)
    }

    pub fn character_spacing_control(&self) -> Option<CharacterSpacingControl> {
        self.settings.as_ref()?.character_spacing_control()
    }

    pub fn set_character_spacing_control(&mut self, value: CharacterSpacingControl) -> Result<()> {
        let mut candidate = self.settings_mutation_candidate()?;
        candidate
            .settings
            .get_or_insert_with(CT_Settings::new)
            .set_character_spacing_control(value)?;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    pub fn remove_character_spacing_control(&mut self) -> Result<Option<CharacterSpacingControl>> {
        if self.character_spacing_control().is_none() {
            return Ok(None);
        }
        let mut candidate = self.settings_mutation_candidate()?;
        let removed = candidate
            .settings
            .as_mut()
            .expect("settings candidate retains its model")
            .remove_character_spacing_control()?;
        candidate.prune_empty_owned_settings();
        self.commit_staged_mutation(candidate);
        Ok(removed)
    }

    pub fn theme_font_language(&self) -> Option<&ThemeFontLanguage> {
        self.settings.as_ref()?.theme_font_language()
    }

    pub fn remove_theme_font_language(&mut self) -> Result<Option<ThemeFontLanguage>> {
        if self.theme_font_language().is_none() {
            return Ok(None);
        }
        let mut candidate = self.settings_mutation_candidate()?;
        let removed = candidate
            .settings
            .as_mut()
            .expect("settings candidate retains its model")
            .remove_theme_font_language()?;
        candidate.prune_empty_owned_settings();
        self.commit_staged_mutation(candidate);
        Ok(removed)
    }

    /// Return the relationship-resolved DrawingML theme.
    pub fn theme(&self) -> Option<&oxml_drawing::theme::CT_OfficeStyleSheet> {
        self.theme.as_ref()
    }

    /// Replace or create the relationship-resolved DrawingML theme atomically.
    pub fn set_theme(&mut self, theme: oxml_drawing::theme::CT_OfficeStyleSheet) -> Result<()> {
        theme
            .to_xml()
            .map_err(|error| Error::Other(format!("theme serialization failed: {error}")))?;
        let mut candidate = self.clone_for_staging();
        candidate
            .identifiers
            .observe_package_graph(&candidate.package)?;
        let part_name = match &candidate.theme_part_name {
            Some(part_name) => part_name.clone(),
            None => candidate
                .identifiers
                .reserve_preferred_part_name(DEFAULT_THEME_PART)?,
        };
        candidate
            .ensure_part_relationship_checked(&part_name, rel_types::THEME, content_types::THEME)
            .map_err(|error| {
                Error::Other(format!("theme relationship allocation failed: {error}"))
            })?;
        candidate.theme = Some(theme);
        candidate.theme_part_name = Some(part_name);
        candidate.theme_dirty = true;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Set the document-wide Latin, East Asian, and bidirectional languages.
    pub fn set_language_defaults(&mut self, value: ThemeFontLanguage) -> Result<()> {
        let mut candidate = self.settings_mutation_candidate()?;
        candidate
            .settings
            .get_or_insert_with(CT_Settings::new)
            .set_theme_font_language(value)?;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Return font-table records with related embedded bytes deobfuscated.
    pub fn fonts(&self) -> Vec<FontDefinition> {
        let Some(table) = &self.font_table else {
            return Vec::new();
        };
        let relationships = self
            .font_table_part_name
            .as_deref()
            .and_then(|part_name| self.package.get_part_rels(part_name));
        table
            .fonts()
            .iter()
            .map(|font| {
                let embedded_fonts = font
                    .embedded_fonts
                    .iter()
                    .map(|embedded| {
                        let data = relationships
                            .and_then(|relationships| {
                                relationships.items.iter().find(|relationship| {
                                    relationship.id == embedded.relationship_id
                                        && relationship.rel_type == rel_types::FONT
                                        && relationship_is_internal(relationship)
                                })
                            })
                            .and_then(|relationship| {
                                let owner = self.font_table_part_name.as_deref()?;
                                let target =
                                    OpcPackage::resolve_rel_target(owner, &relationship.target);
                                self.package.get_part(&target)
                            })
                            .and_then(|bytes| deobfuscate_odttf_with_key(bytes, &embedded.font_key))
                            .unwrap_or_default();
                        EmbeddedFont {
                            kind: embedded.kind.into(),
                            data,
                            font_key: embedded.font_key.clone(),
                            subsetted: embedded.subsetted.unwrap_or(false),
                            license: FontEmbeddingLicense {
                                authorized: embedded.authorized.unwrap_or(false),
                                identity: embedded.license_identity.clone().unwrap_or_default(),
                            },
                        }
                    })
                    .collect();
                FontDefinition {
                    name: font.name.clone(),
                    alternate_name: font.alternate_name.clone(),
                    family: font.family.clone(),
                    pitch: font.pitch.clone(),
                    embedded_fonts,
                }
            })
            .collect()
    }

    /// Add or replace one descriptive font-table record.
    pub fn set_font(&mut self, font: FontDefinition) -> Result<()> {
        if font.name.trim().is_empty() {
            return Err(Error::Other("font name must not be empty".to_owned()));
        }
        if font.family.as_deref().is_some_and(|family| {
            !matches!(
                family,
                "decorative" | "modern" | "roman" | "script" | "swiss" | "auto"
            )
        }) {
            return Err(Error::Other(
                "font family must be a valid OOXML family enumeration".to_owned(),
            ));
        }
        if font
            .pitch
            .as_deref()
            .is_some_and(|pitch| !matches!(pitch, "default" | "fixed" | "variable"))
        {
            return Err(Error::Other(
                "font pitch must be a valid OOXML pitch enumeration".to_owned(),
            ));
        }
        if !font.embedded_fonts.is_empty() {
            return Err(Error::Other(
                "set_font does not accept embedded bytes, use embed_font".to_owned(),
            ));
        }
        let mut candidate = self.font_table_mutation_candidate()?;
        candidate
            .font_table
            .as_mut()
            .expect("font-table mutation candidate has a model")
            .set_font(font.name, font.alternate_name, font.family, font.pitch);
        candidate
            .font_table
            .as_ref()
            .expect("font-table mutation candidate has a model")
            .to_xml()?;
        candidate.font_table_dirty = true;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Remove one font-table record and every embedded face it owns.
    pub fn remove_font(&mut self, name: &str) -> Result<Option<FontDefinition>> {
        let Some(removed) = self.fonts().into_iter().find(|font| font.name == name) else {
            return Ok(None);
        };
        let mut candidate = self.font_table_mutation_candidate()?;
        let references = candidate
            .font_table
            .as_mut()
            .expect("font-table mutation candidate has a model")
            .remove_font(name)
            .expect("font record existed before staging")
            .embedded_fonts;
        for reference in references {
            candidate.remove_embedded_font_part(&reference)?;
        }
        candidate.font_table_dirty = true;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(Some(removed))
    }

    /// Embed caller-provided font bytes after explicit license authorization.
    pub fn embed_font(&mut self, font_name: &str, embedded: EmbeddedFont) -> Result<()> {
        if !embedded.license.authorized {
            return Err(Error::Other(
                "font embedding requires explicit caller authorization".to_owned(),
            ));
        }
        if embedded.license.identity.trim().is_empty() {
            return Err(Error::Other(
                "font embedding requires an exact license identity".to_owned(),
            ));
        }
        if embedded.license.identity.len() > 4096 {
            return Err(Error::Other(
                "font embedding license identity exceeds 4096 bytes".to_owned(),
            ));
        }
        if embedded
            .license
            .identity
            .bytes()
            .any(|byte| matches!(byte, b'\t' | b'\n' | b'\r'))
        {
            return Err(Error::Other(
                "font embedding license identity contains XML-normalized whitespace".to_owned(),
            ));
        }
        if !looks_like_sfnt(&embedded.data) {
            return Err(Error::Other(
                "embedded font bytes are not a supported sfnt font".to_owned(),
            ));
        }
        let guid = font_key_bytes(&embedded.font_key).ok_or_else(|| {
            Error::Other("embedded font key must be a 16-byte OOXML GUID".to_owned())
        })?;

        let mut candidate = self.font_table_mutation_candidate()?;
        let table = candidate
            .font_table
            .as_ref()
            .expect("font-table mutation candidate has a model");
        let font = table
            .fonts()
            .iter()
            .find(|font| font.name == font_name)
            .ok_or_else(|| Error::Other(format!("font-table record {font_name} does not exist")))?;
        let kind = FontFaceKind::from(embedded.kind);
        let existing = font
            .embedded_fonts
            .iter()
            .find(|value| value.kind == kind)
            .cloned();
        let font_table_part = candidate
            .font_table_part_name
            .clone()
            .expect("font-table mutation candidate has a part name");

        let existing_target = if let Some(existing) = &existing {
            let relationship = candidate
                .package
                .get_part_rels(&font_table_part)
                .and_then(|relationships| relationships.get_by_id(&existing.relationship_id))
                .filter(|relationship| {
                    relationship.rel_type == rel_types::FONT
                        && relationship_is_internal(relationship)
                })
                .ok_or_else(|| {
                    Error::Other(format!(
                        "embedded font relationship {} is missing",
                        existing.relationship_id
                    ))
                })?;
            Some(OpcPackage::resolve_rel_target(
                &font_table_part,
                &relationship.target,
            ))
        } else {
            None
        };
        let can_reuse_existing = existing.as_ref().is_some_and(|existing| {
            let relationship_reference_count =
                table.embedded_relationship_reference_count(&existing.relationship_id);
            let target = existing_target
                .as_ref()
                .expect("existing embedded font has a resolved target");
            let target_relationship_count = std::iter::once(("/", &candidate.package.package_rels))
                .chain(
                    candidate
                        .package
                        .part_rels
                        .iter()
                        .map(|(source, relationships)| (source.as_str(), relationships)),
                )
                .flat_map(|(source, relationships)| {
                    relationships
                        .items
                        .iter()
                        .map(move |relationship| (source, relationship))
                })
                .filter(|(source, relationship)| {
                    relationship_is_internal(relationship)
                        && OpcPackage::resolve_rel_target(source, &relationship.target) == *target
                })
                .count();
            relationship_reference_count == 1
                && target_relationship_count == 1
                && !table.retained_raw_mentions_relationship(&existing.relationship_id)
        });

        let (relationship_id, part_name) = if can_reuse_existing {
            (
                existing
                    .as_ref()
                    .expect("reusable embedded font exists")
                    .relationship_id
                    .clone(),
                existing_target.expect("reusable embedded font has a target"),
            )
        } else {
            let part_name = candidate
                .identifiers
                .reserve_preferred_part_name("/word/fonts/font1.odttf")?;
            let relationship_id = candidate
                .identifiers
                .reserve_relationship_id_checked(&font_table_part)?;
            let target = relative_target(&font_table_part, &part_name);
            candidate
                .package
                .get_or_create_part_rels(&font_table_part)
                .add_with_id(&relationship_id, rel_types::FONT, &target);
            (relationship_id, part_name)
        };

        candidate
            .package
            .set_part(&part_name, obfuscate_odttf_with_key(&embedded.data, &guid));
        candidate
            .package
            .content_types
            .add_override(&part_name, EMBEDDED_FONT_CONTENT_TYPE);
        candidate
            .identifiers
            .register_content_type_override(&part_name);
        let reference = EmbeddedFontReference::new(
            kind,
            relationship_id,
            embedded.font_key,
            Some(embedded.subsetted),
            Some(true),
            Some(embedded.license.identity),
        );
        let inserted = candidate
            .font_table
            .as_mut()
            .expect("font-table mutation candidate has a model")
            .set_embedded_font(font_name, reference);
        debug_assert!(inserted);
        if !can_reuse_existing && let Some(existing) = &existing {
            candidate.remove_embedded_font_part(existing)?;
        }
        candidate
            .font_table
            .as_ref()
            .expect("font-table mutation candidate has a model")
            .to_xml()?;
        candidate.font_table_dirty = true;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Remove one embedded face without removing its font-table record.
    pub fn remove_embedded_font(
        &mut self,
        font_name: &str,
        kind: EmbeddedFontKind,
    ) -> Result<Option<EmbeddedFont>> {
        let Some(removed) = self
            .fonts()
            .into_iter()
            .find(|font| font.name == font_name)
            .and_then(|font| {
                font.embedded_fonts
                    .into_iter()
                    .find(|embedded| embedded.kind == kind)
            })
        else {
            return Ok(None);
        };
        let mut candidate = self.font_table_mutation_candidate()?;
        let reference = candidate
            .font_table
            .as_mut()
            .expect("font-table mutation candidate has a model")
            .remove_embedded_font(font_name, kind.into())
            .expect("embedded font existed before staging");
        candidate.remove_embedded_font_part(&reference)?;
        candidate.font_table_dirty = true;
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
        Ok(Some(removed))
    }

    fn font_table_mutation_candidate(&self) -> Result<Self> {
        let mut candidate = self.clone_for_staging();
        candidate
            .identifiers
            .observe_package_graph(&candidate.package)?;
        let part_name = match &candidate.font_table_part_name {
            Some(part_name) => {
                if candidate.font_table.is_none() {
                    return Err(Error::Other(format!(
                        "cannot mutate unmodeled font table at {part_name}"
                    )));
                }
                part_name.clone()
            }
            None => candidate
                .identifiers
                .reserve_preferred_part_name(DEFAULT_FONT_TABLE_PART)?,
        };
        candidate
            .ensure_part_relationship_checked(
                &part_name,
                rel_types::FONT_TABLE,
                FONT_TABLE_CONTENT_TYPE,
            )
            .map_err(|error| {
                Error::Other(format!(
                    "font-table relationship allocation failed: {error}"
                ))
            })?;
        candidate.font_table.get_or_insert_with(FontTable::new);
        candidate.font_table_part_name = Some(part_name);
        Ok(candidate)
    }

    fn remove_embedded_font_part(&mut self, reference: &EmbeddedFontReference) -> Result<()> {
        let owner = self
            .font_table_part_name
            .clone()
            .ok_or_else(|| Error::Other("font-table part name is missing".to_owned()))?;
        if self.font_table.as_ref().is_some_and(|table| {
            table.embedded_relationship_reference_count(&reference.relationship_id) > 0
                || table.retained_raw_mentions_relationship(&reference.relationship_id)
        }) {
            return Ok(());
        }
        let target = self
            .package
            .get_part_rels(&owner)
            .and_then(|relationships| relationships.get_by_id(&reference.relationship_id))
            .filter(|relationship| {
                relationship.rel_type == rel_types::FONT && relationship_is_internal(relationship)
            })
            .map(|relationship| OpcPackage::resolve_rel_target(&owner, &relationship.target));
        let Some(target) = target else {
            return Ok(());
        };
        if let Some(relationships) = self.package.get_part_rels_mut(&owner) {
            relationships.items.retain(|relationship| {
                relationship.id != reference.relationship_id
                    || relationship.rel_type != rel_types::FONT
                    || !relationship_is_internal(relationship)
            });
        }
        self.identifiers
            .retire_authored_story_relationships(&owner, vec![reference.relationship_id.clone()]);
        let still_referenced = std::iter::once(("/", &self.package.package_rels))
            .chain(
                self.package
                    .part_rels
                    .iter()
                    .map(|(source, relationships)| (source.as_str(), relationships)),
            )
            .any(|(source, relationships)| {
                relationships.items.iter().any(|relationship| {
                    relationship_is_internal(relationship)
                        && OpcPackage::resolve_rel_target(source, &relationship.target) == target
                })
            });
        if !still_referenced {
            self.package.remove_part(&target);
            self.package.remove_part_rels(&target);
            self.package.content_types.remove_override(&target);
            self.identifiers.retire_authored_part(&target);
        }
        Ok(())
    }

    fn settings_mutation_candidate(&self) -> Result<Self> {
        let mut candidate = self.clone_for_staging();
        let created = candidate.settings_part_name.is_none();
        candidate
            .identifiers
            .observe_package_graph(&candidate.package)?;
        let part_name = match &candidate.settings_part_name {
            Some(part_name) => part_name.clone(),
            None => candidate
                .identifiers
                .reserve_preferred_part_name(DEFAULT_SETTINGS_PART)?,
        };
        candidate.invalidate_layout();
        candidate.settings_part_name = Some(part_name.clone());
        candidate.settings_owned |= created;
        candidate
            .ensure_part_relationship_checked(
                &part_name,
                rel_types::SETTINGS,
                SETTINGS_CONTENT_TYPE,
            )
            .map_err(|error| {
                Error::Other(format!("settings relationship allocation failed: {error}"))
            })?;
        Ok(candidate)
    }

    fn prune_empty_owned_settings(&mut self) {
        if !self.settings_owned || !self.settings.as_ref().is_some_and(CT_Settings::is_empty) {
            return;
        }
        let Some(part_name) = self.settings_part_name.take() else {
            return;
        };
        let owner = self.doc_part_name.clone();
        let mut removed_ids = Vec::new();
        if let Some(relationships) = self.package.get_part_rels_mut(&owner) {
            relationships.items.retain(|relationship| {
                let remove = relationship.rel_type == rel_types::SETTINGS
                    && relationship_is_internal(relationship)
                    && OpcPackage::resolve_rel_target(&owner, &relationship.target) == part_name;
                if remove {
                    removed_ids.push(relationship.id.clone());
                }
                !remove
            });
        }
        self.identifiers
            .retire_authored_story_relationships(&owner, removed_ids);
        self.package.remove_part(&part_name);
        self.package.remove_part_rels(&part_name);
        self.package.content_types.remove_override(&part_name);
        self.identifiers.retire_authored_part(&part_name);
        self.settings = None;
        self.settings_owned = false;
    }

    // ---- Metadata access ----

    /// Return the complete core-properties model.
    pub fn core_properties(&self) -> Option<&CoreProperties> {
        self.core_properties.as_ref()
    }

    /// Replace the complete core-properties model.
    pub fn set_core_properties(&mut self, properties: CoreProperties) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.reserve_core_properties_bundle()?;
        candidate.core_properties = Some(properties);
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Remove the complete core-properties part and its package relationship.
    pub fn remove_core_properties(&mut self) -> Result<Option<CoreProperties>> {
        let Some(properties) = self.core_properties.clone() else {
            return Ok(None);
        };
        let mut candidate = self.clone_for_staging();
        if let Some(part_name) = candidate.core_properties_part_name.take() {
            candidate.remove_package_properties_bundle(&part_name, rel_types::CORE_PROPERTIES);
        }
        candidate.core_properties = None;
        self.commit_staged_mutation(candidate);
        Ok(Some(properties))
    }

    /// Return the complete application-properties model.
    pub fn application_properties(&self) -> Option<&AppProperties> {
        self.application_properties.as_ref()
    }

    /// Replace the complete application-properties model.
    pub fn set_application_properties(&mut self, properties: AppProperties) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.reserve_application_properties_bundle()?;
        candidate.application_properties = Some(properties);
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Remove the complete application-properties part and package relationship.
    pub fn remove_application_properties(&mut self) -> Result<Option<AppProperties>> {
        let Some(properties) = self.application_properties.clone() else {
            return Ok(None);
        };
        let mut candidate = self.clone_for_staging();
        if let Some(part_name) = candidate.application_properties_part_name.take() {
            candidate.remove_package_properties_bundle(&part_name, rel_types::EXTENDED_PROPERTIES);
        }
        candidate.application_properties = None;
        self.commit_staged_mutation(candidate);
        Ok(Some(properties))
    }

    /// Return all custom properties in package order.
    pub fn custom_properties(&self) -> &[CustomProperty] {
        self.custom_properties
            .as_ref()
            .map(|properties| properties.properties.as_slice())
            .unwrap_or_default()
    }

    /// Return a custom property by its case-sensitive name.
    pub fn custom_property(&self, name: &str) -> Option<&CustomProperty> {
        self.custom_properties()
            .iter()
            .find(|property| property.name.as_deref() == Some(name))
    }

    /// Add or replace one named custom property.
    pub fn set_custom_property(&mut self, property: CustomProperty) -> Result<()> {
        let Some(name) = property.name.as_deref().filter(|name| !name.is_empty()) else {
            return Err(Error::Other(
                "an authored custom property requires a non-empty name".to_owned(),
            ));
        };
        if property.pid < 2 {
            return Err(Error::Other(
                "an authored custom property pid must be at least 2".to_owned(),
            ));
        }
        if self
            .custom_properties()
            .iter()
            .any(|existing| existing.name.as_deref() != Some(name) && existing.pid == property.pid)
        {
            return Err(Error::Other(format!(
                "custom property pid {} is already in use",
                property.pid
            )));
        }

        let mut candidate = self.clone_for_staging();
        let created = candidate.custom_properties_part_name.is_none();
        candidate.reserve_custom_properties_bundle()?;
        let properties = candidate
            .custom_properties
            .get_or_insert_with(CustomProperties::default);
        if let Some(index) = properties
            .properties
            .iter()
            .position(|existing| existing.name.as_deref() == Some(name))
        {
            properties
                .properties
                .retain(|existing| existing.name.as_deref() != Some(name));
            properties.properties.insert(index, property);
        } else {
            properties.properties.push(property);
        }
        candidate.custom_properties_owned |= created;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Remove one named custom property.
    pub fn remove_custom_property(&mut self, name: &str) -> Result<Option<CustomProperty>> {
        let Some(removed) = self
            .custom_properties()
            .iter()
            .find(|property| property.name.as_deref() == Some(name))
            .cloned()
        else {
            return Ok(None);
        };
        let mut candidate = self.clone_for_staging();
        let properties = candidate
            .custom_properties
            .as_mut()
            .expect("custom property lookup proved the model exists");
        properties
            .properties
            .retain(|property| property.name.as_deref() != Some(name));
        if properties.properties.is_empty() && candidate.custom_properties_owned {
            if let Some(part_name) = candidate.custom_properties_part_name.take() {
                candidate
                    .remove_package_properties_bundle(&part_name, rel_types::CUSTOM_PROPERTIES);
            }
            candidate.custom_properties = None;
            candidate.custom_properties_owned = false;
        }
        self.commit_staged_mutation(candidate);
        Ok(Some(removed))
    }

    /// Remove the complete custom-properties part and package relationship.
    pub fn remove_custom_properties(&mut self) -> Result<Option<Vec<CustomProperty>>> {
        let Some(properties) = self.custom_properties.clone() else {
            return Ok(None);
        };
        let mut candidate = self.clone_for_staging();
        if let Some(part_name) = candidate.custom_properties_part_name.take() {
            candidate.remove_package_properties_bundle(&part_name, rel_types::CUSTOM_PROPERTIES);
        }
        candidate.custom_properties = None;
        candidate.custom_properties_owned = false;
        self.commit_staged_mutation(candidate);
        Ok(Some(properties.properties))
    }

    fn remove_package_properties_bundle(&mut self, part_name: &str, rel_type: &str) {
        let mut removed_ids = Vec::new();
        self.package.package_rels.items.retain(|relationship| {
            let remove = relationship.rel_type == rel_type
                && relationship_is_internal(relationship)
                && OpcPackage::resolve_rel_target("/", &relationship.target) == part_name;
            if remove {
                removed_ids.push(relationship.id.clone());
            }
            !remove
        });
        self.identifiers
            .retire_authored_story_relationships("/", removed_ids);
        self.package.remove_part(part_name);
        self.package.remove_part_rels(part_name);
        self.package.content_types.remove_override(part_name);
        self.identifiers.retire_authored_part(part_name);
    }

    /// Get the document title.
    pub fn title(&self) -> Option<&str> {
        self.core_properties.as_ref()?.title.as_deref()
    }

    /// Set the document title.
    pub fn set_title(&mut self, title: &str) {
        let mut candidate = self.clone_for_staging();
        candidate
            .reserve_core_properties_bundle()
            .expect("an in-memory document can allocate a core-properties part");
        candidate.invalidate_layout();
        candidate.ensure_core_properties().title = Some(title.to_string());
        self.commit_staged_mutation(candidate);
    }

    /// Get the document author/creator.
    pub fn author(&self) -> Option<&str> {
        self.core_properties.as_ref()?.creator.as_deref()
    }

    /// Set the document author/creator.
    pub fn set_author(&mut self, author: &str) {
        let mut candidate = self.clone_for_staging();
        candidate
            .reserve_core_properties_bundle()
            .expect("an in-memory document can allocate a core-properties part");
        candidate.invalidate_layout();
        candidate.ensure_core_properties().creator = Some(author.to_string());
        self.commit_staged_mutation(candidate);
    }

    /// Get the document subject.
    pub fn subject(&self) -> Option<&str> {
        self.core_properties.as_ref()?.subject.as_deref()
    }

    /// Set the document subject.
    pub fn set_subject(&mut self, subject: &str) {
        let mut candidate = self.clone_for_staging();
        candidate
            .reserve_core_properties_bundle()
            .expect("an in-memory document can allocate a core-properties part");
        candidate.invalidate_layout();
        candidate.ensure_core_properties().subject = Some(subject.to_string());
        self.commit_staged_mutation(candidate);
    }

    /// Get the document keywords.
    pub fn keywords(&self) -> Option<&str> {
        self.core_properties.as_ref()?.keywords.as_deref()
    }

    /// Set the document keywords.
    pub fn set_keywords(&mut self, keywords: &str) {
        let mut candidate = self.clone_for_staging();
        candidate
            .reserve_core_properties_bundle()
            .expect("an in-memory document can allocate a core-properties part");
        candidate.invalidate_layout();
        candidate.ensure_core_properties().keywords = Some(keywords.to_string());
        self.commit_staged_mutation(candidate);
    }

    fn ensure_core_properties(&mut self) -> &mut CoreProperties {
        self.core_properties
            .get_or_insert_with(CoreProperties::default)
    }

    // ---- Document Merging ----

    /// Append the content of another document to this document.
    ///
    /// Copies all body content (paragraphs and tables) from the other document.
    /// Handles style deduplication and numbering remapping.
    pub fn append(&mut self, other: &Document) {
        let mut candidate = self.clone_for_staging();
        candidate
            .append_document_content(other)
            .expect("document append preflight failed");
        self.commit_staged_mutation(candidate);
    }

    /// Append the content of another document with a section break.
    pub fn append_with_break(&mut self, other: &Document, break_type: crate::SectionBreak) {
        let mut candidate = self.clone_for_staging();
        candidate.invalidate_layout();
        // Insert a section break paragraph before the merged content
        let mut p = CT_P::new();
        let sect_pr = match break_type {
            crate::SectionBreak::NextPage => CT_SectPr::default_letter(),
            crate::SectionBreak::Continuous => {
                let mut sp = CT_SectPr::default_letter();
                sp.section_type = Some(ST_SectionType::Continuous);
                sp
            }
            crate::SectionBreak::EvenPage => {
                let mut sp = CT_SectPr::default_letter();
                sp.section_type = Some(ST_SectionType::EvenPage);
                sp
            }
            crate::SectionBreak::OddPage => {
                let mut sp = CT_SectPr::default_letter();
                sp.section_type = Some(ST_SectionType::OddPage);
                sp
            }
        };
        p.properties = Some(CT_PPr {
            sect_pr: Some(sect_pr),
            ..Default::default()
        });
        candidate
            .document
            .body
            .content
            .push(BodyContent::Paragraph(p));
        candidate
            .append_document_content(other)
            .expect("document append-with-break preflight failed");
        self.commit_staged_mutation(candidate);
    }

    /// Insert the content of another document at a specified body index.
    ///
    /// An `index` past the end is clamped to the end rather than panicking.
    pub fn insert_document(&mut self, index: usize, other: &Document) {
        let mut candidate = self.clone_for_staging();
        candidate
            .insert_document_content_staged(index, other)
            .expect("document insertion preflight failed");
        self.commit_staged_mutation(candidate);
    }

    /// Insert into a caller-owned staged candidate and return any preflight failure.
    pub(crate) fn insert_document_content_staged(
        &mut self,
        index: usize,
        other: &Document,
    ) -> Result<()> {
        self.reserve_merge_bundles(other)?;
        self.invalidate_layout();
        self.merge_styles(other)?;

        let insert_at = index.min(self.document.body.content.len());
        for (i, content) in other.document.body.content.iter().enumerate() {
            self.document
                .body
                .content
                .insert(insert_at + i, content.clone());
        }

        self.remap_merged_numbering(other, insert_at)
    }

    fn append_document_content(&mut self, other: &Document) -> Result<()> {
        self.reserve_merge_bundles(other)?;
        self.invalidate_layout();
        self.merge_styles(other)?;
        let start_idx = self.document.body.content.len();
        self.document
            .body
            .content
            .extend(other.document.body.content.iter().cloned());
        self.remap_merged_numbering(other, start_idx)
    }

    fn reserve_merge_bundles(&mut self, other: &Document) -> Result<()> {
        let needs_styles = other
            .styles
            .styles
            .iter()
            .any(|style| self.styles.get_by_id(&style.style_id).is_none());
        let needs_numbering = other.numbering.is_some();
        let additional_relationships = usize::from(
            needs_styles
                && self.document_bundle_relationship_is_missing(
                    self.styles_part_name.as_deref(),
                    rel_types::STYLES,
                ),
        ) + usize::from(
            needs_numbering
                && self.document_bundle_relationship_is_missing(
                    self.numbering_part_name.as_deref(),
                    rel_types::NUMBERING,
                ),
        );
        self.identifiers
            .ensure_relationship_capacity(&self.doc_part_name, additional_relationships)?;
        if needs_styles {
            self.reserve_styles_bundle()?;
        }
        if needs_numbering {
            self.reserve_numbering_bundle()?;
        }
        Ok(())
    }

    fn document_bundle_relationship_is_missing(
        &self,
        part_name: Option<&str>,
        rel_type: &str,
    ) -> bool {
        let Some(part_name) = part_name else {
            return true;
        };
        !self
            .package
            .get_part_rels(&self.doc_part_name)
            .is_some_and(|relationships| {
                relationships.items.iter().any(|relationship| {
                    relationship.rel_type == rel_type
                        && relationship_is_internal(relationship)
                        && OpcPackage::resolve_rel_target(&self.doc_part_name, &relationship.target)
                            == part_name
                })
            })
    }

    /// Merge styles from another document, avoiding duplicates.
    fn merge_styles(&mut self, other: &Document) -> Result<()> {
        for style in &other.styles.styles {
            if self.styles.get_by_id(&style.style_id).is_none() {
                self.styles.styles.push(style.clone());
            }
        }
        style::validate_style_graph(&self.styles)
    }

    /// Merge numbering from another document and remap IDs in the merged content.
    /// `start_idx` is the index where the other document's content starts in self.
    fn remap_merged_numbering(&mut self, other: &Document, start_idx: usize) -> Result<()> {
        let Some(other_numbering) = &other.numbering else {
            return Ok(());
        };
        let mut abstract_remap = HashMap::new();
        for abs_num in &other_numbering.abstract_nums {
            let new_id = self.identifiers.reserve_abstract_numbering_id()?;
            abstract_remap.insert(abs_num.abstract_num_id, new_id);
        }
        let mut num_remap = HashMap::new();
        for num in &other_numbering.nums {
            num_remap.insert(
                num.num_id,
                self.identifiers.reserve_numbering_instance_id()?,
            );
            if let std::collections::hash_map::Entry::Vacant(entry) =
                abstract_remap.entry(num.abstract_num_id)
            {
                entry.insert(self.identifiers.reserve_abstract_numbering_id()?);
            }
        }

        let numbering = self
            .numbering
            .get_or_insert_with(|| rdocx_oxml::numbering::CT_Numbering {
                abstract_nums: Vec::new(),
                nums: Vec::new(),
                root_attributes: Vec::new(),
                extra_xml: Vec::new(),
            });
        for abs_num in &other_numbering.abstract_nums {
            let mut new_abs = abs_num.clone();
            new_abs.abstract_num_id = abstract_remap[&abs_num.abstract_num_id];
            numbering.abstract_nums.push(new_abs);
        }
        for num in &other_numbering.nums {
            let mut new_num = num.clone();
            new_num.num_id = num_remap[&num.num_id];
            new_num.abstract_num_id = abstract_remap[&num.abstract_num_id];
            numbering.nums.push(new_num);
        }
        let incoming_count = other.document.body.content.len();
        visit_body_paragraphs_mut(
            &mut self.document.body.content[start_idx..start_idx + incoming_count],
            &mut |paragraph| {
                if let Some(num_id) = paragraph
                    .properties
                    .as_mut()
                    .and_then(|properties| properties.num_id.as_mut())
                    && let Some(updated) = num_remap.get(num_id)
                {
                    *num_id = *updated;
                }
            },
        );
        Ok(())
    }

    // ---- Table of Contents ----

    /// Insert a Table of Contents at the given body content index.
    ///
    /// Scans the document for heading paragraphs (Heading1..HeadingN where N <= max_level),
    /// inserts bookmark markers at each heading, and generates TOC entry paragraphs
    /// with internal hyperlinks and dot-leader tab stops.
    ///
    /// # Arguments
    /// * `index` - Body content index at which to insert the TOC
    /// * `max_level` - Maximum heading level to include (1-9, typically 3)
    pub fn insert_toc(&mut self, index: usize, max_level: u32) {
        self.invalidate_layout();
        use rdocx_oxml::borders::{CT_TabStop, CT_Tabs};
        use rdocx_oxml::shared::{ST_TabJc, ST_TabLeader};
        use rdocx_oxml::text::HyperlinkSpan;
        use rdocx_oxml::units::Twips;

        let max_level = max_level.clamp(1, 9);

        // Step 1: Collect heading info from the document body
        struct HeadingInfo {
            content_index: usize,
            level: u32,
            text: String,
            bookmark_name: String,
            bookmark_id: i32,
        }

        // Calling insert_toc twice must not mint bookmarks that collide with
        // the ones the first call left behind — duplicate `w:name` values make
        // the internal links ambiguous. Continue numbering past whatever is
        // already there.
        let mut occupied_suffixes = self.toc_bookmark_suffixes();
        let mut toc_counter = occupied_suffixes.iter().copied().max().unwrap_or(0);
        let mut identifiers = self.identifiers.clone();
        let mut duplicate_typed_id = false;
        let mut typed_ids = HashSet::new();
        visit_body_paragraphs(&self.document.body.content, &mut |paragraph| {
            for marker in &paragraph.bookmark_markers {
                if marker.is_start()
                    && let Some(id) = marker.id()
                    && id >= 0
                    && !typed_ids.insert(id)
                {
                    duplicate_typed_id = true;
                }
            }
        });
        if duplicate_typed_id {
            return;
        }
        identifiers.bookmark_ids.extend(typed_ids);

        let mut headings = Vec::new();

        for (idx, content) in self.document.body.content.iter().enumerate() {
            if let BodyContent::Paragraph(p) = content
                && let Some(level) = Self::detect_heading_level_for_toc(p)
                && level <= max_level
            {
                let text = p.text();
                if !text.trim().is_empty() {
                    let Some(suffix) = next_toc_bookmark_suffix(&occupied_suffixes, toc_counter)
                    else {
                        return;
                    };
                    toc_counter = suffix;
                    occupied_suffixes.insert(suffix);
                    let Some(preferred_id) = suffix
                        .checked_add(99)
                        .and_then(|candidate| i32::try_from(candidate).ok())
                    else {
                        return;
                    };
                    let Ok(bookmark_id) = identifiers.reserve_preferred_bookmark_id(preferred_id)
                    else {
                        return;
                    };
                    headings.push(HeadingInfo {
                        content_index: idx,
                        level,
                        text,
                        bookmark_name: format!("_Toc{suffix}"),
                        bookmark_id,
                    });
                }
            }
        }

        // Step 2: Insert typed bookmark markers at each heading paragraph.
        for heading in &headings {
            if let Some(BodyContent::Paragraph(p)) =
                self.document.body.content.get_mut(heading.content_index)
            {
                let run_count = p.runs.len();
                let inserted_start =
                    p.insert_bookmark_start(0, heading.bookmark_id, &heading.bookmark_name);
                let inserted_end = p.insert_bookmark_end(run_count, heading.bookmark_id);
                debug_assert!(inserted_start && inserted_end);
            }
        }

        // Step 3: Build TOC entry paragraphs.
        // The dot leader runs to the right text margin, which depends on the
        // section's page size and margins rather than being a fixed 6.5".
        let right_tab = CT_Tabs {
            tabs: vec![CT_TabStop {
                val: ST_TabJc::Right,
                pos: Twips(self.text_width_twips()),
                leader: Some(ST_TabLeader::Dot),
                source_occurrence: None,
            }],
        };

        let mut toc_paragraphs: Vec<CT_P> = Vec::new();

        // TOC title
        let mut title_p = CT_P::new();
        let mut title_r = CT_R::new("Table of Contents");
        title_r.properties = Some(CT_RPr {
            bold: Some(true),
            ..Default::default()
        });
        title_p.runs.push(title_r);
        title_p.properties = Some(CT_PPr {
            space_after: Some(Twips(120)),
            ..Default::default()
        });
        toc_paragraphs.push(title_p);

        for heading in &headings {
            let mut p = CT_P::new();

            // Indentation based on heading level (each level indented 360 twips = 0.25")
            let indent = Twips(360 * (heading.level as i32 - 1));

            p.properties = Some(CT_PPr {
                tabs: Some(right_tab.clone()),
                ind_left: if indent.0 > 0 { Some(indent) } else { None },
                ..Default::default()
            });

            // Run with heading text
            let text_run = CT_R::new(&heading.text);
            p.runs.push(text_run);

            // Tab run (separates text from page number)
            p.runs.push(CT_R {
                alt_drawings: Vec::new(),
                properties: None,
                content: vec![rdocx_oxml::text::RunContent::Tab],
                extra_xml: Vec::new(),
                extra_xml_positions: Vec::new(),
            });

            // Wrap the text run in a hyperlink to the bookmark
            p.hyperlinks.push(HyperlinkSpan {
                rel_id: None,
                anchor: Some(heading.bookmark_name.clone()),
                tooltip: None,
                doc_location: None,
                run_start: 0,
                run_end: 1, // Just the text run, not the tab
                extra_attributes: Vec::new(),
                extra_xml: Vec::new(),
                preserved_raw_before: None,
            });

            toc_paragraphs.push(p);
        }

        // Step 4: Insert TOC paragraphs at the specified index
        let insert_at = index.min(self.document.body.content.len());
        for (i, p) in toc_paragraphs.into_iter().enumerate() {
            self.document
                .body
                .content
                .insert(insert_at + i, BodyContent::Paragraph(p));
        }
        self.identifiers = identifiers;
    }

    /// Numeric `_TocN` bookmark suffixes already present in the body.
    fn toc_bookmark_suffixes(&self) -> HashSet<u64> {
        let mut suffixes = HashSet::new();
        visit_body_paragraphs(&self.document.body.content, &mut |p| {
            for marker in &p.bookmark_markers {
                let Some(name) = marker.name() else {
                    continue;
                };
                let Some(after) = name.strip_prefix("_Toc") else {
                    continue;
                };
                if let Ok(suffix) = after.parse::<u64>() {
                    suffixes.insert(suffix);
                }
            }
        });
        suffixes
    }

    /// Width of the text column in twips: page width less both side margins.
    ///
    /// Falls back to the US Letter default (6.5") when the section does not
    /// specify a size, and never returns a non-positive width.
    fn text_width_twips(&self) -> i32 {
        const DEFAULT_TEXT_WIDTH: i32 = 9360;

        let Some(sect) = self.document.body.sect_pr.as_ref() else {
            return DEFAULT_TEXT_WIDTH;
        };
        let page_width = sect.page_width.map(|w| w.0).unwrap_or(12240);
        let left = sect.margin_left.map(|m| m.0).unwrap_or(1440);
        let right = sect.margin_right.map(|m| m.0).unwrap_or(1440);

        let width = page_width - left - right;
        if width > 0 { width } else { DEFAULT_TEXT_WIDTH }
    }

    /// Detect heading level from a paragraph's style ID.
    fn detect_heading_level_for_toc(para: &CT_P) -> Option<u32> {
        let ppr = para.properties.as_ref()?;
        let style_id = ppr.style_id.as_deref()?;
        let rest = style_id.strip_prefix("Heading")?;
        rest.parse::<u32>().ok().filter(|n| (1..=9).contains(n))
    }

    // ---- Placeholder replacement ----

    /// Replace all occurrences of `placeholder` with `replacement` throughout the document.
    ///
    /// Searches body paragraphs, tables (including nested), headers, footers,
    /// text boxes and chart labels. Handles placeholders split across multiple
    /// runs. Returns the total number of replacements made.
    ///
    /// A `replacement` that contains `placeholder` is substituted once, not
    /// repeatedly.
    pub fn replace_text(&mut self, placeholder: &str, replacement: &str) -> usize {
        let mut candidate = self.clone_for_staging();
        let count = candidate
            .replace_batch(&[(placeholder, replacement)])
            .expect("text replacement package preflight failed");
        self.commit_staged_mutation(candidate);
        count
    }

    /// Replace multiple placeholders at once. Returns total replacements.
    ///
    /// Cheaper than calling [`Self::replace_text`] per entry: the document is
    /// serialised and re-parsed once for the whole batch rather than once per
    /// placeholder.
    pub fn replace_all(&mut self, replacements: &std::collections::HashMap<&str, &str>) -> usize {
        let pairs: Vec<(&str, &str)> = replacements.iter().map(|(k, v)| (*k, *v)).collect();
        let mut candidate = self.clone_for_staging();
        let count = candidate
            .replace_batch(&pairs)
            .expect("text replacement package preflight failed");
        self.commit_staged_mutation(candidate);
        count
    }

    /// Render scalar and structural template tags from structured JSON data.
    ///
    /// Tags use `{{ path.to.value }}` syntax and may cross ordinary Word run
    /// boundaries. String, number and boolean leaves render as text, while
    /// `null` renders as an empty string. Missing paths, malformed tags, and
    /// object or array leaves return an error without changing the document.
    /// Dedicated main-body paragraphs and table rows may contain nested
    /// `{% for item in path %}` or `{% if path %}` blocks with their matching
    /// end markers. Loop paths require arrays and introduce lexical scopes.
    /// Conditions use JSON truthiness, where false, null, zero, and empty
    /// strings or collections are false.
    ///
    /// This additive API is native-only. Python, WASM, and CLI binding surfaces
    /// remain unchanged and continue to preserve documents rendered here.
    pub fn render_template(&mut self, data: &serde_json::Value) -> Result<usize> {
        crate::template::render(self, data)
    }

    pub(crate) fn template_numbering_reference_exists(&self, num_id: u32, level: u32) -> bool {
        self.numbering
            .as_ref()
            .and_then(|numbering| numbering.get_abstract_num_for(num_id))
            .is_some_and(|abstract_numbering| {
                abstract_numbering
                    .levels
                    .iter()
                    .any(|candidate| candidate.ilvl == level)
            })
    }

    pub(crate) fn template_sources(&mut self) -> Result<Vec<String>> {
        self.flush_to_package()?;
        let mut sources = crate::template::body_sources(&self.document);

        for (rel_id, is_header) in self.header_footer_rel_ids() {
            if let Some(header_footer) = self.load_header_footer(&rel_id, is_header) {
                sources.extend(crate::template::header_footer_sources(&header_footer));
            }
        }

        for (part_name, _) in self.raw_text_bearing_part_names() {
            if let Some(xml) = self.package.get_part(&part_name) {
                sources.extend(crate::template::text_box_sources(xml)?);
            }
        }

        for part_name in self.chart_part_names() {
            if let Some(xml) = self.package.get_part(&part_name) {
                sources.extend(crate::template::chart_sources(xml)?);
            }
        }

        Ok(sources)
    }

    pub(crate) fn apply_template_pairs(&mut self, pairs: &[(&str, &str)]) -> Result<usize> {
        self.replace_batch(pairs)
    }

    pub(crate) fn commit_template(&mut self, mut candidate: Self) {
        candidate.invalidate_layout();
        self.commit_staged_mutation(candidate);
    }

    /// Apply a batch of literal replacements across the whole document.
    fn replace_batch(&mut self, pairs: &[(&str, &str)]) -> Result<usize> {
        if pairs.is_empty() {
            return Ok(0);
        }

        let mut count = 0;

        // Typed model: body content, then headers and footers.
        for (placeholder, replacement) in pairs {
            count += self.replace_in_body(placeholder, replacement);
        }
        count += self.replace_in_headers_footers(pairs)?;

        // Raw XML: text boxes, shapes and charts live in markup the typed model
        // does not cover, so flush first and work on the serialised parts.
        self.flush_to_package()?;
        count += self.replace_in_xml_parts(pairs);

        Ok(count)
    }

    /// Run the typed replacement over body paragraphs and tables.
    fn replace_in_body(&mut self, placeholder: &str, replacement: &str) -> usize {
        use rdocx_oxml::placeholder;

        let mut count = 0;
        for content in &mut self.document.body.content {
            match content {
                BodyContent::Paragraph(p) => {
                    count += placeholder::replace_in_paragraph(p, placeholder, replacement);
                }
                BodyContent::Table(t) => {
                    count += placeholder::replace_in_table(t, placeholder, replacement);
                }
                BodyContent::ContentControl(_) => {}
                BodyContent::RawXml(_) => {}
            }
        }
        count
    }

    /// Run the typed replacement over every referenced header and footer part.
    fn replace_in_headers_footers(&mut self, pairs: &[(&str, &str)]) -> Result<usize> {
        use rdocx_oxml::placeholder;

        let mut count = 0;
        for (rel_id, is_header) in self.header_footer_rel_ids() {
            let Some(mut hf) = self.load_header_footer(&rel_id, is_header) else {
                continue;
            };
            let mut part_count = 0;
            for (placeholder, replacement) in pairs {
                part_count +=
                    placeholder::replace_in_header_footer(&mut hf, placeholder, replacement);
            }
            if part_count > 0 {
                self.save_header_footer(&rel_id, &hf, is_header)?;
                count += part_count;
            }
        }
        Ok(count)
    }

    /// Relationship IDs of every section's headers and footers, with a flag
    /// saying which kind each one is.
    fn header_footer_rel_ids(&self) -> Vec<(String, bool)> {
        let mut rel_ids = Vec::new();
        let mut seen = HashSet::new();
        let mut collect = |section: &CT_SectPr| {
            for reference in &section.header_refs {
                let rel_id = (reference.rel_id.clone(), true);
                if seen.insert(rel_id.clone()) {
                    rel_ids.push(rel_id);
                }
            }
            for reference in &section.footer_refs {
                let rel_id = (reference.rel_id.clone(), false);
                if seen.insert(rel_id.clone()) {
                    rel_ids.push(rel_id);
                }
            }
        };
        visit_body_paragraphs(&self.document.body.content, &mut |paragraph| {
            if let Some(section) = paragraph
                .properties
                .as_ref()
                .and_then(|properties| properties.sect_pr.as_ref())
            {
                collect(section);
            }
        });
        if let Some(section) = &self.document.body.sect_pr {
            collect(section);
        }

        rel_ids
    }

    fn header_footer_rel_ids_for_layout(&self) -> HashSet<String> {
        let even_headers_enabled = self.even_headers_enabled();
        let mut rel_ids = HashSet::new();
        let sections = self
            .document
            .body
            .content
            .iter()
            .filter_map(|content| match content {
                BodyContent::Paragraph(paragraph) => paragraph
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.sect_pr.as_ref()),
                BodyContent::Table(_) | BodyContent::ContentControl(_) | BodyContent::RawXml(_) => {
                    None
                }
            })
            .chain(self.document.body.sect_pr.iter());
        for section in sections {
            for reference in section.header_refs.iter().chain(&section.footer_refs) {
                if reference.hdr_ftr_type != HdrFtrType::Even || even_headers_enabled {
                    rel_ids.insert(reference.rel_id.clone());
                }
            }
        }
        rel_ids
    }

    fn even_headers_enabled(&self) -> bool {
        let Some(settings) = self.settings.as_ref() else {
            return false;
        };
        settings
            .to_xml()
            .ok()
            .is_some_and(|xml| settings_enable_even_headers(&xml))
    }

    // ---- Regex replacement ----

    /// Replace all regex matches with `replacement` throughout the document.
    ///
    /// The `replacement` string supports capture groups: `$1`, `$2`, etc.
    /// Searches body paragraphs, tables (including nested), headers, and footers.
    /// Returns the total number of replacements made, or an error if the regex is invalid.
    pub fn replace_regex(&mut self, pattern: &str, replacement: &str) -> Result<usize> {
        let re =
            regex::Regex::new(pattern).map_err(|e| Error::Other(format!("invalid regex: {e}")))?;
        let mut candidate = self.clone_for_staging();
        let count = candidate.replace_regex_compiled(&re, replacement)?;
        self.commit_staged_mutation(candidate);
        Ok(count)
    }

    /// Replace multiple regex patterns at once. Returns total replacements.
    pub fn replace_all_regex(&mut self, patterns: &[(String, String)]) -> Result<usize> {
        let compiled = patterns
            .iter()
            .map(|(pattern, replacement)| {
                regex::Regex::new(pattern)
                    .map(|regex| (regex, replacement.as_str()))
                    .map_err(|error| Error::Other(format!("invalid regex: {error}")))
            })
            .collect::<Result<Vec<_>>>()?;
        let mut candidate = self.clone_for_staging();
        let mut count = 0;
        for (regex, replacement) in &compiled {
            count += candidate.replace_regex_compiled(regex, replacement)?;
        }
        self.commit_staged_mutation(candidate);
        Ok(count)
    }

    /// Internal: replace using a pre-compiled regex.
    fn replace_regex_compiled(&mut self, re: &regex::Regex, replacement: &str) -> Result<usize> {
        use rdocx_oxml::placeholder;

        let mut count = 0;

        // Replace in body paragraphs and tables
        for content in &mut self.document.body.content {
            match content {
                BodyContent::Paragraph(p) => {
                    count += placeholder::replace_regex_in_paragraph(p, re, replacement);
                }
                BodyContent::Table(t) => {
                    count += placeholder::replace_regex_in_table(t, re, replacement);
                }
                BodyContent::ContentControl(_) => {}
                BodyContent::RawXml(_) => {}
            }
        }

        // Replace in headers and footers
        for (rel_id, is_header) in self.header_footer_rel_ids() {
            let Some(mut hf) = self.load_header_footer(&rel_id, is_header) else {
                continue;
            };
            let n = placeholder::replace_regex_in_header_footer(&mut hf, re, replacement);
            if n > 0 {
                self.save_header_footer(&rel_id, &hf, is_header)?;
                count += n;
            }
        }

        // Text boxes and shapes live in raw markup the typed model does not
        // reach. `replace_text` has always covered them; do the same here so
        // the two entry points search the same places.
        self.flush_to_package()?;
        count += self.replace_regex_in_xml_parts(re, replacement);

        Ok(count)
    }

    /// Apply a regex replacement to the text-box content of the raw XML parts.
    fn replace_regex_in_xml_parts(&mut self, re: &regex::Regex, replacement: &str) -> usize {
        let mut count = 0;

        for (part_name, _) in self.text_bearing_part_names() {
            let Some(xml) = self.package.get_part(&part_name).map(<[u8]>::to_vec) else {
                continue;
            };
            if let Ok((new_xml, n)) =
                rdocx_oxml::placeholder::replace_regex_in_xml_part(&xml, re, replacement)
                && n > 0
            {
                self.package.set_part(&part_name, new_xml);
                count += n;
            }
        }

        // Re-parse so the in-memory model reflects the edited markup; otherwise
        // the next flush would write the pre-replacement document back out.
        if count > 0
            && let Some(doc_xml) = self.package.get_part(&self.doc_part_name)
            && let Ok(doc) = CT_Document::from_xml(doc_xml)
        {
            self.document = doc;
        }

        count
    }

    /// The main document part plus every header and footer part: everywhere
    /// text boxes and shapes with replaceable text can appear.
    fn text_bearing_part_names(&self) -> Vec<(String, Option<bool>)> {
        let mut names = vec![(self.doc_part_name.clone(), None)];
        for (rel_id, is_header) in self.header_footer_rel_ids() {
            if let Some(part_name) = self.header_footer_part_name(&rel_id, is_header) {
                names.push((part_name, Some(is_header)));
            }
        }
        names
    }

    fn raw_text_bearing_part_names(&self) -> Vec<(String, Option<bool>)> {
        let mut names = vec![(self.doc_part_name.clone(), None)];
        for (rel_id, is_header) in self.header_footer_rel_ids() {
            if let Some(part_name) = self.header_footer_part_name(&rel_id, is_header) {
                names.push((part_name, Some(is_header)));
            }
        }
        names
    }

    fn chart_part_names(&self) -> Vec<String> {
        self.package
            .get_part_rels(&self.doc_part_name)
            .map(|relationships| {
                relationships
                    .get_all_by_type(rel_types::CHART)
                    .iter()
                    .filter(|relationship| relationship_is_internal(relationship))
                    .map(|relationship| {
                        OpcPackage::resolve_rel_target(&self.doc_part_name, &relationship.target)
                    })
                    .collect()
            })
            .unwrap_or_default()
    }

    /// Load a header/footer part by its relationship ID.
    fn header_footer_part_name(&self, rel_id: &str, is_header: bool) -> Option<String> {
        let expected_type = if is_header {
            rel_types::HEADER
        } else {
            rel_types::FOOTER
        };
        let relationship = self
            .package
            .get_part_rels(&self.doc_part_name)?
            .get_by_id(rel_id)
            .filter(|relationship| {
                relationship.rel_type == expected_type && relationship_is_internal(relationship)
            })?;
        Some(OpcPackage::resolve_rel_target(
            &self.doc_part_name,
            &relationship.target,
        ))
    }

    /// Load a header/footer part by its relationship ID and section-reference kind.
    fn load_header_footer(&self, rel_id: &str, is_header: bool) -> Option<CT_HdrFtr> {
        let part_name = self.header_footer_part_name(rel_id, is_header)?;
        let xml = self.package.get_part(&part_name)?;
        CT_HdrFtr::from_xml(xml).ok()
    }

    /// Run raw XML replacement on all XML parts (for text boxes, shapes, charts, etc.).
    ///
    /// This is called after the typed-model replacement and flush_to_package.
    fn replace_in_xml_parts(&mut self, pairs: &[(&str, &str)]) -> usize {
        use rdocx_oxml::placeholder::{replace_many_in_chart_xml, replace_many_in_xml_part};

        let mut count = 0;

        // Collect part names for XML parts to process (text boxes/shapes)
        for (part_name, _) in self.raw_text_bearing_part_names() {
            if let Some(xml) = self.package.get_part(&part_name) {
                let xml = xml.to_vec();
                if let Ok((new_xml, n)) = replace_many_in_xml_part(&xml, pairs)
                    && n > 0
                {
                    self.package.set_part(&part_name, new_xml);
                    count += n;
                }
            }
        }

        // Collect chart part names
        for part_name in self.chart_part_names() {
            if let Some(xml) = self.package.get_part(&part_name) {
                let xml = xml.to_vec();
                if let Ok((new_xml, n)) = replace_many_in_chart_xml(&xml, pairs)
                    && n > 0
                {
                    self.package.set_part(&part_name, new_xml);
                    count += n;
                }
            }
        }

        // Re-parse document from the (possibly modified) package XML
        if count > 0
            && let Some(doc_xml) = self.package.get_part(&self.doc_part_name)
            && let Ok(doc) = CT_Document::from_xml(doc_xml)
        {
            self.document = doc;
        }

        count
    }

    // ---- Layout and PDF conversion ----

    /// Transfer compatible reusable normal-layout work from another document.
    ///
    /// The transfer succeeds only when every retained-work input other than
    /// body content matches exactly. A rejected transfer leaves both engines
    /// unchanged. Completed layout results remain owned by their documents.
    pub fn transfer_reusable_layout_from(&mut self, source: &mut Document) -> bool {
        let input = self.build_layout_input();
        let transferred = {
            let mut source_engine = source
                .normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner);
            rdocx_layout::engine::Engine::take_if_compatible(&mut source_engine, &input)
        };
        let Some(transferred) = transferred else {
            return false;
        };
        let mut receiver_engine = self
            .normal_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        *receiver_engine = Some(transferred);
        true
    }

    /// Transfer compatible bundled-fallback layout work from another document.
    ///
    /// The exact caller-font slice is part of the compatibility check. A
    /// rejected transfer leaves both private engines unchanged. This method
    /// never exposes the engine or permits system-font discovery.
    pub fn transfer_reusable_bundled_fallback_layout_from(
        &mut self,
        source: &mut Document,
        font_files: &[(&str, &[u8])],
    ) -> bool {
        let input = self.build_layout_input_with_fonts(font_files, RenderOptions::default());
        let transferred = {
            let mut source_engine = source
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner);
            rdocx_layout::engine::Engine::take_if_compatible(&mut source_engine, &input)
        };
        let Some(transferred) = transferred else {
            return false;
        };
        let mut receiver_engine = self
            .bundled_fallback_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        *receiver_engine = Some(transferred);
        true
    }

    /// Transfer compatible alias-aware bundled-fallback layout work.
    ///
    /// Caller-font bytes and the exact ordered alias slice are both part of
    /// compatibility. A rejected transfer preserves both private engines.
    pub fn transfer_reusable_bundled_fallback_layout_from_with_aliases(
        &mut self,
        source: &mut Document,
        font_files: &[(&str, &[u8])],
        font_aliases: &[(&str, &str)],
    ) -> bool {
        let input = self.build_layout_input_with_fonts(font_files, RenderOptions::default());
        let font_aliases = owned_font_aliases(font_aliases);
        let transferred = {
            let mut source_engine = source
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner);
            rdocx_layout::engine::Engine::take_if_compatible_with_caller_aliases(
                &mut source_engine,
                &input,
                &font_aliases,
            )
        };
        let Some(transferred) = transferred else {
            return false;
        };
        let mut receiver_engine = self
            .bundled_fallback_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        *receiver_engine = Some(transferred);
        true
    }

    /// Return the cached normal-font layout with its Word source map.
    ///
    /// Repeated calls share the same accepted-view result until the document
    /// is mutated.
    pub fn layout(&self) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
        self.layout_with_options(RenderOptions::default())
    }

    /// Return a normal-font layout with the selected revision view.
    ///
    /// Accepted-view calls share the normal layout cache. Tracked-view calls
    /// remain uncached because they do not replace the accepted-view cache.
    pub fn layout_with_options(
        &self,
        options: RenderOptions,
    ) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
        self.layout_for_options(options, false)
    }

    /// Return the bundled-font-only layout with its Word source map.
    ///
    /// Repeated accepted-view calls share the deterministic layout cache. This
    /// is the same snapshot used by deterministic PDF and raster render helpers.
    pub fn layout_deterministic(&self) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
        self.layout_deterministic_with_options(RenderOptions::default())
    }

    /// Return a bundled-font-only layout with the selected revision view.
    ///
    /// Accepted-view calls share the deterministic layout cache. Tracked-view
    /// calls remain uncached because they do not replace the accepted-view cache.
    pub fn layout_deterministic_with_options(
        &self,
        options: RenderOptions,
    ) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
        self.layout_for_options(options, true)
    }

    /// Return an uncached layout using user-provided font files.
    ///
    /// User-provided fonts take highest priority in font resolution. The
    /// returned owned result retains the exact font bytes and Word source map
    /// used by layout.
    pub fn layout_with_fonts(
        &self,
        font_files: &[(&str, &[u8])],
    ) -> Result<rdocx_layout::WordLayoutResult> {
        self.layout_with_fonts_and_options(font_files, RenderOptions::default())
    }

    /// Return an uncached caller-font layout with the selected revision view.
    pub fn layout_with_fonts_and_options(
        &self,
        font_files: &[(&str, &[u8])],
        options: RenderOptions,
    ) -> Result<rdocx_layout::WordLayoutResult> {
        let input = self.build_layout_input_with_fonts(font_files, options);
        #[cfg(test)]
        record_layout_invocation();
        Ok(rdocx_layout::layout_document_with_caller_fonts_and_provenance(&input)?)
    }

    /// Return an uncached layout using caller fonts over bundled fallbacks.
    ///
    /// Caller faces have highest priority. Missing families resolve from the
    /// deterministic bundled inventory, never from the system-font snapshot.
    /// The returned result owns its source map and shares its heavy page and
    /// font payloads internally.
    pub fn layout_with_fonts_and_bundled_fallback(
        &self,
        font_files: &[(&str, &[u8])],
    ) -> Result<rdocx_layout::WordLayoutResult> {
        self.layout_with_fonts_and_bundled_fallback_and_options(
            font_files,
            RenderOptions::default(),
        )
    }

    /// Return a bundled-fallback caller-font layout for one revision view.
    pub fn layout_with_fonts_and_bundled_fallback_and_options(
        &self,
        font_files: &[(&str, &[u8])],
        options: RenderOptions,
    ) -> Result<rdocx_layout::WordLayoutResult> {
        let input = self.build_layout_input_with_fonts(font_files, options);
        #[cfg(test)]
        record_layout_invocation();
        let mut engine = self
            .bundled_fallback_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        let engine = match engine.as_mut() {
            Some(engine) => engine,
            None => engine.insert(rdocx_layout::engine::Engine::new_deterministic()?),
        };
        engine.set_caller_font_aliases(&[]);
        Ok(rdocx_layout::layout_document_with_reusable_engine(
            engine, &input,
        )?)
    }

    /// Return a reusable bundled-fallback layout with byte-free font aliases.
    ///
    /// Exact embedded families retain priority. A caller alias is tried next,
    /// before the existing mapped and generic fallbacks.
    pub fn layout_with_fonts_aliases_and_bundled_fallback(
        &self,
        font_files: &[(&str, &[u8])],
        font_aliases: &[(&str, &str)],
    ) -> Result<rdocx_layout::WordLayoutResult> {
        self.layout_with_fonts_aliases_and_bundled_fallback_and_options(
            font_files,
            font_aliases,
            RenderOptions::default(),
        )
    }

    /// Return an alias-aware bundled-fallback layout for one revision view.
    pub fn layout_with_fonts_aliases_and_bundled_fallback_and_options(
        &self,
        font_files: &[(&str, &[u8])],
        font_aliases: &[(&str, &str)],
        options: RenderOptions,
    ) -> Result<rdocx_layout::WordLayoutResult> {
        let input = self.build_layout_input_with_fonts(font_files, options);
        let font_aliases = owned_font_aliases(font_aliases);
        #[cfg(test)]
        record_layout_invocation();
        let mut engine = self
            .bundled_fallback_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        let engine = match engine.as_mut() {
            Some(engine) => engine,
            None => engine.insert(rdocx_layout::engine::Engine::new_deterministic()?),
        };
        engine.set_caller_font_aliases(&font_aliases);
        Ok(rdocx_layout::layout_document_with_reusable_engine(
            engine, &input,
        )?)
    }

    fn build_layout_input_with_fonts(
        &self,
        font_files: &[(&str, &[u8])],
        options: RenderOptions,
    ) -> rdocx_layout::LayoutInput {
        let mut input = self.build_layout_input();
        input.revision_view = options.revision_view;
        input.fonts.extend(
            font_files
                .iter()
                .map(|(family, data)| rdocx_layout::FontFile {
                    family: (*family).to_owned(),
                    data: data.to_vec(),
                }),
        );
        input
    }

    /// Render the document to PDF bytes.
    ///
    /// This performs a full layout pass (font shaping, line breaking, pagination)
    /// and then renders the result to a PDF document.
    ///
    /// Font resolution order:
    /// 1. Fonts embedded in the DOCX file (word/fonts/)
    /// 2. System fonts when the default `system-fonts` feature is enabled
    /// 3. Always-available bundled metric-compatible fonts
    pub fn to_pdf(&self) -> Result<Vec<u8>> {
        self.to_pdf_with_options(RenderOptions::default())
    }

    /// Render the document to PDF bytes with the selected revision view.
    pub fn to_pdf_with_options(&self, options: RenderOptions) -> Result<Vec<u8>> {
        let layout = self.layout_with_options(options)?;
        Ok(oxml_pdf::render_to_pdf(&layout.layout))
    }

    /// Render the document to PDF bytes using bundled fonts without system
    /// font discovery.
    ///
    /// The deterministic layout is cached independently from the normal-font
    /// layout and is suitable for reproducible render baselines.
    pub fn to_pdf_deterministic(&self) -> Result<Vec<u8>> {
        self.to_pdf_deterministic_with_options(RenderOptions::default())
    }

    /// Render the document to the selected archival PDF profile using bundled fonts.
    pub fn to_pdfa_deterministic(&self, profile: oxml_pdf::PdfConformance) -> Result<Vec<u8>> {
        let layout = self.layout_for_options(RenderOptions::default(), true)?;
        Ok(oxml_pdf::render_to_pdf_with_options(
            &layout.layout,
            oxml_pdf::PdfOptions::new(profile),
        )?)
    }

    /// Render the selected revision view to deterministic PDF bytes.
    pub fn to_pdf_deterministic_with_options(&self, options: RenderOptions) -> Result<Vec<u8>> {
        let layout = self.layout_for_options(options, true)?;
        Ok(oxml_pdf::render_to_pdf(&layout.layout))
    }

    /// Render the document to PDF bytes with user-provided font files.
    ///
    /// User-provided fonts take highest priority in font resolution.
    ///
    /// # Arguments
    /// * `font_files` - Additional font files to use. Each entry is `(family_name, font_bytes)`.
    ///
    /// Font resolution order:
    /// 1. User-provided fonts (this parameter)
    /// 2. Fonts embedded in the DOCX file (word/fonts/)
    /// 3. System fonts when the default `system-fonts` feature is enabled
    /// 4. Always-available bundled metric-compatible fonts
    pub fn to_pdf_with_fonts(&self, font_files: &[(&str, &[u8])]) -> Result<Vec<u8>> {
        self.to_pdf_with_fonts_and_options(font_files, RenderOptions::default())
    }

    /// Render the selected revision view to PDF with user-provided fonts.
    pub fn to_pdf_with_fonts_and_options(
        &self,
        font_files: &[(&str, &[u8])],
        options: RenderOptions,
    ) -> Result<Vec<u8>> {
        let layout = self.layout_with_fonts_and_options(font_files, options)?;
        Ok(oxml_pdf::render_to_pdf(&layout.layout))
    }

    /// Save the document as a PDF file.
    pub fn save_pdf<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.save_pdf_with_options(path, RenderOptions::default())
    }

    /// Save the selected revision view as a PDF file.
    pub fn save_pdf_with_options<P: AsRef<Path>>(
        &self,
        path: P,
        options: RenderOptions,
    ) -> Result<()> {
        let pdf_bytes = self.to_pdf_with_options(options)?;
        std::fs::write(path, pdf_bytes)?;
        Ok(())
    }

    /// Convert the document to a complete HTML document string.
    pub fn to_html(&self) -> String {
        let input = self.build_html_input();
        rdocx_html::to_html_document(&input, &rdocx_html::HtmlOptions::default())
    }

    /// Convert the document to an HTML fragment (body content only, no `<html>` wrapper).
    pub fn to_html_fragment(&self) -> String {
        let input = self.build_html_input();
        rdocx_html::to_html_fragment(&input, &rdocx_html::HtmlOptions::default())
    }

    /// Convert the document to Markdown.
    pub fn to_markdown(&self) -> String {
        let input = self.build_html_input();
        rdocx_html::to_markdown(&input)
    }

    /// Build an HtmlInput from the document's current state.
    fn build_html_input(&self) -> rdocx_html::HtmlInput {
        use oxml_opc::relationship::rel_types;
        use std::collections::HashMap;

        let mut images: HashMap<String, rdocx_html::ImageData> = HashMap::new();
        let mut hyperlink_urls: HashMap<String, String> = HashMap::new();

        if let Some(rels) = self.package.get_part_rels(&self.doc_part_name) {
            for rel in &rels.items {
                match rel.rel_type.as_str() {
                    t if t == rel_types::IMAGE => {
                        if !relationship_is_internal(rel) {
                            continue;
                        }
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(data) = self.package.get_part(&part_name) {
                            let content_type = oxml_media::resolve(data, &part_name)
                                .content_type()
                                .to_owned();
                            images.insert(
                                rel.id.clone(),
                                rdocx_html::ImageData {
                                    data: data.to_vec(),
                                    content_type,
                                },
                            );
                        }
                    }
                    t if t == rel_types::HYPERLINK
                        && rel.target_mode.as_ref().is_some_and(|m| m == "External") =>
                    {
                        hyperlink_urls.insert(rel.id.clone(), rel.target.clone());
                    }
                    _ => {}
                }
            }
        }

        rdocx_html::HtmlInput {
            document: self.document.clone(),
            styles: self.styles.clone(),
            numbering: self.numbering.clone(),
            images,
            hyperlink_urls,
        }
    }

    /// Render a single page of the document to PNG bytes.
    ///
    /// # Arguments
    /// * `page_index` - 0-based page index
    /// * `dpi` - Resolution (72 = 1:1, 150 = standard, 300 = high quality)
    pub fn render_page_to_png(&self, page_index: usize, dpi: f64) -> Result<Option<Vec<u8>>> {
        self.render_page_to_png_with_options(page_index, dpi, RenderOptions::default())
    }

    /// Render one page to PNG with the selected revision view.
    pub fn render_page_to_png_with_options(
        &self,
        page_index: usize,
        dpi: f64,
        options: RenderOptions,
    ) -> Result<Option<Vec<u8>>> {
        let layout = self.layout_with_options(options)?;
        Ok(oxml_pdf::render_page_to_png(
            &layout.layout,
            page_index,
            dpi,
        ))
    }

    /// Render a single page to PNG using bundled fonts without system font
    /// discovery.
    ///
    /// # Arguments
    /// * `page_index` - 0-based page index
    /// * `dpi` - Resolution (72 = 1:1, 150 = standard, 300 = high quality)
    pub fn render_page_to_png_deterministic(
        &self,
        page_index: usize,
        dpi: f64,
    ) -> Result<Option<Vec<u8>>> {
        self.render_page_to_png_deterministic_with_options(
            page_index,
            dpi,
            RenderOptions::default(),
        )
    }

    /// Render one page to deterministic PNG with the selected revision view.
    pub fn render_page_to_png_deterministic_with_options(
        &self,
        page_index: usize,
        dpi: f64,
        options: RenderOptions,
    ) -> Result<Option<Vec<u8>>> {
        let layout = self.layout_for_options(options, true)?;
        Ok(oxml_pdf::render_page_to_png(
            &layout.layout,
            page_index,
            dpi,
        ))
    }

    /// Render one zero-based page as a self-contained searchable SVG document.
    ///
    /// An index beyond the laid-out document returns `None`. Layout diagnostics
    /// precede path-specific SVG lowering diagnostics in the returned result.
    pub fn render_page_to_svg(&self, page_index: usize) -> Result<Option<crate::SvgRenderResult>> {
        self.render_page_to_svg_with_options(page_index, RenderOptions::default())
    }

    /// Render one page as SVG with the selected revision view.
    pub fn render_page_to_svg_with_options(
        &self,
        page_index: usize,
        options: RenderOptions,
    ) -> Result<Option<crate::SvgRenderResult>> {
        let layout = self.layout_with_options(options)?;
        Ok(crate::svg::render_page(&layout.layout, page_index))
    }

    /// Render one page as SVG using bundled fonts without system font discovery.
    pub fn render_page_to_svg_deterministic(
        &self,
        page_index: usize,
    ) -> Result<Option<crate::SvgRenderResult>> {
        self.render_page_to_svg_deterministic_with_options(page_index, RenderOptions::default())
    }

    /// Render one revision view as deterministic SVG using bundled fonts.
    pub fn render_page_to_svg_deterministic_with_options(
        &self,
        page_index: usize,
        options: RenderOptions,
    ) -> Result<Option<crate::SvgRenderResult>> {
        let layout = self.layout_for_options(options, true)?;
        Ok(crate::svg::render_page(&layout.layout, page_index))
    }

    /// Render selected zero-based pages to the requested image format.
    pub fn render_pages(
        &self,
        page_indices: &[usize],
        raster_options: oxml_pdf::RasterOptions,
    ) -> Result<oxml_pdf::RasterOutput> {
        self.render_pages_with_options(page_indices, raster_options, RenderOptions::default())
    }

    /// Render selected zero-based pages with the selected revision view.
    pub fn render_pages_with_options(
        &self,
        page_indices: &[usize],
        raster_options: oxml_pdf::RasterOptions,
        options: RenderOptions,
    ) -> Result<oxml_pdf::RasterOutput> {
        let layout = self.layout_with_options(options)?;
        Ok(oxml_pdf::render_pages(
            &layout.layout,
            page_indices,
            raster_options,
        )?)
    }

    /// Render selected zero-based pages using bundled fonts without system font discovery.
    pub fn render_pages_deterministic(
        &self,
        page_indices: &[usize],
        raster_options: oxml_pdf::RasterOptions,
    ) -> Result<oxml_pdf::RasterOutput> {
        self.render_pages_deterministic_with_options(
            page_indices,
            raster_options,
            RenderOptions::default(),
        )
    }

    /// Render selected zero-based pages to deterministic images with the selected revision view.
    pub fn render_pages_deterministic_with_options(
        &self,
        page_indices: &[usize],
        raster_options: oxml_pdf::RasterOptions,
        options: RenderOptions,
    ) -> Result<oxml_pdf::RasterOutput> {
        let layout = self.layout_for_options(options, true)?;
        Ok(oxml_pdf::render_pages(
            &layout.layout,
            page_indices,
            raster_options,
        )?)
    }

    /// Render all pages of the document to PNG bytes.
    pub fn render_all_pages(&self, dpi: f64) -> Result<Vec<Vec<u8>>> {
        self.render_all_pages_with_options(dpi, RenderOptions::default())
    }

    /// Render every page to PNG with the selected revision view.
    pub fn render_all_pages_with_options(
        &self,
        dpi: f64,
        options: RenderOptions,
    ) -> Result<Vec<Vec<u8>>> {
        let layout = self.layout_with_options(options)?;
        Ok(oxml_pdf::render_all_pages(&layout.layout, dpi))
    }

    /// Return a cloned positioned page from the cached normal-font layout.
    ///
    /// `page_index` is zero-based. An index beyond the document returns `None`.
    pub fn layout_page(&self, page_index: usize) -> Result<Option<oxml_layout::PageFrame>> {
        self.layout_page_with_options(page_index, RenderOptions::default())
    }

    /// Return one positioned page from the selected revision view.
    pub fn layout_page_with_options(
        &self,
        page_index: usize,
        options: RenderOptions,
    ) -> Result<Option<oxml_layout::PageFrame>> {
        let layout = self.layout_with_options(options)?;
        Ok(layout
            .layout
            .pages
            .get(page_index)
            .map(|page| page.as_ref().clone()))
    }

    /// Build a LayoutInput from the document's current state.
    fn build_layout_input(&self) -> rdocx_layout::LayoutInput {
        use oxml_opc::relationship::rel_types;
        use rdocx_layout::{ImageData, LayoutInput};
        use std::collections::HashMap;

        let mut headers: HashMap<String, CT_HdrFtr> = HashMap::new();
        let mut footers: HashMap<String, CT_HdrFtr> = HashMap::new();
        let mut images: HashMap<String, ImageData> = HashMap::new();
        let mut charts = HashMap::new();
        let mut hyperlink_urls: HashMap<String, String> = HashMap::new();
        let mut footnotes = None;
        let mut endnotes = None;
        let active_header_footer_ids = self.header_footer_rel_ids_for_layout();
        let even_headers_enabled = self.even_headers_enabled();

        // Extract embedded fonts from the DOCX package
        let fonts = self.extract_embedded_fonts();

        if let Some(rels) = self.package.get_part_rels(&self.doc_part_name) {
            for rel in &rels.items {
                match rel.rel_type.as_str() {
                    t if t == rel_types::HEADER => {
                        if !active_header_footer_ids.contains(&rel.id)
                            || !relationship_is_internal(rel)
                        {
                            continue;
                        }
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(xml) = self.package.get_part(&part_name)
                            && let Ok(hf) = CT_HdrFtr::from_xml(xml)
                        {
                            headers.insert(rel.id.clone(), hf);
                        }
                        if let Some(header_relationships) = self.package.get_part_rels(&part_name) {
                            for image_relationship in
                                header_relationships.items.iter().filter(|item| {
                                    item.rel_type == rel_types::IMAGE
                                        && relationship_is_internal(item)
                                })
                            {
                                let image_part = OpcPackage::resolve_rel_target(
                                    &part_name,
                                    &image_relationship.target,
                                );
                                if let Some(data) = self.package.get_part(&image_part) {
                                    images.insert(
                                        format!("{}\0{}", rel.id, image_relationship.id),
                                        ImageData {
                                            data: data.to_vec(),
                                            content_type: oxml_media::resolve(data, &image_part)
                                                .content_type()
                                                .to_owned(),
                                        },
                                    );
                                }
                            }
                        }
                    }
                    t if t == rel_types::FOOTER => {
                        if !active_header_footer_ids.contains(&rel.id)
                            || !relationship_is_internal(rel)
                        {
                            continue;
                        }
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(xml) = self.package.get_part(&part_name)
                            && let Ok(hf) = CT_HdrFtr::from_xml(xml)
                        {
                            footers.insert(rel.id.clone(), hf);
                        }
                    }
                    t if t == rel_types::IMAGE => {
                        if !relationship_is_internal(rel) {
                            continue;
                        }
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(data) = self.package.get_part(&part_name) {
                            let content_type = oxml_media::resolve(data, &part_name)
                                .content_type()
                                .to_owned();
                            images.insert(
                                rel.id.clone(),
                                ImageData {
                                    data: data.to_vec(),
                                    content_type,
                                },
                            );
                        }
                    }
                    t if t == rel_types::CHART => {
                        let chart = if !relationship_is_internal(rel) {
                            Err(format!("non-internal target {}", rel.target))
                        } else {
                            let part_name =
                                OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                            match self.package.get_part(&part_name) {
                                Some(xml) => {
                                    CT_ChartSpace::from_xml(xml).map(Box::new).map_err(|error| {
                                        format!("malformed target {part_name}: {error}")
                                    })
                                }
                                None => Err(format!("missing target {part_name}")),
                            }
                        };
                        charts.insert(rel.id.clone(), chart);
                    }
                    t if t == rel_types::HYPERLINK => {
                        if rel.target_mode.as_ref().is_some_and(|m| m == "External") {
                            hyperlink_urls.insert(rel.id.clone(), rel.target.clone());
                        }
                    }
                    t if t == rel_types::FOOTNOTES => {
                        if !relationship_is_internal(rel) {
                            continue;
                        }
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(xml) = self.package.get_part(&part_name) {
                            footnotes = rdocx_oxml::footnotes::CT_Footnotes::from_xml(xml).ok();
                        }
                    }
                    t if t == rel_types::ENDNOTES => {
                        if !relationship_is_internal(rel) {
                            continue;
                        }
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(xml) = self.package.get_part(&part_name) {
                            endnotes = rdocx_oxml::footnotes::CT_Footnotes::from_xml(xml).ok();
                        }
                    }
                    _ => {}
                }
            }
        }

        let chart_theme = if self.theme_dirty {
            self.theme.clone()
        } else {
            self.package
                .get_part_rels(&self.doc_part_name)
                .and_then(|relationships| {
                    relationships.items.iter().rev().find(|relationship| {
                        relationship.rel_type == rel_types::THEME
                            && relationship_is_internal(relationship)
                    })
                })
                .map(|relationship| {
                    OpcPackage::resolve_rel_target(&self.doc_part_name, &relationship.target)
                })
                .as_deref()
                .and_then(|part_name| self.package.get_part(part_name))
                .and_then(|xml| oxml_drawing::theme::CT_OfficeStyleSheet::from_xml(xml).ok())
                .or_else(|| self.theme.clone())
        }
        .unwrap_or_else(oxml_drawing::theme::CT_OfficeStyleSheet::office_default);
        let theme = Some(rdocx_oxml::theme::Theme::from(&chart_theme));

        let mut document = self.document.clone();
        if !even_headers_enabled {
            for content in &mut document.body.content {
                if let BodyContent::Paragraph(paragraph) = content
                    && let Some(section) = paragraph
                        .properties
                        .as_mut()
                        .and_then(|properties| properties.sect_pr.as_mut())
                {
                    section
                        .header_refs
                        .retain(|reference| reference.hdr_ftr_type != HdrFtrType::Even);
                    section
                        .footer_refs
                        .retain(|reference| reference.hdr_ftr_type != HdrFtrType::Even);
                }
            }
            if let Some(section) = document.body.sect_pr.as_mut() {
                section
                    .header_refs
                    .retain(|reference| reference.hdr_ftr_type != HdrFtrType::Even);
                section
                    .footer_refs
                    .retain(|reference| reference.hdr_ftr_type != HdrFtrType::Even);
            }
        }
        materialize_header_footer_inheritance(
            &mut document,
            even_headers_enabled,
            &headers.keys().cloned().collect(),
            &footers.keys().cloned().collect(),
        );

        LayoutInput {
            revision_view: rdocx_layout::RevisionView::Accepted,
            document,
            automatic_hyphenation: self
                .settings
                .as_ref()
                .is_some_and(CT_Settings::automatic_hyphenation),
            math_properties: self
                .settings
                .as_ref()
                .and_then(CT_Settings::math_properties)
                .cloned(),
            styles: self.styles.clone(),
            numbering: self.numbering.clone(),
            headers,
            footers,
            images,
            charts,
            chart_theme,
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: self.core_properties.clone(),
            hyperlink_urls,
            footnotes,
            endnotes,
            theme,
            fonts,
        }
    }

    /// Extract embedded fonts from the DOCX package.
    ///
    /// Word can embed fonts as `.odttf` (obfuscated TrueType) or regular `.ttf`/`.otf`
    /// files in the `word/fonts/` directory. ODTTF files have the first 32 bytes
    /// XOR'd with a 16-byte GUID derived from the font's relationship ID.
    fn extract_embedded_fonts(&self) -> Vec<rdocx_layout::FontFile> {
        if let (Some(table), Some(owner)) = (&self.font_table, &self.font_table_part_name) {
            let relationships = self.package.get_part_rels(owner);
            return table
                .fonts()
                .iter()
                .flat_map(|font| {
                    font.embedded_fonts.iter().filter_map(|embedded| {
                        let relationship = relationships?.items.iter().find(|relationship| {
                            relationship.id == embedded.relationship_id
                                && relationship.rel_type == rel_types::FONT
                                && relationship_is_internal(relationship)
                        })?;
                        let part_name = OpcPackage::resolve_rel_target(owner, &relationship.target);
                        let data = self.package.get_part(&part_name)?;
                        let data = deobfuscate_odttf_with_key(data, &embedded.font_key)?;
                        Some(rdocx_layout::FontFile {
                            family: font.name.clone(),
                            data,
                        })
                    })
                })
                .collect();
        }

        let mut fonts = Vec::new();

        // Look for font parts in word/fonts/ directory
        for (part_name, data) in &self.package.parts {
            let lower = part_name.to_lowercase();
            if !lower.contains("/word/fonts/") && !lower.contains("/word/font") {
                continue;
            }

            // Determine font family name from the file name
            let file_name = part_name.rsplit('/').next().unwrap_or(part_name);
            let family = file_name.split('.').next().unwrap_or(file_name).to_string();

            if lower.ends_with(".odttf") {
                // Deobfuscate ODTTF: XOR first 32 bytes with GUID from the file name
                if let Some(deobfuscated) = deobfuscate_odttf(data, file_name) {
                    fonts.push(rdocx_layout::FontFile {
                        family,
                        data: deobfuscated,
                    });
                }
            } else if lower.ends_with(".ttf") || lower.ends_with(".otf") || lower.ends_with(".ttc")
            {
                fonts.push(rdocx_layout::FontFile {
                    family,
                    data: data.clone(),
                });
            }
        }

        fonts
    }

    /// Load font files from a directory and return them as FontFile entries.
    ///
    /// This is useful for CLI tools that accept a `--font-dir` argument.
    /// Supports `.ttf`, `.otf`, and `.ttc` files.
    pub fn load_fonts_from_dir<P: AsRef<Path>>(dir: P) -> Vec<rdocx_layout::FontFile> {
        let mut fonts = Vec::new();
        let dir = dir.as_ref();
        if let Ok(entries) = std::fs::read_dir(dir) {
            for entry in entries.flatten() {
                let path = entry.path();
                let ext = path
                    .extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("")
                    .to_lowercase();
                if (ext == "ttf" || ext == "otf" || ext == "ttc")
                    && let Ok(data) = std::fs::read(&path)
                {
                    let family = path
                        .file_stem()
                        .and_then(|s| s.to_str())
                        .unwrap_or("Unknown")
                        .to_string();
                    fonts.push(rdocx_layout::FontFile { family, data });
                }
            }
        }
        fonts
    }

    /// Save a header/footer part back to the OPC package.
    fn save_header_footer(&mut self, rel_id: &str, hf: &CT_HdrFtr, is_header: bool) -> Result<()> {
        #[cfg(test)]
        if FAIL_NEXT_HEADER_FOOTER_SERIALIZATION.replace(false) {
            return Err(Error::Other(
                "injected header or footer serialization failure".to_owned(),
            ));
        }
        let part_name = self
            .header_footer_part_name(rel_id, is_header)
            .ok_or_else(|| {
                Error::Other(format!(
                    "header or footer relationship {rel_id} is missing, non-internal, or has the wrong type"
                ))
            })?;
        let xml = Self::serialize_hdr_ftr(hf, is_header)?;
        self.package.set_part(&part_name, xml);
        Ok(())
    }

    // ---- Document Intelligence API ----

    /// Get all headings in the document as (level, text) pairs.
    ///
    /// Detects heading paragraphs by their style ID (e.g. "Heading1", "Heading2").
    pub fn headings(&self) -> Vec<(u32, String)> {
        let mut result = Vec::new();
        for content in &self.document.body.content {
            if let BodyContent::Paragraph(p) = content
                && let Some(level) = Self::detect_heading_level_for_toc(p)
            {
                result.push((level, p.text()));
            }
        }
        result
    }

    /// Get a hierarchical outline of the document headings.
    ///
    /// Returns a tree structure where each node contains the heading level,
    /// text, and children (sub-headings).
    pub fn document_outline(&self) -> Vec<OutlineNode> {
        let headings = self.headings();
        build_outline_tree(&headings)
    }

    /// Get information about all images in the document.
    ///
    /// Returns metadata for each inline and anchored image found in body paragraphs.
    pub fn images(&self) -> Vec<ImageInfo> {
        let mut result = Vec::new();

        for content in &self.document.body.content {
            Self::collect_images_from_content(content, &mut result);
        }
        result
    }

    fn collect_images_from_content(content: &BodyContent, result: &mut Vec<ImageInfo>) {
        match content {
            BodyContent::Paragraph(p) => Self::collect_images_from_paragraph(p, result),
            BodyContent::Table(tbl) => Self::collect_images_from_table(tbl, result),
            BodyContent::ContentControl(_) => {}
            BodyContent::RawXml(_) => {}
        }
    }

    fn collect_images_from_paragraph(p: &CT_P, result: &mut Vec<ImageInfo>) {
        for run in &p.runs {
            for rc in &run.content {
                let RunContent::Drawing(drawing) = rc else {
                    continue;
                };
                if let Some(inline) = &drawing.inline {
                    result.push(ImageInfo {
                        embed_id: inline.embed_id.clone(),
                        name: inline.name.clone(),
                        description: inline.description.clone(),
                        width_emu: inline.extent_cx.0,
                        height_emu: inline.extent_cy.0,
                        is_anchor: false,
                    });
                }
                if let Some(anchor) = &drawing.anchor {
                    result.push(ImageInfo {
                        embed_id: anchor.embed_id.clone(),
                        name: anchor.name.clone(),
                        description: anchor.description.clone(),
                        width_emu: anchor.extent_cx.0,
                        height_emu: anchor.extent_cy.0,
                        is_anchor: true,
                    });
                }
            }
        }
    }

    fn collect_images_from_table(tbl: &CT_Tbl, result: &mut Vec<ImageInfo>) {
        use rdocx_oxml::table::CellContent;

        for row in &tbl.rows {
            for cell in &row.cells {
                for cc in &cell.content {
                    match cc {
                        CellContent::Paragraph(p) => Self::collect_images_from_paragraph(p, result),
                        CellContent::Table(nested) => {
                            Self::collect_images_from_table(nested, result)
                        }
                        CellContent::ContentControl(_) => {}
                    }
                }
            }
        }
    }

    /// Get information about all hyperlinks in the document.
    ///
    /// Resolves hyperlink relationship IDs to their target URLs where possible.
    pub fn links(&self) -> Vec<LinkInfo> {
        use oxml_opc::relationship::rel_types;

        // Build a map of hyperlink rel_id -> target URL
        let mut url_map = std::collections::HashMap::new();
        if let Some(rels) = self.package.get_part_rels(&self.doc_part_name) {
            for rel in &rels.items {
                if rel.rel_type == rel_types::HYPERLINK
                    && rel.target_mode.as_ref().is_some_and(|m| m == "External")
                {
                    url_map.insert(rel.id.clone(), rel.target.clone());
                }
            }
        }

        let mut result = Vec::new();
        for content in &self.document.body.content {
            if let BodyContent::Paragraph(p) = content {
                for hl in &p.hyperlinks {
                    // `HyperlinkSpan`'s bounds are public and can be set by
                    // hand, so clamp rather than slice-panic on a bad range.
                    let start = hl.run_start.min(p.runs.len());
                    let end = hl.run_end.clamp(start, p.runs.len());
                    let text: String = p.runs[start..end].iter().map(|r| r.text()).collect();

                    let url = hl.rel_id.as_ref().and_then(|id| url_map.get(id)).cloned();

                    result.push(LinkInfo {
                        text,
                        url,
                        anchor: hl.anchor.clone(),
                        rel_id: hl.rel_id.clone(),
                    });
                }
                for field in p.complex_field_hyperlinks() {
                    let start = field.run_start.min(p.runs.len());
                    let end = field.run_end.clamp(start, p.runs.len());
                    let text: String = p.runs[start..end].iter().map(|run| run.text()).collect();
                    result.push(LinkInfo {
                        text,
                        url: Some(field.target),
                        anchor: None,
                        rel_id: None,
                    });
                }
            }
        }
        result
    }

    /// Count the number of words in the document.
    ///
    /// Counts whitespace-separated tokens across all paragraphs (including
    /// paragraphs inside table cells).
    pub fn word_count(&self) -> usize {
        let mut count = 0;
        for content in &self.document.body.content {
            count += Self::word_count_in_content(content);
        }
        count
    }

    fn word_count_in_content(content: &BodyContent) -> usize {
        match content {
            BodyContent::Paragraph(p) => p.text().split_whitespace().count(),
            BodyContent::Table(tbl) => Self::word_count_in_table(tbl),
            BodyContent::ContentControl(_) => 0,
            BodyContent::RawXml(_) => 0,
        }
    }

    fn word_count_in_table(tbl: &CT_Tbl) -> usize {
        use rdocx_oxml::table::CellContent;

        let mut count = 0;
        for row in &tbl.rows {
            for cell in &row.cells {
                for cc in &cell.content {
                    match cc {
                        CellContent::Paragraph(p) => {
                            count += p.text().split_whitespace().count();
                        }
                        CellContent::Table(nested) => {
                            count += Self::word_count_in_table(nested);
                        }
                        CellContent::ContentControl(_) => {}
                    }
                }
            }
        }
        count
    }

    /// Audit the document for accessibility issues.
    ///
    /// Checks for common problems: missing image alt text, heading level gaps,
    /// empty paragraphs, missing document metadata.
    pub fn audit_accessibility(&self) -> Vec<AccessibilityIssue> {
        let mut issues = Vec::new();

        // Check: missing document title
        if self.title().is_none() {
            issues.push(AccessibilityIssue {
                severity: IssueSeverity::Warning,
                message: "Document has no title".to_string(),
            });
        }

        // Check: missing document language (author as a proxy for basic metadata)
        if self.author().is_none() {
            issues.push(AccessibilityIssue {
                severity: IssueSeverity::Info,
                message: "Document has no author".to_string(),
            });
        }

        // Check: images without alt text
        let images = self.images();
        for img in &images {
            let has_alt = img
                .description
                .as_ref()
                .is_some_and(|d| !d.is_empty() && d != "Background");
            if !has_alt {
                let name = img
                    .name
                    .as_deref()
                    .or(Some(&img.embed_id))
                    .unwrap_or("unknown");
                issues.push(AccessibilityIssue {
                    severity: IssueSeverity::Error,
                    message: format!("Image \"{name}\" has no alt text"),
                });
            }
        }

        // Check: heading level gaps
        let headings = self.headings();
        let mut prev_level: Option<u32> = None;
        for (level, text) in &headings {
            if let Some(prev) = prev_level
                && *level > prev + 1
            {
                issues.push(AccessibilityIssue {
                    severity: IssueSeverity::Warning,
                    message: format!(
                        "Heading level gap: h{prev} -> h{level} (\"{}\")",
                        truncate_str(text, 40)
                    ),
                });
            }
            prev_level = Some(*level);
        }

        // Check: excessive empty paragraphs
        let mut consecutive_empty = 0u32;
        for content in &self.document.body.content {
            if let BodyContent::Paragraph(p) = content {
                if p.text().trim().is_empty() {
                    consecutive_empty += 1;
                    if consecutive_empty >= 3 {
                        issues.push(AccessibilityIssue {
                            severity: IssueSeverity::Info,
                            message: format!(
                                "{consecutive_empty} consecutive empty paragraphs (consider using spacing instead)"
                            ),
                        });
                    }
                } else {
                    consecutive_empty = 0;
                }
            } else {
                consecutive_empty = 0;
            }
        }

        issues
    }
}

pub(crate) fn write_atomic_file(
    path: &Path,
    bytes: &[u8],
    invalid_name_message: &'static str,
    exhausted_message: &'static str,
) -> std::io::Result<()> {
    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let file_name = path.file_name().ok_or_else(|| {
        std::io::Error::new(std::io::ErrorKind::InvalidInput, invalid_name_message)
    })?;
    for attempt in 0..128_u8 {
        let mut temporary_name = std::ffi::OsString::from(".");
        temporary_name.push(file_name);
        temporary_name.push(format!(".rdocx-{}-{attempt}.tmp", std::process::id()));
        let temporary = parent.join(temporary_name);
        let mut file = match std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temporary)
        {
            Ok(file) => file,
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error),
        };
        let result = std::io::Write::write_all(&mut file, bytes).and_then(|()| file.sync_all());
        drop(file);
        let result = result.and_then(|()| replace_file(&temporary, path));
        if result.is_err() {
            let _ = std::fs::remove_file(&temporary);
        }
        return result;
    }
    Err(std::io::Error::new(
        std::io::ErrorKind::AlreadyExists,
        exhausted_message,
    ))
}

#[cfg(not(target_os = "windows"))]
pub(crate) fn replace_file(source: &Path, destination: &Path) -> std::io::Result<()> {
    std::fs::rename(source, destination)
}

#[cfg(target_os = "windows")]
pub(crate) fn replace_file(source: &Path, destination: &Path) -> std::io::Result<()> {
    use std::os::windows::ffi::OsStrExt;

    const MOVEFILE_REPLACE_EXISTING: u32 = 0x1;
    const MOVEFILE_WRITE_THROUGH: u32 = 0x8;

    #[link(name = "kernel32")]
    unsafe extern "system" {
        fn MoveFileExW(
            existing_file_name: *const u16,
            new_file_name: *const u16,
            flags: u32,
        ) -> i32;
    }

    let source: Vec<u16> = source.as_os_str().encode_wide().chain(Some(0)).collect();
    let destination: Vec<u16> = destination
        .as_os_str()
        .encode_wide()
        .chain(Some(0))
        .collect();
    // SAFETY: both path buffers are NUL-terminated and remain alive for the call.
    let replaced = unsafe {
        MoveFileExW(
            source.as_ptr(),
            destination.as_ptr(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH,
        )
    };
    if replaced == 0 {
        Err(std::io::Error::last_os_error())
    } else {
        Ok(())
    }
}

impl Default for Document {
    fn default() -> Self {
        Self::new()
    }
}

fn next_toc_bookmark_suffix(occupied: &HashSet<u64>, after: u64) -> Option<u64> {
    after
        .checked_add(1)
        .filter(|candidate| !occupied.contains(candidate))
        .or_else(|| (1..=u64::MAX).find(|candidate| !occupied.contains(candidate)))
}

fn chart_with_workbook_relationship(
    chart: &CT_ChartSpace,
    relationship_id: &str,
) -> Result<Vec<u8>> {
    const END: &[u8] = b"</c:chartSpace>";
    let mut xml = chart
        .to_xml()
        .map_err(|error| Error::Other(format!("invalid chart part: {error}")))?;
    if chart_has_external_data(&xml)? {
        return Err(Error::Other(
            "chart already contains an external workbook relationship".to_owned(),
        ));
    }
    if !xml.ends_with(END) {
        return Err(Error::Other(
            "serialized chart lacks the chartSpace closing element".to_owned(),
        ));
    }

    xml.truncate(xml.len() - END.len());
    xml.extend_from_slice(
        format!(
            r#"<c:externalData r:id="{relationship_id}"><c:autoUpdate val="0"/></c:externalData>"#
        )
        .as_bytes(),
    );
    xml.extend_from_slice(END);
    CT_ChartSpace::from_xml(&xml)
        .and_then(|validated| validated.to_xml())
        .map_err(|error| Error::Other(format!("invalid chart part: {error}")))
}

fn chart_has_external_data(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut chart_space_depth = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid chart part XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                if chart_space_depth.is_none()
                    && chart_namespace_matches(&namespace)
                    && matches_local_name(element.name().as_ref(), b"chartSpace")
                {
                    chart_space_depth = Some(depth);
                } else if chart_space_depth.is_some_and(|root| depth == root + 1)
                    && chart_namespace_matches(&namespace)
                    && matches_local_name(element.name().as_ref(), b"externalData")
                {
                    return Ok(true);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::Other("chart part XML depth overflow".to_owned()))?;
            }
            Event::Empty(ref element)
                if chart_space_depth.is_some_and(|root| depth == root + 1)
                    && chart_namespace_matches(&namespace)
                    && matches_local_name(element.name().as_ref(), b"externalData") =>
            {
                return Ok(true);
            }
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn chart_namespace_matches(namespace: &ResolveResult<'_>) -> bool {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => *uri == oxml_chart::C_NS.as_bytes(),
        ResolveResult::Unknown(prefix) => *prefix == b"c",
        ResolveResult::Unbound => false,
    }
}

fn settings_enable_even_headers(xml: &[u8]) -> bool {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        let Ok((namespace, event)) = reader.read_resolved_event_into(&mut buffer) else {
            return false;
        };
        match event {
            Event::Start(ref element) => {
                if depth == 1
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == rdocx_oxml::namespace::W_NS.as_bytes())
                    && matches_local_name(element.name().as_ref(), b"evenAndOddHeaders")
                {
                    return word_on_off_value(&reader, element);
                }
                depth += 1;
            }
            Event::Empty(ref element)
                if depth == 1
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == rdocx_oxml::namespace::W_NS.as_bytes())
                    && matches_local_name(element.name().as_ref(), b"evenAndOddHeaders") =>
            {
                return word_on_off_value(&reader, element);
            }
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => return false,
            _ => {}
        }
        buffer.clear();
    }
}

fn header_type_index(hdr_type: HdrFtrType) -> usize {
    match hdr_type {
        HdrFtrType::Default => 0,
        HdrFtrType::First => 1,
        HdrFtrType::Even => 2,
    }
}

fn materialize_header_footer_inheritance(
    document: &mut CT_Document,
    even_headers_enabled: bool,
    available_headers: &HashSet<String>,
    available_footers: &HashSet<String>,
) {
    let mut effective_headers: [Option<HdrFtrRef>; 3] = [None, None, None];
    let mut effective_footers: [Option<HdrFtrRef>; 3] = [None, None, None];
    let inherit = |references: &mut Vec<HdrFtrRef>,
                   effective: &mut [Option<HdrFtrRef>; 3],
                   available: &HashSet<String>,
                   materialize_blank_even: bool| {
        for hdr_type in [HdrFtrType::Default, HdrFtrType::First, HdrFtrType::Even] {
            let index = header_type_index(hdr_type);
            if let Some(reference) = references
                .iter()
                .find(|reference| {
                    reference.hdr_ftr_type == hdr_type && available.contains(&reference.rel_id)
                })
                .cloned()
            {
                effective[index] = Some(reference);
            } else if let Some(reference) = effective[index].clone() {
                references.push(reference);
            } else if hdr_type == HdrFtrType::Even && materialize_blank_even {
                references.push(HdrFtrRef {
                    hdr_ftr_type: HdrFtrType::Even,
                    rel_id: String::new(),
                });
            }
        }
    };
    for content in &mut document.body.content {
        if let BodyContent::Paragraph(paragraph) = content
            && let Some(section) = paragraph
                .properties
                .as_mut()
                .and_then(|properties| properties.sect_pr.as_mut())
        {
            inherit(
                &mut section.header_refs,
                &mut effective_headers,
                available_headers,
                even_headers_enabled,
            );
            inherit(
                &mut section.footer_refs,
                &mut effective_footers,
                available_footers,
                false,
            );
        }
    }
    if let Some(section) = document.body.sect_pr.as_mut() {
        inherit(
            &mut section.header_refs,
            &mut effective_headers,
            available_headers,
            even_headers_enabled,
        );
        inherit(
            &mut section.footer_refs,
            &mut effective_footers,
            available_footers,
            false,
        );
    }
}

fn word_on_off_value(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> bool {
    element
        .attributes()
        .flatten()
        .find_map(|attribute| {
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == rdocx_oxml::namespace::W_NS.as_bytes())
                && local.as_ref() == b"val"
            {
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                    .ok()
                    .map(|value| !matches!(value.as_ref(), "0" | "false" | "off"))
            } else {
                None
            }
        })
        .unwrap_or(true)
}

/// Express `target_part` relative to the directory holding `source_part`.
///
/// Falls back to the absolute part name when the two live in different
/// directories, which OPC also permits.
fn relative_target(source_part: &str, target_part: &str) -> String {
    let dir = match source_part.rfind('/') {
        Some(pos) => &source_part[..=pos],
        None => "/",
    };
    match target_part.strip_prefix(dir) {
        Some(rest) if !rest.contains('/') => rest.to_string(),
        _ => target_part.to_string(),
    }
}

fn relative_descendant_target(source_part: &str, target_part: &str) -> String {
    let directory = source_part
        .rfind('/')
        .map_or("/", |position| &source_part[..=position]);
    target_part
        .strip_prefix(directory)
        .unwrap_or(target_part)
        .to_owned()
}

fn section_has_layout(properties: &CT_SectPr) -> bool {
    properties.page_width.is_some()
        || properties.page_height.is_some()
        || properties.orientation.is_some()
        || properties.margin_top.is_some()
        || properties.margin_right.is_some()
        || properties.margin_bottom.is_some()
        || properties.margin_left.is_some()
        || properties.gutter.is_some()
        || properties.header_distance.is_some()
        || properties.footer_distance.is_some()
        || properties.section_type.is_some()
        || properties.columns.is_some()
        || properties.title_pg.is_some()
}

/// Numbering format for one level of a custom list definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ListNumberFormat {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
    Ordinal,
    CardinalText,
    OrdinalText,
    Hex,
    Chicago,
    IdeographDigital,
    JapaneseCounting,
    Aiueo,
    Iroha,
    DecimalFullWidth,
    DecimalHalfWidth,
    JapaneseLegal,
    JapaneseDigitalTenThousand,
    DecimalEnclosedCircle,
    DecimalFullWidth2,
    AiueoFullWidth,
    IrohaFullWidth,
    DecimalZero,
    Bullet,
    Ganada,
    Chosung,
    DecimalEnclosedFullstop,
    DecimalEnclosedParen,
    DecimalEnclosedCircleChinese,
    IdeographEnclosedCircle,
    IdeographTraditional,
    IdeographZodiac,
    IdeographZodiacTraditional,
    TaiwaneseCounting,
    IdeographLegalTraditional,
    TaiwaneseCountingThousand,
    TaiwaneseDigital,
    ChineseCounting,
    ChineseLegalSimplified,
    ChineseCountingThousand,
    KoreanDigital,
    KoreanCounting,
    KoreanLegal,
    KoreanDigital2,
    Hebrew1,
    ArabicAlpha,
    Hebrew2,
    ArabicAbjad,
    HindiVowels,
    HindiConsonants,
    HindiNumbers,
    HindiCounting,
    ThaiLetters,
    ThaiNumbers,
    ThaiCounting,
    VietnameseCounting,
    NumberInDash,
    RussianLower,
    RussianUpper,
    None,
    /// A producer-defined format retained across inspection and mutation.
    Other(String),
}

impl ListNumberFormat {
    fn to_st(&self) -> ST_NumberFormat {
        match self {
            Self::Decimal => ST_NumberFormat::Decimal,
            Self::UpperRoman => ST_NumberFormat::UpperRoman,
            Self::LowerRoman => ST_NumberFormat::LowerRoman,
            Self::UpperLetter => ST_NumberFormat::UpperLetter,
            Self::LowerLetter => ST_NumberFormat::LowerLetter,
            Self::Ordinal => ST_NumberFormat::Ordinal,
            Self::CardinalText => ST_NumberFormat::CardinalText,
            Self::OrdinalText => ST_NumberFormat::OrdinalText,
            Self::Hex => ST_NumberFormat::Hex,
            Self::Chicago => ST_NumberFormat::Chicago,
            Self::IdeographDigital => ST_NumberFormat::IdeographDigital,
            Self::JapaneseCounting => ST_NumberFormat::JapaneseCounting,
            Self::Aiueo => ST_NumberFormat::Aiueo,
            Self::Iroha => ST_NumberFormat::Iroha,
            Self::DecimalFullWidth => ST_NumberFormat::DecimalFullWidth,
            Self::DecimalHalfWidth => ST_NumberFormat::DecimalHalfWidth,
            Self::JapaneseLegal => ST_NumberFormat::JapaneseLegal,
            Self::JapaneseDigitalTenThousand => ST_NumberFormat::JapaneseDigitalTenThousand,
            Self::DecimalEnclosedCircle => ST_NumberFormat::DecimalEnclosedCircle,
            Self::DecimalFullWidth2 => ST_NumberFormat::DecimalFullWidth2,
            Self::AiueoFullWidth => ST_NumberFormat::AiueoFullWidth,
            Self::IrohaFullWidth => ST_NumberFormat::IrohaFullWidth,
            Self::DecimalZero => ST_NumberFormat::DecimalZero,
            Self::Bullet => ST_NumberFormat::Bullet,
            Self::Ganada => ST_NumberFormat::Ganada,
            Self::Chosung => ST_NumberFormat::Chosung,
            Self::DecimalEnclosedFullstop => ST_NumberFormat::DecimalEnclosedFullstop,
            Self::DecimalEnclosedParen => ST_NumberFormat::DecimalEnclosedParen,
            Self::DecimalEnclosedCircleChinese => ST_NumberFormat::DecimalEnclosedCircleChinese,
            Self::IdeographEnclosedCircle => ST_NumberFormat::IdeographEnclosedCircle,
            Self::IdeographTraditional => ST_NumberFormat::IdeographTraditional,
            Self::IdeographZodiac => ST_NumberFormat::IdeographZodiac,
            Self::IdeographZodiacTraditional => ST_NumberFormat::IdeographZodiacTraditional,
            Self::TaiwaneseCounting => ST_NumberFormat::TaiwaneseCounting,
            Self::IdeographLegalTraditional => ST_NumberFormat::IdeographLegalTraditional,
            Self::TaiwaneseCountingThousand => ST_NumberFormat::TaiwaneseCountingThousand,
            Self::TaiwaneseDigital => ST_NumberFormat::TaiwaneseDigital,
            Self::ChineseCounting => ST_NumberFormat::ChineseCounting,
            Self::ChineseLegalSimplified => ST_NumberFormat::ChineseLegalSimplified,
            Self::ChineseCountingThousand => ST_NumberFormat::ChineseCountingThousand,
            Self::KoreanDigital => ST_NumberFormat::KoreanDigital,
            Self::KoreanCounting => ST_NumberFormat::KoreanCounting,
            Self::KoreanLegal => ST_NumberFormat::KoreanLegal,
            Self::KoreanDigital2 => ST_NumberFormat::KoreanDigital2,
            Self::Hebrew1 => ST_NumberFormat::Hebrew1,
            Self::ArabicAlpha => ST_NumberFormat::ArabicAlpha,
            Self::Hebrew2 => ST_NumberFormat::Hebrew2,
            Self::ArabicAbjad => ST_NumberFormat::ArabicAbjad,
            Self::HindiVowels => ST_NumberFormat::HindiVowels,
            Self::HindiConsonants => ST_NumberFormat::HindiConsonants,
            Self::HindiNumbers => ST_NumberFormat::HindiNumbers,
            Self::HindiCounting => ST_NumberFormat::HindiCounting,
            Self::ThaiLetters => ST_NumberFormat::ThaiLetters,
            Self::ThaiNumbers => ST_NumberFormat::ThaiNumbers,
            Self::ThaiCounting => ST_NumberFormat::ThaiCounting,
            Self::VietnameseCounting => ST_NumberFormat::VietnameseCounting,
            Self::NumberInDash => ST_NumberFormat::NumberInDash,
            Self::RussianLower => ST_NumberFormat::RussianLower,
            Self::RussianUpper => ST_NumberFormat::RussianUpper,
            Self::None => ST_NumberFormat::None,
            Self::Other(value) => ST_NumberFormat::Other(value.clone()),
        }
    }
}

/// The numbering format reported by the read-only numbering projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberingFormat<'a> {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
    Ordinal,
    CardinalText,
    OrdinalText,
    Hex,
    Chicago,
    IdeographDigital,
    JapaneseCounting,
    Aiueo,
    Iroha,
    DecimalFullWidth,
    DecimalHalfWidth,
    JapaneseLegal,
    JapaneseDigitalTenThousand,
    DecimalEnclosedCircle,
    DecimalFullWidth2,
    AiueoFullWidth,
    IrohaFullWidth,
    DecimalZero,
    Bullet,
    Ganada,
    Chosung,
    DecimalEnclosedFullstop,
    DecimalEnclosedParen,
    DecimalEnclosedCircleChinese,
    IdeographEnclosedCircle,
    IdeographTraditional,
    IdeographZodiac,
    IdeographZodiacTraditional,
    TaiwaneseCounting,
    IdeographLegalTraditional,
    TaiwaneseCountingThousand,
    TaiwaneseDigital,
    ChineseCounting,
    ChineseLegalSimplified,
    ChineseCountingThousand,
    KoreanDigital,
    KoreanCounting,
    KoreanLegal,
    KoreanDigital2,
    Hebrew1,
    ArabicAlpha,
    Hebrew2,
    ArabicAbjad,
    HindiVowels,
    HindiConsonants,
    HindiNumbers,
    HindiCounting,
    ThaiLetters,
    ThaiNumbers,
    ThaiCounting,
    VietnameseCounting,
    NumberInDash,
    RussianLower,
    RussianUpper,
    None,
    /// A producer-defined format value retained by rdocx.
    Other(&'a str),
}

impl<'a> NumberingFormat<'a> {
    fn from_st(value: &'a ST_NumberFormat) -> Self {
        match value {
            ST_NumberFormat::Decimal => Self::Decimal,
            ST_NumberFormat::UpperRoman => Self::UpperRoman,
            ST_NumberFormat::LowerRoman => Self::LowerRoman,
            ST_NumberFormat::UpperLetter => Self::UpperLetter,
            ST_NumberFormat::LowerLetter => Self::LowerLetter,
            ST_NumberFormat::Ordinal => Self::Ordinal,
            ST_NumberFormat::CardinalText => Self::CardinalText,
            ST_NumberFormat::OrdinalText => Self::OrdinalText,
            ST_NumberFormat::Hex => Self::Hex,
            ST_NumberFormat::Chicago => Self::Chicago,
            ST_NumberFormat::IdeographDigital => Self::IdeographDigital,
            ST_NumberFormat::JapaneseCounting => Self::JapaneseCounting,
            ST_NumberFormat::Aiueo => Self::Aiueo,
            ST_NumberFormat::Iroha => Self::Iroha,
            ST_NumberFormat::DecimalFullWidth => Self::DecimalFullWidth,
            ST_NumberFormat::DecimalHalfWidth => Self::DecimalHalfWidth,
            ST_NumberFormat::JapaneseLegal => Self::JapaneseLegal,
            ST_NumberFormat::JapaneseDigitalTenThousand => Self::JapaneseDigitalTenThousand,
            ST_NumberFormat::DecimalEnclosedCircle => Self::DecimalEnclosedCircle,
            ST_NumberFormat::DecimalFullWidth2 => Self::DecimalFullWidth2,
            ST_NumberFormat::AiueoFullWidth => Self::AiueoFullWidth,
            ST_NumberFormat::IrohaFullWidth => Self::IrohaFullWidth,
            ST_NumberFormat::DecimalZero => Self::DecimalZero,
            ST_NumberFormat::Bullet => Self::Bullet,
            ST_NumberFormat::Ganada => Self::Ganada,
            ST_NumberFormat::Chosung => Self::Chosung,
            ST_NumberFormat::DecimalEnclosedFullstop => Self::DecimalEnclosedFullstop,
            ST_NumberFormat::DecimalEnclosedParen => Self::DecimalEnclosedParen,
            ST_NumberFormat::DecimalEnclosedCircleChinese => Self::DecimalEnclosedCircleChinese,
            ST_NumberFormat::IdeographEnclosedCircle => Self::IdeographEnclosedCircle,
            ST_NumberFormat::IdeographTraditional => Self::IdeographTraditional,
            ST_NumberFormat::IdeographZodiac => Self::IdeographZodiac,
            ST_NumberFormat::IdeographZodiacTraditional => Self::IdeographZodiacTraditional,
            ST_NumberFormat::TaiwaneseCounting => Self::TaiwaneseCounting,
            ST_NumberFormat::IdeographLegalTraditional => Self::IdeographLegalTraditional,
            ST_NumberFormat::TaiwaneseCountingThousand => Self::TaiwaneseCountingThousand,
            ST_NumberFormat::TaiwaneseDigital => Self::TaiwaneseDigital,
            ST_NumberFormat::ChineseCounting => Self::ChineseCounting,
            ST_NumberFormat::ChineseLegalSimplified => Self::ChineseLegalSimplified,
            ST_NumberFormat::ChineseCountingThousand => Self::ChineseCountingThousand,
            ST_NumberFormat::KoreanDigital => Self::KoreanDigital,
            ST_NumberFormat::KoreanCounting => Self::KoreanCounting,
            ST_NumberFormat::KoreanLegal => Self::KoreanLegal,
            ST_NumberFormat::KoreanDigital2 => Self::KoreanDigital2,
            ST_NumberFormat::Hebrew1 => Self::Hebrew1,
            ST_NumberFormat::ArabicAlpha => Self::ArabicAlpha,
            ST_NumberFormat::Hebrew2 => Self::Hebrew2,
            ST_NumberFormat::ArabicAbjad => Self::ArabicAbjad,
            ST_NumberFormat::HindiVowels => Self::HindiVowels,
            ST_NumberFormat::HindiConsonants => Self::HindiConsonants,
            ST_NumberFormat::HindiNumbers => Self::HindiNumbers,
            ST_NumberFormat::HindiCounting => Self::HindiCounting,
            ST_NumberFormat::ThaiLetters => Self::ThaiLetters,
            ST_NumberFormat::ThaiNumbers => Self::ThaiNumbers,
            ST_NumberFormat::ThaiCounting => Self::ThaiCounting,
            ST_NumberFormat::VietnameseCounting => Self::VietnameseCounting,
            ST_NumberFormat::NumberInDash => Self::NumberInDash,
            ST_NumberFormat::RussianLower => Self::RussianLower,
            ST_NumberFormat::RussianUpper => Self::RussianUpper,
            ST_NumberFormat::None => Self::None,
            ST_NumberFormat::Other(value) => Self::Other(value),
        }
    }
}

/// The layout item emitted between a list marker and paragraph content.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListLevelSuffix {
    Tab,
    Space,
    Nothing,
}

/// Explicit restart behavior for a numbering level.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListLevelRestart {
    /// Continue across every more-significant level.
    Never,
    /// Restart after the given zero-based, more-significant level.
    After(u32),
}

impl ListLevelRestart {
    fn from_st(value: u32) -> Self {
        if value == 0 {
            Self::Never
        } else {
            Self::After(value - 1)
        }
    }
}

impl ListLevelSuffix {
    fn from_st(value: ST_LvlSuffix) -> Self {
        match value {
            ST_LvlSuffix::Tab => Self::Tab,
            ST_LvlSuffix::Space => Self::Space,
            ST_LvlSuffix::Nothing => Self::Nothing,
        }
    }
}

/// Resolved reader metadata for one numbering definition level.
#[derive(Debug, Clone, PartialEq)]
pub struct NumberingLevel<'a> {
    /// The zero-based list level.
    pub level: u32,
    /// The marker format at this level.
    pub format: NumberingFormat<'a>,
    /// The exact OOXML format name, including producer-defined values.
    pub format_name: &'a str,
    /// The first marker value, defaulting to one when Word omits it.
    pub start: u32,
    /// The item emitted between the marker and paragraph content.
    pub suffix: ListLevelSuffix,
    /// The Word level-text template or bullet glyph.
    pub level_text: Option<&'a str>,
    /// The marker alignment, when explicitly configured.
    pub alignment: Option<crate::paragraph::Alignment>,
    /// The paragraph style associated with this numbering level.
    pub paragraph_style: Option<&'a str>,
    /// Explicit restart behavior. Absence means the OOXML default.
    pub restart: Option<ListLevelRestart>,
    /// Whether inherited placeholders use decimal legal numbering.
    pub legal_numbering: Option<bool>,
    /// Left indentation in twips.
    pub indent_left: Option<oxml_core::Twips>,
    /// Hanging indentation in twips.
    pub indent_hanging: Option<oxml_core::Twips>,
    /// First-line indentation in twips.
    pub indent_first_line: Option<oxml_core::Twips>,
    /// Typed properties applied only to the numbering marker.
    pub marker_properties: Option<&'a CT_RPr>,
    /// Producer template code from `w:tplc`.
    pub template_code: Option<&'a str>,
    /// Whether the producer marked this level tentative.
    pub tentative: Option<bool>,
    /// Whether this level retains semantic facts this projection does not model.
    pub has_unmodeled_properties: bool,
    /// Whether the list item has nonstandard paragraph-level presentation.
    pub has_paragraph_presentation: bool,
    /// Whether the marker has run-level presentation properties.
    pub has_marker_presentation: bool,
}

fn list_level_alignment(value: ST_Jc) -> crate::paragraph::Alignment {
    match value {
        ST_Jc::Center => crate::paragraph::Alignment::Center,
        ST_Jc::Right | ST_Jc::End => crate::paragraph::Alignment::Right,
        ST_Jc::Both | ST_Jc::Distribute => crate::paragraph::Alignment::Justify,
        _ => crate::paragraph::Alignment::Left,
    }
}

/// One level of a custom list definition for [`Document::add_list_definition`].
#[derive(Debug, Clone)]
pub struct ListLevel {
    /// Numbering format for this level.
    pub format: ListNumberFormat,
    /// Starting number (defaults to 1; ignored for bullet levels).
    pub start: Option<u32>,
    level_text: Option<String>,
    suffix: Option<ListLevelSuffix>,
    alignment: Option<crate::paragraph::Alignment>,
    indent_left: Option<oxml_core::Twips>,
    indent_hanging: Option<oxml_core::Twips>,
    indent_first_line: Option<oxml_core::Twips>,
    marker_properties: Option<CT_RPr>,
    legal_numbering: Option<bool>,
    restart: Option<ListLevelRestart>,
    template_code: Option<String>,
    tentative: Option<bool>,
    source_level: Option<Box<CT_Lvl>>,
}

impl PartialEq for ListLevel {
    fn eq(&self, other: &Self) -> bool {
        self.format == other.format
            && self.start == other.start
            && self.level_text == other.level_text
            && self.suffix == other.suffix
            && self.alignment == other.alignment
            && self.indent_left == other.indent_left
            && self.indent_hanging == other.indent_hanging
            && self.indent_first_line == other.indent_first_line
            && self.marker_properties == other.marker_properties
            && self.legal_numbering == other.legal_numbering
            && self.restart == other.restart
            && self.template_code == other.template_code
            && self.tentative == other.tentative
    }
}

impl ListLevel {
    /// A level with the given format, starting at 1.
    pub fn new(format: ListNumberFormat) -> Self {
        ListLevel {
            format,
            start: None,
            level_text: None,
            suffix: None,
            alignment: None,
            indent_left: None,
            indent_hanging: None,
            indent_first_line: None,
            marker_properties: None,
            legal_numbering: None,
            restart: None,
            template_code: None,
            tentative: None,
            source_level: None,
        }
    }

    /// A bullet level.
    pub fn bullet() -> Self {
        Self::new(ListNumberFormat::Bullet)
    }

    /// A decimal-numbered level.
    pub fn decimal() -> Self {
        Self::new(ListNumberFormat::Decimal)
    }

    /// Override the starting number for this level.
    pub fn start(mut self, start: u32) -> Self {
        self.start = Some(start);
        self
    }

    /// Set the level-text template or bullet glyph.
    pub fn level_text(mut self, value: impl Into<String>) -> Self {
        self.level_text = Some(value.into());
        self
    }

    /// Select the content emitted after the marker.
    pub fn suffix(mut self, value: ListLevelSuffix) -> Self {
        self.suffix = Some(value);
        self
    }

    /// Set marker alignment.
    pub fn alignment(mut self, value: crate::paragraph::Alignment) -> Self {
        self.alignment = Some(value);
        self
    }

    /// Set left, hanging, and first-line indentation in twips.
    pub fn indentation(
        mut self,
        left: Option<oxml_core::Twips>,
        hanging: Option<oxml_core::Twips>,
        first_line: Option<oxml_core::Twips>,
    ) -> Self {
        self.indent_left = left;
        self.indent_hanging = hanging;
        self.indent_first_line = first_line;
        self
    }

    /// Set properties applied only to the numbering marker.
    pub fn marker_properties(mut self, value: CT_RPr) -> Self {
        self.marker_properties = Some(value);
        self
    }

    /// Select legal-numbering placeholder behavior.
    pub fn legal_numbering(mut self, value: bool) -> Self {
        self.legal_numbering = Some(value);
        self
    }

    /// Set explicit restart behavior.
    pub fn restart(mut self, value: ListLevelRestart) -> Self {
        self.restart = Some(value);
        self
    }

    /// Set the eight-hex-digit producer template code.
    pub fn template_code(mut self, value: impl Into<String>) -> Self {
        self.template_code = Some(value.into());
        self
    }

    /// Set the producer tentative flag.
    pub fn tentative(mut self, value: bool) -> Self {
        self.tentative = Some(value);
        self
    }

    /// Return the configured level-text template, when one was supplied.
    pub fn level_text_value(&self) -> Option<&str> {
        self.level_text.as_deref()
    }

    /// Return the configured marker suffix, when one was supplied.
    pub fn suffix_value(&self) -> Option<ListLevelSuffix> {
        self.suffix
    }

    /// Return the configured marker alignment, when one was supplied.
    pub fn alignment_value(&self) -> Option<crate::paragraph::Alignment> {
        self.alignment
    }

    /// Return the configured left, hanging, and first-line indentation.
    pub fn indentation_value(
        &self,
    ) -> (
        Option<oxml_core::Twips>,
        Option<oxml_core::Twips>,
        Option<oxml_core::Twips>,
    ) {
        (
            self.indent_left,
            self.indent_hanging,
            self.indent_first_line,
        )
    }

    /// Return the configured marker run properties.
    pub fn marker_properties_value(&self) -> Option<&CT_RPr> {
        self.marker_properties.as_ref()
    }

    /// Return the configured legal-numbering flag.
    pub fn legal_numbering_value(&self) -> Option<bool> {
        self.legal_numbering
    }

    /// Return the configured restart behavior.
    pub fn restart_value(&self) -> Option<ListLevelRestart> {
        self.restart
    }

    /// Return the configured producer template code.
    pub fn template_code_value(&self) -> Option<&str> {
        self.template_code.as_deref()
    }

    /// Return the configured tentative flag.
    pub fn tentative_value(&self) -> Option<bool> {
        self.tentative
    }
}

/// Owned public projection of an abstract numbering definition.
#[derive(Debug, Clone, PartialEq)]
pub struct NumberingDefinition {
    /// The abstract numbering definition identifier.
    pub id: u32,
    /// The definition's ordered levels.
    pub levels: Vec<NumberingDefinitionLevel>,
    /// Imported style links, indexed in parallel with `levels`.
    pub paragraph_style_links: Vec<Option<String>>,
    /// Whether the definition contains imported properties outside this projection.
    pub has_unmodeled_properties: bool,
}

/// One explicitly identified level in an inspected numbering definition.
#[derive(Debug, Clone, PartialEq)]
pub struct NumberingDefinitionLevel {
    /// The zero-based `w:ilvl` identifier.
    pub level: u32,
    /// The modeled properties owned by this level.
    pub properties: ListLevel,
}

/// One typed override owned by a numbering instance.
#[derive(Debug, Clone, PartialEq)]
pub struct NumberingLevelOverride {
    /// The zero-based level receiving the override.
    pub level: u32,
    /// An optional starting-value override.
    pub start: Option<u32>,
    /// An optional replacement level definition.
    pub replacement: Option<ListLevel>,
    /// Imported style link for the replacement level.
    pub paragraph_style_link: Option<String>,
    /// Whether the override contains imported properties outside this projection.
    pub has_unmodeled_properties: bool,
}

impl NumberingLevelOverride {
    /// Create an empty override for a zero-based level.
    pub fn new(level: u32) -> Self {
        Self {
            level,
            start: None,
            replacement: None,
            paragraph_style_link: None,
            has_unmodeled_properties: false,
        }
    }

    /// Set the override's starting value.
    pub fn start(mut self, value: u32) -> Self {
        self.start = Some(value);
        self
    }

    /// Replace the level definition for this instance.
    pub fn replacement(mut self, value: ListLevel) -> Self {
        self.replacement = Some(value);
        self
    }
}

/// Owned public projection of one numbering instance.
#[derive(Debug, Clone, PartialEq)]
pub struct NumberingInstance {
    /// The numbering instance identifier used by document paragraphs.
    pub id: u32,
    /// The referenced abstract numbering definition identifier.
    pub definition_id: u32,
    /// Typed overrides applied by this instance.
    pub level_overrides: Vec<NumberingLevelOverride>,
    /// Whether the instance contains imported properties outside this projection.
    pub has_unmodeled_properties: bool,
}

fn list_alignment_to_st(value: crate::paragraph::Alignment) -> ST_Jc {
    match value {
        crate::paragraph::Alignment::Left => ST_Jc::Left,
        crate::paragraph::Alignment::Center => ST_Jc::Center,
        crate::paragraph::Alignment::Right => ST_Jc::Right,
        crate::paragraph::Alignment::Justify => ST_Jc::Both,
    }
}

fn list_level_restart_to_st(value: ListLevelRestart) -> Option<u32> {
    match value {
        ListLevelRestart::Never => Some(0),
        ListLevelRestart::After(level) => level.checked_add(1),
    }
}

fn list_level_from_ct(level: &CT_Lvl) -> ListLevel {
    let mut value = list_level_values_from_ct(level);
    value.source_level = Some(Box::new(level.clone()));
    value
}

fn list_level_values_from_ct(level: &CT_Lvl) -> ListLevel {
    ListLevel {
        format: level
            .num_fmt
            .clone()
            .map(list_number_format_from_st)
            .unwrap_or(ListNumberFormat::Decimal),
        start: level.start,
        level_text: level.lvl_text.clone(),
        suffix: level.suffix.map(ListLevelSuffix::from_st),
        alignment: level.lvl_jc.map(list_level_alignment),
        indent_left: level.ppr.as_ref().and_then(|value| value.ind_left),
        indent_hanging: level.ppr.as_ref().and_then(|value| value.ind_hanging),
        indent_first_line: level.ppr.as_ref().and_then(|value| value.ind_first_line),
        marker_properties: level.rpr.clone(),
        legal_numbering: level.legal,
        restart: level.restart.map(ListLevelRestart::from_st),
        template_code: level.template_code.clone(),
        tentative: level.tentative,
        source_level: None,
    }
}

pub(crate) fn list_number_format_from_st(value: ST_NumberFormat) -> ListNumberFormat {
    // Both enums intentionally have the same standard token set.
    match value.to_str() {
        "decimal" => ListNumberFormat::Decimal,
        "upperRoman" => ListNumberFormat::UpperRoman,
        "lowerRoman" => ListNumberFormat::LowerRoman,
        "upperLetter" => ListNumberFormat::UpperLetter,
        "lowerLetter" => ListNumberFormat::LowerLetter,
        "ordinal" => ListNumberFormat::Ordinal,
        "cardinalText" => ListNumberFormat::CardinalText,
        "ordinalText" => ListNumberFormat::OrdinalText,
        "hex" => ListNumberFormat::Hex,
        "chicago" => ListNumberFormat::Chicago,
        "ideographDigital" => ListNumberFormat::IdeographDigital,
        "japaneseCounting" => ListNumberFormat::JapaneseCounting,
        "Aiueo" => ListNumberFormat::Aiueo,
        "Iroha" => ListNumberFormat::Iroha,
        "decimalFullWidth" => ListNumberFormat::DecimalFullWidth,
        "decimalHalfWidth" => ListNumberFormat::DecimalHalfWidth,
        "japaneseLegal" => ListNumberFormat::JapaneseLegal,
        "japaneseDigitalTenThousand" => ListNumberFormat::JapaneseDigitalTenThousand,
        "decimalEnclosedCircle" => ListNumberFormat::DecimalEnclosedCircle,
        "decimalFullWidth2" => ListNumberFormat::DecimalFullWidth2,
        "aiueoFullWidth" => ListNumberFormat::AiueoFullWidth,
        "irohaFullWidth" => ListNumberFormat::IrohaFullWidth,
        "decimalZero" => ListNumberFormat::DecimalZero,
        "bullet" => ListNumberFormat::Bullet,
        "ganada" => ListNumberFormat::Ganada,
        "chosung" => ListNumberFormat::Chosung,
        "decimalEnclosedFullstop" => ListNumberFormat::DecimalEnclosedFullstop,
        "decimalEnclosedParen" => ListNumberFormat::DecimalEnclosedParen,
        "decimalEnclosedCircleChinese" => ListNumberFormat::DecimalEnclosedCircleChinese,
        "ideographEnclosedCircle" => ListNumberFormat::IdeographEnclosedCircle,
        "ideographTraditional" => ListNumberFormat::IdeographTraditional,
        "ideographZodiac" => ListNumberFormat::IdeographZodiac,
        "ideographZodiacTraditional" => ListNumberFormat::IdeographZodiacTraditional,
        "taiwaneseCounting" => ListNumberFormat::TaiwaneseCounting,
        "ideographLegalTraditional" => ListNumberFormat::IdeographLegalTraditional,
        "taiwaneseCountingThousand" => ListNumberFormat::TaiwaneseCountingThousand,
        "taiwaneseDigital" => ListNumberFormat::TaiwaneseDigital,
        "chineseCounting" => ListNumberFormat::ChineseCounting,
        "chineseLegalSimplified" => ListNumberFormat::ChineseLegalSimplified,
        "chineseCountingThousand" => ListNumberFormat::ChineseCountingThousand,
        "koreanDigital" => ListNumberFormat::KoreanDigital,
        "koreanCounting" => ListNumberFormat::KoreanCounting,
        "koreanLegal" => ListNumberFormat::KoreanLegal,
        "koreanDigital2" => ListNumberFormat::KoreanDigital2,
        "hebrew1" => ListNumberFormat::Hebrew1,
        "arabicAlpha" => ListNumberFormat::ArabicAlpha,
        "hebrew2" => ListNumberFormat::Hebrew2,
        "arabicAbjad" => ListNumberFormat::ArabicAbjad,
        "hindiVowels" => ListNumberFormat::HindiVowels,
        "hindiConsonants" => ListNumberFormat::HindiConsonants,
        "hindiNumbers" => ListNumberFormat::HindiNumbers,
        "hindiCounting" => ListNumberFormat::HindiCounting,
        "thaiLetters" => ListNumberFormat::ThaiLetters,
        "thaiNumbers" => ListNumberFormat::ThaiNumbers,
        "thaiCounting" => ListNumberFormat::ThaiCounting,
        "vietnameseCounting" => ListNumberFormat::VietnameseCounting,
        "numberInDash" => ListNumberFormat::NumberInDash,
        "russianLower" => ListNumberFormat::RussianLower,
        "russianUpper" => ListNumberFormat::RussianUpper,
        "none" => ListNumberFormat::None,
        other => ListNumberFormat::Other(other.to_owned()),
    }
}

fn numbering_level_has_unmodeled(level: &CT_Lvl) -> bool {
    let unprojected_paragraph_properties = level.ppr.as_ref().is_some_and(|properties| {
        let projected = CT_PPr {
            ind_left: properties.ind_left,
            ind_hanging: properties.ind_hanging,
            ind_first_line: properties.ind_first_line,
            ..CT_PPr::default()
        };
        properties != &projected
    });
    !level.extra_xml.is_empty()
        || !level.extra_attributes.is_empty()
        || level
            .start_raw
            .as_ref()
            .is_some_and(|(_, raw, prefixes)| typed_numbering_leaf_has_unmodeled(raw, prefixes))
        || level
            .num_fmt_raw
            .as_ref()
            .is_some_and(|(_, raw, prefixes)| typed_numbering_leaf_has_unmodeled(raw, prefixes))
        || level
            .p_style_raw
            .as_ref()
            .is_some_and(|(_, raw, prefixes)| typed_numbering_leaf_has_unmodeled(raw, prefixes))
        || level
            .restart_raw
            .as_ref()
            .is_some_and(|(_, raw, prefixes)| typed_numbering_leaf_has_unmodeled(raw, prefixes))
        || level
            .legal_raw
            .as_ref()
            .is_some_and(|(_, raw, prefixes)| typed_numbering_leaf_has_unmodeled(raw, prefixes))
        || level
            .suffix_raw
            .as_ref()
            .is_some_and(|(_, raw, prefixes)| typed_numbering_leaf_has_unmodeled(raw, prefixes))
        || level
            .lvl_text_raw
            .as_ref()
            .is_some_and(|(_, raw, prefixes)| typed_numbering_leaf_has_unmodeled(raw, prefixes))
        || level
            .lvl_jc_raw
            .as_ref()
            .is_some_and(|(_, raw, prefixes)| typed_numbering_leaf_has_unmodeled(raw, prefixes))
        || level.ppr_raw.is_some()
        || level.rpr_raw.is_some()
        || unprojected_paragraph_properties
        || matches!(level.num_fmt.as_ref(), Some(ST_NumberFormat::Other(_)))
}

fn numbering_override_has_unmodeled(value: &CT_NumLvl) -> bool {
    !value.extra_xml.is_empty()
        || !value.extra_attributes.is_empty()
        || value
            .start_override_raw
            .as_ref()
            .is_some_and(|(_, raw, prefixes)| typed_numbering_leaf_has_unmodeled(raw, prefixes))
        || value
            .level
            .as_ref()
            .is_some_and(numbering_level_has_unmodeled)
}

fn validate_list_level(level: u32, value: &ListLevel) -> Result<()> {
    if level > 8 {
        return Err(Error::Other(format!(
            "numbering level {level} is outside the supported range 0 through 8"
        )));
    }
    if value.indent_hanging.is_some() && value.indent_first_line.is_some() {
        return Err(Error::Other(format!(
            "numbering level {level} cannot have both hanging and first-line indentation"
        )));
    }
    if let Some(ListLevelRestart::After(owner)) = value.restart
        && owner >= level
    {
        return Err(Error::Other(format!(
            "numbering level {level} cannot restart after level {owner}"
        )));
    }
    if let Some(template_code) = &value.template_code
        && (template_code.len() != 8 || !template_code.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        return Err(Error::Other(format!(
            "numbering level {level} template code must contain eight hexadecimal digits"
        )));
    }
    if let Some(text) = &value.level_text {
        let bytes = text.as_bytes();
        let mut position = 0usize;
        while position < bytes.len() {
            if bytes[position] != b'%' {
                position += 1;
                continue;
            }
            let Some(next) = bytes.get(position + 1).copied() else {
                return Err(Error::Other(format!(
                    "numbering level {level} has an incomplete placeholder"
                )));
            };
            if next == b'%' {
                position += 2;
                continue;
            }
            if !(b'1'..=b'9').contains(&next) || u32::from(next - b'1') > level {
                return Err(Error::Other(format!(
                    "numbering level {level} has an invalid placeholder %{}",
                    char::from(next)
                )));
            }
            position += 2;
        }
    }
    Ok(())
}

fn default_list_level_text(level: u32, format: &ListNumberFormat) -> String {
    const BULLETS: [&str; 3] = ["\u{2022}", "\u{25E6}", "\u{25AA}"];
    match format {
        ListNumberFormat::Bullet => BULLETS[level as usize % BULLETS.len()].to_owned(),
        ListNumberFormat::None => String::new(),
        _ => format!("%{}.", level + 1),
    }
}

fn ct_level_from_list(level: u32, value: &ListLevel) -> Result<CT_Lvl> {
    validate_list_level(level, value)?;
    let mut result = CT_Lvl::new(level);
    result.template_code.clone_from(&value.template_code);
    result.tentative = value.tentative;
    result.start = Some(value.start.unwrap_or(1));
    result.num_fmt = Some(value.format.to_st());
    result.restart = value.restart.and_then(list_level_restart_to_st);
    result.legal = value.legal_numbering;
    result.suffix = value.suffix.map(|suffix| match suffix {
        ListLevelSuffix::Tab => ST_LvlSuffix::Tab,
        ListLevelSuffix::Space => ST_LvlSuffix::Space,
        ListLevelSuffix::Nothing => ST_LvlSuffix::Nothing,
    });
    result.lvl_text = Some(
        value
            .level_text
            .clone()
            .unwrap_or_else(|| default_list_level_text(level, &value.format)),
    );
    result.lvl_jc = Some(
        value
            .alignment
            .map(list_alignment_to_st)
            .unwrap_or(ST_Jc::Left),
    );
    let default_left = (level + 1)
        .checked_mul(720)
        .and_then(|value| i32::try_from(value).ok())
        .map(oxml_core::Twips)
        .ok_or_else(|| Error::Other(format!("numbering level {level} indentation overflow")))?;
    result.ppr = Some(CT_PPr {
        ind_left: Some(value.indent_left.unwrap_or(default_left)),
        ind_hanging: Some(value.indent_hanging.unwrap_or(oxml_core::Twips(360))),
        ind_first_line: value.indent_first_line,
        ..CT_PPr::default()
    });
    if value.indent_first_line.is_some() && value.indent_hanging.is_none() {
        result
            .ppr
            .as_mut()
            .expect("created paragraph properties")
            .ind_hanging = None;
    }
    result.rpr.clone_from(&value.marker_properties);
    Ok(result)
}

fn merge_list_level(existing: &CT_Lvl, value: &ListLevel) -> Result<CT_Lvl> {
    if value.source_level.as_deref() == Some(existing) {
        let source = list_level_values_from_ct(existing);
        let mut merged = existing.clone();
        if value.format != source.format {
            merged.num_fmt = Some(value.format.to_st());
        }
        if value.start != source.start {
            merged.start = value.start;
        }
        if value.level_text != source.level_text {
            merged.lvl_text.clone_from(&value.level_text);
        }
        if value.suffix != source.suffix {
            merged.suffix = value.suffix.map(|suffix| match suffix {
                ListLevelSuffix::Tab => ST_LvlSuffix::Tab,
                ListLevelSuffix::Space => ST_LvlSuffix::Space,
                ListLevelSuffix::Nothing => ST_LvlSuffix::Nothing,
            });
        }
        if value.alignment != source.alignment {
            merged.lvl_jc = value.alignment.map(list_alignment_to_st);
        }
        if value.restart != source.restart {
            merged.restart = value.restart.and_then(list_level_restart_to_st);
        }
        if value.legal_numbering != source.legal_numbering {
            merged.legal = value.legal_numbering;
        }
        if value.template_code != source.template_code {
            merged.template_code.clone_from(&value.template_code);
        }
        if value.tentative != source.tentative {
            merged.tentative = value.tentative;
        }
        if value.indentation_value() != source.indentation_value() {
            let properties = merged.ppr.get_or_insert_with(CT_PPr::default);
            properties.ind_left = value.indent_left;
            properties.ind_hanging = value.indent_hanging;
            properties.ind_first_line = value.indent_first_line;
        }
        if value.marker_properties != source.marker_properties {
            merged.rpr.clone_from(&value.marker_properties);
        }
        return Ok(merged);
    }
    let mut authored = ct_level_from_list(existing.ilvl, value)?;
    authored.start_raw.clone_from(&existing.start_raw);
    authored.num_fmt_raw.clone_from(&existing.num_fmt_raw);
    authored.p_style.clone_from(&existing.p_style);
    authored.p_style_raw.clone_from(&existing.p_style_raw);
    authored.restart_raw.clone_from(&existing.restart_raw);
    authored.legal_raw.clone_from(&existing.legal_raw);
    authored.suffix_raw.clone_from(&existing.suffix_raw);
    authored.lvl_text_raw.clone_from(&existing.lvl_text_raw);
    authored.lvl_jc_raw.clone_from(&existing.lvl_jc_raw);
    authored.extra_xml.clone_from(&existing.extra_xml);
    authored
        .extra_attributes
        .clone_from(&existing.extra_attributes);
    authored.ppr_raw.clone_from(&existing.ppr_raw);
    authored.rpr_raw.clone_from(&existing.rpr_raw);
    Ok(authored)
}

fn reject_conflicting_raw_level_attribute(level: &CT_Lvl, local_name: &[u8]) -> Result<()> {
    if level.extra_attributes.iter().any(|(name, _)| {
        name.as_bytes()
            .rsplit(|byte| *byte == b':')
            .next()
            .is_some_and(|local| local == local_name)
    }) {
        return Err(Error::Other(format!(
            "cannot replace unmodeled numbering level attribute '{}'",
            String::from_utf8_lossy(local_name)
        )));
    }
    Ok(())
}

fn validate_numbering_level_slice(levels: &[ListLevel]) -> Result<()> {
    if levels.is_empty() || levels.len() > 9 {
        return Err(Error::Other(format!(
            "numbering definitions require 1 through 9 levels, received {}",
            levels.len()
        )));
    }
    for (level, value) in levels.iter().enumerate() {
        validate_list_level(level as u32, value)?;
    }
    Ok(())
}

fn validate_numbering_definition_levels(levels: &[NumberingDefinitionLevel]) -> Result<()> {
    if levels.is_empty() || levels.len() > 9 {
        return Err(Error::Other(format!(
            "numbering definitions require 1 through 9 levels, received {}",
            levels.len()
        )));
    }
    let mut seen = HashSet::new();
    for value in levels {
        if !seen.insert(value.level) {
            return Err(Error::Other(format!(
                "numbering definition contains duplicate level {}",
                value.level
            )));
        }
        validate_list_level(value.level, &value.properties)?;
    }
    Ok(())
}

fn validate_ct_numbering_level(level: &CT_Lvl) -> Result<()> {
    if level.template_code.is_some() {
        reject_conflicting_raw_level_attribute(level, b"tplc")?;
    }
    if level.tentative.is_some() {
        reject_conflicting_raw_level_attribute(level, b"tentative")?;
    }
    validate_list_level(level.ilvl, &list_level_from_ct(level))
}

fn numbering_override_from_public(value: &NumberingLevelOverride) -> Result<CT_NumLvl> {
    if value.level > 8 {
        return Err(Error::Other(format!(
            "numbering override level {} is outside the supported range 0 through 8",
            value.level
        )));
    }
    if value.paragraph_style_link.is_some() {
        return Err(Error::Other(
            "standalone numbering style-link mutation is deferred to F-248".to_owned(),
        ));
    }
    if value.has_unmodeled_properties {
        return Err(Error::Other(
            "a new numbering override cannot claim imported unmodeled properties".to_owned(),
        ));
    }
    let mut result = CT_NumLvl::new(value.level);
    result.start_override = value.start;
    result.level = value
        .replacement
        .as_ref()
        .map(|level| ct_level_from_list(value.level, level))
        .transpose()?;
    Ok(result)
}

fn numbering_override_for_update(
    existing: Option<&CT_NumLvl>,
    value: &NumberingLevelOverride,
) -> Result<CT_NumLvl> {
    let Some(existing) = existing else {
        return numbering_override_from_public(value);
    };
    let existing_style = existing
        .level
        .as_ref()
        .and_then(|level| level.p_style.as_deref());
    if value.paragraph_style_link.as_deref() != existing_style {
        return Err(Error::Other(
            "standalone numbering style-link mutation is deferred to F-248".to_owned(),
        ));
    }
    let mut result = existing.clone();
    result.start_override = value.start;
    result.level = match (&existing.level, &value.replacement) {
        (Some(existing), Some(replacement)) => Some(merge_list_level(existing, replacement)?),
        (None, Some(replacement)) => Some(ct_level_from_list(value.level, replacement)?),
        (_, None) => None,
    };
    Ok(result)
}

fn xml_paragraph_numbering_properties(xml: &[u8]) -> Result<Vec<CT_PPr>> {
    struct ParagraphState {
        depth: usize,
        ppr_depth: Option<usize>,
        numpr_depth: Option<usize>,
        properties: CT_PPr,
    }

    fn word_value(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Option<String>> {
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            if matches!(
                namespace,
                ResolveResult::Bound(value) if value.as_ref() == WORD_NAMESPACE.as_bytes()
            ) && local.as_ref() == b"val"
            {
                return attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map(|value| Some(value.into_owned()))
                    .map_err(|error| Error::Other(error.to_string()));
            }
        }
        Ok(None)
    }

    fn apply_leaf(
        reader: &NsReader<&[u8]>,
        state: &mut ParagraphState,
        depth: usize,
        local_name: &[u8],
        element: &BytesStart<'_>,
    ) -> Result<()> {
        if local_name == b"pStyle" && state.ppr_depth == depth.checked_sub(1) {
            state.properties.style_id = word_value(reader, element)?;
        } else if state.numpr_depth == depth.checked_sub(1) {
            if local_name == b"numId" {
                state.properties.num_id =
                    word_value(reader, element)?.and_then(|value| value.parse::<u32>().ok());
            } else if local_name == b"ilvl" {
                state.properties.num_ilvl =
                    word_value(reader, element)?.and_then(|value| value.parse::<u32>().ok());
            }
        }
        Ok(())
    }

    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0_usize;
    let mut paragraphs = Vec::new();
    let mut result = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(error.to_string()))?;
        match event {
            Event::Start(ref element) => {
                let is_word = matches!(
                    namespace,
                    ResolveResult::Bound(value)
                        if value.as_ref() == WORD_NAMESPACE.as_bytes()
                );
                let local_name = element.local_name();
                if is_word && local_name.as_ref() == b"p" {
                    paragraphs.push(ParagraphState {
                        depth,
                        ppr_depth: None,
                        numpr_depth: None,
                        properties: CT_PPr::default(),
                    });
                } else if is_word && let Some(state) = paragraphs.last_mut() {
                    if local_name.as_ref() == b"pPr" && depth == state.depth + 1 {
                        state.ppr_depth = Some(depth);
                    } else if local_name.as_ref() == b"numPr"
                        && state.ppr_depth == depth.checked_sub(1)
                    {
                        state.numpr_depth = Some(depth);
                    } else {
                        apply_leaf(&reader, state, depth, local_name.as_ref(), element)?;
                    }
                }
                depth += 1;
            }
            Event::Empty(ref element) => {
                let is_word = matches!(
                    namespace,
                    ResolveResult::Bound(value)
                        if value.as_ref() == WORD_NAMESPACE.as_bytes()
                );
                let local_name = element.local_name();
                if is_word && local_name.as_ref() == b"p" {
                    result.push(CT_PPr::default());
                } else if is_word && let Some(state) = paragraphs.last_mut() {
                    apply_leaf(&reader, state, depth, local_name.as_ref(), element)?;
                }
            }
            Event::End(ref element) => {
                depth = depth.saturating_sub(1);
                let is_word = matches!(
                    namespace,
                    ResolveResult::Bound(value)
                        if value.as_ref() == WORD_NAMESPACE.as_bytes()
                );
                if is_word && element.local_name().as_ref() == b"p" {
                    if let Some(paragraph) = paragraphs.pop_if(|state| state.depth == depth) {
                        result.push(paragraph.properties);
                    }
                } else if let Some(state) = paragraphs.last_mut() {
                    if state.numpr_depth == Some(depth) {
                        state.numpr_depth = None;
                    }
                    if state.ppr_depth == Some(depth) {
                        state.ppr_depth = None;
                        state.numpr_depth = None;
                    }
                }
            }
            Event::Eof => return Ok(result),
            _ => {}
        }
        buffer.clear();
    }
}

fn xml_references_numbering_instance(xml: &[u8], id: u32) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(error.to_string()))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if matches!(
                    namespace,
                    ResolveResult::Bound(value)
                        if value.as_ref() == WORD_NAMESPACE.as_bytes()
                ) && element.local_name().as_ref() == b"numId" =>
            {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if !matches!(
                        namespace,
                        ResolveResult::Bound(value)
                            if value.as_ref() == WORD_NAMESPACE.as_bytes()
                    ) || local.as_ref() != b"val"
                    {
                        continue;
                    }
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Other(error.to_string()))?;
                    if value.parse::<u32>().ok() == Some(id) {
                        return Ok(true);
                    }
                }
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

/// A node in the document outline tree.
#[derive(Debug, Clone, PartialEq)]
pub struct OutlineNode {
    /// The heading level (1-9).
    pub level: u32,
    /// The heading text.
    pub text: String,
    /// Child headings (sub-headings).
    pub children: Vec<OutlineNode>,
}

/// Information about an image in the document.
#[derive(Debug, Clone, PartialEq)]
pub struct ImageInfo {
    /// The relationship ID for the embedded image.
    pub embed_id: String,
    /// Optional name attribute.
    pub name: Option<String>,
    /// Optional description (alt text).
    pub description: Option<String>,
    /// Width in EMUs (English Metric Units, 914400 EMU = 1 inch).
    pub width_emu: i64,
    /// Height in EMUs.
    pub height_emu: i64,
    /// Whether this is an anchored (floating) image vs inline.
    pub is_anchor: bool,
}

/// Information about a hyperlink in the document.
#[derive(Debug, Clone, PartialEq)]
pub struct LinkInfo {
    /// The display text of the hyperlink.
    pub text: String,
    /// The resolved target URL (if external).
    pub url: Option<String>,
    /// Internal document anchor (if any).
    pub anchor: Option<String>,
    /// The relationship ID.
    pub rel_id: Option<String>,
}

/// Severity level for accessibility issues.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IssueSeverity {
    /// Informational suggestion.
    Info,
    /// Potential problem.
    Warning,
    /// Definite accessibility barrier.
    Error,
}

/// An accessibility issue found during audit.
#[derive(Debug, Clone, PartialEq)]
pub struct AccessibilityIssue {
    /// How severe the issue is.
    pub severity: IssueSeverity,
    /// Human-readable description of the issue.
    pub message: String,
}

/// Build a hierarchical outline tree from a flat list of (level, text) headings.
fn build_outline_tree(headings: &[(u32, String)]) -> Vec<OutlineNode> {
    let mut root: Vec<OutlineNode> = Vec::new();
    let mut stack: Vec<(u32, usize)> = Vec::new(); // (level, index in parent's children)

    for (level, text) in headings {
        let node = OutlineNode {
            level: *level,
            text: text.clone(),
            children: Vec::new(),
        };

        // Pop stack until we find a parent with a lower level
        while let Some(&(stack_level, _)) = stack.last() {
            if stack_level >= *level {
                stack.pop();
            } else {
                break;
            }
        }

        if stack.is_empty() {
            root.push(node);
            let idx = root.len() - 1;
            stack.push((*level, idx));
        } else {
            // Navigate to the correct parent in the tree
            let target = get_outline_parent_mut(&mut root, &stack);
            target.children.push(node);
            let idx = target.children.len() - 1;
            stack.push((*level, idx));
        }
    }

    root
}

/// Navigate to the parent node indicated by the stack.
fn get_outline_parent_mut<'a>(
    root: &'a mut [OutlineNode],
    stack: &[(u32, usize)],
) -> &'a mut OutlineNode {
    let mut current = &mut root[stack[0].1];
    for &(_, idx) in &stack[1..] {
        current = &mut current.children[idx];
    }
    current
}

/// Truncate a string to at most `max_len` characters, appending "..." if it
/// was cut short.
///
/// Both the comparison and the cut are in characters; mixing byte length with
/// character counts would truncate non-ASCII text earlier than asked.
fn truncate_str(s: &str, max_len: usize) -> String {
    if s.chars().count() <= max_len {
        return s.to_string();
    }
    let truncated: String = s.chars().take(max_len.saturating_sub(3)).collect();
    format!("{truncated}...")
}

/// Deobfuscate an ODTTF (obfuscated TrueType) font file.
///
/// Word embeds fonts as `.odttf` files whose first 32 bytes are XOR'd with a
/// 16-byte key derived from the GUID in the part name (ECMA-376 Part 1,
/// "Embedded Font Obfuscation"). The GUID hex is read into the key *backwards*,
/// but implementations differ in whether they reverse the raw hex string or the
/// mixed-endian layout .NET's `Guid.ToByteArray` produces — the two agree on
/// the first eight key bytes and disagree on the rest.
///
/// Rather than pick one and hope, both orders are tried and the result is only
/// accepted if it starts with a recognised sfnt version. A wrong key yields
/// bytes that no font parser can use, so validating here means a bad guess
/// degrades to "font not embedded" instead of feeding garbage downstream.
fn deobfuscate_odttf(data: &[u8], file_name: &str) -> Option<Vec<u8>> {
    if data.len() < 32 {
        return None;
    }

    // Extract GUID from file name: "00112233-4455-6677-8899-AABBCCDDEEFF.odttf"
    // or "{00112233-4455-6677-8899-AABBCCDDEEFF}.odttf"
    let name = file_name
        .split('.')
        .next()
        .unwrap_or("")
        .trim_start_matches('{')
        .trim_end_matches('}');
    let hex = name
        .chars()
        .filter(|character| *character != '-')
        .collect::<String>()
        .to_ascii_uppercase();
    if hex.len() != 32 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    let key = format!(
        "{{{}-{}-{}-{}-{}}}",
        &hex[0..8],
        &hex[8..12],
        &hex[12..16],
        &hex[16..20],
        &hex[20..32]
    );
    deobfuscate_odttf_with_key(data, &key)
}

fn font_key_bytes(value: &str) -> Option<[u8; 16]> {
    let value_bytes = value.as_bytes();
    if value_bytes.len() != 38
        || value_bytes.first() != Some(&b'{')
        || value_bytes.last() != Some(&b'}')
        || ![9, 14, 19, 24]
            .into_iter()
            .all(|index| value_bytes.get(index) == Some(&b'-'))
        || !value_bytes.iter().enumerate().all(|(index, byte)| {
            matches!(index, 0 | 9 | 14 | 19 | 24 | 37) || matches!(byte, b'0'..=b'9' | b'A'..=b'F')
        })
    {
        return None;
    }
    let hex: String = value[1..37]
        .chars()
        .filter(|character| *character != '-')
        .collect();
    let mut guid = [0u8; 16];
    for (index, byte) in guid.iter_mut().enumerate() {
        *byte = u8::from_str_radix(&hex[index * 2..index * 2 + 2], 16).ok()?;
    }
    Some(guid)
}

fn obfuscate_odttf_with_key(data: &[u8], guid: &[u8; 16]) -> Vec<u8> {
    let key = odttf_key_candidates(guid)[0];
    let mut result = data.to_vec();
    for (index, byte) in result.iter_mut().take(32).enumerate() {
        *byte ^= key[index % 16];
    }
    result
}

fn deobfuscate_odttf_with_key(data: &[u8], font_key: &str) -> Option<Vec<u8>> {
    if data.len() < 32 {
        return None;
    }
    let guid = font_key_bytes(font_key)?;

    let candidates = odttf_key_candidates(&guid);
    let decoded: Vec<Vec<u8>> = candidates
        .iter()
        .map(|key| {
            let mut result = data.to_vec();
            // XOR the first 32 bytes with the 16-byte key, applied twice.
            for (i, byte) in result.iter_mut().take(32).enumerate() {
                *byte ^= key[i % 16];
            }
            result
        })
        .collect();

    // A well-formed table directory pins down which key was used. Fall back to
    // the weaker signature check for fonts whose header arithmetic is wrong —
    // subsetting tools do emit those — so they still load rather than being
    // dropped entirely.
    decoded
        .iter()
        .find(|d| has_consistent_sfnt_header(d))
        .or_else(|| decoded.iter().find(|d| looks_like_sfnt(d)))
        .cloned()
}

/// The two candidate XOR keys for ODTTF deobfuscation, most likely first.
fn odttf_key_candidates(guid: &[u8; 16]) -> [[u8; 16]; 2] {
    // Read the hex string end-first, as the spec prose describes.
    let mut plain_reversed = *guid;
    plain_reversed.reverse();

    // The .NET route: `Guid.ToByteArray` byte-swaps the first three groups,
    // and the whole array is then reversed.
    let dotnet = [
        guid[3], guid[2], guid[1], guid[0], guid[5], guid[4], guid[7], guid[6], guid[8], guid[9],
        guid[10], guid[11], guid[12], guid[13], guid[14], guid[15],
    ];
    let mut dotnet_reversed = dotnet;
    dotnet_reversed.reverse();

    [plain_reversed, dotnet_reversed]
}

/// Check that `data` opens with a plausible sfnt (TrueType/OpenType) header.
///
/// This is the weak test: signature plus printable table tags. It cannot always
/// tell the two ODTTF key conventions apart, since they produce identical
/// output for the first eight bytes.
fn looks_like_sfnt(data: &[u8]) -> bool {
    let Some(signature) = data.first_chunk::<4>() else {
        return false;
    };
    match signature {
        b"\x00\x01\x00\x00" | b"OTTO" | b"true" => {}
        // A collection header has a different layout; take it on signature.
        b"ttcf" => return true,
        _ => return false,
    }

    if data.len() < 32 {
        return false;
    }

    let num_tables = u16::from_be_bytes([data[4], data[5]]);
    if num_tables == 0 || num_tables > 512 {
        return false;
    }

    // Table records begin at offset 12 and are 16 bytes each, so the first
    // record's tag is at 12..16 and the second record's tag at 28..32 — both
    // inside the 32 bytes the obfuscation touches.
    let is_tag = |tag: &[u8]| tag.iter().all(|b| (0x20..=0x7E).contains(b));
    is_tag(&data[12..16]) && (num_tables < 2 || is_tag(&data[28..32]))
}

/// The strong test: the sfnt header's binary-search hints must agree with the
/// table count.
///
/// `searchRange`, `entrySelector` and `rangeShift` are all derived from
/// `numTables`, and `entrySelector`/`rangeShift` sit in the byte range where
/// the two ODTTF key conventions differ — so this identifies the right key
/// outright whenever the font's header is spec-conformant.
fn has_consistent_sfnt_header(data: &[u8]) -> bool {
    if !looks_like_sfnt(data) || data.len() < 12 {
        return false;
    }
    if data.first_chunk::<4>() == Some(b"ttcf") {
        return false; // no table directory at this offset
    }

    let num_tables = u16::from_be_bytes([data[4], data[5]]);
    let search_range = u16::from_be_bytes([data[6], data[7]]);
    let entry_selector = u16::from_be_bytes([data[8], data[9]]);
    let range_shift = u16::from_be_bytes([data[10], data[11]]);

    let expected_selector = num_tables.ilog2() as u16;
    let expected_search_range = (1u16 << expected_selector) * 16;
    let expected_range_shift = num_tables
        .wrapping_mul(16)
        .wrapping_sub(expected_search_range);

    search_range == expected_search_range
        && entry_selector == expected_selector
        && range_shift == expected_range_shift
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::FieldEvaluationContext;
    use crate::paragraph::Alignment;
    use oxml_chart::{
        Axis, AxisData, AxisId, AxisKind, AxisPosition, BarDirection, BarGrouping, CT_PlotArea,
        ChartData, ChartKind, NumericData, Plot, Series, StringRef,
    };
    use oxml_sml::Column;
    use rdocx_oxml::text::Field;
    use rdocx_oxml::units::{HalfPoint, Twips};
    use std::fs;
    use std::io::Cursor;
    use std::process::Command;

    const WORD_VERSION: &str = "16.104";
    const WORD_BUILD: &str = "16.104.25121423";
    const WORD_CHART_CANDIDATE_SHA256: &str =
        "79e9b9ff9e7557dbd09a365bb8c189806e700ed48ca768b27d7158cf2b41370b";
    const FX087_WORD_VERSION: &str = "16.112.3";
    const FX087_WORD_BUILD: &str = "16.112.26083020";
    const FX087_PAGES_VERSION: &str = "15.1.1";
    const FX087_PAGES_BUILD: &str = "7044.0.273";
    const FX087_CANDIDATE_SHA256: &str =
        "54faeec0d56767577afa014564d56571c46d00df11c73baaa38889999a39b3f9";

    fn replace_numbering_xml(document: &mut Document, xml: Vec<u8>) -> Vec<u8> {
        let mut package =
            OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        package.set_part(DEFAULT_NUMBERING_PART, xml);
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        bytes.into_inner()
    }

    #[cfg(all(feature = "digital-signatures", not(target_arch = "wasm32")))]
    fn signature_fixture(name: &str) -> Vec<u8> {
        use base64::Engine as _;

        let source = include_str!("../../oxml-opc/src/signature.rs");
        let prefix = format!("const {name}: &str = \"");
        let encoded = source
            .split_once(&prefix)
            .and_then(|(_, remainder)| remainder.split_once("\";").map(|(value, _)| value))
            .expect("signature fixture remains available");
        base64::engine::general_purpose::STANDARD
            .decode(encoded)
            .unwrap()
    }

    #[test]
    fn identifier_reservation_reports_relationship_overflow() {
        let mut occupied = HashSet::from([format!("rId{}", u32::MAX)]);
        let mut cursor = u32::MAX;
        let error = reserve_relationship_from_cursor(&mut occupied, &mut cursor).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("relationship id range is exhausted")
        );
        assert_eq!(occupied.len(), 1, "a failed reservation must be atomic");
    }

    #[test]
    fn provisional_relationships_reserve_numeric_capacity() {
        let owner = "/word/document.xml";
        let mut package = OpcPackage::new();
        package.get_or_create_part_rels(owner).add_with_id(
            &format!("rId{}", u32::MAX - 1),
            "urn:producer",
            "producer.bin",
        );
        let mut identifiers = DocumentIdentifiers::scan(&package).unwrap();
        let first = identifiers
            .reserve_bundle_relationship_id_checked(owner, rel_types::STYLES)
            .unwrap();
        assert!(first.starts_with("rdocxDeferredStyles"));
        assert!(
            identifiers
                .reserve_bundle_relationship_id_checked(owner, rel_types::NUMBERING)
                .unwrap_err()
                .to_string()
                .contains("relationship id range is exhausted")
        );
        assert!(identifiers.reserve_relationship_id_checked(owner).is_err());
        assert_eq!(identifiers.provisional_relationship_ids[owner].len(), 1);
    }

    #[test]
    fn encoded_identifier_aliases_collide_during_package_scan() {
        let mut package =
            OpcPackage::with_main_part("word/document.xml", content_types::WORD_DOCUMENT);
        package.set_part(
            "/word/document.xml",
            format!(
                r#"<w:document xmlns:w="{WORD_NAMESPACE}" xmlns:wp="{}"><w:body><wp:docPr id="1"/><wp:docPr id="&#49;"/></w:body></w:document>"#,
                drawing_ns::WP,
            )
            .into_bytes(),
        );
        let error = DocumentIdentifiers::scan(&package).unwrap_err();
        assert!(error.to_string().contains("duplicate drawing id 1"));
    }

    #[test]
    fn merge_preflights_all_required_bundle_relationships() {
        let owner = "/word/document.xml";
        let mut source =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let mut package = OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
        package
            .get_or_create_part_rels(owner)
            .items
            .retain(|relationship| relationship.rel_type != rel_types::STYLES);
        package.get_or_create_part_rels(owner).add_with_id(
            &format!("rId{}", u32::MAX - 1),
            "urn:producer",
            "producer.bin",
        );
        package.parts.remove(DEFAULT_STYLES_PART);
        package.content_types.overrides.remove(DEFAULT_STYLES_PART);
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let mut target = Document::from_bytes(bytes.get_ref()).unwrap();

        let mut other = Document::new();
        other
            .add_style(StyleBuilder::paragraph("Merged", "Merged"))
            .unwrap();
        other.add_list_definition(&[ListLevel::decimal()]);
        let before = {
            let mut bytes = Cursor::new(Vec::new());
            target.package.write_to(&mut bytes).unwrap();
            bytes.into_inner()
        };
        assert!(target.reserve_merge_bundles(&other).is_err());
        let after = {
            let mut bytes = Cursor::new(Vec::new());
            target.package.write_to(&mut bytes).unwrap();
            bytes.into_inner()
        };
        assert_eq!(after, before);
        assert!(target.styles_part_name.is_none());
        assert!(target.numbering_part_name.is_none());
    }

    #[test]
    fn preferred_part_fallback_uses_the_greatest_positive_family_suffix() {
        fn reserve(existing: &[&str], preferred: &str) -> String {
            let mut package = OpcPackage::new();
            for part_name in existing {
                package.set_part(part_name, Vec::new());
            }
            DocumentIdentifiers::scan(&package)
                .unwrap()
                .reserve_preferred_part_name(preferred)
                .unwrap()
        }

        assert_eq!(
            reserve(&["/word/header1.xml"], "/word/header1.xml"),
            "/word/header2.xml"
        );
        assert_eq!(
            reserve(&["/WORD/HEADER1.XML"], "/word/header1.xml"),
            "/word/header2.xml"
        );
        let mut relationship_only = OpcPackage::new();
        relationship_only
            .get_or_create_part_rels("/word/document.xml")
            .add("urn:producer", "HEADER1.XML");
        assert_eq!(
            DocumentIdentifiers::scan(&relationship_only)
                .unwrap()
                .reserve_preferred_part_name("/word/header1.xml")
                .unwrap(),
            "/word/header2.xml"
        );
        assert_eq!(
            reserve(
                &[
                    "/word/comments.xml",
                    "/word/comments4.xml",
                    "/word/comments0.xml",
                    "/word/comments-9.xml",
                    "/word/commentsx.xml",
                ],
                "/word/comments.xml",
            ),
            "/word/comments5.xml"
        );
        assert_eq!(
            reserve(&["/word/footnotes.xml"], "/word/footnotes.xml"),
            "/word/footnotes1.xml"
        );

        let mut identifiers = DocumentIdentifiers::scan(&OpcPackage::new()).unwrap();
        for malformed in [
            "comments.xml",
            "/word/comments",
            "/word/.xml",
            "/word/1.xml",
        ] {
            assert!(identifiers.reserve_preferred_part_name(malformed).is_err());
        }
        assert!(identifiers.part_names.is_empty());
    }

    #[test]
    fn relationship_only_targets_stay_reserved_when_facades_materialize_them() {
        fn reopen(package: OpcPackage) -> Document {
            let mut bytes = Cursor::new(Vec::new());
            package.write_to(&mut bytes).unwrap();
            Document::from_bytes(bytes.get_ref()).unwrap()
        }

        let mut settings_source = Document::new();
        let mut settings_package =
            OpcPackage::from_reader(Cursor::new(settings_source.to_bytes().unwrap())).unwrap();
        settings_package.parts.remove(DEFAULT_SETTINGS_PART);
        settings_package
            .content_types
            .overrides
            .remove(DEFAULT_SETTINGS_PART);
        let mut settings = reopen(settings_package);
        settings.set_auto_hyphenation(true).unwrap();
        assert_ne!(
            settings
                .identifiers
                .reserve_fragment_part_name(DEFAULT_SETTINGS_PART)
                .unwrap(),
            DEFAULT_SETTINGS_PART
        );

        let mut comments_source = Document::new();
        comments_source.add_paragraph("commented");
        let mut comments_package =
            OpcPackage::from_reader(Cursor::new(comments_source.to_bytes().unwrap())).unwrap();
        comments_package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id("producerComments", rel_types::COMMENTS, "comments7.xml");
        comments_package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id(
                "producerCommentsExtended",
                crate::comments::COMMENTS_EXTENDED_REL_TYPE,
                "commentsExtended7.xml",
            );
        let mut comments = reopen(comments_package);
        comments
            .add_comment(
                crate::comments::RunRange {
                    start: crate::comments::RunPosition {
                        body_index: 0,
                        run_index: 0,
                    },
                    end: crate::comments::RunPosition {
                        body_index: 0,
                        run_index: 1,
                    },
                },
                "Reviewer",
                None,
                "note",
            )
            .unwrap();
        for target in ["/word/comments7.xml", "/word/commentsExtended7.xml"] {
            assert_ne!(
                comments
                    .identifiers
                    .reserve_fragment_part_name(target)
                    .unwrap(),
                target
            );
        }

        let mut header_source = Document::new();
        header_source.set_header("producer");
        let mut header_package =
            OpcPackage::from_reader(Cursor::new(header_source.to_bytes().unwrap())).unwrap();
        header_package.parts.remove("/word/header1.xml");
        header_package
            .content_types
            .overrides
            .remove("/word/header1.xml");
        let mut header = reopen(header_package);
        header.set_header("replacement");
        assert_ne!(
            header
                .identifiers
                .reserve_fragment_part_name("/word/header1.xml")
                .unwrap(),
            "/word/header1.xml"
        );
    }

    #[test]
    fn authored_bundles_preserve_unrelated_conventional_parts() {
        let mut source =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let mut package = OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
        package
            .get_or_create_part_rels("/word/document.xml")
            .items
            .retain(|relationship| relationship.rel_type != rel_types::STYLES);
        let orphans = [
            (DEFAULT_STYLES_PART, b"producer-styles".as_slice()),
            (DEFAULT_NUMBERING_PART, b"producer-numbering".as_slice()),
            (DEFAULT_CORE_PROPERTIES_PART, b"producer-core".as_slice()),
        ];
        for (part_name, bytes) in orphans {
            package.set_part(part_name, bytes.to_vec());
            package
                .content_types
                .add_override(part_name, "application/example+xml");
        }
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let mut document = Document::from_bytes(bytes.get_ref()).unwrap();

        document
            .add_style(StyleBuilder::paragraph("Custom", "Custom"))
            .unwrap();
        document.add_list_definition(&[ListLevel::decimal()]);
        document.set_title("Title");
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(Cursor::new(saved)).unwrap();

        for (part_name, expected) in orphans {
            assert_eq!(package.get_part(part_name), Some(expected));
            assert_eq!(
                package.content_types.content_type_for(part_name),
                Some("application/example+xml")
            );
        }
        let document_relationships = package.get_part_rels("/word/document.xml").unwrap();
        for (rel_type, expected) in [
            (rel_types::STYLES, "/word/styles1.xml"),
            (rel_types::NUMBERING, "/word/numbering1.xml"),
        ] {
            let relationship = document_relationships.get_by_type(rel_type).unwrap();
            assert_eq!(
                OpcPackage::resolve_rel_target("/word/document.xml", &relationship.target),
                expected
            );
        }
        let core_relationship = package
            .package_rels
            .get_by_type(CORE_PROPERTIES_REL_TYPE)
            .unwrap();
        assert_eq!(
            OpcPackage::resolve_rel_target("/", &core_relationship.target),
            "/docProps/core1.xml"
        );
    }

    #[test]
    fn identifier_content_type_registry_uses_ascii_case_insensitive_keys() {
        let mut package = OpcPackage::new();
        package.content_types.add_default("PNG", "image/png");
        package
            .content_types
            .add_override("/WORD/DOCUMENT.XML", content_types::WORD_DOCUMENT);
        let mut identifiers = DocumentIdentifiers::scan(&package).unwrap();
        assert!(identifiers.content_type_defaults.contains("png"));
        assert!(
            identifiers
                .content_type_overrides
                .contains("/word/document.xml")
        );

        identifiers.register_content_type_default("JPG");
        identifiers.register_content_type_override("/WORD/HEADER1.XML");
        assert!(identifiers.content_type_defaults.contains("jpg"));
        assert!(identifiers.authored_content_type_defaults.contains("jpg"));
        assert!(
            identifiers
                .content_type_overrides
                .contains("/word/header1.xml")
        );
        identifiers.retire_authored_part("/WORD/HEADER1.XML");
        assert!(
            !identifiers
                .content_type_overrides
                .contains("/word/header1.xml")
        );
    }

    #[test]
    fn read_only_save_canonicalizes_a_missing_styles_edge() {
        let mut source =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let mut package = OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
        package
            .get_or_create_part_rels("/word/document.xml")
            .items
            .retain(|relationship| relationship.rel_type != rel_types::STYLES);
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let mut document = Document::from_bytes(bytes.get_ref()).unwrap();

        let saved = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let styles = saved
            .get_part_rels("/word/document.xml")
            .and_then(|relationships| relationships.get_by_type(rel_types::STYLES))
            .unwrap();
        assert!(
            styles
                .id
                .strip_prefix("rId")
                .is_some_and(|suffix| suffix.parse::<u32>().is_ok())
        );
        assert!(!styles.id.starts_with("rdocxDeferred"));
    }

    #[cfg(all(feature = "digital-signatures", not(target_arch = "wasm32")))]
    #[test]
    fn signing_and_subsequent_save_share_canonical_package_preparation() {
        let mut source =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        source.add_paragraph("signed");
        let mut package = OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
        package
            .get_or_create_part_rels("/word/document.xml")
            .items
            .retain(|relationship| relationship.rel_type != rel_types::STYLES);
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let mut document = Document::from_bytes(bytes.get_ref()).unwrap();

        let report = document
            .sign(
                &signature_fixture("TEST_PRIVATE_KEY_PKCS8_BASE64"),
                &signature_fixture("TEST_CERTIFICATE_DER_BASE64"),
            )
            .unwrap();
        assert!(report.cryptographically_valid);
        assert!(report.coverage_complete);
        let styles = document
            .package
            .get_part_rels("/word/document.xml")
            .and_then(|relationships| {
                relationships.items.iter().find(|relationship| {
                    relationship.rel_type == rel_types::STYLES
                        && relationship_is_internal(relationship)
                })
            })
            .unwrap();
        assert!(
            styles
                .id
                .strip_prefix("rId")
                .is_some_and(|suffix| suffix.parse::<u32>().is_ok())
        );

        let saved = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&saved).unwrap();
        let reports = reopened.verify_signatures().unwrap();
        assert_eq!(reports.len(), 1);
        assert!(reports[0].cryptographically_valid);
        assert!(reports[0].coverage_complete);
    }

    #[test]
    fn staged_numbering_bundle_uses_a_collision_safe_nonsemantic_relationship_id() {
        fn image_relationship_id(bytes: Vec<u8>) -> String {
            OpcPackage::from_reader(Cursor::new(bytes))
                .unwrap()
                .get_part_rels("/word/document.xml")
                .and_then(|relationships| relationships.get_by_type(rel_types::IMAGE))
                .expect("image relationship")
                .id
                .clone()
        }

        let mut control =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        control.add_picture(b"image", "image.png", Length::pt(1.0), Length::pt(1.0));
        control.add_list_definition(&[ListLevel::decimal()]);
        let control_image_id = image_relationship_id(control.to_bytes().unwrap());

        let mut document =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        document.add_list_definition(&[ListLevel::decimal()]);
        let provisional = document
            .package
            .get_part_rels("/word/document.xml")
            .and_then(|relationships| relationships.get_by_type(rel_types::NUMBERING))
            .expect("staged numbering relationship")
            .id
            .clone();
        assert!(provisional.starts_with("rdocxDeferredNumbering"));
        document.add_picture(b"image", "image.png", Length::pt(1.0), Length::pt(1.0));
        assert_eq!(
            image_relationship_id(document.to_bytes().unwrap()),
            control_image_id
        );

        let mut staged = document.clone_for_staging();
        staged.canonicalize_authored_identifiers().unwrap();
        let numeric = staged
            .package
            .get_part_rels("/word/document.xml")
            .and_then(|relationships| relationships.get_by_type(rel_types::NUMBERING))
            .expect("canonical numbering relationship")
            .id
            .clone();
        assert!(
            numeric
                .strip_prefix("rId")
                .is_some_and(|value| value.parse::<u32>().is_ok())
        );
        let occupied = staged
            .identifiers
            .relationship_ids
            .get("/word/document.xml")
            .unwrap();
        assert!(occupied.contains(&numeric));
        assert!(!occupied.contains(&provisional));
        let authored = staged
            .identifiers
            .authored_bundle_relationship_ids
            .get("/word/document.xml")
            .unwrap();
        assert!(authored.contains(&numeric));
        assert!(!authored.contains(&provisional));
        staged.canonicalize_authored_identifiers().unwrap();
        assert_eq!(
            staged
                .package
                .get_part_rels("/word/document.xml")
                .and_then(|relationships| relationships.get_by_type(rel_types::NUMBERING))
                .unwrap()
                .id,
            numeric
        );

        let mut collision =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        collision
            .package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id("rId99", "urn:producer", "producer.bin");
        collision.identifiers = DocumentIdentifiers::scan(&collision.package).unwrap();
        collision.add_list_definition(&[ListLevel::decimal()]);
        let collision_bytes = collision.to_bytes().unwrap();
        let collision_package = OpcPackage::from_reader(Cursor::new(collision_bytes)).unwrap();
        let relationships = collision_package
            .get_part_rels("/word/document.xml")
            .unwrap();
        assert_eq!(
            relationships.get_by_id("rId99").unwrap().rel_type,
            "urn:producer"
        );
        assert!(
            relationships
                .get_by_type(rel_types::NUMBERING)
                .expect("numbering fallback relationship")
                .id
                .starts_with("rId")
        );

        let mut producer =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        producer
            .package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id("rdocxDeferredNumbering", "urn:producer", "producer.bin");
        producer.identifiers = DocumentIdentifiers::scan(&producer.package).unwrap();
        producer.add_list_definition(&[ListLevel::decimal()]);
        assert_eq!(
            producer
                .package
                .get_part_rels("/word/document.xml")
                .and_then(|relationships| relationships.get_by_type(rel_types::NUMBERING))
                .unwrap()
                .id,
            "rdocxDeferredNumbering1"
        );
        let saved = OpcPackage::from_reader(Cursor::new(producer.to_bytes().unwrap())).unwrap();
        let relationships = saved.get_part_rels("/word/document.xml").unwrap();
        assert_eq!(
            relationships
                .get_by_id("rdocxDeferredNumbering")
                .unwrap()
                .rel_type,
            "urn:producer"
        );
        let numbering = relationships.get_by_type(rel_types::NUMBERING).unwrap();
        assert!(
            numbering
                .id
                .strip_prefix("rId")
                .is_some_and(|value| value.parse::<u32>().is_ok())
        );
        assert!(!relationships.items.iter().any(|relationship| {
            relationship.id.starts_with("rdocxDeferredNumbering")
                && relationship.rel_type == rel_types::NUMBERING
        }));
    }

    #[test]
    fn prepared_reopen_preserves_authored_bundle_provenance_and_final_order() {
        fn build(image_before_reopen: bool) -> Vec<u8> {
            let mut document = Document::new_with_profile(WordCreationProfile::Minimal(
                WordPackageClass::Document,
            ));
            document.add_list_definition(&[ListLevel::decimal()]);
            let provisional = document
                .package
                .get_part_rels("/word/document.xml")
                .and_then(|relationships| relationships.get_by_type(rel_types::NUMBERING))
                .unwrap()
                .id
                .clone();
            assert!(provisional.starts_with("rdocxDeferredNumbering"));
            if image_before_reopen {
                document.add_picture(b"image", "image.png", Length::pt(1.0), Length::pt(1.0));
            }

            let mut document = document.prepare_and_reopen_staged().unwrap();
            let numbering = document
                .package
                .get_part_rels("/word/document.xml")
                .and_then(|relationships| relationships.get_by_type(rel_types::NUMBERING))
                .unwrap()
                .id
                .clone();
            assert!(
                numbering
                    .strip_prefix("rId")
                    .is_some_and(|suffix| suffix.parse::<u32>().is_ok())
            );
            assert!(
                document.identifiers.authored_bundle_relationship_ids["/word/document.xml"]
                    .contains(&numbering)
            );
            assert!(
                !document.identifiers.preserved_relationship_ids["/word/document.xml"]
                    .contains(&numbering)
            );
            if !image_before_reopen {
                document.add_picture(b"image", "image.png", Length::pt(1.0), Length::pt(1.0));
            }
            document.to_bytes().unwrap()
        }

        assert_eq!(build(false), build(true));
    }

    #[test]
    fn external_bundle_edges_do_not_shadow_or_satisfy_internal_parts() {
        let mut source = Document::new();
        let mut package = OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
        for (id, mode) in [
            ("producerExternalStyles", "External"),
            ("producerLowercaseStyles", "internal"),
            ("producerMalformedStyles", "ProducerMode"),
        ] {
            package
                .get_or_create_part_rels("/word/document.xml")
                .items
                .insert(
                    0,
                    oxml_opc::relationship::Relationship {
                        id: id.to_owned(),
                        rel_type: rel_types::STYLES.to_owned(),
                        target: format!("https://example.com/{id}.xml"),
                        target_mode: Some(mode.to_owned()),
                    },
                );
        }
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let document = Document::from_bytes(bytes.get_ref()).unwrap();
        assert_eq!(
            document.styles_part_name.as_deref(),
            Some(DEFAULT_STYLES_PART)
        );

        let mut external_only =
            OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
        external_only
            .get_or_create_part_rels("/word/document.xml")
            .items
            .retain(|relationship| relationship.rel_type != rel_types::STYLES);
        external_only
            .get_or_create_part_rels("/word/document.xml")
            .items
            .push(oxml_opc::relationship::Relationship {
                id: "rdocxStyles".to_owned(),
                rel_type: rel_types::STYLES.to_owned(),
                target: "https://example.com/styles.xml".to_owned(),
                target_mode: Some("External".to_owned()),
            });
        external_only.parts.remove(DEFAULT_STYLES_PART);
        external_only
            .content_types
            .overrides
            .remove(DEFAULT_STYLES_PART);
        let mut bytes = Cursor::new(Vec::new());
        external_only.write_to(&mut bytes).unwrap();
        let mut document = Document::from_bytes(bytes.get_ref()).unwrap();
        document
            .add_style(StyleBuilder::paragraph("Custom", "Custom"))
            .unwrap();
        let saved = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let relationships = saved.get_part_rels("/word/document.xml").unwrap();
        assert!(relationships.items.iter().any(|relationship| {
            relationship.id == "rdocxStyles"
                && relationship.target_mode.as_deref() == Some("External")
        }));
        assert!(relationships.items.iter().any(|relationship| {
            relationship.id.starts_with("rId")
                && relationship.rel_type == rel_types::STYLES
                && relationship_is_internal(relationship)
        }));

        let mut malformed_header = Document::new();
        malformed_header.set_header("producer header");
        let header_id = malformed_header
            .document
            .body
            .sect_pr
            .as_ref()
            .unwrap()
            .header_refs[0]
            .rel_id
            .clone();
        malformed_header
            .package
            .get_or_create_part_rels("/word/document.xml")
            .items
            .iter_mut()
            .find(|relationship| relationship.id == header_id)
            .unwrap()
            .target_mode = Some("internal".to_owned());
        assert_eq!(malformed_header.header_text(), None);
        assert!(malformed_header.active_header_footer_parts().is_empty());
        malformed_header.set_header("replacement header");
        assert_eq!(
            malformed_header.header_text().as_deref(),
            Some("replacement header")
        );
        let active_relationship = malformed_header
            .package
            .get_part_rels("/word/document.xml")
            .unwrap()
            .items
            .iter()
            .find(|relationship| {
                relationship.rel_type == rel_types::HEADER && relationship_is_internal(relationship)
            })
            .unwrap();
        assert_ne!(active_relationship.id, header_id);
    }

    #[test]
    fn malformed_header_modes_do_not_shadow_or_enter_raw_mutations() {
        let mut document = Document::new();
        document.set_header("valid secret");
        let valid_id = document.section_properties().unwrap().header_refs[0]
            .rel_id
            .clone();
        document.package.set_part(
            "/word/producer-header.xml",
            format!(
                r#"<w:hdr xmlns:w="{}"><w:p><w:r><w:t>producer secret</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
        );
        document
            .package
            .get_or_create_part_rels("/word/document.xml")
            .items
            .insert(
                0,
                oxml_opc::relationship::Relationship {
                    id: "producerHeader".to_owned(),
                    rel_type: rel_types::HEADER.to_owned(),
                    target: "producer-header.xml".to_owned(),
                    target_mode: Some("internal".to_owned()),
                },
            );
        document.section_properties_mut().header_refs.insert(
            0,
            HdrFtrRef {
                hdr_ftr_type: HdrFtrType::Default,
                rel_id: "producerHeader".to_owned(),
            },
        );

        assert_eq!(document.replace_regex("secret", "changed").unwrap(), 1);
        assert_eq!(document.header_text().as_deref(), Some("valid changed"));
        assert!(
            String::from_utf8_lossy(
                document
                    .package
                    .get_part("/word/producer-header.xml")
                    .unwrap()
            )
            .contains("producer secret")
        );
        assert!(document.load_header_footer(&valid_id, true).is_some());

        let mut external_only = Document::new();
        external_only.set_header("external secret");
        let relationship = external_only
            .package
            .get_or_create_part_rels("/word/document.xml")
            .items
            .iter_mut()
            .find(|relationship| relationship.rel_type == rel_types::HEADER)
            .unwrap();
        relationship.target_mode = Some("External".to_owned());
        assert_eq!(external_only.replace_regex("secret", "changed").unwrap(), 0);
        assert!(external_only.header_text().is_none());

        let mut watermark = Document::new();
        watermark.set_header("producer header");
        let producer_id = watermark.section_properties().unwrap().header_refs[0]
            .rel_id
            .clone();
        watermark
            .package
            .get_or_create_part_rels("/word/document.xml")
            .items
            .iter_mut()
            .find(|relationship| relationship.id == producer_id)
            .unwrap()
            .target_mode = Some("ProducerDefined".to_owned());
        watermark.set_text_watermark("DRAFT").unwrap();
        assert!(
            watermark
                .package
                .get_part_rels("/word/document.xml")
                .unwrap()
                .items
                .iter()
                .any(|relationship| {
                    relationship.rel_type == rel_types::HEADER
                        && relationship.id != producer_id
                        && relationship_is_internal(relationship)
                })
        );
    }

    #[test]
    fn cross_type_header_footer_references_never_mutate_the_target_story() {
        let mut document = Document::new();
        document.set_header("header secret");
        document.set_footer("footer secret");
        let section = document.section_properties().unwrap();
        let header_id = section.header_refs[0].rel_id.clone();
        let footer_id = section.footer_refs[0].rel_id.clone();
        let header_part = document.header_footer_part_name(&header_id, true).unwrap();
        let footer_part = document.header_footer_part_name(&footer_id, false).unwrap();
        let section = document.section_properties_mut();
        section.header_refs[0].rel_id.clone_from(&footer_id);
        section.footer_refs[0].rel_id.clone_from(&header_id);

        assert_eq!(document.replace_regex("secret", "changed").unwrap(), 0);
        assert!(
            String::from_utf8_lossy(document.package.get_part(&header_part).unwrap())
                .contains("header secret")
        );
        assert!(
            String::from_utf8_lossy(document.package.get_part(&footer_part).unwrap())
                .contains("footer secret")
        );
        assert!(document.load_header_footer(&footer_id, true).is_none());
        assert!(document.load_header_footer(&header_id, false).is_none());

        fn assert_setter_reserves_a_fresh_part(
            is_header: bool,
            hdr_type: HdrFtrType,
            setter: impl FnOnce(&mut Document),
        ) {
            let producer_part = "/word/producer-cross-type.xml";
            let producer_bytes = b"producer bytes that must not change";
            let mut document = Document::new();
            document
                .package
                .set_part(producer_part, producer_bytes.to_vec());
            let cross_type = if is_header {
                rel_types::FOOTER
            } else {
                rel_types::HEADER
            };
            document
                .package
                .get_or_create_part_rels("/word/document.xml")
                .add_with_id("producerCrossType", cross_type, "producer-cross-type.xml");
            let reference = HdrFtrRef {
                hdr_ftr_type: hdr_type,
                rel_id: "producerCrossType".to_owned(),
            };
            if is_header {
                document
                    .section_properties_mut()
                    .header_refs
                    .push(reference);
            } else {
                document
                    .section_properties_mut()
                    .footer_refs
                    .push(reference);
            }

            setter(&mut document);

            assert_eq!(
                document.package.get_part(producer_part),
                Some(producer_bytes.as_slice())
            );
            let section = document.section_properties().unwrap();
            let installed = if is_header {
                &section.header_refs
            } else {
                &section.footer_refs
            }
            .iter()
            .find(|reference| reference.hdr_ftr_type == hdr_type)
            .unwrap();
            assert_ne!(installed.rel_id, "producerCrossType");
            let installed_part = document
                .header_footer_part_name(&installed.rel_id, is_header)
                .unwrap();
            assert_ne!(installed_part, producer_part);
        }

        let png = super::watermark_tests::PNG;
        let size = Length::pt(12.0);
        assert_setter_reserves_a_fresh_part(true, HdrFtrType::Default, |document| {
            document.set_header("new header")
        });
        assert_setter_reserves_a_fresh_part(false, HdrFtrType::Default, |document| {
            document.set_footer("new footer")
        });
        assert_setter_reserves_a_fresh_part(true, HdrFtrType::First, |document| {
            document.set_first_page_header("new first header")
        });
        assert_setter_reserves_a_fresh_part(false, HdrFtrType::First, |document| {
            document.set_first_page_footer("new first footer")
        });
        assert_setter_reserves_a_fresh_part(true, HdrFtrType::Default, |document| {
            document.set_header_image(png, "image.png", size, size)
        });
        assert_setter_reserves_a_fresh_part(false, HdrFtrType::Default, |document| {
            document.set_footer_image(png, "image.png", size, size)
        });
        assert_setter_reserves_a_fresh_part(true, HdrFtrType::Default, |document| {
            document.set_raw_header_with_images(
                format!(r#"<w:hdr xmlns:w="{}"/>"#, rdocx_oxml::namespace::W_NS).into_bytes(),
                &[],
                HdrFtrType::Default,
            )
        });
        assert_setter_reserves_a_fresh_part(false, HdrFtrType::Default, |document| {
            document.set_raw_footer_with_images(
                format!(r#"<w:ftr xmlns:w="{}"/>"#, rdocx_oxml::namespace::W_NS).into_bytes(),
                &[],
                HdrFtrType::Default,
            )
        });
        assert_setter_reserves_a_fresh_part(true, HdrFtrType::Default, |document| {
            document.set_header_image_with_background(png, "image.png", size, size, "000000")
        });
        assert_setter_reserves_a_fresh_part(true, HdrFtrType::First, |document| {
            document.set_first_page_header_image(png, "image.png", size, size)
        });
    }

    #[test]
    fn unrelated_header_shaped_target_does_not_shadow_or_enter_replacement() {
        let mut document = Document::new();
        document.set_header("valid secret");
        document.package.set_part(
            "/word/producer-shaped.xml",
            format!(
                r#"<w:hdr xmlns:w="{}"><w:p><w:r><w:t>producer secret</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
        );
        document
            .package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id("producerShaped", rel_types::IMAGE, "producer-shaped.xml");
        document.section_properties_mut().header_refs.insert(
            0,
            HdrFtrRef {
                hdr_ftr_type: HdrFtrType::Default,
                rel_id: "producerShaped".to_owned(),
            },
        );

        let replacements = HashMap::from([("secret", "changed")]);
        assert_eq!(document.replace_all(&replacements), 1);
        assert_eq!(document.header_text().as_deref(), Some("valid changed"));
        assert!(
            String::from_utf8_lossy(
                document
                    .package
                    .get_part("/word/producer-shaped.xml")
                    .unwrap()
            )
            .contains("producer secret")
        );
    }

    #[test]
    fn bundle_and_semantic_relationships_are_canonicalized_together() {
        fn build(reverse: bool) -> Vec<u8> {
            let mut document = Document::new_with_profile(WordCreationProfile::Minimal(
                WordPackageClass::Document,
            ));
            document.add_paragraph("review");
            document.add_paragraph("link: ");
            document
                .package
                .get_or_create_part_rels("/word/document.xml")
                .add_with_id("rId40", "urn:producer", "producer.bin");
            document.identifiers = DocumentIdentifiers::scan(&document.package).unwrap();
            if reverse {
                document
                    .add_comment(
                        crate::comments::RunRange {
                            start: crate::comments::RunPosition {
                                body_index: 0,
                                run_index: 0,
                            },
                            end: crate::comments::RunPosition {
                                body_index: 0,
                                run_index: 1,
                            },
                        },
                        "Reviewer",
                        None,
                        "note",
                    )
                    .unwrap();
                document.set_first_page_header("first");
                document.set_header("default");
                document.add_picture(b"image", "image.png", Length::pt(1.0), Length::pt(1.0));
                let hyperlink = document.add_hyperlink_relationship("https://example.com");
                document
                    .paragraph_mut(1)
                    .unwrap()
                    .add_hyperlink("example", &hyperlink);
                document.add_list_definition(&[ListLevel::decimal()]);
                document
                    .add_style(StyleBuilder::paragraph("Custom", "Custom"))
                    .unwrap();
                document.set_auto_hyphenation(true).unwrap();
            } else {
                document.set_auto_hyphenation(true).unwrap();
                document
                    .add_style(StyleBuilder::paragraph("Custom", "Custom"))
                    .unwrap();
                document.add_list_definition(&[ListLevel::decimal()]);
                let hyperlink = document.add_hyperlink_relationship("https://example.com");
                document
                    .paragraph_mut(1)
                    .unwrap()
                    .add_hyperlink("example", &hyperlink);
                document.add_picture(b"image", "image.png", Length::pt(1.0), Length::pt(1.0));
                document.set_header("default");
                document.set_first_page_header("first");
                document
                    .add_comment(
                        crate::comments::RunRange {
                            start: crate::comments::RunPosition {
                                body_index: 0,
                                run_index: 0,
                            },
                            end: crate::comments::RunPosition {
                                body_index: 0,
                                run_index: 1,
                            },
                        },
                        "Reviewer",
                        None,
                        "note",
                    )
                    .unwrap();
            }
            document.to_bytes().unwrap()
        }

        assert_eq!(build(false), build(true));
        let package = OpcPackage::from_reader(Cursor::new(build(false))).unwrap();
        let relationships = package.get_part_rels("/word/document.xml").unwrap();
        assert_eq!(
            relationships.get_by_id("rId40").unwrap().rel_type,
            "urn:producer"
        );
        assert!(
            relationships
                .get_by_type(rel_types::NUMBERING)
                .unwrap()
                .id
                .starts_with("rId")
        );
        assert!(
            relationships
                .get_by_type(rel_types::IMAGE)
                .unwrap()
                .id
                .starts_with("rId")
        );
    }

    #[test]
    fn deferred_bundle_facades_panic_before_live_mutation_on_exhaustion() {
        fn document_relationship_exhausted(_base: &str) -> Document {
            let mut source = Document::new_with_profile(WordCreationProfile::Minimal(
                WordPackageClass::Document,
            ));
            let mut package =
                OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
            package
                .get_or_create_part_rels("/word/document.xml")
                .items
                .retain(|relationship| relationship.rel_type != rel_types::STYLES);
            package
                .get_or_create_part_rels("/word/document.xml")
                .add_with_id(
                    &format!("rId{}", u32::MAX),
                    "urn:exhaustion",
                    "unchanged.bin",
                );
            let mut bytes = Cursor::new(Vec::new());
            package.write_to(&mut bytes).unwrap();
            Document::from_bytes(bytes.get_ref()).unwrap()
        }

        fn package_bytes(document: &Document) -> Vec<u8> {
            let mut bytes = Cursor::new(Vec::new());
            document.package.write_to(&mut bytes).unwrap();
            bytes.into_inner()
        }

        for (base, mutation) in [
            (
                "rdocxNumbering",
                (|document: &mut Document| {
                    document.add_bullet_list_item("item", 0);
                }) as fn(&mut Document),
            ),
            ("rdocxNumbering", |document| {
                document.add_numbered_list_item("item", 0);
            }),
            ("rdocxNumbering", |document| {
                document.add_list_definition(&[ListLevel::decimal()]);
            }),
            ("rdocxStyles", |document| {
                document
                    .add_style(StyleBuilder::paragraph("Custom", "Custom"))
                    .unwrap();
            }),
            ("rdocxFootnotes", |document| {
                document.add_footnote("note");
            }),
        ] {
            let mut document = document_relationship_exhausted(base);
            let before = package_bytes(&document);
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                mutation(&mut document);
            }));
            assert!(result.is_err());
            assert_eq!(package_bytes(&document), before);
            assert_eq!(document.paragraph_count(), 0);
            assert!(document.numbering.is_none());
            assert!(document.footnotes().is_empty());
            assert!(document.style("Custom").is_none());
        }

        for mutation in [
            (|document: &mut Document| document.set_title("Title")) as fn(&mut Document),
            |document| document.set_author("Author"),
            |document| document.set_subject("Subject"),
            |document| document.set_keywords("Keywords"),
        ] {
            let mut document = document_relationship_exhausted("rdocxUnused");
            document.package.package_rels.add_with_id(
                &format!("rId{}", u32::MAX),
                "urn:exhaustion",
                "unchanged.bin",
            );
            document.identifiers = DocumentIdentifiers::scan(&document.package).unwrap();
            let before = package_bytes(&document);
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                mutation(&mut document);
            }));
            assert!(result.is_err());
            assert_eq!(package_bytes(&document), before);
            assert!(document.core_properties.is_none());
        }
    }

    #[test]
    fn failed_save_identifier_reservation_leaves_document_unchanged() {
        let mut document = Document::new();
        document.add_picture(
            b"save-rollback",
            "rollback.png",
            Length::pt(1.0),
            Length::pt(1.0),
        );
        let before = document.to_bytes().unwrap();
        let owner = document.doc_part_name.clone();
        let exhausted = format!("rId{}", u32::MAX);
        document
            .identifiers
            .preserved_relationship_ids
            .entry(owner.clone())
            .or_default()
            .insert(exhausted.clone());

        let error = document.to_bytes().unwrap_err();
        assert!(
            error
                .to_string()
                .contains("relationship id range is exhausted")
        );

        document
            .identifiers
            .preserved_relationship_ids
            .get_mut(&owner)
            .unwrap()
            .remove(&exhausted);
        assert_eq!(document.to_bytes().unwrap(), before);
    }

    #[test]
    fn image_facades_fail_before_mutating_on_relationship_exhaustion() {
        let mut source = Document::new();
        let mut package = OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id(
                &format!("rId{}", u32::MAX),
                "urn:exhaustion",
                "unchanged.bin",
            );
        let mut output = Cursor::new(Vec::new());
        package.write_to(&mut output).unwrap();
        let source_bytes = output.into_inner();

        let package_bytes = |document: &Document| {
            let mut output = Cursor::new(Vec::new());
            document.package.write_to(&mut output).unwrap();
            output.into_inner()
        };
        let mut picture = Document::from_bytes(&source_bytes).unwrap();
        let before = package_bytes(&picture);
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            picture.add_picture(b"image", "image.png", Length::pt(1.0), Length::pt(1.0));
        }));
        assert!(result.is_err());
        assert_eq!(package_bytes(&picture), before);
        assert_eq!(picture.paragraph_count(), 0);

        let mut embedded = Document::from_bytes(&source_bytes).unwrap();
        let before = package_bytes(&embedded);
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            embedded.embed_image(b"image", "image.png");
        }));
        assert!(result.is_err());
        assert_eq!(package_bytes(&embedded), before);

        let mut automatic = Document::from_bytes(&source_bytes).unwrap();
        let before = package_bytes(&automatic);
        assert!(
            automatic
                .add_picture_auto(super::watermark_tests::PNG, "image.png")
                .is_err()
        );
        assert_eq!(package_bytes(&automatic), before);
        assert_eq!(automatic.paragraph_count(), 0);

        let mut background = Document::from_bytes(&source_bytes).unwrap();
        let before = package_bytes(&background);
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            background.add_background_image(b"image", "image.png");
        }));
        assert!(result.is_err());
        assert_eq!(package_bytes(&background), before);
        assert_eq!(background.paragraph_count(), 0);

        let mut anchored = Document::from_bytes(&source_bytes).unwrap();
        let before = package_bytes(&anchored);
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            anchored.add_anchored_image(
                b"image",
                "image.png",
                Length::pt(1.0),
                Length::pt(1.0),
                true,
            );
        }));
        assert!(result.is_err());
        assert_eq!(package_bytes(&anchored), before);
        assert_eq!(anchored.paragraph_count(), 0);
    }

    #[test]
    fn background_and_anchored_images_roll_back_drawing_id_exhaustion() {
        fn assert_rollback(mutation: impl FnOnce(&mut Document)) {
            let mut document = Document::new();
            document.identifiers.drawing_ids.insert(u32::MAX);
            let before = document.clone_for_staging();
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                mutation(&mut document);
            }));
            assert!(result.is_err());
            assert_eq!(document.document, before.document);
            assert_eq!(document.package.parts, before.package.parts);
            assert_eq!(
                document.package.part_rels.len(),
                before.package.part_rels.len()
            );
        }

        assert_rollback(|document| {
            document.add_background_image(b"image", "image.png");
        });
        assert_rollback(|document| {
            document.add_anchored_image(
                b"image",
                "image.png",
                Length::pt(1.0),
                Length::pt(1.0),
                false,
            );
        });
    }

    #[test]
    fn gapped_imported_bookmarks_keep_lowest_free_allocation() {
        let mut source = Document::new();
        source.add_paragraph("zero");
        source.add_paragraph("two");
        for (index, id) in [(0, 0), (1, 2)] {
            let BodyContent::Paragraph(paragraph) = &mut source.document.body.content[index] else {
                panic!("paragraph");
            };
            assert!(paragraph.insert_bookmark_start(0, id, &format!("b{id}")));
            assert!(paragraph.insert_bookmark_end(1, id));
        }
        let mut reopened = Document::from_bytes(&source.to_bytes().unwrap()).unwrap();
        reopened.add_paragraph("one");
        let id = reopened
            .add_bookmark(
                "b1",
                crate::comments::RunRange {
                    start: crate::comments::RunPosition {
                        body_index: 2,
                        run_index: 0,
                    },
                    end: crate::comments::RunPosition {
                        body_index: 2,
                        run_index: 1,
                    },
                },
            )
            .unwrap();
        assert_eq!(id, 1);
    }

    #[test]
    fn toc_reservation_observes_nested_bookmark_ids() {
        let mut document = Document::new();
        {
            let mut table = document.add_table(1, 1);
            let mut cell = table.cell(0, 0).unwrap();
            cell.set_text("nested");
        }
        let BodyContent::Table(table) = &mut document.document.body.content[0] else {
            panic!("table");
        };
        let CellContent::Paragraph(paragraph) = &mut table.rows[0].cells[0].content[0] else {
            panic!("paragraph");
        };
        assert!(paragraph.insert_bookmark_start(0, 100, "nested"));
        assert!(paragraph.insert_bookmark_end(1, 100));
        document.add_paragraph("Chapter").style("Heading1");

        document.insert_toc(0, 1);

        let toc = document
            .bookmarks()
            .into_iter()
            .find(|bookmark| bookmark.name() == Some("_Toc1"))
            .unwrap();
        assert_eq!(toc.id(), Some(0));
    }

    #[test]
    fn cell_and_body_pictures_receive_distinct_document_order_ids() {
        const IMAGE: &[u8] = b"not-a-real-png";
        let mut document = Document::new();
        document.add_picture(IMAGE, "body.png", Length::pt(1.0), Length::pt(1.0));
        let relationship = document.embed_image(IMAGE, "cell.png");
        document.add_table(1, 1).cell(0, 0).unwrap().add_picture(
            &relationship,
            Length::pt(1.0),
            Length::pt(1.0),
        );

        let package = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert_eq!(xml.matches(r#"wp:docPr id="1""#).count(), 1, "{xml}");
        assert_eq!(xml.matches(r#"wp:docPr id="2""#).count(), 1, "{xml}");
    }

    #[test]
    fn hyperlink_and_drawing_relationships_follow_final_paragraph_order() {
        fn build(reverse_allocation: bool) -> Vec<u8> {
            let mut document = Document::new();
            let (hyperlink, picture) = if reverse_allocation {
                let picture = document.embed_image(b"image", "image.png");
                let hyperlink = document.add_hyperlink_relationship("https://example.com");
                (hyperlink, picture)
            } else {
                let hyperlink = document.add_hyperlink_relationship("https://example.com");
                let picture = document.embed_image(b"image", "image.png");
                (hyperlink, picture)
            };
            let mut paragraph = document.add_paragraph("");
            paragraph.add_hyperlink("link", &hyperlink);
            paragraph.add_picture(&picture, Length::pt(1.0), Length::pt(1.0));
            document.to_bytes().unwrap()
        }
        assert_eq!(build(false), build(true));
    }

    #[test]
    fn current_unpreserved_relationships_join_semantic_canonicalization() {
        fn add_theme(document: &mut Document, id: &str) {
            document.package.set_part(
                DEFAULT_THEME_PART,
                oxml_drawing::theme::OFFICE_DEFAULT_XML.as_bytes().to_vec(),
            );
            document
                .package
                .content_types
                .add_override(DEFAULT_THEME_PART, content_types::THEME);
            document
                .package
                .get_or_create_part_rels("/word/document.xml")
                .add_with_id(id, rel_types::THEME, "theme/theme1.xml");
        }

        let mut authored =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        authored
            .add_chart(
                ChartKind::Bar,
                Length::inches(2.0),
                Length::inches(1.0),
                &f158_chart_data(1),
            )
            .unwrap();
        add_theme(&mut authored, "rId40");
        let reopened = Document::from_bytes(&authored.to_bytes().unwrap()).unwrap();
        let relationships = reopened
            .package
            .get_part_rels("/word/document.xml")
            .unwrap();
        assert_eq!(
            relationships.get_by_type(rel_types::CHART).unwrap().id,
            "rId1"
        );
        assert_eq!(
            relationships.get_by_type(rel_types::THEME).unwrap().id,
            "rId2"
        );

        let mut producer =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let mut producer = Document::from_bytes(&producer.to_bytes().unwrap()).unwrap();
        add_theme(&mut producer, "rId9");
        let mut producer_bytes = Cursor::new(Vec::new());
        producer.package.write_to(&mut producer_bytes).unwrap();
        let mut producer = Document::from_bytes(producer_bytes.get_ref()).unwrap();
        producer
            .add_chart(
                ChartKind::Bar,
                Length::inches(2.0),
                Length::inches(1.0),
                &f158_chart_data(1),
            )
            .unwrap();
        let reopened = Document::from_bytes(&producer.to_bytes().unwrap()).unwrap();
        let relationships = reopened
            .package
            .get_part_rels("/word/document.xml")
            .unwrap();
        assert_eq!(
            relationships.get_by_type(rel_types::THEME).unwrap().id,
            "rId9"
        );
        assert_eq!(
            relationships.get_by_type(rel_types::CHART).unwrap().id,
            "rId10"
        );

        let mut opaque =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let relationships = opaque.package.get_or_create_part_rels("/word/document.xml");
        relationships.add_with_id("rId41", "urn:producer", "opaque.bin");
        relationships
            .items
            .push(oxml_opc::relationship::Relationship {
                id: "rId40".to_owned(),
                rel_type: rel_types::HYPERLINK.to_owned(),
                target: "https://example.com/opaque".to_owned(),
                target_mode: Some("External".to_owned()),
            });
        let reopened = Document::from_bytes(&opaque.to_bytes().unwrap()).unwrap();
        let relationships = reopened
            .package
            .get_part_rels("/word/document.xml")
            .unwrap();
        assert!(relationships.items.iter().any(|relationship| {
            relationship.id == "rId1"
                && relationship.rel_type == "urn:producer"
                && relationship.target == "opaque.bin"
                && relationship_is_internal(relationship)
        }));
        assert!(relationships.items.iter().any(|relationship| {
            relationship.id == "rId2"
                && relationship.rel_type == rel_types::HYPERLINK
                && relationship.target == "https://example.com/opaque"
                && relationship.target_mode.as_deref() == Some("External")
        }));
    }

    #[test]
    fn opaque_xml_relationship_ids_are_fixed_across_opposite_allocation_order() {
        const RAW: &[u8] = br#"<producer:item xmlns:producer="urn:producer" r:embed="rId40"/>"#;

        fn build(raw_relationship_first: bool) -> Vec<u8> {
            let mut document = Document::new_with_profile(WordCreationProfile::Minimal(
                WordPackageClass::Document,
            ));
            document
                .body_namespace_bindings
                .push(("xmlns:r".to_owned(), drawing_ns::R.to_owned()));
            let add_raw_relationship = |document: &mut Document| {
                document
                    .package
                    .set_part("/word/opaque.bin", b"opaque relationship target".to_vec());
                document
                    .package
                    .content_types
                    .add_default("bin", "application/octet-stream");
                document
                    .package
                    .get_or_create_part_rels("/word/document.xml")
                    .add_with_id("rId40", "urn:producer:opaque", "opaque.bin");
            };
            if raw_relationship_first {
                add_raw_relationship(&mut document);
            }
            let hyperlink = document.add_hyperlink_relationship("https://example.com");
            if !raw_relationship_first {
                add_raw_relationship(&mut document);
            }
            document.add_paragraph("").add_hyperlink("link", &hyperlink);
            document
                .document
                .body
                .content
                .push(BodyContent::RawXml(RAW.to_vec()));
            document.to_bytes().unwrap()
        }

        let raw_first = build(true);
        let modeled_first = build(false);
        assert_eq!(raw_first, modeled_first);

        let package = OpcPackage::from_reader(Cursor::new(raw_first)).unwrap();
        let relationships = package
            .get_part_rels("/word/document.xml")
            .expect("document relationships");
        assert!(relationships.items.iter().any(|relationship| {
            relationship.id == "rId40"
                && relationship.rel_type == "urn:producer:opaque"
                && relationship.target == "opaque.bin"
        }));
        assert!(relationships.items.iter().any(|relationship| {
            relationship.id == "rId1" && relationship.rel_type == rel_types::HYPERLINK
        }));
        assert!(
            package
                .get_part("/word/document.xml")
                .is_some_and(|xml| xml.windows(RAW.len()).any(|window| window == RAW))
        );
    }

    #[test]
    fn paragraph_local_binding_covers_unsupported_run_child_relationship() {
        const RAW: &[u8] = br#"<w:custom localRel:id="rId&#52;0"/>"#;
        let xml = format!(
            r#"<w:document xmlns:w="{WORD_NAMESPACE}"><w:body><w:p xmlns:localRel="{}"><w:r><w:t>producer</w:t>{}</w:r></w:p><w:sectPr/></w:body></w:document>"#,
            drawing_ns::R,
            std::str::from_utf8(RAW).unwrap(),
        );
        let seed =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let mut package = seed.package;
        package.set_part("/word/document.xml", xml.into_bytes());
        package.set_part("/word/opaque.bin", b"opaque relationship target".to_vec());
        package
            .content_types
            .add_default("bin", "application/octet-stream");
        package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id("rId40", "urn:producer:opaque", "opaque.bin");
        let mut document = Document::from_package(package).unwrap();
        let hyperlink = document.add_hyperlink_relationship("https://example.com/modeled");
        document
            .add_paragraph("")
            .add_hyperlink("modeled", &hyperlink);

        let saved = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let relationships = saved.get_part_rels("/word/document.xml").unwrap();
        assert!(relationships.get_by_id("rId40").is_some());
        assert_eq!(
            relationships.get_by_type(rel_types::HYPERLINK).unwrap().id,
            "rId41"
        );
        assert!(
            saved
                .get_part("/word/document.xml")
                .is_some_and(|xml| xml.windows(RAW.len()).any(|window| window == RAW))
        );
    }

    #[test]
    fn word_main_part_and_relationship_owner_resolve_case_equivalent_spelling() {
        let mut document =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let mut package =
            OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let document_xml = package.parts.remove("/word/document.xml").unwrap();
        package
            .parts
            .insert("/WORD/DOCUMENT.XML".to_owned(), document_xml);
        let relationships = package.part_rels.remove("/word/document.xml").unwrap();
        package
            .part_rels
            .insert("/WORD/DOCUMENT.XML".to_owned(), relationships);
        package.set_part("/word/occupied.bin", b"occupied".to_vec());
        package
            .content_types
            .add_default("bin", "application/octet-stream");
        package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id("rId40", "urn:producer:occupied", "occupied.bin");

        let mut reopened = Document::from_package(package).unwrap();
        assert!(
            reopened
                .identifiers
                .relationship_ids
                .contains_key("/word/document.xml")
        );
        assert!(
            !reopened
                .identifiers
                .relationship_ids
                .contains_key("/WORD/DOCUMENT.XML")
        );
        let hyperlink = reopened.add_hyperlink_relationship("https://example.com/mixed-case");
        assert_eq!(hyperlink, "rId41");
        reopened
            .add_paragraph("")
            .add_hyperlink("mixed case", &hyperlink);
        let saved = OpcPackage::from_reader(Cursor::new(reopened.to_bytes().unwrap())).unwrap();
        assert!(saved.parts.contains_key("/WORD/DOCUMENT.XML"));
        assert!(saved.part_rels.contains_key("/WORD/DOCUMENT.XML"));
        let relationships = saved.get_part_rels("/word/document.xml").unwrap();
        assert!(relationships.get_by_id("rId40").is_some());
        assert_eq!(
            relationships.get_by_type(rel_types::HYPERLINK).unwrap().id,
            "rId41"
        );
    }

    #[test]
    fn custom_xml_identity_decoys_do_not_block_open() {
        let mut document = Document::new();
        let bytes = document.to_bytes().unwrap();
        let mut package = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
        package.set_part(
            "/customXml/item1.xml",
            br#"<x:root xmlns:x="urn:custom" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:bookmarkStart w:id="0"/><w:bookmarkStart w:id="0"/><wp:docPr id="bad">"#.to_vec(),
        );
        let mut output = Cursor::new(Vec::new());
        package.write_to(&mut output).unwrap();
        Document::from_bytes(&output.into_inner()).expect("opaque custom XML is not a story");
    }

    #[test]
    fn inserted_numbering_is_registered_before_later_authoring() {
        let mut source = Document::new();
        let source_id = source.add_list_definition(&[ListLevel::decimal()]);
        source.add_paragraph("source").set_numbering(source_id, 0);
        let mut destination = Document::new();
        destination.insert_document(0, &source);
        let inserted_id = destination
            .numbering
            .as_ref()
            .unwrap()
            .nums
            .first()
            .unwrap()
            .num_id;
        let authored_id = destination.add_list_definition(&[ListLevel::bullet()]);
        assert_ne!(inserted_id, authored_id);
    }

    #[test]
    fn unsupported_names_come_from_the_accepted_parser_event() {
        assert_eq!(
            raw_element_names(b"<?producer keep?><x:item xmlns:x=\"urn:test\"/>")
                .as_ref()
                .map(|(qualified, local)| (qualified.as_str(), local.as_str())),
            Some(("x:item", "item")),
        );
    }

    #[test]
    fn namespace_declaration_decoding_keeps_xml_error_classification() {
        let valid = format!(
            "<w:document xmlns:w=\"{WORD_NAMESPACE}\" xmlns:x=\"urn:plain\" \
             xmlns:y=\"urn:a&amp;b\"><w:body/></w:document>"
        );
        let scopes = document_namespace_scopes(valid.as_bytes()).unwrap();
        assert!(
            scopes
                .root_declarations
                .contains(&("xmlns:x".to_owned(), "urn:plain".to_owned()))
        );
        assert!(
            scopes
                .root_declarations
                .contains(&("xmlns:y".to_owned(), "urn:a&b".to_owned()))
        );

        let malformed = format!(
            "<w:document xmlns:w=\"{WORD_NAMESPACE}\" \
             xmlns:x=\"urn:&bad;\"><w:body/></w:document>"
        );
        assert!(matches!(
            document_namespace_scopes(malformed.as_bytes()),
            Err(Error::Oxml(_))
        ));
    }

    fn minimal_chart_workbook() -> Workbook {
        Workbook::new(
            "Sheet1",
            vec![
                Column::Text {
                    header: "Category".to_owned(),
                    values: vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
                },
                Column::Number {
                    header: "Revenue".to_owned(),
                    values: vec![12.5, 19.0, 14.25],
                    number_format: None,
                },
            ],
        )
        .expect("valid chart workbook")
    }

    fn minimal_typed_chart() -> CT_ChartSpace {
        let values = NumericData::new(
            "Sheet1!$B$2:$B$4".to_owned(),
            "General".to_owned(),
            vec![12.5, 19.0, 14.25],
        )
        .expect("valid values");
        let mut series = Series::new(0, 0, values);
        series.name = Some(
            StringRef::new("Sheet1!$B$1".to_owned(), vec!["Revenue".to_owned()])
                .expect("valid series name"),
        );
        series.categories = Some(AxisData::String(
            StringRef::new(
                "Sheet1!$A$2:$A$4".to_owned(),
                vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
            )
            .expect("valid categories"),
        ));
        let category_axis = AxisId::new(48_650_112).expect("valid category axis ID");
        let value_axis = AxisId::new(48_672_768).expect("valid value axis ID");
        let plot = Plot::bar(
            BarDirection::Column,
            BarGrouping::Clustered,
            vec![series],
            [category_axis, value_axis],
        )
        .expect("valid bar plot");
        let axes = vec![
            Axis::new(
                AxisKind::Category,
                category_axis,
                AxisPosition::Bottom,
                value_axis,
            ),
            Axis::new(
                AxisKind::Value,
                value_axis,
                AxisPosition::Left,
                category_axis,
            ),
        ];
        let mut chart = CT_ChartSpace::from_xml(
            format!(
                r#"<c:chartSpace xmlns:c="{}" xmlns:a="{}" xmlns:r="{}"><c:chart><c:plotArea/></c:chart></c:chartSpace>"#,
                oxml_chart::C_NS,
                oxml_chart::A_NS,
                oxml_chart::R_NS,
            )
            .as_bytes(),
        )
        .expect("valid chart shell");
        chart.chart.auto_title_deleted = true;
        chart.chart.plot_area = CT_PlotArea::new(vec![plot], axes).expect("valid plot area");
        chart
    }

    #[cfg(feature = "agile-encryption")]
    #[test]
    fn native_encrypted_save_round_trips_without_live_package_mutation() {
        let mut document = Document::new();
        document.add_paragraph("F-170 encrypted round trip");
        let original_parts = document.package.parts.clone();
        let original_relationships = document.package.package_rels.to_xml().unwrap();

        let error = document.to_encrypted_bytes("").unwrap_err();
        assert!(matches!(
            error,
            Error::Opc(oxml_opc::OpcError::InvalidPassword)
        ));
        assert_eq!(document.package.parts, original_parts);
        assert_eq!(
            document.package.package_rels.to_xml().unwrap(),
            original_relationships
        );

        let encrypted = document.to_encrypted_bytes("rdocx-f170").unwrap();
        let reopened = Document::from_encrypted_bytes(&encrypted, "rdocx-f170").unwrap();
        assert_eq!(reopened.text(), "F-170 encrypted round trip\n");
        assert_eq!(document.package.parts, original_parts);
        assert_eq!(
            document.package.package_rels.to_xml().unwrap(),
            original_relationships
        );

        let path = std::env::temp_dir().join(format!(
            "rdocx-f170-atomic-save-{}.docx",
            std::process::id()
        ));
        fs::write(&path, b"existing destination").unwrap();
        assert!(document.save_encrypted(&path, "").is_err());
        assert_eq!(fs::read(&path).unwrap(), b"existing destination");
        document.save_encrypted(&path, "rdocx-f170").unwrap();
        let saved = fs::read(&path).unwrap();
        assert_eq!(
            Document::from_encrypted_bytes(&saved, "rdocx-f170")
                .unwrap()
                .text(),
            "F-170 encrypted round trip\n"
        );
        fs::remove_file(path).unwrap();
    }

    #[cfg(feature = "agile-encryption")]
    #[test]
    #[ignore = "requires pinned Microsoft Word 16.104 and human password evidence"]
    fn word_opens_the_written_agile_document() {
        let plist = "/Applications/Microsoft Word.app/Contents/Info.plist";
        let version = Command::new("plutil")
            .args(["-extract", "CFBundleShortVersionString", "raw", plist])
            .output()
            .unwrap();
        assert!(version.status.success());
        assert_eq!(
            String::from_utf8_lossy(&version.stdout).trim(),
            WORD_VERSION
        );
        let build = Command::new("plutil")
            .args(["-extract", "CFBundleVersion", "raw", plist])
            .output()
            .unwrap();
        assert!(build.status.success());
        assert_eq!(String::from_utf8_lossy(&build.stdout).trim(), WORD_BUILD);

        let mut document = Document::new();
        document.add_paragraph("F-170 Microsoft Word encryption oracle");
        let path = Path::new("/private/tmp/F-170-word-16.104-encrypted.docx");
        document.save_encrypted(path, "rdocx-f170").unwrap();
        assert_eq!(
            std::env::var("RDOCX_F170_WORD_CORRECT_PASSWORD").as_deref(),
            Ok("opened"),
            "open {} in Word {WORD_VERSION} build {WORD_BUILD} with password rdocx-f170, then rerun with RDOCX_F170_WORD_CORRECT_PASSWORD=opened",
            path.display()
        );
        assert_eq!(
            std::env::var("RDOCX_F170_WORD_WRONG_PASSWORD").as_deref(),
            Ok("rejected"),
            "open {} with a wrong password, then rerun with RDOCX_F170_WORD_WRONG_PASSWORD=rejected",
            path.display()
        );
    }

    fn document_with_minimal_chart() -> Document {
        let mut document =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        document
            .add_chart_package(
                ChartPackageSource::Typed {
                    chart: &minimal_typed_chart(),
                    workbook: &minimal_chart_workbook(),
                },
                Length::inches(5.0),
                Length::inches(3.0),
            )
            .expect("assemble chart package");
        document
    }

    fn reset_layout_invocations() {
        LAYOUT_INVOCATIONS.set(0);
    }

    fn layout_invocations() -> usize {
        LAYOUT_INVOCATIONS.get()
    }

    fn caller_only_font() -> (&'static str, Vec<u8>) {
        const FAMILY: &str = "Callira";
        const SOURCE: &[u8] =
            include_bytes!("../../oxml-layout/fonts/Carlito-Regular.ttf").as_slice();

        fn replace_all_same_length(data: &mut [u8], from: &[u8], to: &[u8]) -> usize {
            assert_eq!(from.len(), to.len());
            let mut replaced = 0;
            let mut offset = 0;
            while let Some(index) = data[offset..]
                .windows(from.len())
                .position(|window| window == from)
            {
                let start = offset + index;
                data[start..start + from.len()].copy_from_slice(to);
                offset = start + from.len();
                replaced += 1;
            }
            replaced
        }

        let mut bytes = SOURCE.to_vec();
        let ascii = replace_all_same_length(&mut bytes, b"Carlito", FAMILY.as_bytes());
        let source_utf16 = "Carlito"
            .encode_utf16()
            .flat_map(u16::to_be_bytes)
            .collect::<Vec<_>>();
        let family_utf16 = FAMILY
            .encode_utf16()
            .flat_map(u16::to_be_bytes)
            .collect::<Vec<_>>();
        let utf16 = replace_all_same_length(&mut bytes, &source_utf16, &family_utf16);
        assert!(ascii > 0 && utf16 > 0, "font family records were renamed");
        assert!(
            oxml_layout::bundled_fonts::bundled_font_data()
                .iter()
                .all(|(family, bundled)| *family != FAMILY && *bundled != bytes),
            "the caller-only font must not match bundled family names or bytes"
        );
        (FAMILY, bytes)
    }

    #[test]
    fn body_items_preserve_paragraph_table_control_and_raw_order() {
        let mut doc = Document::new();
        doc.add_paragraph("first");
        doc.add_table(1, 1);
        doc.document
            .body
            .content
            .push(BodyContent::RawXml(b"<w:custom/>".to_vec()));

        let mut reader = quick_xml::Reader::from_reader(
            br#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:sdtContent></w:sdt>"#
                .as_slice(),
        );
        let mut buffer = Vec::new();
        let control = match reader.read_event_into(&mut buffer).unwrap() {
            Event::Start(start) => CT_Sdt::from_xml(&mut reader, &start).unwrap(),
            event => panic!("expected content control start, got {event:?}"),
        };
        doc.document
            .body
            .content
            .push(BodyContent::ContentControl(control));
        doc.add_paragraph("last");

        let items = doc
            .body_items()
            .map(|item| match item {
                BodyItemRef::Paragraph(paragraph) => format!("paragraph:{}", paragraph.text()),
                BodyItemRef::Table(_) => "table".to_owned(),
                BodyItemRef::ContentControl(control) => format!("control:{}", control.text()),
                BodyItemRef::UnsupportedXml(raw) => {
                    format!("raw:{}", std::str::from_utf8(raw).unwrap())
                }
            })
            .collect::<Vec<_>>();

        assert_eq!(
            items,
            [
                "paragraph:first",
                "table",
                "raw:<w:custom/>",
                "control:inside",
                "paragraph:last",
            ]
        );
    }

    #[test]
    fn reader_reports_background_and_section_completeness_facts() {
        let mut doc = Document::new();
        assert!(!doc.has_document_background());
        assert!(doc.has_section_layout_formatting());
        assert!(!doc.has_unmodeled_section_properties());

        doc.document.background_xml = Some(b"<w:background/>".to_vec());
        doc.document
            .body
            .sect_pr
            .as_mut()
            .expect("default section properties")
            .extra_xml
            .push(b"<w:printerSettings r:id=\"rId1\"/>".to_vec());

        assert!(doc.has_document_background());
        assert!(doc.has_unmodeled_section_properties());
    }

    #[test]
    fn rendering_all_pages_performs_one_layout() {
        let mut doc = Document::new();
        doc.add_paragraph("Page 1");
        for page in 2..=20 {
            doc.add_paragraph(&format!("Page {page}"))
                .page_break_before(true);
        }

        reset_layout_invocations();
        for page_index in 0..20 {
            assert!(
                doc.render_page_to_png_deterministic(page_index, 1.0)
                    .expect("deterministic layout should succeed")
                    .is_some(),
                "page {page_index} should exist"
            );
        }

        assert_eq!(layout_invocations(), 1);
    }

    #[test]
    fn tracked_option_layouts_are_not_cached() {
        let mut doc = Document::new();
        doc.add_paragraph("tracked view");
        let options = RenderOptions {
            revision_view: rdocx_layout::RevisionView::Tracked,
        };

        reset_layout_invocations();
        doc.render_page_to_png_deterministic_with_options(0, 1.0, options)
            .unwrap();
        doc.render_page_to_png_deterministic_with_options(0, 1.0, options)
            .unwrap();
        assert_eq!(layout_invocations(), 2);
    }

    #[test]
    fn document_mutation_invalidates_cached_layout() {
        let mut doc = Document::new();
        doc.add_paragraph("Before mutation");

        reset_layout_invocations();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 1);

        doc.add_paragraph("After mutation");
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 2);
    }

    #[test]
    fn field_update_batch_invalidates_cached_layout_once() {
        let mut doc = Document::new();
        let mut paragraph = CT_P::new();
        paragraph.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field(Field::new("PAGE", "4"))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        doc.document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        reset_layout_invocations();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 1);

        assert_eq!(
            doc.update_fields(&FieldEvaluationContext::default())
                .unwrap(),
            1
        );
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 2);
    }

    #[test]
    fn mutable_accessor_invalidates_cached_layout() {
        let mut doc = Document::new();
        doc.add_paragraph("Before wrapper mutation");

        reset_layout_invocations();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 1);

        doc.paragraph_mut(0)
            .expect("paragraph should exist")
            .add_run(" changed");
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 2);

        let mut table = doc.add_table(1, 1);
        table
            .cell(0, 0)
            .expect("cell should exist")
            .set_text("table mutation");
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 3);
    }

    #[test]
    fn immutable_run_accessors_preserve_cached_layout() {
        let mut doc = Document::new();
        doc.add_paragraph("Before immutable access")
            .add_run(" remains cached");

        reset_layout_invocations();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 1);

        let paragraph = doc.paragraph(0).expect("paragraph should exist");
        assert_eq!(paragraph.run_count(), 2);
        assert_eq!(
            paragraph.run(1).expect("run should exist").text(),
            " remains cached"
        );
        assert!(paragraph.run(2).is_none());

        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 1);
    }

    #[test]
    fn font_modes_use_isolated_layout_caches() {
        let mut doc = Document::new();
        doc.add_paragraph("Font mode isolation");

        reset_layout_invocations();
        doc.render_page_to_png(0, 1.0).unwrap();
        doc.render_page_to_png(0, 1.0).unwrap();
        assert!(doc.layout_page(0).unwrap().is_some());
        assert!(doc.layout_page(usize::MAX).unwrap().is_none());
        assert_eq!(doc.render_all_pages(1.0).unwrap().len(), 1);
        assert!(!doc.to_pdf().unwrap().is_empty());
        assert_eq!(layout_invocations(), 1);

        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 2);

        doc.render_page_to_png(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 2);

        let (family, font_data) = oxml_layout::bundled_fonts::bundled_font_data()[0];
        doc.to_pdf_with_fonts(&[(family, font_data)]).unwrap();
        doc.to_pdf_with_fonts(&[(family, font_data)]).unwrap();
        assert_eq!(layout_invocations(), 4);
    }

    #[test]
    fn officemath_document_settings_reach_pdf_and_raster_backends() {
        fn has_direct_rule(element: &oxml_layout::PositionedElement) -> bool {
            matches!(element, oxml_layout::PositionedElement::Line { .. })
        }

        fn fraction_group_mut(
            elements: &mut [oxml_layout::PositionedElement],
        ) -> Option<&mut oxml_layout::GroupElement> {
            for element in elements {
                match element {
                    oxml_layout::PositionedElement::Group(group) => {
                        if group.children.iter().any(has_direct_rule) {
                            return Some(group);
                        }
                        if let Some(found) = fraction_group_mut(&mut group.children) {
                            return Some(found);
                        }
                    }
                    oxml_layout::PositionedElement::MarkedContent { children, .. } => {
                        if let Some(found) = fraction_group_mut(children) {
                            return Some(found);
                        }
                    }
                    _ => {}
                }
            }
            None
        }

        let word = rdocx_oxml::namespace::W_NS;
        let math = rdocx_oxml::namespace::M_NS;
        let document_xml = format!(
            r#"<w:document xmlns:w="{word}" xmlns:m="{math}"><w:body><w:p><w:r><w:t>before</w:t></w:r><m:oMath><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath><w:r><w:t>after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#
        );
        let settings_xml = format!(
            r#"<w:settings xmlns:w="{word}" xmlns:m="{math}"><m:mathPr><m:mathFont m:val="Caladea"/><m:jc m:val="right"/><m:preSp m:val="80"/><m:postSp m:val="120"/></m:mathPr></w:settings>"#
        );
        let mut document = Document::new();
        document.document =
            CT_Document::from_xml(document_xml.as_bytes()).expect("document parses");
        document.settings =
            Some(CT_Settings::from_xml(settings_xml.as_bytes()).expect("math settings parse"));

        let input = document.build_layout_input();
        assert_eq!(
            input
                .math_properties
                .as_ref()
                .and_then(|properties| properties.math_font.as_deref()),
            Some("Caladea")
        );
        let pdf = document
            .to_pdf_deterministic()
            .expect("OfficeMath PDF render");
        assert!(pdf.starts_with(b"%PDF-"));
        let png = document
            .render_page_to_png_deterministic(0, 150.0)
            .expect("OfficeMath raster render")
            .expect("OfficeMath page exists");
        assert!(png.starts_with(b"\x89PNG\r\n\x1a\n"));

        let mut mutated = document
            .cached_deterministic_layout()
            .expect("deterministic OfficeMath layout")
            .layout
            .clone();
        let page = Arc::make_mut(&mut mutated.pages[0]);
        let fraction = fraction_group_mut(&mut page.elements).expect("rendered fraction group");
        let original_y = fraction.transform.f;
        fraction.transform.f += 1.01;
        assert!((fraction.transform.f - original_y).abs() > 1.0);
        let mutated_png = oxml_pdf::render_page_to_png(&mutated, 0, 150.0)
            .expect("mutated OfficeMath page renders");
        assert_ne!(png, mutated_png, "1.01 point page mutation must be visible");
    }

    #[test]
    fn native_word_pdfa_method_selects_the_requested_profile() {
        let mut doc = Document::new();
        doc.add_paragraph("Archival Word document");

        for (profile, part) in [
            (oxml_pdf::PdfConformance::PdfA2b, "2"),
            (oxml_pdf::PdfConformance::PdfA3b, "3"),
        ] {
            let pdf = doc.to_pdfa_deterministic(profile).unwrap();
            assert!(
                String::from_utf8_lossy(&pdf)
                    .contains(&format!("<pdfaid:part>{part}</pdfaid:part>"))
            );
        }
    }

    #[test]
    fn full_layout_exposes_resolvable_font_data_and_reuses_the_cache() {
        let mut doc = Document::new();
        doc.add_paragraph("complete layout result");

        reset_layout_invocations();
        let first = doc.layout().expect("normal layout should succeed");
        let second = doc
            .layout_with_options(RenderOptions::default())
            .expect("accepted layout should succeed");
        assert!(Arc::ptr_eq(&first, &second));
        assert_eq!(layout_invocations(), 1);

        let mut sourced_runs = 0;
        for page in &first.layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| {
                if let oxml_layout::PositionedElement::Text(run) = element {
                    assert!(first.layout.fonts.iter().any(|font| font.id == run.font_id));
                    if let Some(source) = run.source {
                        assert!(first.source_node(source.node).is_some());
                        sourced_runs += 1;
                    }
                }
            });
        }
        assert!(sourced_runs > 0);

        assert!(
            !doc.to_pdf()
                .expect("PDF should use cached layout")
                .is_empty()
        );
        assert_eq!(layout_invocations(), 1);
    }

    #[test]
    fn layout_with_fonts_returns_the_caller_font_mapping_without_caching() {
        let mut doc = Document::new();
        let (family, bytes) = caller_only_font();
        doc.add_paragraph("")
            .add_run("caller font result")
            .font(family);

        reset_layout_invocations();
        let first = doc
            .layout_with_fonts(&[(family, &bytes)])
            .expect("caller-font layout should succeed");
        let second = doc
            .layout_with_fonts_and_options(&[(family, &bytes)], RenderOptions::default())
            .expect("caller-font options layout should succeed");
        assert_eq!(layout_invocations(), 2);

        for result in [&first, &second] {
            let mut positioned_runs = Vec::new();
            for page in &result.layout.pages {
                oxml_layout::walk(&page.elements, &mut |element, _| {
                    if let oxml_layout::PositionedElement::Text(run) = element
                        && run.source.is_some()
                    {
                        positioned_runs.push(run.clone());
                    }
                });
            }
            assert!(
                !positioned_runs.is_empty(),
                "caller-font text was positioned"
            );
            for run in positioned_runs {
                let font = result
                    .layout
                    .fonts
                    .iter()
                    .find(|font| font.id == run.font_id)
                    .expect("caller-font glyph run should resolve its font id");
                assert_eq!(font.family, family);
                assert_eq!(font.data.as_ref(), bytes.as_slice());
                let source = run
                    .source
                    .expect("caller-font text should retain source provenance");
                assert!(result.source_node(source.node).is_some());
            }
        }
    }

    #[test]
    fn caller_font_layout_does_not_fall_through_to_system_or_bundled_fonts() {
        let mut doc = Document::new();
        doc.add_paragraph("caller isolation requires an explicit font universe");
        assert!(doc.layout_with_fonts(&[]).is_err());
        assert!(
            doc.normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_none()
        );
    }

    #[test]
    fn bundled_fallback_completes_an_incomplete_caller_font_set() {
        let (caller_family, caller_bytes) = caller_only_font();
        let mut document = Document::new();
        document.add_paragraph("Calibri must resolve from bundled fonts");
        document
            .add_paragraph("")
            .add_run("caller face must win")
            .font(caller_family);

        assert!(
            document
                .layout_with_fonts(&[(caller_family, &caller_bytes)])
                .is_err(),
            "strict caller-only layout must not borrow bundled Calibri"
        );
        let result = document
            .layout_with_fonts_and_bundled_fallback(&[(caller_family, &caller_bytes)])
            .expect("bundled fallback completes the caller font set");

        let mut caller_runs = 0;
        let mut bundled_runs = 0;
        for page in &result.layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| {
                let oxml_layout::PositionedElement::Text(run) = element else {
                    return;
                };
                let font = result
                    .layout
                    .fonts
                    .iter()
                    .find(|font| font.id == run.font_id)
                    .expect("every glyph run resolves its font");
                if font.family == caller_family {
                    assert_eq!(font.data.as_ref(), caller_bytes.as_slice());
                    caller_runs += 1;
                } else if oxml_layout::bundled_fonts::bundled_font_data()
                    .iter()
                    .any(|(_, bytes)| *bytes == font.data.as_ref())
                {
                    bundled_runs += 1;
                }
            });
        }
        assert!(caller_runs > 0, "the matching caller face has priority");
        assert!(bundled_runs > 0, "the missing family uses bundled fallback");
        assert!(
            document
                .normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_none(),
            "bundled fallback stays isolated from the normal engine"
        );
    }

    #[test]
    fn bundled_fallback_engine_reuses_and_transfers_exact_context() {
        let (caller_family, caller_bytes) = caller_only_font();
        let fonts = [(caller_family, caller_bytes.as_slice())];
        let mut source = Document::new();
        for index in 0..140 {
            source.add_paragraph(&format!("paragraph {index:03} stable line"));
        }
        let initial = source
            .layout_with_fonts_and_bundled_fallback(&fonts)
            .expect("initial bundled-fallback layout");
        source
            .paragraph_mut(70)
            .expect("middle paragraph")
            .add_run(" changed");
        let warm = source
            .layout_with_fonts_and_bundled_fallback(&fonts)
            .expect("warm bundled-fallback layout");
        assert!(
            warm.layout
                .pages
                .iter()
                .zip(&initial.layout.pages)
                .any(|(current, previous)| Arc::ptr_eq(current, previous)),
            "a bounded edit reuses at least one retained page"
        );

        let mut receiver = source.clone_for_staging();
        assert!(receiver.transfer_reusable_bundled_fallback_layout_from(&mut source, &fonts));
        assert!(
            source
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_none()
        );
        assert!(
            receiver
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );

        let mut font_source = receiver.clone_for_staging();
        let mut font_receiver = receiver.clone_for_staging();
        font_source
            .layout_with_fonts_and_bundled_fallback(&fonts)
            .expect("prime font source");
        font_receiver
            .layout_with_fonts_and_bundled_fallback(&fonts)
            .expect("prime font receiver");
        let mut changed_bytes = caller_bytes.clone();
        let last = changed_bytes.last_mut().expect("caller font has bytes");
        *last ^= 1;
        assert!(
            !font_receiver.transfer_reusable_bundled_fallback_layout_from(
                &mut font_source,
                &[(caller_family, &changed_bytes)],
            )
        );
        assert!(
            font_source
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );
        assert!(
            font_receiver
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );

        let mut context_source = receiver.clone_for_staging();
        let mut context_receiver = receiver.clone_for_staging();
        context_receiver.set_header("different retained context");
        context_source
            .layout_with_fonts_and_bundled_fallback(&fonts)
            .expect("prime context source");
        context_receiver
            .layout_with_fonts_and_bundled_fallback(&fonts)
            .expect("prime context receiver");
        assert!(
            !context_receiver
                .transfer_reusable_bundled_fallback_layout_from(&mut context_source, &fonts)
        );
        assert!(
            context_source
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );
        assert!(
            context_receiver
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );
    }

    #[test]
    fn bundled_fallback_warm_layout_equals_fresh_layout() {
        let (caller_family, caller_bytes) = caller_only_font();
        let fonts = [(caller_family, caller_bytes.as_slice())];
        let mut warm_document = Document::new();
        for index in 0..140 {
            warm_document.add_paragraph(&format!("paragraph {index:03} stable line"));
        }
        warm_document
            .layout_with_fonts_and_bundled_fallback(&fonts)
            .expect("prime warm engine");
        warm_document
            .paragraph_mut(70)
            .expect("middle paragraph")
            .add_run(" changed");
        let fresh_document = warm_document.clone_for_staging();

        let warm = warm_document
            .layout_with_fonts_and_bundled_fallback_and_options(&fonts, RenderOptions::default())
            .expect("warm layout");
        let fresh = fresh_document
            .layout_with_fonts_and_bundled_fallback_and_options(&fonts, RenderOptions::default())
            .expect("fresh layout");

        assert_eq!(warm.revision_view, fresh.revision_view);
        assert_eq!(
            format!("{:?}", warm.layout.pages),
            format!("{:?}", fresh.layout.pages)
        );
        assert_eq!(
            format!("{:?}", warm.layout.fonts),
            format!("{:?}", fresh.layout.fonts)
        );
        assert_eq!(
            format!("{:?}", warm.layout.diagnostics),
            format!("{:?}", fresh.layout.diagnostics)
        );
        assert_eq!(
            format!("{:?}", warm.layout.outlines),
            format!("{:?}", fresh.layout.outlines)
        );
        assert_eq!(format!("{warm:?}"), format!("{fresh:?}"));
        assert_eq!(
            oxml_pdf::render_to_pdf(&warm.layout),
            oxml_pdf::render_to_pdf(&fresh.layout)
        );
    }

    #[test]
    fn changed_alias_context_cannot_reuse_stale_layout_work() {
        let (caller_family, caller_bytes) = caller_only_font();
        let fonts = [(caller_family, caller_bytes.as_slice())];
        let first_aliases = [("Document Sans", caller_family)];
        let changed_aliases = [("Document Sans", "Carlito")];
        let mut source = Document::new();
        source
            .add_paragraph("")
            .add_run("alias-sensitive paragraph")
            .font("Document Sans");

        let first = source
            .layout_with_fonts_aliases_and_bundled_fallback(&fonts, &first_aliases)
            .expect("prime alias-aware engine");
        let first_repeat = source
            .layout_with_fonts_aliases_and_bundled_fallback(&fonts, &first_aliases)
            .expect("equal alias context reuses safely");
        assert!(
            first
                .layout
                .pages
                .iter()
                .zip(&first_repeat.layout.pages)
                .all(|(left, right)| Arc::ptr_eq(left, right)),
            "equal aliases reuse completed page work"
        );
        let changed = source
            .layout_with_fonts_aliases_and_bundled_fallback(&fonts, &changed_aliases)
            .expect("changed alias context relayouts");
        assert_ne!(
            format!("{:?}", first_repeat.layout.fonts),
            format!("{:?}", changed.layout.fonts)
        );

        let mut receiver = source.clone_for_staging();
        receiver
            .layout_with_fonts_aliases_and_bundled_fallback(&fonts, &changed_aliases)
            .expect("prime receiver engine");
        assert!(
            !receiver.transfer_reusable_bundled_fallback_layout_from_with_aliases(
                &mut source,
                &fonts,
                &first_aliases,
            )
        );
        assert!(
            source
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );
        assert!(
            receiver
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );

        let mut compatible_source = receiver.clone_for_staging();
        compatible_source
            .layout_with_fonts_aliases_and_bundled_fallback(&fonts, &changed_aliases)
            .expect("prime compatible source engine");
        let mut compatible_receiver = compatible_source.clone_for_staging();
        assert!(
            compatible_receiver.transfer_reusable_bundled_fallback_layout_from_with_aliases(
                &mut compatible_source,
                &fonts,
                &changed_aliases,
            )
        );
        assert!(
            compatible_source
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_none()
        );
        assert!(
            compatible_receiver
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );
    }

    #[test]
    fn caller_alias_warm_layout_equals_cold_layout() {
        let (caller_family, caller_bytes) = caller_only_font();
        let fonts = [(caller_family, caller_bytes.as_slice())];
        let aliases = [
            ("Document Sans", caller_family),
            ("Document UI", caller_family),
        ];
        let mut warm_document = Document::new();
        for index in 0..140 {
            warm_document
                .add_paragraph("")
                .add_run(&format!("alias paragraph {index:03}"))
                .font(if index % 2 == 0 {
                    "Document Sans"
                } else {
                    "Document UI"
                });
        }
        warm_document
            .layout_with_fonts_aliases_and_bundled_fallback(&fonts, &aliases)
            .expect("prime alias-aware engine");
        warm_document
            .paragraph_mut(70)
            .expect("middle paragraph")
            .add_run(" changed");
        let cold_document = warm_document.clone_for_staging();

        let warm = warm_document
            .layout_with_fonts_aliases_and_bundled_fallback_and_options(
                &fonts,
                &aliases,
                RenderOptions::default(),
            )
            .expect("warm alias-aware layout");
        let cold = cold_document
            .layout_with_fonts_aliases_and_bundled_fallback_and_options(
                &fonts,
                &aliases,
                RenderOptions::default(),
            )
            .expect("cold alias-aware layout");

        assert_eq!(format!("{warm:?}"), format!("{cold:?}"));
        assert_eq!(
            oxml_pdf::render_to_pdf(&warm.layout),
            oxml_pdf::render_to_pdf(&cold.layout)
        );
        assert_eq!(
            oxml_pdf::render_page_to_png(&warm.layout, 0, 72.0),
            oxml_pdf::render_page_to_png(&cold.layout, 0, 72.0)
        );
    }

    #[test]
    fn substituted_page_reuse_keeps_pdf_and_raster_backends_identical() {
        let mut warm_document = Document::new();
        let mut fields = CT_P::new();
        for instruction in ["PAGE", "NUMPAGES"] {
            let mut run = CT_R::new("");
            run.content = vec![RunContent::Field(Field::new(instruction, "cached"))];
            fields.runs.push(run);
        }
        warm_document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(fields));
        for index in 0..140 {
            warm_document.add_paragraph(&format!("paragraph {index:03} stable line"));
        }
        let fresh_document = warm_document.clone_for_staging();

        warm_document
            .layout_with_fonts_and_bundled_fallback(&[])
            .expect("prime substituted pages");
        let warm = warm_document
            .layout_with_fonts_and_bundled_fallback(&[])
            .expect("warm substituted pages");
        let fresh = fresh_document
            .layout_with_fonts_and_bundled_fallback(&[])
            .expect("cold substituted pages");

        assert_eq!(format!("{warm:?}"), format!("{fresh:?}"));
        assert_eq!(
            oxml_pdf::render_to_pdf(&warm.layout),
            oxml_pdf::render_to_pdf(&fresh.layout)
        );
        assert_eq!(
            oxml_pdf::render_all_pages(&warm.layout, 72.0),
            oxml_pdf::render_all_pages(&fresh.layout, 72.0)
        );
    }

    #[test]
    fn cached_header_footer_warm_equals_cold() {
        let mut warm_document = Document::new();
        warm_document.set_header("default cached header");
        warm_document.set_footer("default cached footer");
        warm_document.set_first_page_header("first cached header");
        warm_document.set_first_page_footer("first cached footer");
        warm_document
            .set_text_watermark("DRAFT")
            .expect("watermark is valid");
        for index in 0..140 {
            warm_document.add_paragraph(&format!("paragraph {index:03} stable line"));
        }
        warm_document
            .layout_with_fonts_and_bundled_fallback(&[])
            .expect("prime deterministic header/footer engine");
        warm_document
            .paragraph_mut(70)
            .expect("middle paragraph")
            .add_run(" changed");
        let fresh_document = warm_document.clone_for_staging();

        let warm = warm_document
            .layout_with_fonts_and_bundled_fallback(&[])
            .expect("warm deterministic layout");
        let fresh = fresh_document
            .layout_with_fonts_and_bundled_fallback(&[])
            .expect("cold deterministic layout");

        assert_eq!(format!("{warm:?}"), format!("{fresh:?}"));
        assert_eq!(
            oxml_pdf::render_to_pdf(&warm.layout),
            oxml_pdf::render_to_pdf(&fresh.layout)
        );
    }

    #[test]
    fn staged_mutations_preserve_valid_bundled_fallback_work() {
        let (caller_family, caller_bytes) = caller_only_font();
        let fonts = [(caller_family, caller_bytes.as_slice())];
        let mut document = Document::new();
        document.add_paragraph("staged mutation retains reusable work");
        document
            .layout_with_fonts_and_bundled_fallback(&fonts)
            .expect("prime bundled-fallback engine");
        document
            .set_text_watermark("DRAFT")
            .expect("successful staged mutation");
        assert!(
            document
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some(),
            "successful staging retains the private engine"
        );
        document
            .layout_with_fonts_and_bundled_fallback(&fonts)
            .expect("engine remains usable after staging");

        let header_part = document
            .package
            .get_part_rels(&document.doc_part_name)
            .and_then(|relationships| relationships.get_by_type(rel_types::HEADER))
            .map(|relationship| {
                OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target)
            })
            .expect("watermark header relationship");
        document.package.set_part(&header_part, b"<".to_vec());
        let before = document.package.get_part(&header_part).unwrap().to_vec();
        assert!(document.set_text_watermark("FAIL").is_err());
        assert_eq!(
            document.package.get_part(&header_part),
            Some(before.as_slice())
        );
        assert!(
            document
                .bundled_fallback_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some(),
            "failed staging preserves the published engine"
        );

        let mut poisoned = Document::new();
        poisoned.add_paragraph("poison recovery");
        poisoned
            .layout_with_fonts_and_bundled_fallback(&fonts)
            .expect("prime engine before poison");
        let poisoned = Arc::new(poisoned);
        let poison_owner = Arc::clone(&poisoned);
        assert!(
            std::thread::spawn(move || {
                let _engine = poison_owner.bundled_fallback_layout_engine.lock().unwrap();
                panic!("poison bundled fallback engine lock");
            })
            .join()
            .is_err()
        );
        poisoned
            .layout_with_fonts_and_bundled_fallback(&fonts)
            .expect("bundled-fallback layout recovers from poison");
    }

    #[test]
    fn layout_options_keep_tracked_and_accepted_cache_ownership_separate() {
        let mut doc = Document::new();
        doc.add_paragraph("revision cache separation");
        let tracked = RenderOptions {
            revision_view: rdocx_layout::RevisionView::Tracked,
        };

        reset_layout_invocations();
        let accepted_first = doc.layout().expect("accepted layout should succeed");
        assert_eq!(layout_invocations(), 1);

        let tracked_first = doc
            .layout_with_options(tracked)
            .expect("tracked layout should succeed");
        let tracked_second = doc
            .layout_with_options(tracked)
            .expect("second tracked layout should succeed");
        assert!(!Arc::ptr_eq(&tracked_first, &tracked_second));
        assert_eq!(layout_invocations(), 3);
        assert!(
            doc.normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );

        let accepted_second = doc.layout().expect("accepted cache should succeed");
        assert!(Arc::ptr_eq(&accepted_first, &accepted_second));
        assert_eq!(layout_invocations(), 3);
        assert_eq!(
            tracked_first.revision_view,
            rdocx_layout::RevisionView::Tracked
        );
        assert_eq!(
            accepted_first.revision_view,
            rdocx_layout::RevisionView::Accepted
        );
    }

    #[test]
    fn tracked_normal_layouts_retain_the_document_engine_without_caching_results() {
        let mut doc = Document::new();
        doc.add_paragraph("tracked engine reuse");
        let tracked = RenderOptions {
            revision_view: rdocx_layout::RevisionView::Tracked,
        };

        let first = doc
            .layout_with_options(tracked)
            .expect("first tracked layout");
        let second = doc
            .layout_with_options(tracked)
            .expect("second tracked layout");
        assert!(!Arc::ptr_eq(&first, &second));
        assert!(
            doc.normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );
        assert!(
            doc.layout_cache
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_none()
        );
    }

    #[test]
    fn document_remains_send_and_sync() {
        fn assert_send_and_sync<T: Send + Sync>() {}
        assert_send_and_sync::<Document>();
    }

    #[test]
    fn relayout_caches_are_bounded_and_recover_from_poison() {
        let mut document = Document::new();
        document.add_paragraph("layout after engine lock poison");
        let document = Arc::new(document);
        let poison = Arc::clone(&document);
        assert!(
            std::thread::spawn(move || {
                let _engine = poison.normal_layout_engine.lock().unwrap();
                panic!("poison normal layout engine lock for recovery coverage");
            })
            .join()
            .is_err()
        );

        let first = document
            .layout()
            .expect("layout recovers from poisoned engine lock");
        let second = document.layout().expect("recovered layout remains cached");
        assert!(Arc::ptr_eq(&first, &second));
    }

    #[test]
    fn all_backends_accept_shared_layout_payloads_without_output_change() {
        let mut document = Document::new();
        document.add_paragraph("Shared layout").style("Heading1");
        document.add_paragraph("Backend and provenance coverage");

        let original = document.layout().expect("normal layout succeeds");
        let shared = original.layout.clone();
        assert!(
            original
                .layout
                .pages
                .iter()
                .zip(&shared.pages)
                .all(|(left, right)| Arc::ptr_eq(left, right))
        );
        assert!(
            original
                .layout
                .fonts
                .iter()
                .zip(&shared.fonts)
                .all(|(left, right)| Arc::ptr_eq(&left.data, &right.data))
        );
        assert_eq!(
            oxml_pdf::render_to_pdf(&original.layout),
            oxml_pdf::render_to_pdf(&shared)
        );
        assert_eq!(
            oxml_pdf::render_page_to_png(&original.layout, 0, 72.0),
            oxml_pdf::render_page_to_png(&shared, 0, 72.0)
        );
        assert_eq!(
            format!(
                "{:?}",
                document
                    .layout_page(0)
                    .expect("page access")
                    .expect("first page")
            ),
            format!("{:?}", shared.pages[0])
        );
        assert_eq!(original.layout.diagnostics, shared.diagnostics);
        assert_eq!(
            format!("{:?}", original.layout.outlines),
            format!("{:?}", shared.outlines)
        );

        let (family, bytes) = caller_only_font();
        let mut caller_document = Document::new();
        caller_document
            .add_paragraph("")
            .add_run("caller-owned font")
            .font(family);
        let caller = caller_document
            .layout_with_fonts(&[(family, &bytes)])
            .expect("caller-font layout succeeds");
        let caller_shared = caller.layout.clone();
        assert!(
            caller
                .layout
                .fonts
                .iter()
                .zip(&caller_shared.fonts)
                .all(|(left, right)| Arc::ptr_eq(&left.data, &right.data))
        );
        let source = caller.layout.pages.iter().find_map(|page| {
            let mut source = None;
            oxml_layout::walk(&page.elements, &mut |element, _| {
                if let oxml_layout::PositionedElement::Text(run) = element {
                    source = source.or(run.source);
                }
            });
            source
        });
        assert!(source.is_some_and(|span| caller.source_node(span.node).is_some()));
    }

    #[test]
    fn compatible_document_transfer_reuses_normal_layout_work() {
        let mut source = Document::new();
        source.add_paragraph("unchanged paragraph");
        source.add_paragraph("old second paragraph");
        let source_result = source.layout().expect("prime source engine");

        let mut receiver = source.clone_for_staging();
        let receiver_result = receiver.layout().expect("prime receiver engine");
        assert!(receiver.transfer_reusable_layout_from(&mut source));
        assert!(
            source
                .normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_none()
        );
        assert!(
            receiver
                .normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );
        assert!(Arc::ptr_eq(
            &source_result,
            &source.layout().expect("source result cache remains")
        ));
        assert!(Arc::ptr_eq(
            &receiver_result,
            &receiver.layout().expect("receiver result cache remains")
        ));

        receiver
            .paragraph_mut(1)
            .expect("second paragraph")
            .add_run(" changed");
        assert!(receiver.layout().is_ok());
    }

    #[test]
    fn incompatible_or_failed_transfer_preserves_both_engines() {
        let mut source = Document::new();
        source.add_paragraph("source paragraph");
        let mut receiver = source.clone_for_staging();
        source.layout().expect("prime source engine");
        receiver.layout().expect("prime receiver engine");

        receiver.styles = CT_Styles::new();
        receiver.invalidate_layout();
        assert!(!receiver.transfer_reusable_layout_from(&mut source));
        assert!(
            source
                .normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );
        assert!(
            receiver
                .normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );

        let poisoned = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            let _engine = source.normal_layout_engine.lock().unwrap();
            panic!("poison source engine lock");
        }));
        assert!(poisoned.is_err());
        let mut compatible = source.clone_for_staging();
        assert!(compatible.transfer_reusable_layout_from(&mut source));
        assert!(compatible.layout().is_ok());
    }

    #[test]
    fn transferred_warm_layout_equals_fresh_layout() {
        let mut source = Document::new();
        source.add_paragraph("unchanged paragraph");
        source.add_paragraph("old second paragraph");
        source.layout().expect("prime source engine");

        let mut warm = source.clone_for_staging();
        warm.paragraph_mut(1)
            .expect("second paragraph")
            .add_run(" changed");
        let fresh = warm.clone_for_staging();
        assert!(warm.transfer_reusable_layout_from(&mut source));

        let warm_result = warm.layout().expect("transferred warm layout");
        let fresh_result = fresh.layout().expect("fresh cold layout");
        assert_eq!(format!("{warm_result:?}"), format!("{fresh_result:?}"));
        assert_eq!(
            oxml_pdf::render_to_pdf(&warm_result.layout),
            oxml_pdf::render_to_pdf(&fresh_result.layout)
        );
    }

    #[test]
    fn word_chart_part_and_workbook_round_trip() {
        let workbook_bytes = minimal_chart_workbook()
            .to_xlsx_bytes()
            .expect("serialize expected workbook");
        let mut document = document_with_minimal_chart();
        let bytes = document.to_bytes().expect("save document");
        let reopened = Document::from_bytes(&bytes).expect("reopen document");
        let chart_part = "/word/charts/chart1.xml";
        let workbook_part = "/word/embeddings/Workbook1.xlsx";
        let chart_xml = reopened.package.get_part(chart_part).expect("chart part");
        assert!(CT_ChartSpace::from_xml(chart_xml).is_ok());
        assert_eq!(
            reopened.package.get_part(workbook_part),
            Some(workbook_bytes.as_slice())
        );
        assert_eq!(
            reopened.package.content_types.content_type_for(chart_part),
            Some(content_types::CHART)
        );
        assert_eq!(
            reopened
                .package
                .content_types
                .content_type_for(workbook_part),
            Some(content_types::EMBEDDED_WORKBOOK)
        );
        let document_relationship = reopened
            .package
            .get_part_rels(&reopened.doc_part_name)
            .and_then(|relationships| relationships.get_by_type(rel_types::CHART))
            .expect("document to chart relationship");
        assert_eq!(
            OpcPackage::resolve_rel_target(&reopened.doc_part_name, &document_relationship.target,),
            chart_part
        );
        let workbook_relationship = reopened
            .package
            .get_part_rels(chart_part)
            .and_then(|relationships| relationships.get_by_type(rel_types::PACKAGE))
            .expect("chart to workbook relationship");
        assert_eq!(
            OpcPackage::resolve_rel_target(chart_part, &workbook_relationship.target),
            workbook_part
        );
        assert!(
            std::str::from_utf8(chart_xml)
                .expect("chart XML is utf8")
                .contains(&format!(
                    r#"<c:externalData r:id="{}">"#,
                    workbook_relationship.id
                ))
        );
        let document_xml = std::str::from_utf8(
            reopened
                .package
                .get_part(&reopened.doc_part_name)
                .expect("document part"),
        )
        .expect("document XML is utf8");
        assert!(document_xml.contains(&format!(r#"r:id="{}""#, document_relationship.id)));
    }

    #[test]
    fn word_chart_parts_allocate_after_sparse_suffixes() {
        let mut document = Document::new();
        document
            .package
            .set_part("/word/charts/chart3.xml", b"occupied".to_vec());
        document
            .package
            .set_part("/word/embeddings/Workbook7.xlsx", b"occupied".to_vec());
        document
            .add_chart_package(
                ChartPackageSource::Typed {
                    chart: &minimal_typed_chart(),
                    workbook: &minimal_chart_workbook(),
                },
                Length::inches(5.0),
                Length::inches(3.0),
            )
            .expect("assemble after sparse suffixes");

        assert!(
            document
                .package
                .get_part("/word/charts/chart4.xml")
                .is_some()
        );
        assert!(
            document
                .package
                .get_part("/word/embeddings/Workbook8.xlsx")
                .is_some()
        );
        assert_eq!(
            document.package.get_part("/word/charts/chart3.xml"),
            Some(b"occupied".as_slice())
        );
        assert_eq!(
            document.package.get_part("/word/embeddings/Workbook7.xlsx"),
            Some(b"occupied".as_slice())
        );
    }

    #[test]
    fn invalid_chart_package_assembly_is_atomic() {
        let mut document = Document::new();
        let before = document
            .to_bytes()
            .expect("serialize before failed mutation");
        let chart = CT_ChartSpace::from_xml(
            format!(
                r#"<c:chartSpace xmlns:c="{}" xmlns:a="{}" xmlns:r="{}" xmlns:q="{}"><c:chart><c:plotArea/></c:chart><q:externalData r:id="rId99"/></c:chartSpace>"#,
                oxml_chart::C_NS,
                oxml_chart::A_NS,
                oxml_chart::R_NS,
                oxml_chart::C_NS,
            )
            .as_bytes(),
        )
        .expect("valid chart with an occupied workbook relationship");
        let error = document
            .add_chart_package(
                ChartPackageSource::Typed {
                    chart: &chart,
                    workbook: &minimal_chart_workbook(),
                },
                Length::inches(5.0),
                Length::inches(3.0),
            )
            .expect_err("a second workbook relationship is invalid");
        assert!(error.to_string().contains("already contains"));
        assert_eq!(
            document
                .to_bytes()
                .expect("serialize after failed mutation"),
            before
        );
    }

    #[test]
    fn nested_external_data_lookalike_does_not_occupy_workbook_slot() {
        let chart = CT_ChartSpace::from_xml(
            format!(
                r#"<c:chartSpace xmlns:c="{}" xmlns:a="{}" xmlns:r="{}" xmlns:q="{}"><c:chart><c:plotArea/></c:chart><c:extLst><c:ext uri="urn:producer"><q:externalData r:id="rId99"/></c:ext></c:extLst></c:chartSpace>"#,
                oxml_chart::C_NS,
                oxml_chart::A_NS,
                oxml_chart::R_NS,
                oxml_chart::C_NS,
            )
            .as_bytes(),
        )
        .expect("valid chart with a nested producer lookalike");
        let xml = chart_with_workbook_relationship(&chart, "rId1")
            .expect("nested lookalike leaves the workbook slot available");
        let xml = std::str::from_utf8(&xml).expect("chart XML is utf8");
        assert!(xml.contains(r#"<q:externalData r:id="rId99"/>"#));
        assert!(xml.contains(r#"<c:externalData r:id="rId1">"#));
    }

    fn f158_chart_data(series_count: usize) -> ChartData {
        ChartData {
            categories: vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
            series: vec![
                ("Revenue".to_owned(), vec![12.5, 19.0, 14.25]),
                ("Cost".to_owned(), vec![8.0, 11.5, 9.75]),
            ]
            .into_iter()
            .take(series_count)
            .collect(),
            number_format: Some("0.00".to_owned()),
            ..ChartData::default()
        }
    }

    fn portable_chart_data(series_count: usize) -> ChartData {
        ChartData {
            categories: vec!["August".to_owned(), "September".to_owned()],
            series: vec![
                ("Gold".to_owned(), vec![12.5, 19.2]),
                ("Silver".to_owned(), vec![8.0, 11.5]),
            ]
            .into_iter()
            .take(series_count)
            .collect(),
            number_format: Some(r"0.##\%".to_owned()),
            category_axis_title: Some("Month".to_owned()),
            value_axis_title: Some("Change".to_owned()),
            palette: vec![
                oxml_chart::RgbColor::parse("2B6FE3").unwrap(),
                oxml_chart::RgbColor::parse("F0761F").unwrap(),
            ],
        }
    }

    fn chart_part_names(document: &Document) -> Vec<String> {
        let mut parts = document
            .package
            .content_types
            .overrides
            .iter()
            .filter_map(|(part, content_type)| {
                (content_type == content_types::CHART).then_some(part.clone())
            })
            .collect::<Vec<_>>();
        parts.sort_unstable();
        parts
    }

    fn assert_editable_chart_workbook(document: &Document, chart_part: &str) {
        let relationship = document
            .package
            .get_part_rels(chart_part)
            .and_then(|relationships| relationships.get_by_type(rel_types::PACKAGE))
            .expect("editable workbook relationship");
        let workbook_part = OpcPackage::resolve_rel_target(chart_part, &relationship.target);
        let workbook = document
            .package
            .get_part(&workbook_part)
            .expect("editable workbook part");
        OpcPackage::from_reader(Cursor::new(workbook)).expect("valid editable workbook");
    }

    #[test]
    fn word_authored_line_and_bar_charts_reopen_with_axes_and_palette() {
        let data = portable_chart_data(2);
        let mut document = Document::new();
        for kind in [ChartKind::Line, ChartKind::Bar] {
            document
                .add_chart(kind, Length::inches(6.0), Length::inches(3.5), &data)
                .expect("author portable axis chart");
        }

        let bytes = document.to_bytes().expect("save portable axis charts");
        let reopened = Document::from_bytes(&bytes).expect("reopen portable axis charts");
        let parts = chart_part_names(&reopened);
        assert_eq!(parts.len(), 2);
        for part in parts {
            let xml = reopened.package.get_part(&part).expect("chart part");
            let text = std::str::from_utf8(xml).expect("chart XML is utf8");
            assert_eq!(text.matches(r#"<c:delete val="0"/>"#).count(), 2);
            assert!(text.contains("<a:t>Month</a:t>"));
            assert!(text.contains("<a:t>Change</a:t>"));
            assert!(text.contains(r#"<c:numFmt formatCode="0.##\%" sourceLinked="0"/>"#));
            assert!(text.contains(r#"<a:srgbClr val="2B6FE3"/>"#));
            assert!(text.contains(r#"<a:srgbClr val="F0761F"/>"#));
            let chart = CT_ChartSpace::from_xml(xml).expect("typed portable chart");
            assert_authored_chart_matches(&chart, &data);
            assert_editable_chart_workbook(&reopened, &part);
        }
    }

    #[test]
    fn word_authored_pie_and_doughnut_charts_reopen_with_point_colours() {
        let data = portable_chart_data(1);
        let mut document = Document::new();
        for kind in [ChartKind::Pie, ChartKind::Doughnut] {
            document
                .add_chart(kind, Length::inches(6.0), Length::inches(3.5), &data)
                .expect("author portable circular chart");
        }

        let bytes = document.to_bytes().expect("save portable circular charts");
        let reopened = Document::from_bytes(&bytes).expect("reopen portable circular charts");
        let parts = chart_part_names(&reopened);
        assert_eq!(parts.len(), 2);
        for part in parts {
            let xml = reopened.package.get_part(&part).expect("chart part");
            let text = std::str::from_utf8(xml).expect("chart XML is utf8");
            assert!(text.contains("<c:legend"));
            assert!(text.contains(r#"<c:showPercent val="1"/>"#));
            assert_eq!(text.matches("<c:dPt>").count(), 2);
            assert!(
                text.contains(r#"<c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="2B6FE3"/>"#)
            );
            assert!(
                text.contains(r#"<c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="F0761F"/>"#)
            );
            if text.contains("<c:doughnutChart>") {
                assert!(text.contains(r#"<c:holeSize val="50"/>"#));
            }
            CT_ChartSpace::from_xml(xml).expect("typed portable circular chart");
            assert_editable_chart_workbook(&reopened, &part);
        }
    }

    #[test]
    #[ignore = "requires pinned Microsoft Word and Pages Creator Studio interoperability evidence"]
    fn word_and_pages_open_portable_authored_charts() {
        let output_dir = std::env::var_os("RDOCX_FX087_ORACLE_DIR")
            .map(std::path::PathBuf::from)
            .expect("set RDOCX_FX087_ORACLE_DIR to an existing evidence directory");
        let candidate = output_dir.join("portable-authored-charts.docx");
        let pages_output = output_dir.join("portable-authored-charts-pages.docx");
        let _ = fs::remove_file(&pages_output);
        portable_chart_document()
            .save(&candidate)
            .expect("write portable chart candidate");
        assert_eq!(sha256(&candidate), FX087_CANDIDATE_SHA256);

        for (application, version, build) in [
            ("Microsoft Word", FX087_WORD_VERSION, FX087_WORD_BUILD),
            (
                "Pages Creator Studio",
                FX087_PAGES_VERSION,
                FX087_PAGES_BUILD,
            ),
        ] {
            let plist = format!("/Applications/{application}.app/Contents/Info.plist");
            assert_eq!(plist_value(&plist, "CFBundleShortVersionString"), version);
            assert_eq!(plist_value(&plist, "CFBundleVersion"), build);
        }

        let word_script = format!(
            "with timeout of 120 seconds\ntell application \"Microsoft Word\"\nactivate\nopen POSIX file \"{}\"\ndelay 5\nclose active document saving no\nend tell\nend timeout\n",
            candidate.display()
        );
        let word = Command::new("osascript")
            .args(["-e", &word_script])
            .output()
            .expect("launch Word chart oracle");
        assert!(
            word.status.success(),
            "Word chart oracle failed: {}",
            String::from_utf8_lossy(&word.stderr)
        );

        let pages_open = Command::new("/usr/bin/open")
            .args(["-a", "/Applications/Pages Creator Studio.app"])
            .arg(&candidate)
            .output()
            .expect("open chart candidate through LaunchServices");
        assert!(
            pages_open.status.success(),
            "Pages Creator Studio launch failed: {}",
            String::from_utf8_lossy(&pages_open.stderr)
        );
        let pages_script = format!(
            "with timeout of 120 seconds\ntell application \"Pages Creator Studio\"\nactivate\ndelay 15\nset chartDocument to first document whose name is \"portable-authored-charts\"\nexport chartDocument to POSIX file \"{}\" as Microsoft Word\nclose chartDocument saving no\nend tell\nend timeout\n",
            pages_output.display()
        );
        let pages = Command::new("osascript")
            .args(["-e", &pages_script])
            .output()
            .expect("launch Pages Creator Studio chart oracle");
        assert!(
            pages.status.success(),
            "Pages Creator Studio chart oracle failed: {}",
            String::from_utf8_lossy(&pages.stderr)
        );
        eprintln!(
            "Pages Creator Studio export SHA-256: {}",
            sha256(&pages_output)
        );
        assert_portable_chart_data(
            &Document::open(&pages_output).expect("open Pages Creator Studio DOCX"),
        );
    }

    #[test]
    #[ignore = "requires the SHA-recorded Pages Creator Studio export artifact"]
    fn recorded_pages_export_preserves_portable_authored_charts() {
        let candidate = std::env::var_os("RDOCX_FX087_ORACLE_CANDIDATE")
            .map(std::path::PathBuf::from)
            .expect("set RDOCX_FX087_ORACLE_CANDIDATE to the SHA-bound candidate");
        let pages_output = std::env::var_os("RDOCX_FX087_PAGES_EXPORT")
            .map(std::path::PathBuf::from)
            .expect("set RDOCX_FX087_PAGES_EXPORT to the recorded Pages export");
        let plist = "/Applications/Pages Creator Studio.app/Contents/Info.plist";

        assert_eq!(sha256(&candidate), FX087_CANDIDATE_SHA256);
        assert_eq!(
            plist_value(plist, "CFBundleShortVersionString"),
            FX087_PAGES_VERSION
        );
        assert_eq!(plist_value(plist, "CFBundleVersion"), FX087_PAGES_BUILD);
        eprintln!(
            "recorded Pages Creator Studio export SHA-256: {}",
            sha256(&pages_output)
        );
        assert_portable_chart_data(
            &Document::open(&pages_output).expect("open recorded Pages Creator Studio DOCX"),
        );
    }

    fn portable_chart_document() -> Document {
        let data = portable_chart_data(1);
        let mut document = Document::new();
        for kind in [
            ChartKind::Line,
            ChartKind::Bar,
            ChartKind::Pie,
            ChartKind::Doughnut,
        ] {
            document
                .add_chart(kind, Length::inches(6.0), Length::inches(3.5), &data)
                .expect("author portable chart");
        }
        document
    }

    fn assert_portable_chart_data(document: &Document) {
        let relationships = document
            .package
            .get_part_rels(&document.doc_part_name)
            .expect("document relationships")
            .get_all_by_type(rel_types::CHART);
        assert_eq!(relationships.len(), 4);
        let mut chart_families = std::collections::BTreeSet::new();
        let mut workbook_parts = std::collections::BTreeSet::new();
        for relationship in relationships {
            let part =
                OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
            let chart = std::str::from_utf8(
                document
                    .package
                    .get_part(&part)
                    .expect("chart part after Pages Creator Studio"),
            )
            .expect("chart XML after Pages Creator Studio");
            if let Err(error) = CT_ChartSpace::from_xml(chart.as_bytes()) {
                assert!(
                    matches!(
                        error,
                        oxml_chart::ChartError::MissingElement(ref element)
                            if element == "c:val/c:numRef"
                    ) && chart.contains(
                        r#"<c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>12.500000</c:v></c:pt><c:pt idx="1"><c:v>19.200000</c:v></c:pt></c:numLit></c:val>"#
                    ),
                    "unexpected typed chart result after Pages Creator Studio: {error}"
                );
            }
            let family = [
                ("line", "<c:lineChart>"),
                ("bar", "<c:barChart>"),
                ("pie", "<c:pieChart>"),
                ("doughnut", "<c:doughnutChart>"),
            ]
            .into_iter()
            .find_map(|(family, marker)| chart.contains(marker).then_some(family))
            .expect("known portable chart family after Pages Creator Studio");
            assert!(chart_families.insert(family), "duplicate {family} chart");
            for value in ["August", "September", "12.500000", "19.200000", "2B6FE3"] {
                assert!(chart.contains(value), "missing {value} in {part}");
            }
            if family == "line" || family == "bar" {
                for title in ["Month", "Change"] {
                    assert!(
                        chart.contains(title),
                        "missing {title} axis title in {part}"
                    );
                }
            }
            if family != "line" {
                assert!(
                    chart.contains("F0761F"),
                    "missing second point colour in {part}"
                );
            }
            if family == "pie" || family == "doughnut" {
                assert!(chart.contains("<c:legend"), "missing legend in {part}");
                assert!(
                    chart.contains(r#"<c:showPercent val="1"/>"#),
                    "missing percentage labels in {part}"
                );
            }
            if family == "doughnut" {
                assert!(
                    chart.contains(r#"<c:holeSize val="50"/>"#),
                    "missing doughnut hole size in {part}"
                );
            }
            let workbook_relationship = document
                .package
                .get_part_rels(&part)
                .and_then(|relationships| relationships.get_by_type(rel_types::PACKAGE))
                .expect("editable workbook relationship after Pages Creator Studio");
            let workbook = OpcPackage::resolve_rel_target(&part, &workbook_relationship.target);
            assert!(
                workbook_parts.insert(workbook.clone()),
                "duplicate editable workbook target {workbook}"
            );
            let workbook = OpcPackage::from_reader(Cursor::new(
                document
                    .package
                    .get_part(&workbook)
                    .expect("editable workbook after Pages Creator Studio"),
            ))
            .expect("open editable workbook after Pages Creator Studio");
            for value in ["Gold", "August", "September", "12.5", "19.2"] {
                assert!(
                    workbook
                        .parts
                        .values()
                        .any(|part| String::from_utf8_lossy(part).contains(value)),
                    "missing {value} from editable workbook"
                );
            }
        }
        assert_eq!(chart_families.len(), 4);
        assert_eq!(workbook_parts.len(), 4);
    }

    fn assert_valid_related_theme(document: &Document) -> String {
        let relationship = document
            .package
            .get_part_rels(&document.doc_part_name)
            .and_then(|relationships| relationships.get_by_type(rel_types::THEME))
            .expect("document theme relationship");
        assert!(relationship_is_internal(relationship));
        let part = OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
        assert_eq!(
            document.package.content_types.content_type_for(&part),
            Some(content_types::THEME)
        );
        let bytes = document.package.get_part(&part).expect("theme part");
        oxml_drawing::theme::CT_OfficeStyleSheet::from_xml(bytes).expect("typed theme part");
        part
    }

    #[test]
    fn word_chart_theme_validation_is_atomic() {
        let mut missing =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        missing
            .add_chart(
                ChartKind::Line,
                Length::inches(5.0),
                Length::inches(3.0),
                &portable_chart_data(1),
            )
            .expect("replace missing theme");
        assert_valid_related_theme(&missing);

        for (bytes, content_type) in [
            (b"mistyped retained theme".as_slice(), None),
            (b"<a:theme>malformed".as_slice(), Some(content_types::THEME)),
        ] {
            let mut document = Document::new_with_profile(WordCreationProfile::Minimal(
                WordPackageClass::Document,
            ));
            let source_part = "/word/theme/theme1.xml";
            document.package.set_part(source_part, bytes.to_vec());
            if let Some(content_type) = content_type {
                document
                    .package
                    .content_types
                    .add_override(source_part, content_type);
            }
            document
                .package
                .get_or_create_part_rels(&document.doc_part_name)
                .add(rel_types::THEME, "theme/theme1.xml");

            let before = document.to_bytes().expect("serialize invalid-theme source");
            let invalid = ChartData {
                categories: Vec::new(),
                ..ChartData::default()
            };
            assert!(
                document
                    .add_chart(
                        ChartKind::Line,
                        Length::inches(5.0),
                        Length::inches(3.0),
                        &invalid,
                    )
                    .is_err()
            );
            assert_eq!(
                document.to_bytes().expect("serialize after failed chart"),
                before
            );

            document
                .add_chart(
                    ChartKind::Line,
                    Length::inches(5.0),
                    Length::inches(3.0),
                    &portable_chart_data(1),
                )
                .expect("replace invalid theme");
            let replacement = assert_valid_related_theme(&document);
            assert_ne!(replacement, source_part);
            assert_eq!(document.package.get_part(source_part), Some(bytes));
        }
    }

    #[test]
    fn word_chart_theme_reuse_keeps_typed_document_state_consistent() {
        let mut source =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let part = "/word/theme/custom-chart-theme.xml";
        let theme = oxml_drawing::theme::CT_OfficeStyleSheet::office_default();
        let bytes = theme.to_xml().expect("serialize custom theme");
        source.package.set_part(part, bytes.clone());
        source
            .package
            .content_types
            .add_override(part, content_types::THEME);
        source
            .package
            .get_or_create_part_rels(&source.doc_part_name)
            .add(rel_types::THEME, "theme/custom-chart-theme.xml");
        source.theme = Some(theme.clone());
        source.theme_part_name = Some(part.to_owned());
        source
            .identifiers
            .observe_package_graph(&source.package)
            .expect("observe themed source");
        let mut document = source;
        assert_eq!(assert_valid_related_theme(&document), part);

        document
            .set_theme(theme.clone())
            .expect("stage replacement theme model");
        assert_eq!(assert_valid_related_theme(&document), part);
        document
            .add_chart(
                ChartKind::Bar,
                Length::inches(5.0),
                Length::inches(3.0),
                &portable_chart_data(1),
            )
            .expect("reuse valid related theme");
        assert_eq!(document.theme(), Some(&theme));
        assert_eq!(assert_valid_related_theme(&document), part);
        assert_eq!(document.package.get_part(part), Some(bytes.as_slice()));

        let saved = document.to_bytes().expect("save reused theme");
        let reopened = Document::from_bytes(&saved).expect("reopen reused theme");
        assert_eq!(reopened.theme(), Some(&theme));
        assert_eq!(assert_valid_related_theme(&reopened), part);
    }

    #[test]
    fn word_chart_allocation_is_deterministic() {
        fn construct(reverse: bool) -> Vec<u8> {
            let mut document = Document::new_with_profile(WordCreationProfile::Minimal(
                WordPackageClass::Document,
            ));
            let mut occupied = [
                ("/word/charts/chart3.xml", b"chart".as_slice()),
                ("/word/embeddings/Workbook7.xlsx", b"workbook".as_slice()),
                ("/word/theme/theme3.xml", b"theme".as_slice()),
            ];
            if reverse {
                occupied.reverse();
            }
            for (part, bytes) in occupied {
                document.package.set_part(part, bytes.to_vec());
            }
            document
                .add_chart(
                    ChartKind::Doughnut,
                    Length::inches(5.0),
                    Length::inches(3.0),
                    &portable_chart_data(1),
                )
                .expect("deterministic chart allocation");
            assert!(
                document
                    .package
                    .get_part("/word/theme/theme1.xml")
                    .is_some()
            );
            assert!(
                document
                    .package
                    .get_part("/word/charts/chart4.xml")
                    .is_some()
            );
            assert!(
                document
                    .package
                    .get_part("/word/embeddings/Workbook8.xlsx")
                    .is_some()
            );
            document
                .to_bytes()
                .expect("serialize deterministic package")
        }

        assert_eq!(construct(false), construct(true));
    }

    fn assert_authored_chart_matches(chart: &CT_ChartSpace, data: &ChartData) {
        let series = chart.chart.plot_area.series().expect("typed chart series");
        assert_eq!(series.len(), data.series.len());
        for (actual, (name, values)) in series.iter().zip(&data.series) {
            assert_eq!(
                actual.name.as_ref().expect("series name").values.as_slice(),
                std::slice::from_ref(name)
            );
            let AxisData::String(categories) = actual.categories.as_ref().expect("categories")
            else {
                panic!("Word bar, line, and pie charts use string categories");
            };
            assert_eq!(categories.values, data.categories);
            assert_eq!(actual.values.values, *values);
            assert_eq!(
                actual.values.format_code,
                data.number_format.as_deref().unwrap_or("General")
            );
        }
    }

    fn deterministic_chart_layout(document: &Document) -> oxml_layout::LayoutResult {
        rdocx_layout::layout_document_deterministic(&document.build_layout_input())
            .expect("deterministic Word chart layout")
    }

    fn compatibility_page_elements(
        elements: &[oxml_layout::PositionedElement],
    ) -> Vec<&oxml_layout::PositionedElement> {
        fn collect<'a>(
            elements: &'a [oxml_layout::PositionedElement],
            output: &mut Vec<&'a oxml_layout::PositionedElement>,
        ) {
            for element in elements {
                match element {
                    oxml_layout::PositionedElement::MarkedContent { children, .. } => {
                        collect(children, output);
                    }
                    other => output.push(other),
                }
            }
        }

        let mut output = Vec::new();
        collect(elements, &mut output);
        output
    }

    fn chart_leaf_counts(layout: &oxml_layout::LayoutResult) -> (usize, usize, usize) {
        let mut paths = 0;
        let mut text = 0;
        let mut images = 0;
        oxml_layout::walk(&layout.pages[0].elements, &mut |element, _| match element {
            oxml_layout::PositionedElement::Path(_) => paths += 1,
            oxml_layout::PositionedElement::Text(_) => text += 1,
            oxml_layout::PositionedElement::Image { .. } => images += 1,
            _ => {}
        });
        (paths, text, images)
    }

    #[test]
    fn inline_word_chart_renders_backend_neutral_group() {
        let mut document = Document::new();
        document
            .add_chart(
                ChartKind::Bar,
                Length::inches(5.0),
                Length::inches(3.0),
                &f158_chart_data(2),
            )
            .expect("author inline chart");

        let layout = deterministic_chart_layout(&document);
        assert!(layout.diagnostics.is_empty());
        assert!(
            compatibility_page_elements(&layout.pages[0].elements)
                .iter()
                .any(|element| matches!(element, oxml_layout::PositionedElement::Group(_)))
        );
        let (paths, text, images) = chart_leaf_counts(&layout);
        assert!(paths > 0, "chart should lower to backend-neutral paths");
        assert!(text > 0, "chart should lower labels to shaped text");
        assert_eq!(images, 0, "chart must not be rasterized before pagination");
    }

    #[test]
    fn anchored_word_chart_uses_existing_wrap_and_z_order() {
        let mut document = Document::new();
        document
            .add_chart(
                ChartKind::Line,
                Length::inches(3.0),
                Length::inches(2.0),
                &f158_chart_data(2),
            )
            .expect("author chart");
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            panic!("chart paragraph");
        };
        let RunContent::Drawing(drawing) = &mut paragraph.runs[0].content[0] else {
            panic!("chart drawing");
        };
        let inline = drawing.inline.take().expect("authored inline chart");
        let mut anchor = rdocx_oxml::drawing::CT_Anchor::new_chart(
            inline.chart_rel_id.as_deref().expect("chart relationship"),
            inline.extent_cx.0,
            inline.extent_cy.0,
        );
        anchor.behind_doc = true;
        anchor.wrap = rdocx_oxml::drawing::WrapType::Square;
        anchor.pos_h_relative_from = rdocx_oxml::drawing::ST_RelativeFromH::Page;
        anchor.pos_v_relative_from = rdocx_oxml::drawing::ST_RelativeFromV::Page;
        anchor.pos_h_offset = Length::inches(1.0).as_emu();
        anchor.pos_v_offset = Length::inches(0.5).as_emu();
        anchor.dist_l = Length::pt(12.0).as_emu();
        anchor.dist_r = Length::pt(12.0).as_emu();
        drawing.anchor = Some(anchor);
        paragraph.add_run("Foreground text");

        let layout = deterministic_chart_layout(&document);
        assert!(layout.diagnostics.is_empty());
        let elements = compatibility_page_elements(&layout.pages[0].elements);
        let group_index = elements
            .iter()
            .position(|element| matches!(element, oxml_layout::PositionedElement::Group(_)))
            .expect("anchored chart group");
        let text_index = elements
            .iter()
            .position(|element| matches!(element, oxml_layout::PositionedElement::Text(_)))
            .expect("chart label text");
        assert!(
            group_index <= text_index,
            "behind-text group must be emitted first"
        );
        let oxml_layout::PositionedElement::Group(group) = elements[group_index] else {
            unreachable!()
        };
        assert_eq!((group.transform.e, group.transform.f), (72.0, 36.0));
        let oxml_layout::PositionedElement::Text(foreground) = elements[text_index] else {
            unreachable!()
        };
        assert!(
            foreground.origin.x >= 300.0,
            "12 point wrap distance should clear the chart's 288 point right edge"
        );
    }

    #[test]
    fn word_chart_uses_document_theme_and_default_color_map() {
        let mut document = Document::new();
        document
            .add_chart(
                ChartKind::Bar,
                Length::inches(5.0),
                Length::inches(3.0),
                &f158_chart_data(1),
            )
            .expect("author themed chart");
        let themed = oxml_drawing::theme::OFFICE_DEFAULT_XML.replace("156082", "12AB34");
        document
            .package
            .set_part("/word/theme/custom-chart-theme.xml", themed.into_bytes());
        document.package.parts.remove("/word/theme/theme1.xml");
        document
            .package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::THEME, "theme/custom-chart-theme.xml");

        let input = document.build_layout_input();
        assert_eq!(
            input.chart_color_map,
            oxml_drawing::color::ColorMap::default()
        );
        let layout =
            rdocx_layout::layout_document_deterministic(&input).expect("layout themed Word chart");
        let expected = oxml_layout::Color::from_hex("12AB34");
        let mut found = false;
        oxml_layout::walk(&layout.pages[0].elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Path(path) = element
                && path.fill == Some(oxml_layout::Paint::Solid(expected))
            {
                found = true;
            }
        });
        assert!(found, "chart series should use the document theme accent");
    }

    #[test]
    fn missing_or_malformed_word_chart_is_visible() {
        for (label, mutate) in [
            ("missing", 0_u8),
            ("malformed", 1_u8),
            ("external", 2_u8),
            ("unsupported", 3_u8),
        ] {
            let mut document = Document::new();
            document
                .add_chart(
                    ChartKind::Bar,
                    Length::inches(5.0),
                    Length::inches(3.0),
                    &f158_chart_data(1),
                )
                .expect("author chart before corrupting target");
            match mutate {
                0 => {
                    document.package.parts.remove("/word/charts/chart1.xml");
                }
                1 => document
                    .package
                    .set_part("/word/charts/chart1.xml", b"<c:chartSpace".to_vec()),
                2 => {
                    let relationship = document
                        .package
                        .get_or_create_part_rels("/word/document.xml")
                        .items
                        .iter_mut()
                        .find(|relationship| relationship.rel_type == rel_types::CHART)
                        .expect("chart relationship");
                    relationship.target = "https://example.invalid/chart.xml".to_owned();
                    relationship.target_mode = Some("External".to_owned());
                }
                3 => document.package.set_part(
                    "/word/charts/chart1.xml",
                    format!(
                        r#"<c:chartSpace xmlns:c="{}"><c:chart><c:plotArea/></c:chart></c:chartSpace>"#,
                        oxml_chart::C_NS,
                    )
                    .into_bytes(),
                ),
                _ => unreachable!(),
            }

            let layout = deterministic_chart_layout(&document);
            assert_eq!(layout.diagnostics.len(), 1, "{label}");
            assert!(
                layout.diagnostics[0]
                    .message
                    .contains("Word chart relationship")
            );
            let (_, text, images) = chart_leaf_counts(&layout);
            assert!(text > 0, "{label} chart fallback should be visible");
            assert_eq!(images, 0, "{label} chart must not become an empty image");
        }
    }

    #[test]
    fn word_and_powerpoint_chart_pixels_are_identical() {
        const DPI: &str = "150";
        const RASTERIZER: &str = "pdftoppm version 26.01.0";
        const CROP_X: &str = "150";
        const CROP_Y: &str = "150";
        const CROP_WIDTH: &str = "750";
        const CROP_HEIGHT: &str = "450";
        const WORD_SHA256: &str =
            "9acea62539e90e39078a3502c2f2a109073d60497a50db89bac065bd2b4785cf";
        const POWERPOINT_SHA256: &str =
            "f8bcefb13777e423714493a292966d0214298adc826d5487e4bfbea0f36e4582";

        let data = f158_chart_data(2);
        let evidence_dir = std::env::temp_dir();
        let word_path =
            evidence_dir.join(format!("rdocx-f159-word-chart-{}.docx", std::process::id()));
        let powerpoint_path = evidence_dir.join(format!(
            "rdocx-f159-powerpoint-chart-{}.pptx",
            std::process::id()
        ));

        let mut word =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        word.add_chart(
            ChartKind::Bar,
            Length::inches(5.0),
            Length::inches(3.0),
            &data,
        )
        .expect("author Word chart from shared data");
        let mut powerpoint = rpptx::Presentation::new().expect("open bundled PowerPoint template");
        powerpoint
            .set_slide_size(Length::inches(8.5).as_emu(), Length::inches(11.0).as_emu())
            .expect("match the Word page size");
        powerpoint.add_slide(0).expect("add chart slide");
        powerpoint
            .add_chart(
                0,
                ChartKind::Bar,
                Length::inches(1.0).as_emu(),
                Length::inches(1.0).as_emu(),
                Length::inches(5.0).as_emu(),
                Length::inches(3.0).as_emu(),
                &data,
            )
            .expect("author PowerPoint chart from shared data");
        powerpoint
            .save(&powerpoint_path)
            .expect("save PowerPoint chart artifact");
        let powerpoint_package = OpcPackage::from_reader(Cursor::new(
            powerpoint
                .to_bytes()
                .expect("serialize PowerPoint artifact"),
        ))
        .expect("open PowerPoint artifact package");
        let effective_theme = powerpoint_package
            .get_part("/ppt/theme/theme1.xml")
            .expect("PowerPoint effective theme")
            .to_vec();
        word.package
            .set_part("/word/theme/theme1.xml", effective_theme);
        word.save(&word_path).expect("save Word chart artifact");

        let word_sha = sha256(&word_path);
        let powerpoint_sha = sha256(&powerpoint_path);
        assert_eq!(word_sha, WORD_SHA256);
        assert_eq!(powerpoint_sha, POWERPOINT_SHA256);

        let version = Command::new("pdftoppm")
            .arg("-v")
            .output()
            .expect("run pinned rasterizer");
        let reported = String::from_utf8_lossy(&version.stderr);
        assert!(
            reported
                .lines()
                .next()
                .is_some_and(|line| line == RASTERIZER)
        );

        let word_pdf =
            std::env::temp_dir().join(format!("rdocx-f159-word-chart-{}.pdf", std::process::id()));
        let powerpoint_pdf = std::env::temp_dir().join(format!(
            "rdocx-f159-powerpoint-chart-{}.pdf",
            std::process::id()
        ));
        fs::write(
            &word_pdf,
            word.to_pdf_deterministic().expect("render Word chart PDF"),
        )
        .expect("write Word chart PDF");
        fs::write(
            &powerpoint_pdf,
            powerpoint
                .to_pdf_deterministic()
                .expect("render PowerPoint chart PDF"),
        )
        .expect("write PowerPoint chart PDF");

        let word_crop =
            std::env::temp_dir().join(format!("rdocx-f159-word-chart-crop-{}", std::process::id()));
        let powerpoint_crop = std::env::temp_dir().join(format!(
            "rdocx-f159-powerpoint-chart-crop-{}",
            std::process::id()
        ));
        for (pdf, crop) in [(&word_pdf, &word_crop), (&powerpoint_pdf, &powerpoint_crop)] {
            let output = Command::new("pdftoppm")
                .args([
                    "-f",
                    "1",
                    "-l",
                    "1",
                    "-singlefile",
                    "-r",
                    DPI,
                    "-x",
                    CROP_X,
                    "-y",
                    CROP_Y,
                    "-W",
                    CROP_WIDTH,
                    "-H",
                    CROP_HEIGHT,
                    "-png",
                ])
                .arg(pdf)
                .arg(crop)
                .output()
                .expect("rasterize chart crop");
            assert!(
                output.status.success(),
                "pdftoppm chart crop failed: {}",
                String::from_utf8_lossy(&output.stderr)
            );
        }

        let word_png = word_crop.with_extension("png");
        let powerpoint_png = powerpoint_crop.with_extension("png");
        let comparison = Command::new("python3")
            .args([
                "-c",
                "import sys\nsys.path.insert(0, sys.argv[3])\nfrom scripts.golden_png_harness import decode_png\na=decode_png(__import__('pathlib').Path(sys.argv[1]))\nb=decode_png(__import__('pathlib').Path(sys.argv[2]))\nassert a[:2] == b[:2] == (750, 450), (a[:2], b[:2])\ndiff=sum(a[2][i:i+4] != b[2][i:i+4] for i in range(0, len(a[2]), 4))\nprint(f'{a[0]}x{a[1]} differing={diff}')\nassert diff == 0, diff",
            ])
            .arg(&word_png)
            .arg(&powerpoint_png)
            .arg(Path::new(env!("CARGO_MANIFEST_DIR")).join("../.."))
            .output()
            .expect("decode and compare chart RGBA pixels");
        assert!(
            comparison.status.success(),
            "chart pixel comparison failed: {}{}",
            String::from_utf8_lossy(&comparison.stdout),
            String::from_utf8_lossy(&comparison.stderr)
        );
        assert_eq!(
            String::from_utf8(comparison.stdout)
                .expect("pixel comparison output is utf8")
                .trim(),
            "750x450 differing=0"
        );

        for path in [
            word_path,
            powerpoint_path,
            word_pdf,
            powerpoint_pdf,
            word_png,
            powerpoint_png,
        ] {
            fs::remove_file(path).expect("remove temporary golden evidence");
        }
    }

    #[test]
    fn added_bar_line_and_pie_charts_keep_source_data() {
        let mut document = Document::new();
        let expected = [
            (ChartKind::Bar, f158_chart_data(2), "<c:barChart>"),
            (ChartKind::Line, f158_chart_data(2), "<c:lineChart>"),
            (ChartKind::Pie, f158_chart_data(1), "<c:pieChart>"),
        ];
        for (kind, data, _) in &expected {
            document
                .add_chart(*kind, Length::inches(5.0), Length::inches(3.0), data)
                .expect("public Word chart authoring");
        }

        let bytes = document.to_bytes().expect("save authored Word charts");
        let reopened = Document::from_bytes(&bytes).expect("reopen authored Word charts");
        let mut chart_parts = reopened
            .package
            .content_types
            .overrides
            .iter()
            .filter_map(|(part, content_type)| {
                (content_type == content_types::CHART).then_some(part.as_str())
            })
            .collect::<Vec<_>>();
        chart_parts.sort_unstable();
        assert_eq!(chart_parts.len(), expected.len());
        for (part, (_, data, plot_tag)) in chart_parts.into_iter().zip(expected) {
            let xml = reopened
                .package
                .get_part(part)
                .expect("authored chart part");
            assert!(
                std::str::from_utf8(xml)
                    .expect("chart XML is utf8")
                    .contains(plot_tag)
            );
            let chart = CT_ChartSpace::from_xml(xml).expect("typed authored chart");
            assert_authored_chart_matches(&chart, &data);
        }
    }

    #[test]
    fn word_add_chart_writes_cache_and_workbook_from_one_source() {
        let data = f158_chart_data(2);
        let mut document = Document::new();
        document
            .add_chart(
                ChartKind::Bar,
                Length::inches(5.0),
                Length::inches(3.0),
                &data,
            )
            .expect("author Word chart");
        let bytes = document.to_bytes().expect("save Word chart");
        let package = Document::from_bytes(&bytes)
            .expect("reopen Word chart")
            .package;
        let chart_part = "/word/charts/chart1.xml";
        let chart = CT_ChartSpace::from_xml(package.get_part(chart_part).expect("chart part"))
            .expect("parse chart part");
        assert_authored_chart_matches(&chart, &data);
        let series = chart.chart.plot_area.series().expect("typed chart series");
        for (actual, column) in series.iter().zip(['B', 'C']) {
            assert_eq!(
                actual.name.as_ref().expect("series name").formula,
                format!("Sheet1!${column}$1")
            );
            let AxisData::String(categories) = actual.categories.as_ref().expect("categories")
            else {
                panic!("authored bar chart uses string categories");
            };
            assert_eq!(categories.formula, "Sheet1!$A$2:$A$4");
            assert_eq!(
                actual.values.formula,
                format!("Sheet1!${column}$2:${column}$4")
            );
        }
        let workbook_relationship = package
            .get_part_rels(chart_part)
            .and_then(|relationships| relationships.get_by_type(rel_types::PACKAGE))
            .expect("chart workbook relationship");
        let workbook_part =
            OpcPackage::resolve_rel_target(chart_part, &workbook_relationship.target);
        let workbook = OpcPackage::from_reader(Cursor::new(
            package.get_part(&workbook_part).expect("workbook part"),
        ))
        .expect("open embedded workbook");
        let worksheet = std::str::from_utf8(
            workbook
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet"),
        )
        .expect("worksheet XML is utf8");
        let shared_strings = std::str::from_utf8(
            workbook
                .get_part("/xl/sharedStrings.xml")
                .expect("shared strings"),
        )
        .expect("shared strings XML is utf8");
        let shared_values = shared_strings
            .split("<si>")
            .skip(1)
            .map(|item| {
                item.split_once("<t>")
                    .expect("shared string text start")
                    .1
                    .split_once("</t>")
                    .expect("shared string text end")
                    .0
            })
            .collect::<Vec<_>>();
        for (address, value) in [
            ("A1", "Category"),
            ("A2", "North"),
            ("A3", "South"),
            ("A4", "West"),
            ("B1", "Revenue"),
            ("C1", "Cost"),
        ] {
            let index = worksheet_cell_value(worksheet, address)
                .parse::<usize>()
                .expect("shared string index");
            assert_eq!(shared_values[index], value);
        }
        for (address, value) in [
            ("B2", "12.5"),
            ("B3", "19"),
            ("B4", "14.25"),
            ("C2", "8"),
            ("C3", "11.5"),
            ("C4", "9.75"),
        ] {
            assert_eq!(worksheet_cell_value(worksheet, address), value);
        }
    }

    fn worksheet_cell_value<'a>(worksheet: &'a str, address: &str) -> &'a str {
        let marker = format!(r#"<c r="{address}""#);
        worksheet
            .split_once(&marker)
            .unwrap_or_else(|| panic!("missing worksheet cell {address}"))
            .1
            .split_once("</c>")
            .expect("worksheet cell end")
            .0
            .split_once("<v>")
            .expect("worksheet value start")
            .1
            .split_once("</v>")
            .expect("worksheet value end")
            .0
    }

    #[test]
    fn word_add_chart_rejects_invalid_data_without_mutation() {
        let invalid = [
            ChartData {
                categories: Vec::new(),
                series: vec![("Revenue".to_owned(), Vec::new())],
                number_format: None,
                ..ChartData::default()
            },
            ChartData {
                categories: vec!["North".to_owned()],
                series: Vec::new(),
                number_format: None,
                ..ChartData::default()
            },
            ChartData {
                categories: vec!["North".to_owned(), "South".to_owned()],
                series: vec![("Revenue".to_owned(), vec![12.5])],
                number_format: None,
                ..ChartData::default()
            },
            ChartData {
                categories: vec!["North".to_owned()],
                series: vec![("Revenue".to_owned(), vec![f64::NAN])],
                number_format: None,
                ..ChartData::default()
            },
            ChartData {
                categories: vec!["North".to_owned()],
                series: vec![("Revenue".to_owned(), vec![12.5])],
                number_format: Some(String::new()),
                ..ChartData::default()
            },
        ];
        for data in invalid {
            let mut document = Document::new();
            let before = document.to_bytes().expect("state before invalid chart");
            assert!(
                document
                    .add_chart(
                        ChartKind::Bar,
                        Length::inches(5.0),
                        Length::inches(3.0),
                        &data,
                    )
                    .is_err()
            );
            assert_eq!(
                document.to_bytes().expect("state after invalid chart"),
                before
            );
        }

        let invalid_cases = [
            (ChartKind::Pie, Length::inches(5.0), Length::inches(3.0)),
            (ChartKind::Bar, Length::emu(0), Length::inches(3.0)),
            (ChartKind::Bar, Length::emu(-1), Length::inches(3.0)),
            (ChartKind::Bar, Length::inches(5.0), Length::emu(0)),
            (ChartKind::Bar, Length::inches(5.0), Length::emu(-1)),
        ];
        for (kind, width, height) in invalid_cases {
            let mut document = Document::new();
            let before = document.to_bytes().expect("state before invalid chart");
            assert!(
                document
                    .add_chart(kind, width, height, &f158_chart_data(2))
                    .is_err()
            );
            assert_eq!(
                document.to_bytes().expect("state after invalid chart"),
                before
            );
        }
    }

    #[test]
    fn word_add_chart_uses_inline_flow_placement() {
        let mut document = Document::new();
        let width = Length::inches(5.0);
        let height = Length::inches(3.0);
        let paragraph = document
            .add_chart(ChartKind::Line, width, height, &f158_chart_data(2))
            .expect("author inline Word chart");
        assert_eq!(paragraph.run_count(), 1);
        assert_eq!(document.content_count(), 1);
        let BodyContent::Paragraph(paragraph) = &document.document.body.content[0] else {
            panic!("Word chart must be inline flow content");
        };
        let RunContent::Drawing(drawing) = &paragraph.runs[0].content[0] else {
            panic!("Word chart paragraph must carry a drawing run");
        };
        let inline = drawing.inline.as_ref().expect("inline chart drawing");
        assert!(drawing.anchor.is_none());
        assert_eq!(inline.extent_cx.0, width.to_emu());
        assert_eq!(inline.extent_cy.0, height.to_emu());
        assert!(inline.chart_rel_id.is_some());
    }

    #[test]
    #[ignore = "requires pinned Microsoft Word and human Edit Data evidence"]
    fn word_opens_native_chart_without_repair() {
        let output = std::env::var_os("RDOCX_WORD_CHART_GATE_OUTPUT")
            .map(std::path::PathBuf::from)
            .expect("set RDOCX_WORD_CHART_GATE_OUTPUT to the SHA-bound .docx path");
        let mut document = document_with_minimal_chart();
        document.save(&output).expect("write Word chart candidate");
        assert_eq!(sha256(&output), WORD_CHART_CANDIDATE_SHA256);
        let plist = "/Applications/Microsoft Word.app/Contents/Info.plist";
        assert_eq!(
            plist_value(plist, "CFBundleShortVersionString"),
            WORD_VERSION
        );
        assert_eq!(plist_value(plist, "CFBundleVersion"), WORD_BUILD);
        let script = format!(
            "with timeout of 120 seconds\ntell application \"Microsoft Word\"\nactivate\nopen POSIX file \"{}\"\ndelay 3\nset gateDocument to active document\ntry\nset openedPath to POSIX path of (full name of gateDocument as alias)\nif openedPath is not \"{}\" then error \"Word chart candidate path mismatch\"\nclose gateDocument saving no\non error errorMessage number errorNumber\ntry\nclose gateDocument saving no\nend try\nerror errorMessage number errorNumber\nend try\nend tell\nend timeout\n",
            output.display(),
            output.display(),
        );
        let result = Command::new("osascript")
            .args(["-e", &script])
            .output()
            .expect("launch Word acceptance script");
        assert!(
            result.status.success(),
            "Microsoft Word F-157 acceptance failed: {}",
            String::from_utf8_lossy(&result.stderr)
        );
    }

    #[test]
    fn word_chart_candidate_is_bound_to_recorded_sha() {
        let output =
            std::env::temp_dir().join(format!("rdocx-f157-word-chart-{}.docx", std::process::id()));
        document_with_minimal_chart()
            .save(&output)
            .expect("write SHA-bound candidate");
        assert_eq!(sha256(&output), WORD_CHART_CANDIDATE_SHA256);
        fs::remove_file(output).expect("remove temporary candidate");
    }

    fn sha256(path: &Path) -> String {
        let output = Command::new("shasum")
            .args(["-a", "256"])
            .arg(path)
            .output()
            .unwrap_or_else(|error| panic!("{}: run shasum: {error}", path.display()));
        assert!(
            output.status.success(),
            "shasum failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        String::from_utf8(output.stdout)
            .expect("shasum output is utf8")
            .split_whitespace()
            .next()
            .expect("shasum digest")
            .to_owned()
    }

    fn plist_value(path: &str, key: &str) -> String {
        let output = Command::new("defaults")
            .args(["read", path.trim_end_matches(".plist"), key])
            .output()
            .expect("read application plist");
        assert!(output.status.success());
        String::from_utf8(output.stdout)
            .expect("plist value is utf8")
            .trim()
            .to_owned()
    }

    #[test]
    fn html_and_layout_media_use_sniffed_content_type() {
        let jpeg = [0xff, 0xd8, 0xff, 0xd9];
        let mut document = Document::new();
        document
            .package
            .set_part("/word/media/misleading.png", jpeg.to_vec());
        let relationship_id = document
            .package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::IMAGE, "media/misleading.png");

        let html_input = document.build_html_input();
        let layout_input = document.build_layout_input();

        assert_eq!(
            html_input.images[&relationship_id].content_type,
            "image/jpeg"
        );
        assert_eq!(
            layout_input.images[&relationship_id].content_type,
            "image/jpeg"
        );
    }

    #[test]
    fn deterministic_render_is_independent_of_system_fonts() {
        let mut doc = Document::new();
        doc.add_paragraph("Deterministic rendering");

        let input = doc.build_layout_input();
        let layout = rdocx_layout::layout_document_deterministic(&input)
            .expect("deterministic layout should succeed");
        let bundled_fonts = oxml_layout::bundled_fonts::bundled_font_data();

        assert!(!layout.fonts.is_empty());
        for font in &layout.fonts {
            assert!(!font.data.is_empty());
            assert!(
                bundled_fonts
                    .iter()
                    .any(|(_family, data)| *data == font.data.as_ref()),
                "resolved font '{}' did not come from the bundled font set",
                font.family
            );
        }

        let inspected = oxml_pdf::render_page_to_png(&layout, 0, 150.0)
            .expect("document should have a first page");
        let facade = doc
            .render_page_to_png_deterministic(0, 150.0)
            .expect("deterministic layout should succeed")
            .expect("document should have a first page");

        assert!(!inspected.is_empty());
        assert_eq!(facade, inspected);
    }

    #[test]
    fn deterministic_pdf_facade_reuses_bundled_font_layout() {
        let mut doc = Document::new();
        doc.add_paragraph("Deterministic PDF rendering");

        reset_layout_invocations();
        let first = doc
            .to_pdf_deterministic()
            .expect("deterministic PDF rendering should succeed");
        let second = doc
            .to_pdf_deterministic()
            .expect("cached deterministic PDF rendering should succeed");

        assert!(first.starts_with(b"%PDF-"));
        assert!(second.starts_with(b"%PDF-"));
        assert_eq!(layout_invocations(), 1);
    }

    #[test]
    fn create_new_document() {
        let doc = Document::new();
        assert_eq!(doc.paragraph_count(), 0);
        assert!(doc.section_properties().is_some());
    }

    #[test]
    fn add_paragraphs() {
        let mut doc = Document::new();
        doc.add_paragraph("First paragraph");
        doc.add_paragraph("Second paragraph");
        assert_eq!(doc.paragraph_count(), 2);

        let paras = doc.paragraphs();
        assert_eq!(paras[0].text(), "First paragraph");
        assert_eq!(paras[1].text(), "Second paragraph");
    }

    #[test]
    fn document_text_preserves_body_and_table_order() {
        let mut doc = Document::new();
        doc.add_paragraph("Before");
        let mut table = doc.add_table(1, 2);
        table.cell(0, 0).unwrap().set_text("Left");
        table.cell(0, 1).unwrap().set_text("Right");
        doc.add_paragraph("After");

        assert_eq!(doc.text(), "Before\nLeft\tRight\t\nAfter\n");
    }

    #[test]
    fn paragraph_formatting() {
        let mut doc = Document::new();
        doc.add_paragraph("Centered").alignment(Alignment::Center);

        let paras = doc.paragraphs();
        assert_eq!(paras[0].alignment(), Some(Alignment::Center));
    }

    #[test]
    fn run_formatting() {
        let mut doc = Document::new();
        let mut para = doc.add_paragraph("");
        para.add_run("Bold text").bold(true).size(14.0);

        let paras = doc.paragraphs();
        let runs: Vec<_> = paras[0].runs().collect();
        assert!(runs[0].is_bold());
        assert_eq!(runs[0].size(), Some(14.0));
    }

    #[test]
    fn round_trip_in_memory() {
        let mut doc = Document::new();
        doc.add_paragraph("Hello, World!");
        doc.add_paragraph("Second paragraph")
            .alignment(Alignment::Center);

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();

        assert_eq!(doc2.paragraph_count(), 2);
        let paras = doc2.paragraphs();
        assert_eq!(paras[0].text(), "Hello, World!");
        assert_eq!(paras[1].text(), "Second paragraph");
        assert_eq!(paras[1].alignment(), Some(Alignment::Center));
    }

    #[test]
    fn styles_present() {
        let doc = Document::new();
        assert!(doc.style("Normal").is_some());
        assert!(doc.style("Heading1").is_some());
    }

    #[test]
    fn paragraph_with_style() {
        let mut doc = Document::new();
        doc.add_paragraph("Title").style("Heading1");

        let paras = doc.paragraphs();
        assert_eq!(paras[0].style_id(), Some("Heading1"));
    }

    #[test]
    fn multiple_runs_in_paragraph() {
        let mut doc = Document::new();
        let mut para = doc.add_paragraph("");
        para.add_run("Normal ");
        para.add_run("bold ").bold(true);
        para.add_run("italic").italic(true);

        let paras = doc.paragraphs();
        assert_eq!(paras[0].text(), "Normal bold italic");
        let runs: Vec<_> = paras[0].runs().collect();
        assert_eq!(runs.len(), 3);
        assert!(!runs[0].is_bold());
        assert!(runs[1].is_bold());
        assert!(runs[2].is_italic());
    }

    #[test]
    fn add_custom_style() {
        let mut doc = Document::new();
        doc.add_style(StyleBuilder::paragraph("MyCustom", "My Custom Style").based_on("Normal"))
            .unwrap();
        assert!(doc.style("MyCustom").is_some());
        let s = doc.style("MyCustom").unwrap();
        assert_eq!(s.name(), Some("My Custom Style"));
        assert_eq!(s.based_on(), Some("Normal"));
    }

    #[test]
    fn resolve_style_properties() {
        let doc = Document::new();
        // Heading1 should inherit from docDefaults and have its own overrides
        let ppr = doc.resolve_paragraph_properties(Some("Heading1"));
        assert_eq!(ppr.keep_next, Some(true));
        assert_eq!(ppr.space_before, Some(Twips(240)));

        // Default (None) should apply Normal style
        let ppr = doc.resolve_paragraph_properties(None);
        assert_eq!(ppr.space_after, Some(Twips(160)));
    }

    #[test]
    fn resolve_run_style_properties() {
        let doc = Document::new();
        let rpr = doc.resolve_run_properties(Some("Heading1"), None);
        assert_eq!(rpr.bold, Some(true));
        assert_eq!(rpr.sz, Some(HalfPoint(32)));
        assert_eq!(rpr.font_ascii, Some("Calibri".to_string()));
    }

    #[test]
    fn set_landscape() {
        let mut doc = Document::new();
        doc.set_landscape();
        let sect = doc.section_properties().unwrap();
        assert_eq!(sect.orientation, Some(ST_PageOrientation::Landscape));
        // Width should be > height in landscape
        assert!(sect.page_width.unwrap().0 > sect.page_height.unwrap().0);
    }

    #[test]
    fn set_margins() {
        let mut doc = Document::new();
        doc.set_margins(
            Length::inches(0.5),
            Length::inches(0.75),
            Length::inches(0.5),
            Length::inches(0.75),
        );
        let sect = doc.section_properties().unwrap();
        assert_eq!(sect.margin_top, Some(Twips(720)));
        assert_eq!(sect.margin_right, Some(Twips(1080)));
    }

    #[test]
    fn set_columns() {
        let mut doc = Document::new();
        doc.set_columns(2, Length::inches(0.5));
        let sect = doc.section_properties().unwrap();
        let cols = sect.columns.as_ref().unwrap();
        assert_eq!(cols.num, Some(2));
        assert_eq!(cols.space, Some(Twips(720)));
        assert_eq!(cols.equal_width, Some(true));
    }

    #[test]
    fn set_page_size() {
        let mut doc = Document::new();
        doc.set_page_size(Length::cm(21.0), Length::cm(29.7));
        let sect = doc.section_properties().unwrap();
        // A4: ~11906tw x ~16838tw
        let w = sect.page_width.unwrap().0;
        let h = sect.page_height.unwrap().0;
        assert!((w - 11906).abs() < 5);
        assert!((h - 16838).abs() < 5);
    }

    #[test]
    fn set_different_first_page() {
        let mut doc = Document::new();
        doc.set_different_first_page(true);
        assert_eq!(doc.section_properties().unwrap().title_pg, Some(true));
    }

    #[test]
    fn content_insertion_api() {
        let mut doc = Document::new();
        doc.add_paragraph("First");
        doc.add_paragraph("Third");

        // Insert in middle
        doc.insert_paragraph(1, "Second");
        assert_eq!(doc.content_count(), 3);
        let paras = doc.paragraphs();
        assert_eq!(paras[0].text(), "First");
        assert_eq!(paras[1].text(), "Second");
        assert_eq!(paras[2].text(), "Third");

        // Insert at beginning
        doc.insert_paragraph(0, "Zeroth");
        assert_eq!(doc.content_count(), 4);
        assert_eq!(doc.paragraphs()[0].text(), "Zeroth");
    }

    #[test]
    fn find_content_index_and_remove() {
        let mut doc = Document::new();
        doc.add_paragraph("Hello");
        doc.add_paragraph("{{PLACEHOLDER}}");
        doc.add_paragraph("World");

        assert_eq!(doc.find_content_index("{{PLACEHOLDER}}"), Some(1));
        assert_eq!(doc.find_content_index("NONEXISTENT"), None);

        assert!(doc.remove_content(1));
        assert_eq!(doc.content_count(), 2);
        assert_eq!(doc.paragraphs()[1].text(), "World");

        // Out of bounds
        assert!(!doc.remove_content(10));
    }

    #[test]
    fn insert_table_at_index() {
        let mut doc = Document::new();
        doc.add_paragraph("Before");
        doc.add_paragraph("After");

        doc.insert_table(1, 2, 3);
        assert_eq!(doc.content_count(), 3);
        assert_eq!(doc.table_count(), 1);
        // Paragraphs are still in correct order
        let paras = doc.paragraphs();
        assert_eq!(paras[0].text(), "Before");
        assert_eq!(paras[1].text(), "After");
    }

    #[test]
    fn replace_text_in_body() {
        let mut doc = Document::new();
        doc.add_paragraph("Hello {{name}}!");
        doc.add_paragraph("Welcome to {{company}}.");

        let count = doc.replace_text("{{name}}", "Alice");
        assert_eq!(count, 1);
        assert_eq!(doc.paragraphs()[0].text(), "Hello Alice!");

        let count = doc.replace_text("{{company}}", "Acme");
        assert_eq!(count, 1);
        assert_eq!(doc.paragraphs()[1].text(), "Welcome to Acme.");
    }

    #[test]
    fn a_tag_split_across_five_formatted_runs_preserves_surrounding_formatting() {
        let mut doc = Document::new();
        let mut paragraph = CT_P::new();
        for (text, properties) in [
            (
                "Before {",
                CT_RPr {
                    bold: Some(true),
                    ..Default::default()
                },
            ),
            (
                "{ pro",
                CT_RPr {
                    italic: Some(true),
                    ..Default::default()
                },
            ),
            (
                "file.",
                CT_RPr {
                    color: Some("112233".to_owned()),
                    ..Default::default()
                },
            ),
            (
                "name ",
                CT_RPr {
                    strike: Some(true),
                    ..Default::default()
                },
            ),
            (
                "}} after",
                CT_RPr {
                    italic: Some(false),
                    ..Default::default()
                },
            ),
        ] {
            let mut run = CT_R::new(text);
            run.properties = Some(properties);
            paragraph.runs.push(run);
        }
        doc.document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        let count = doc
            .render_template(&serde_json::json!({"profile": {"name": "Ada"}}))
            .expect("valid template should render");

        assert_eq!(count, 1);
        let paragraph = doc.paragraph(0).expect("rendered paragraph");
        assert_eq!(paragraph.text(), "Before Ada after");
        assert_eq!(paragraph.run_count(), 2);
        assert_eq!(paragraph.run(0).unwrap().text(), "Before Ada");
        assert!(paragraph.run(0).unwrap().is_bold());
        assert_eq!(paragraph.run(1).unwrap().text(), " after");
        assert_eq!(paragraph.run(1).unwrap().italic_value(), Some(false));
    }

    #[test]
    fn dotted_scalar_paths_render_supported_json_leaves() {
        let mut doc = Document::new();
        doc.add_paragraph("{{ person.name }}");
        doc.add_paragraph("{{person.age}}");
        doc.add_paragraph("{{ person.active }}");
        doc.add_paragraph("x{{person.middle}}y");

        let count = doc
            .render_template(&serde_json::json!({
                "person": {
                    "name": "Ada",
                    "age": 37,
                    "active": true,
                    "middle": null
                }
            }))
            .expect("scalar leaves should render");

        assert_eq!(count, 4);
        assert_eq!(
            doc.paragraphs()
                .into_iter()
                .map(|paragraph| paragraph.text())
                .collect::<Vec<_>>(),
            ["Ada", "37", "true", "xy"]
        );
    }

    #[test]
    fn invalid_template_input_leaves_the_document_unchanged() {
        for (template, data) in [
            (
                "{{present}} {{missing}}",
                serde_json::json!({"present": "value"}),
            ),
            ("{{value}}", serde_json::json!({"value": [1, 2]})),
            ("{{value}}", serde_json::json!({"value": {"nested": 1}})),
            ("{{ malformed", serde_json::json!({"malformed": "value"})),
        ] {
            let mut doc = Document::new();
            doc.add_paragraph(template);
            let before_xml = doc.document.to_xml().expect("serialize before render");
            let before_parts = doc.package.parts.clone();

            reset_layout_invocations();
            doc.render_page_to_png_deterministic(0, 1.0)
                .expect("warm deterministic layout cache");
            assert_eq!(layout_invocations(), 1);

            assert!(doc.render_template(&data).is_err());
            assert_eq!(
                doc.document.to_xml().expect("serialize after rejection"),
                before_xml
            );
            assert_eq!(doc.package.parts, before_parts);
            doc.render_page_to_png_deterministic(0, 1.0)
                .expect("reuse deterministic layout after rejection");
            assert_eq!(layout_invocations(), 1);
        }
    }

    #[test]
    fn template_scalar_coverage_matches_literal_replacement() {
        let mut doc = Document::new();
        doc.add_paragraph("Body {{body}}");
        doc.set_header("Header {{header}}");
        doc.set_footer("Footer {{footer}}");
        doc.add_table(1, 1)
            .cell(0, 0)
            .expect("template table cell")
            .set_text("Table {{table}}");
        doc.document.body.content.push(BodyContent::RawXml(
            br#"<w:custom><w:txbxContent><w:p><w:r><w:t>Box {{box}}</w:t></w:r></w:p></w:txbxContent></w:custom>"#.to_vec(),
        ));

        let chart_part = "/word/charts/chart99.xml";
        doc.package.set_part(
            chart_part,
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:title><a:p><a:r><a:t>{{chart}}</a:t></a:r></a:p></c:title><c:strCache><c:pt><c:v>{{cache}}</c:v></c:pt></c:strCache></c:chartSpace>"#.to_vec(),
        );
        doc.package
            .get_or_create_part_rels(&doc.doc_part_name.clone())
            .add(rel_types::CHART, "charts/chart99.xml");

        let count = doc
            .render_template(&serde_json::json!({
                "body": "B",
                "header": "H",
                "footer": "F",
                "table": "T",
                "box": "X",
                "chart": "C",
                "cache": "K"
            }))
            .expect("all literal-replacement locations should render");

        assert_eq!(count, 7);
        assert_eq!(doc.paragraphs()[0].text(), "Body B");
        assert_eq!(
            doc.table(0)
                .expect("template table")
                .cell(0, 0)
                .expect("rendered table cell")
                .text(),
            "Table T"
        );
        assert_eq!(doc.header_text().as_deref(), Some("Header H"));
        assert_eq!(doc.footer_text().as_deref(), Some("Footer F"));
        let document_xml = std::str::from_utf8(
            doc.package
                .get_part(&doc.doc_part_name)
                .expect("main document package part"),
        )
        .expect("main document XML");
        assert!(document_xml.contains("Box X"));
        let chart_xml = std::str::from_utf8(
            doc.package
                .get_part(chart_part)
                .expect("chart package part"),
        )
        .expect("chart XML");
        assert!(chart_xml.contains("<a:t>C</a:t>"));
        assert!(chart_xml.contains("<c:v>K</c:v>"));
    }

    #[test]
    fn template_values_are_not_recursively_interpreted() {
        let mut doc = Document::new();
        doc.add_paragraph("{{first}} {{second}}");
        doc.add_paragraph("{% if later %}");
        doc.add_paragraph("hidden");
        doc.add_paragraph("{% endif %}");

        assert_eq!(
            doc.render_template(&serde_json::json!({
                "first": "{{second}}",
                "second": "done",
                "later": false
            }))
            .expect("scalar values and controls should render"),
            2
        );
        assert_eq!(doc.paragraphs().len(), 1);
        assert_eq!(doc.paragraphs()[0].text(), "{{second}} done");
    }

    #[test]
    fn template_render_preserves_unmodelled_paragraph_xml() {
        let raw = br#"<w:proofErr w:type="spellStart"/>"#.to_vec();
        let mut doc = Document::new();
        let mut paragraph = CT_P::new();
        paragraph.runs.push(CT_R::new("Value: {{ value }}"));
        paragraph.extra_xml.push((1, raw.clone()));
        doc.document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        assert_eq!(
            doc.render_template(&serde_json::json!({"value": "kept"}))
                .expect("valid template should render"),
            1
        );
        let reopened = Document::from_bytes(&doc.to_bytes().expect("save rendered document"))
            .expect("reopen rendered document");
        let BodyContent::Paragraph(paragraph) = &reopened.document.body.content[0] else {
            panic!("expected paragraph");
        };
        assert_eq!(paragraph.text(), "Value: kept");
        assert_eq!(paragraph.extra_xml, vec![(1, raw)]);
    }

    #[test]
    fn mismatched_or_cross_container_blocks_fail_without_mutation() {
        let cases = [
            vec!["{% if show %}", "content"],
            vec!["{% endif %}"],
            vec![
                "{% for item in items %}",
                "{% if item.show %}",
                "{% endfor %}",
                "{% endif %}",
            ],
        ];
        for paragraphs in cases {
            let mut document = Document::new();
            for paragraph in paragraphs {
                document.add_paragraph(paragraph);
            }
            let before = document.document.to_xml().unwrap();
            assert!(
                document
                    .render_template(&serde_json::json!({
                        "show": true,
                        "items": [{"show": true}]
                    }))
                    .is_err()
            );
            assert_eq!(document.document.to_xml().unwrap(), before);
        }

        let mut document = Document::new();
        document.add_paragraph("{% for item in items %}");
        document
            .add_table(1, 1)
            .cell(0, 0)
            .unwrap()
            .set_text("{% endfor %}");
        let before = document.document.to_xml().unwrap();
        assert!(
            document
                .render_template(&serde_json::json!({"items": [1]}))
                .is_err()
        );
        assert_eq!(document.document.to_xml().unwrap(), before);
    }

    #[test]
    fn loop_scopes_shadow_root_values_and_restore_after_exit() {
        let mut document = Document::new();
        for text in [
            "{% for item in items %}",
            "outer {{ item.name }}",
            "{% for item in item.children %}",
            "inner {{ item.name }}",
            "{% endfor %}",
            "restored {{ item.name }}",
            "{% endfor %}",
            "root {{ item.name }}",
        ] {
            document.add_paragraph(text);
        }

        document
            .render_template(&serde_json::json!({
                "item": {"name": "root"},
                "items": [{
                    "name": "outer",
                    "children": [{"name": "inner-a"}, {"name": "inner-b"}]
                }]
            }))
            .unwrap();
        assert_eq!(
            document
                .paragraphs()
                .into_iter()
                .map(|paragraph| paragraph.text())
                .collect::<Vec<_>>(),
            [
                "outer outer",
                "inner inner-a",
                "inner inner-b",
                "restored outer",
                "root root"
            ]
        );
    }

    #[test]
    fn structural_generation_preserves_schema_order_and_raw_xml() {
        let mut document = Document::new();
        document.add_paragraph("{% for item in items %}");
        let mut generated = CT_P::new();
        generated.runs.push(CT_R::new("section {{ item }}"));
        generated
            .extra_xml
            .push((1, br#"<w:proofErr w:type="spellStart"/>"#.to_vec()));
        generated.properties = Some(CT_PPr {
            sect_pr: Some(CT_SectPr::default_letter()),
            ..Default::default()
        });
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(generated));
        document.add_paragraph("{% endfor %}");

        let row = |text: &str, raw: Option<Vec<u8>>| {
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            let mut cell = CT_Tc::new();
            cell.content = vec![CellContent::Paragraph(paragraph)];
            let mut row = CT_Row::new();
            row.cells.push(cell);
            if let Some(raw) = raw {
                row.extra_xml.push((0, raw));
            }
            row
        };
        let mut table = CT_Tbl::new();
        table.rows.push(row("{% for item in items %}", None));
        table.rows.push(row(
            "row {{ item }}",
            Some(br#"<w:customRow w:val="kept"/>"#.to_vec()),
        ));
        table.rows.push(row("{% endfor %}", None));
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));

        document
            .render_template(&serde_json::json!({"items": ["one", "two"]}))
            .unwrap();
        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.document.body.content.len(), 3);
        for content in &reopened.document.body.content[..2] {
            let BodyContent::Paragraph(paragraph) = content else {
                panic!("expected generated section-ending paragraph");
            };
            assert_eq!(
                paragraph.extra_xml,
                vec![(1, br#"<w:proofErr w:type="spellStart"/>"#.to_vec())]
            );
            assert!(
                paragraph
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.sect_pr.as_ref())
                    .is_some()
            );
        }
        let BodyContent::Table(table) = &reopened.document.body.content[2] else {
            panic!("expected generated table");
        };
        assert_eq!(table.rows.len(), 2);
        assert_eq!(
            table
                .rows
                .iter()
                .map(|row| row.cells[0].text())
                .collect::<Vec<_>>(),
            ["row one", "row two"]
        );
        assert!(
            table.rows.iter().all(|row| {
                row.extra_xml == vec![(0, br#"<w:customRow w:val="kept"/>"#.to_vec())]
            })
        );
        let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(xml.rfind("<w:p>").unwrap() < xml.rfind("<w:sectPr>").unwrap());
    }

    #[test]
    fn repeated_rows_and_lists_preserve_properties_and_raw_xml() {
        use rdocx_oxml::table::{
            CT_TblGrid, CT_TblGridCol, CT_TblLook, CT_TblPr, CT_TcPr, CT_TrPr, VMerge,
        };

        let mut document = Document::new();
        document.add_paragraph("{% for item in items %}");
        document.add_numbered_list_item("List {{ item }}", 1);
        document.add_paragraph("Detail {{ item }}");
        document.add_paragraph("{% endfor %}");

        let row_raw = br#"<q:rowData xmlns:q="urn:rdocx:f165" q:value="kept"/>"#.to_vec();
        let cell_property_raw =
            br#"<q:cellProperty xmlns:q="urn:rdocx:f165" q:value="kept"/>"#.to_vec();
        let cell_raw = br#"<q:cellData xmlns:q="urn:rdocx:f165" q:value="kept"/>"#.to_vec();
        let table_raw = br#"<q:tableData xmlns:q="urn:rdocx:f165" q:value="kept"/>"#.to_vec();
        let control_xml = br#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sdt><w:sdtPr><w:tag w:val="row-control"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>controlled</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr/></w:body></w:document>"#;
        let mut control_document = CT_Document::from_xml(control_xml).unwrap();
        let BodyContent::ContentControl(control) = control_document.body.content.remove(0) else {
            panic!("expected parsed content control");
        };

        let marker_row = |text: &str| {
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            let mut cell = CT_Tc::new();
            cell.content = vec![CellContent::Paragraph(paragraph)];
            let mut row = CT_Row::new();
            row.cells.push(cell);
            row
        };
        let template_row = |text: &str, merge: VMerge, header: bool| {
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            let mut properties = CT_TcPr {
                grid_span: Some(2),
                v_merge: Some(merge),
                ..Default::default()
            };
            properties.extra_xml.push((3, cell_property_raw.clone()));
            let mut cell = CT_Tc::new();
            cell.properties = Some(properties);
            cell.content = vec![
                CellContent::Paragraph(paragraph),
                CellContent::ContentControl(control.clone()),
            ];
            cell.extra_xml.push((0, cell_raw.clone()));
            let mut row = CT_Row::new();
            row.properties = Some(CT_TrPr {
                header: Some(header),
                cnf_style: Some("100000000000".to_owned()),
                ..Default::default()
            });
            row.cells.push(cell);
            row.extra_xml.push((0, row_raw.clone()));
            row
        };
        let mut table = CT_Tbl::new();
        table.properties = Some(CT_TblPr {
            style_id: Some("BandedRows".to_owned()),
            look: Some(CT_TblLook {
                first_row: Some(true),
                no_h_band: Some(false),
                ..Default::default()
            }),
            ..Default::default()
        });
        table.grid = Some(CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(1800) },
                CT_TblGridCol { width: Twips(1800) },
                CT_TblGridCol { width: Twips(1800) },
            ],
            ..Default::default()
        });
        table.rows.push(marker_row("{% for item in items %}"));
        table
            .rows
            .push(template_row("A {{ item }}", VMerge::Restart, true));
        table
            .rows
            .push(template_row("B {{ item }}", VMerge::Continue, false));
        table
            .rows
            .push(template_row("C {{ item }}", VMerge::Continue, false));
        table.rows.push(marker_row("{% endfor %}"));
        table.extra_xml.push((table.rows.len(), table_raw.clone()));
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));

        assert_eq!(
            document
                .render_template(&serde_json::json!({"items": ["one", "two"]}))
                .unwrap(),
            10
        );
        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        let paragraphs = reopened.paragraphs();
        let numbered = paragraphs
            .iter()
            .filter_map(|paragraph| paragraph.numbering())
            .collect::<Vec<_>>();
        assert_eq!(numbered.len(), 2);
        assert_eq!(numbered[0], numbered[1]);
        assert_eq!(numbered[0].1, 1);

        let table = reopened.document.body.tables().next().unwrap();
        assert_eq!(table.rows.len(), 6);
        assert_eq!(table.grid.as_ref().unwrap().columns.len(), 3);
        assert_eq!(
            table
                .properties
                .as_ref()
                .unwrap()
                .look
                .as_ref()
                .unwrap()
                .no_h_band,
            Some(false)
        );
        assert_eq!(table.extra_xml, vec![(6, table_raw)]);
        for (index, row) in table.rows.iter().enumerate() {
            assert_eq!(row.extra_xml, vec![(0, row_raw.clone())]);
            assert_eq!(
                row.properties.as_ref().unwrap().header,
                (index % 3 == 0).then_some(true)
            );
            let cell = &row.cells[0];
            assert_eq!(cell.extra_xml, vec![(0, cell_raw.clone())]);
            let CellContent::ContentControl(control) = &cell.content[1] else {
                panic!("expected repeated cell content control");
            };
            assert_eq!(
                control.properties.as_ref().unwrap().tag.as_deref(),
                Some("row-control")
            );
            let paragraph = control
                .content
                .iter()
                .find_map(|content| match content {
                    SdtContent::Paragraph(paragraph) => Some(paragraph),
                    _ => None,
                })
                .expect("expected controlled paragraph");
            assert_eq!(paragraph.text(), "controlled");
            let properties = cell.properties.as_ref().unwrap();
            assert_eq!(properties.grid_span, Some(2));
            assert_eq!(
                properties.v_merge,
                Some(if index % 3 == 0 {
                    VMerge::Restart
                } else {
                    VMerge::Continue
                })
            );
            assert_eq!(properties.extra_xml, vec![(3, cell_property_raw.clone())]);
        }

        let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(xml.find("<w:trPr>").unwrap() < xml.find("<q:rowData").unwrap());
        assert!(xml.find("<w:gridSpan").unwrap() < xml.find("<q:cellProperty").unwrap());
        assert!(xml.find("<q:cellProperty").unwrap() < xml.find("<w:vMerge").unwrap());
        assert!(xml.rfind("<w:tr>").unwrap() < xml.rfind("<q:tableData").unwrap());
        assert!(xml.rfind("<q:tableData").unwrap() < xml.rfind("</w:tbl>").unwrap());
    }

    #[test]
    fn conditions_use_explicit_json_truthiness() {
        for (value, included) in [
            (serde_json::json!(false), false),
            (serde_json::Value::Null, false),
            (serde_json::json!(0), false),
            (serde_json::json!(""), false),
            (serde_json::json!([]), false),
            (serde_json::json!({}), false),
            (serde_json::json!(true), true),
            (serde_json::json!(-1), true),
            (serde_json::json!("x"), true),
            (serde_json::json!([1]), true),
            (serde_json::json!({"x": 1}), true),
        ] {
            let mut document = Document::new();
            document.add_paragraph("{% if value %}");
            document.add_paragraph("included");
            document.add_paragraph("{% endif %}");
            assert_eq!(
                document
                    .render_template(&serde_json::json!({"value": value}))
                    .unwrap(),
                0
            );
            assert_eq!(document.paragraphs().len(), usize::from(included));
        }
    }

    #[test]
    fn structural_only_table_blocks_commit_the_candidate() {
        let row = |text: &str| {
            let mut cell = CT_Tc::new();
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            cell.content = vec![CellContent::Paragraph(paragraph)];
            let mut row = CT_Row::new();
            row.cells.push(cell);
            row
        };
        let mut table = CT_Tbl::new();
        table.rows.push(row("{% if show %}"));
        table.rows.push(row("included"));
        table.rows.push(row("{% endif %}"));
        let mut document = Document::new();
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));

        assert_eq!(
            document
                .render_template(&serde_json::json!({"show": false}))
                .unwrap(),
            0
        );
        let BodyContent::Table(table) = &document.document.body.content[0] else {
            panic!("expected table");
        };
        assert!(table.rows.is_empty());
    }

    #[test]
    fn nested_table_controls_stay_with_their_direct_table() {
        let row = |text: &str| {
            let mut cell = CT_Tc::new();
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            cell.content = vec![CellContent::Paragraph(paragraph)];
            let mut row = CT_Row::new();
            row.cells.push(cell);
            row
        };
        let mut nested = CT_Tbl::new();
        nested.rows.push(row("{% if show %}"));
        nested.rows.push(row("nested"));
        nested.rows.push(row("{% endif %}"));
        let mut outer_cell = CT_Tc::new();
        let mut outer_paragraph = CT_P::new();
        outer_paragraph.runs.push(CT_R::new("outer"));
        outer_cell
            .content
            .push(CellContent::Paragraph(outer_paragraph));
        outer_cell.content.push(CellContent::Table(nested));
        let mut outer_row = CT_Row::new();
        outer_row.cells.push(outer_cell);
        let mut outer = CT_Tbl::new();
        outer.rows.push(outer_row);
        let mut document = Document::new();
        document
            .document
            .body
            .content
            .push(BodyContent::Table(outer));

        document
            .render_template(&serde_json::json!({"show": true}))
            .unwrap();
        let BodyContent::Table(outer) = &document.document.body.content[0] else {
            panic!("expected outer table");
        };
        let nested = outer.rows[0].cells[0]
            .content
            .iter()
            .find_map(|content| match content {
                CellContent::Table(table) => Some(table),
                CellContent::Paragraph(_) | CellContent::ContentControl(_) => None,
            })
            .expect("expected nested table");
        assert_eq!(nested.rows.len(), 1);
        assert_eq!(nested.rows[0].cells[0].text(), "nested");
    }

    #[test]
    fn direct_row_markers_reject_nested_table_content() {
        let mut nested_cell = CT_Tc::new();
        let mut nested_paragraph = CT_P::new();
        nested_paragraph.runs.push(CT_R::new("must remain"));
        nested_cell.content = vec![CellContent::Paragraph(nested_paragraph)];
        let mut nested_row = CT_Row::new();
        nested_row.cells.push(nested_cell);
        let mut nested = CT_Tbl::new();
        nested.rows.push(nested_row);

        let mut marker_cell = CT_Tc::new();
        let mut marker = CT_P::new();
        marker.runs.push(CT_R::new("{% if show %}"));
        marker_cell.content = vec![CellContent::Paragraph(marker)];
        let mut nested_cell = CT_Tc::new();
        nested_cell.content = vec![CellContent::Table(nested)];
        let mut marker_row = CT_Row::new();
        marker_row.cells.push(marker_cell);
        marker_row.cells.push(nested_cell);
        let mut table = CT_Tbl::new();
        table.rows.push(marker_row);
        for text in ["included", "{% endif %}"] {
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            let mut cell = CT_Tc::new();
            cell.content = vec![CellContent::Paragraph(paragraph)];
            let mut row = CT_Row::new();
            row.cells.push(cell);
            table.rows.push(row);
        }
        let mut document = Document::new();
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));
        let before = document.document.to_xml().unwrap();

        assert!(
            document
                .render_template(&serde_json::json!({"show": true}))
                .is_err()
        );
        assert_eq!(document.document.to_xml().unwrap(), before);
    }

    #[test]
    fn false_condition_scalar_leaves_are_preflighted() {
        let mut document = Document::new();
        document.add_paragraph("{% if show %}");
        document.add_paragraph("{{ missing }}");
        document.add_paragraph("{% endif %}");
        let before = document.document.to_xml().unwrap();

        assert!(
            document
                .render_template(&serde_json::json!({"show": false}))
                .is_err()
        );
        assert_eq!(document.document.to_xml().unwrap(), before);
    }

    #[test]
    fn empty_loop_bodies_are_preflighted_without_requiring_an_item() {
        let mut valid = Document::new();
        valid.add_paragraph("{% for item in items %}");
        valid.add_paragraph("{{ item.name }}");
        valid.add_paragraph("{% endfor %}");
        assert_eq!(
            valid
                .render_template(&serde_json::json!({"items": []}))
                .unwrap(),
            0
        );
        assert!(valid.paragraphs().is_empty());

        for invalid_body in ["{{ item.name", "{{ config.missing }}"] {
            let mut document = Document::new();
            document.add_paragraph("{% for item in items %}");
            document.add_paragraph(invalid_body);
            document.add_paragraph("{% endfor %}");
            let before = document.document.to_xml().unwrap();

            assert!(
                document
                    .render_template(&serde_json::json!({"items": [], "config": {}}))
                    .is_err()
            );
            assert_eq!(document.document.to_xml().unwrap(), before);
        }

        let mut nested = Document::new();
        for text in [
            "{% for item in items %}",
            "{% for child in config.missing %}",
            "{{ child.name }}",
            "{% endfor %}",
            "{% endfor %}",
        ] {
            nested.add_paragraph(text);
        }
        let before = nested.document.to_xml().unwrap();
        assert!(
            nested
                .render_template(&serde_json::json!({"items": [], "config": {}}))
                .is_err()
        );
        assert_eq!(nested.document.to_xml().unwrap(), before);

        let mut nested_scope = Document::new();
        for text in [
            "{% for item in items %}",
            "{% for child in item.children %}",
            "{{ settings.missing }}",
            "{% endfor %}",
            "{% endfor %}",
        ] {
            nested_scope.add_paragraph(text);
        }
        let before = nested_scope.document.to_xml().unwrap();
        assert!(
            nested_scope
                .render_template(&serde_json::json!({
                    "items": [{"children": [], "settings": {}}]
                }))
                .is_err()
        );
        assert_eq!(nested_scope.document.to_xml().unwrap(), before);
    }

    #[test]
    fn repeated_table_level_controls_use_the_row_loop_scope() {
        let xml = br#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>{% for item in items %}</w:t></w:r></w:p></w:tc></w:tr><w:sdt><w:sdtPr><w:tag w:val="row-control"/></w:sdtPr><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>control {{ item.name }}</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt><w:tr><w:tc><w:p><w:r><w:t>row {{ item.name }}</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>{% endfor %}</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#;
        let mut document = Document::new();
        document.document = CT_Document::from_xml(xml).unwrap();

        assert_eq!(
            document
                .render_template(&serde_json::json!({
                    "items": [{"name": "one"}, {"name": "two"}]
                }))
                .unwrap(),
            4
        );
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let BodyContent::Table(table) = &reopened.document.body.content[0] else {
            panic!("expected table");
        };
        assert_eq!(
            table
                .rows
                .iter()
                .map(|row| row.cells[0].text())
                .collect::<Vec<_>>(),
            ["row one", "row two"]
        );
        assert_eq!(table.content_controls.len(), 2);
        assert_eq!(
            table
                .content_controls
                .iter()
                .map(|(_, _, control)| {
                    control
                        .content
                        .iter()
                        .find_map(|content| match content {
                            SdtContent::Row(row) => Some(row.cells[0].text()),
                            _ => None,
                        })
                        .expect("expected controlled row")
                })
                .collect::<Vec<_>>(),
            ["control one", "control two"]
        );
    }

    fn template_test_row(text: &str) -> CT_Row {
        let mut paragraph = CT_P::new();
        paragraph.runs.push(CT_R::new(text));
        let mut cell = CT_Tc::new();
        cell.content.push(CellContent::Paragraph(paragraph));
        let mut row = CT_Row::new();
        row.cells.push(cell);
        row
    }

    fn template_test_row_control(row: CT_Row) -> CT_Sdt {
        let xml = br#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sdt><w:sdtContent><w:p><w:r><w:t>placeholder</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr/></w:body></w:document>"#;
        let mut parsed = CT_Document::from_xml(xml).unwrap();
        let BodyContent::ContentControl(mut control) = parsed.body.content.remove(0) else {
            panic!("expected parsed content control");
        };
        control.content = vec![SdtContent::Row(row)];
        control
    }

    fn template_test_table_with_control(open: &str, control: CT_Sdt, close: &str) -> Document {
        let mut table = CT_Tbl::new();
        table.rows.push(template_test_row(open));
        table.rows.push(template_test_row("source"));
        table.rows.push(template_test_row(close));
        table.content_controls.push((1, 0, control));
        let mut document = Document::new();
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));
        document
    }

    #[test]
    fn excluded_table_level_controls_are_preflighted() {
        let cases = [
            (
                "{% for item in items %}",
                "{% endfor %}",
                "{{ config.missing }}",
                serde_json::json!({"items": [], "config": {}}),
            ),
            (
                "{% if show %}",
                "{% endif %}",
                "{{ unclosed",
                serde_json::json!({"show": false}),
            ),
        ];
        for (open, close, control_text, data) in cases {
            let control = template_test_row_control(template_test_row(control_text));
            let mut document = template_test_table_with_control(open, control, close);
            let before = document.document.to_xml().unwrap();

            assert!(document.render_template(&data).is_err());
            assert_eq!(document.document.to_xml().unwrap(), before);
        }
    }

    #[test]
    fn repeated_table_level_controls_validate_numbering() {
        let mut numbered_row = template_test_row("{{ item.name }}");
        let CellContent::Paragraph(paragraph) = &mut numbered_row.cells[0].content[0] else {
            panic!("expected paragraph");
        };
        paragraph.properties = Some(CT_PPr {
            num_id: Some(99),
            num_ilvl: Some(0),
            ..Default::default()
        });
        let control = template_test_row_control(numbered_row);
        let mut document =
            template_test_table_with_control("{% for item in items %}", control, "{% endfor %}");
        let before = document.document.to_xml().unwrap();

        assert!(
            document
                .render_template(&serde_json::json!({"items": [{"name": "one"}]}))
                .is_err()
        );
        assert_eq!(document.document.to_xml().unwrap(), before);
    }

    #[test]
    fn replace_text_in_header_and_footer() {
        let mut doc = Document::new();
        assert!(!doc.has_header_footer_content());
        doc.set_header("Header: {{title}}");
        doc.set_footer("Footer: {{title}}");
        doc.add_paragraph("Body: {{title}}");

        assert!(doc.has_header_footer_content());

        let count = doc.replace_text("{{title}}", "My Doc");
        assert_eq!(count, 3);

        assert_eq!(doc.paragraphs()[0].text(), "Body: My Doc");
        assert_eq!(doc.header_text().unwrap(), "Header: My Doc");
        assert_eq!(doc.footer_text().unwrap(), "Footer: My Doc");

        let reopened = Document::from_bytes(&doc.to_bytes().unwrap()).unwrap();
        assert!(reopened.has_header_footer_content());
    }

    #[test]
    fn header_footer_content_includes_earlier_sections_after_round_trip() {
        let mut doc = Document::new();
        doc.set_header("Earlier section header");
        doc.add_paragraph("First section");

        let header_reference = doc
            .document
            .body
            .sect_pr
            .as_mut()
            .expect("final section")
            .header_refs
            .pop()
            .expect("header reference");
        let BodyContent::Paragraph(paragraph) = doc
            .document
            .body
            .content
            .last_mut()
            .expect("first section paragraph")
        else {
            panic!("expected paragraph");
        };
        paragraph.properties = Some(CT_PPr {
            sect_pr: Some(CT_SectPr {
                header_refs: vec![header_reference],
                ..CT_SectPr::default_letter()
            }),
            ..Default::default()
        });

        assert!(doc.has_header_footer_content());

        let reopened = Document::from_bytes(&doc.to_bytes().unwrap()).unwrap();
        assert!(reopened.has_header_footer_content());
    }

    #[test]
    fn replace_all_batch() {
        let mut doc = Document::new();
        doc.add_paragraph("{{a}} and {{b}}");

        let mut map = std::collections::HashMap::new();
        map.insert("{{a}}", "X");
        map.insert("{{b}}", "Y");
        let count = doc.replace_all(&map);
        assert_eq!(count, 2);
        assert_eq!(doc.paragraphs()[0].text(), "X and Y");
    }

    #[test]
    fn replacement_flush_failures_leave_the_live_document_unchanged() {
        fn exhausted() -> Document {
            let mut source = Document::new_with_profile(WordCreationProfile::Minimal(
                WordPackageClass::Document,
            ));
            source.add_paragraph("before {{value}}");
            let mut package =
                OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
            let relationships = package.get_or_create_part_rels("/word/document.xml");
            relationships
                .items
                .retain(|relationship| relationship.rel_type != rel_types::STYLES);
            relationships.add_with_id(&format!("rId{}", u32::MAX), "urn:producer", "producer.bin");
            package.parts.remove(DEFAULT_STYLES_PART);
            package.content_types.overrides.remove(DEFAULT_STYLES_PART);
            let mut bytes = Cursor::new(Vec::new());
            package.write_to(&mut bytes).unwrap();
            Document::from_bytes(bytes.get_ref()).unwrap()
        }

        let mut literal = exhausted();
        let before = literal.paragraphs()[0].text();
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            literal.replace_text("{{value}}", "after");
        }));
        assert!(result.is_err());
        assert_eq!(literal.paragraphs()[0].text(), before);

        let mut regex = exhausted();
        let before = regex.paragraphs()[0].text();
        let error = regex.replace_regex(r"\{\{value\}\}", "after").unwrap_err();
        assert!(
            error
                .to_string()
                .contains("styles relationship allocation failed")
        );
        assert_eq!(regex.paragraphs()[0].text(), before);
    }

    #[test]
    fn header_footer_serialization_failures_abort_literal_and_regex_replacement() {
        let mut literal = Document::new();
        literal.set_header("before {{value}}");
        FAIL_NEXT_HEADER_FOOTER_SERIALIZATION.set(true);
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            literal.replace_text("{{value}}", "after");
        }));
        assert!(result.is_err());
        assert_eq!(literal.header_text().as_deref(), Some("before {{value}}"));

        let mut regex = Document::new();
        regex.set_header("before {{value}}");
        FAIL_NEXT_HEADER_FOOTER_SERIALIZATION.set(true);
        let error = regex.replace_regex(r"\{\{value\}\}", "after").unwrap_err();
        assert!(
            error
                .to_string()
                .contains("injected header or footer serialization failure")
        );
        assert_eq!(regex.header_text().as_deref(), Some("before {{value}}"));
    }

    #[test]
    fn template_workflow_round_trip() {
        let mut doc = Document::new();
        doc.add_paragraph("Company: {{company}}");
        doc.add_paragraph("Date: {{date}}");

        doc.replace_text("{{company}}", "Acme Corp");
        doc.replace_text("{{date}}", "2026-02-22");

        // Round-trip
        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();
        assert_eq!(doc2.paragraphs()[0].text(), "Company: Acme Corp");
        assert_eq!(doc2.paragraphs()[1].text(), "Date: 2026-02-22");
    }

    #[test]
    fn template_commit_retains_the_complete_staged_bundle_state() {
        let mut source =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        source.add_paragraph("Hello {{name}}");
        let mut package = OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
        package
            .get_or_create_part_rels("/word/document.xml")
            .items
            .retain(|relationship| relationship.rel_type != rel_types::STYLES);
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let mut document = Document::from_bytes(bytes.get_ref()).unwrap();

        assert_eq!(
            document
                .render_template(&serde_json::json!({"name": "Ada"}))
                .unwrap(),
            1
        );
        let staged_target = document.styles_part_name.clone().unwrap();
        let staged_relationships = document
            .package
            .get_part_rels("/word/document.xml")
            .unwrap();
        assert_eq!(
            staged_relationships
                .items
                .iter()
                .filter(|relationship| {
                    relationship.rel_type == rel_types::STYLES
                        && relationship_is_internal(relationship)
                })
                .count(),
            1
        );

        let saved = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let styles = saved
            .get_part_rels("/word/document.xml")
            .unwrap()
            .items
            .iter()
            .filter(|relationship| {
                relationship.rel_type == rel_types::STYLES && relationship_is_internal(relationship)
            })
            .collect::<Vec<_>>();
        assert_eq!(styles.len(), 1);
        assert_eq!(
            OpcPackage::resolve_rel_target("/word/document.xml", &styles[0].target),
            staged_target
        );
        assert_eq!(document.paragraphs()[0].text(), "Hello Ada");
    }

    #[test]
    fn add_background_image_round_trip() {
        // Create a minimal 1x1 PNG
        let png_data: Vec<u8> = vec![
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, // PNG signature
            0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49,
            0x44, 0x41, 0x54, // IDAT chunk
            0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21,
            0xbc, 0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, // IEND chunk
            0xae, 0x42, 0x60, 0x82,
        ];

        let mut doc = Document::new();
        doc.add_paragraph("Hello World");
        doc.add_background_image(&png_data, "bg.png");

        // Background image paragraph should be at index 0
        assert_eq!(doc.content_count(), 2);

        // Round-trip
        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();

        // Should still have 2 content items
        assert_eq!(doc2.content_count(), 2);
        // The second paragraph should have our text
        assert_eq!(doc2.paragraphs().last().unwrap().text(), "Hello World");
    }

    #[test]
    fn add_anchored_image() {
        let png_data: Vec<u8> = vec![
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
        ];

        let mut doc = Document::new();
        doc.add_paragraph("Content");
        doc.add_anchored_image(
            &png_data,
            "overlay.png",
            Length::inches(4.0),
            Length::inches(3.0),
            false,
        );
        assert_eq!(doc.content_count(), 2);
    }

    #[test]
    fn insert_toc_basic() {
        let mut doc = Document::new();
        doc.add_paragraph("Introduction");
        doc.add_paragraph("Chapter 1").style("Heading1");
        doc.add_paragraph("Some text in chapter 1.");
        doc.add_paragraph("Section 1.1").style("Heading2");
        doc.add_paragraph("Text in section 1.1.");
        doc.add_paragraph("Chapter 2").style("Heading1");
        doc.add_paragraph("Text in chapter 2.");

        // Before TOC: 7 content elements
        assert_eq!(doc.content_count(), 7);

        // Insert TOC at index 0 with max_level 2
        doc.insert_toc(0, 2);

        // TOC adds: 1 title + 3 heading entries (Ch1, Sec1.1, Ch2) = 4 paragraphs
        assert_eq!(doc.content_count(), 11);

        // Verify TOC title
        let paras = doc.paragraphs();
        assert_eq!(paras[0].text(), "Table of Contents");

        // Verify TOC entries contain heading text
        assert_eq!(paras[1].text(), "Chapter 1\t");
        assert_eq!(paras[2].text(), "Section 1.1\t");
        assert_eq!(paras[3].text(), "Chapter 2\t");
        assert_eq!(
            doc.bookmarks()
                .iter()
                .filter_map(|bookmark| bookmark.name())
                .collect::<Vec<_>>(),
            vec!["_Toc1", "_Toc2", "_Toc3"]
        );

        // Verify round-trip: save and re-open
        let bytes = doc.to_bytes().expect("should serialize");
        let doc2 = Document::from_bytes(&bytes).expect("should open");
        assert_eq!(doc2.content_count(), 11);
        let paras2 = doc2.paragraphs();
        assert_eq!(paras2[0].text(), "Table of Contents");
    }

    #[test]
    fn insert_toc_avoids_an_existing_bookmark_id() {
        let mut doc = Document::new();
        doc.add_paragraph("Chapter").style("Heading1");
        let BodyContent::Paragraph(paragraph) = &mut doc.document.body.content[0] else {
            panic!("heading paragraph");
        };
        assert!(paragraph.insert_bookmark_start(0, 100, "existing"));
        assert!(paragraph.insert_bookmark_end(1, 100));

        doc.insert_toc(0, 1);

        let bookmarks = doc.bookmarks();
        let existing = bookmarks
            .iter()
            .find(|bookmark| bookmark.name() == Some("existing"))
            .expect("existing bookmark");
        let toc = bookmarks
            .iter()
            .find(|bookmark| bookmark.name() == Some("_Toc1"))
            .expect("TOC bookmark");
        assert_eq!(existing.id(), Some(100));
        assert_ne!(toc.id(), existing.id());
        assert_eq!(toc.id(), Some(0));
    }

    #[test]
    fn insert_toc_handles_a_maximum_numeric_suffix_without_panicking() {
        let mut doc = Document::new();
        doc.add_paragraph("Chapter").style("Heading1");
        let BodyContent::Paragraph(paragraph) = &mut doc.document.body.content[0] else {
            panic!("heading paragraph");
        };
        assert!(paragraph.insert_bookmark_start(0, 9, "_Toc18446744073709551615"));
        assert!(paragraph.insert_bookmark_end(1, 9));

        doc.insert_toc(0, 1);

        assert!(
            doc.bookmarks()
                .iter()
                .any(|bookmark| bookmark.name() == Some("_Toc1"))
        );
    }

    #[test]
    fn append_documents() {
        let mut doc_a = Document::new();
        doc_a.add_paragraph("Paragraph A1");
        doc_a.add_paragraph("Paragraph A2");

        let mut doc_b = Document::new();
        doc_b.add_paragraph("Paragraph B1");
        doc_b.add_paragraph("Paragraph B2");
        doc_b.add_paragraph("Paragraph B3");

        assert_eq!(doc_a.content_count(), 2);
        doc_a.append(&doc_b);
        assert_eq!(doc_a.content_count(), 5);

        let paras = doc_a.paragraphs();
        assert_eq!(paras[0].text(), "Paragraph A1");
        assert_eq!(paras[1].text(), "Paragraph A2");
        assert_eq!(paras[2].text(), "Paragraph B1");
        assert_eq!(paras[3].text(), "Paragraph B2");
        assert_eq!(paras[4].text(), "Paragraph B3");

        // Verify round-trip
        let bytes = doc_a.to_bytes().expect("serialize");
        let reopened = Document::from_bytes(&bytes).expect("open");
        assert_eq!(reopened.content_count(), 5);
    }

    #[test]
    fn append_variants_panic_before_mutation_when_bundle_preflight_fails() {
        fn receiver() -> Document {
            let mut source = Document::new_with_profile(WordCreationProfile::Minimal(
                WordPackageClass::Document,
            ));
            source.add_paragraph("receiver");
            let mut package =
                OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
            let relationships = package.get_or_create_part_rels("/word/document.xml");
            relationships
                .items
                .retain(|relationship| relationship.rel_type != rel_types::STYLES);
            relationships.add_with_id(&format!("rId{}", u32::MAX), "urn:producer", "producer.bin");
            package.parts.remove(DEFAULT_STYLES_PART);
            package.content_types.overrides.remove(DEFAULT_STYLES_PART);
            let mut bytes = Cursor::new(Vec::new());
            package.write_to(&mut bytes).unwrap();
            Document::from_bytes(bytes.get_ref()).unwrap()
        }

        let mut other = Document::new();
        other
            .add_style(StyleBuilder::paragraph("Merged", "Merged"))
            .unwrap();
        other.add_paragraph("other").style("Merged");
        for mutation in [
            (|document: &mut Document, other: &Document| document.append(other))
                as fn(&mut Document, &Document),
            |document, other| document.append_with_break(other, crate::SectionBreak::NextPage),
            |document, other| document.insert_document(0, other),
        ] {
            let mut document = receiver();
            let before = document.paragraphs()[0].text();
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                mutation(&mut document, &other);
            }));
            assert!(result.is_err());
            assert_eq!(document.content_count(), 1);
            assert_eq!(document.paragraphs()[0].text(), before);
        }
    }

    fn assert_style_graph_merge_failure_is_atomic(destination: &Document, source: &Document) {
        destination.validate_style_graph().unwrap();
        source.validate_style_graph().unwrap();
        for mutation in [
            (|document: &mut Document, other: &Document| document.append(other))
                as fn(&mut Document, &Document),
            |document, other| document.append_with_break(other, crate::SectionBreak::NextPage),
            |document, other| document.insert_document(0, other),
        ] {
            let mut document = destination.clone_for_staging();
            let before = document.to_bytes().unwrap();
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                mutation(&mut document, source);
            }));
            assert!(result.is_err());
            assert_eq!(document.to_bytes().unwrap(), before);
            document.validate_style_graph().unwrap();
        }
    }

    #[test]
    fn append_variants_reject_conflicting_style_graphs_before_mutation() {
        let mut destination = Document::new();
        destination.add_paragraph("destination");

        let mut conflicting_default = Document::new();
        conflicting_default
            .add_style(StyleBuilder::paragraph("SourceDefault", "Source Default"))
            .unwrap();
        conflicting_default
            .set_default_style(StyleType::Paragraph, "SourceDefault")
            .unwrap();
        conflicting_default
            .add_paragraph("source default")
            .style("SourceDefault");
        assert_style_graph_merge_failure_is_atomic(&destination, &conflicting_default);

        destination
            .add_style(StyleBuilder::paragraph(
                "CollisionParagraph",
                "Destination Collision",
            ))
            .unwrap();
        let mut conflicting_link = Document::new();
        conflicting_link
            .add_style(StyleBuilder::character(
                "CollisionCharacter",
                "Collision Character",
            ))
            .unwrap();
        conflicting_link
            .add_style(
                StyleBuilder::paragraph("CollisionParagraph", "Source Collision")
                    .linked_style("CollisionCharacter"),
            )
            .unwrap();
        conflicting_link
            .add_paragraph("source link")
            .style("CollisionParagraph");
        assert_style_graph_merge_failure_is_atomic(&destination, &conflicting_link);
    }

    #[test]
    fn append_with_section_break() {
        let mut doc_a = Document::new();
        doc_a.add_paragraph("A1");

        let mut doc_b = Document::new();
        doc_b.add_paragraph("B1");

        doc_a.append_with_break(&doc_b, crate::SectionBreak::Continuous);
        // 1 original + 1 section break paragraph + 1 merged = 3
        assert_eq!(doc_a.content_count(), 3);
    }

    #[test]
    fn insert_document_at_index() {
        let mut doc_a = Document::new();
        doc_a.add_paragraph("First");
        doc_a.add_paragraph("Last");

        let mut doc_b = Document::new();
        doc_b.add_paragraph("Middle 1");
        doc_b.add_paragraph("Middle 2");

        doc_a.insert_document(1, &doc_b);
        assert_eq!(doc_a.content_count(), 4);

        let paras = doc_a.paragraphs();
        assert_eq!(paras[0].text(), "First");
        assert_eq!(paras[1].text(), "Middle 1");
        assert_eq!(paras[2].text(), "Middle 2");
        assert_eq!(paras[3].text(), "Last");
    }

    #[test]
    fn merge_deduplicates_styles() {
        let mut doc_a = Document::new();
        doc_a.add_paragraph("A").style("Heading1");

        let mut doc_b = Document::new();
        doc_b.add_paragraph("B").style("Heading1");
        doc_b
            .add_style(
                crate::style::StyleBuilder::paragraph("CustomB", "Custom B").based_on("Normal"),
            )
            .unwrap();
        doc_b.add_paragraph("C").style("CustomB");

        let styles_before = doc_a.styles.styles.len();
        doc_a.append(&doc_b);
        let styles_after = doc_a.styles.styles.len();
        doc_a.validate_style_graph().unwrap();

        // Heading1 already existed, so only CustomB should be added
        assert_eq!(styles_after, styles_before + 1);
    }

    #[test]
    fn headings_and_outline() {
        let mut doc = Document::new();
        doc.add_paragraph("Intro");
        doc.add_paragraph("Chapter 1").style("Heading1");
        doc.add_paragraph("Section 1.1").style("Heading2");
        doc.add_paragraph("Section 1.2").style("Heading2");
        doc.add_paragraph("Chapter 2").style("Heading1");
        doc.add_paragraph("Section 2.1").style("Heading2");
        doc.add_paragraph("Sub 2.1.1").style("Heading3");

        let headings = doc.headings();
        assert_eq!(headings.len(), 6);
        assert_eq!(headings[0], (1, "Chapter 1".to_string()));
        assert_eq!(headings[1], (2, "Section 1.1".to_string()));
        assert_eq!(headings[5], (3, "Sub 2.1.1".to_string()));

        let outline = doc.document_outline();
        assert_eq!(outline.len(), 2); // Two h1 nodes
        assert_eq!(outline[0].text, "Chapter 1");
        assert_eq!(outline[0].children.len(), 2); // 1.1 and 1.2
        assert_eq!(outline[1].text, "Chapter 2");
        assert_eq!(outline[1].children.len(), 1); // 2.1
        assert_eq!(outline[1].children[0].children.len(), 1); // 2.1.1
    }

    #[test]
    fn word_count_basic() {
        let mut doc = Document::new();
        doc.add_paragraph("Hello world");
        doc.add_paragraph("Three more words");
        assert_eq!(doc.word_count(), 5);
    }

    #[test]
    fn audit_accessibility_missing_metadata() {
        let doc = Document::new();
        let issues = doc.audit_accessibility();
        // New document has no title or author
        assert!(issues.iter().any(|i| i.message.contains("no title")));
        assert!(issues.iter().any(|i| i.message.contains("no author")));
    }

    #[test]
    fn audit_heading_level_gap() {
        let mut doc = Document::new();
        doc.set_title("Test");
        doc.set_author("Test");
        doc.add_paragraph("Ch 1").style("Heading1");
        doc.add_paragraph("Skip to 3").style("Heading3");

        let issues = doc.audit_accessibility();
        assert!(
            issues
                .iter()
                .any(|i| i.message.contains("Heading level gap"))
        );
    }

    #[test]
    fn links_returns_empty_for_no_hyperlinks() {
        let mut doc = Document::new();
        doc.add_paragraph("No links here.");
        assert!(doc.links().is_empty());
    }

    #[test]
    fn links_exposes_a_complex_field_hyperlink_target() {
        let mut doc = Document::new();
        let xml = concat!(
            r#"<w:p>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText> HYPERLINK &quot;https://example.test&quot; </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>Cached link</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"</w:p>"#,
        );
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buffer = Vec::new();
        let paragraph = loop {
            match reader.read_event_into(&mut buffer).unwrap() {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"p") => {
                    break CT_P::from_xml(&mut reader).unwrap();
                }
                Event::Eof => panic!("missing paragraph"),
                _ => {}
            }
            buffer.clear();
        };
        doc.document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        assert_eq!(
            doc.links(),
            vec![LinkInfo {
                text: "Cached link".to_owned(),
                url: Some("https://example.test".to_owned()),
                anchor: None,
                rel_id: None,
            }]
        );
    }

    #[test]
    fn images_returns_empty_for_text_only() {
        let mut doc = Document::new();
        doc.add_paragraph("Just text.");
        assert!(doc.images().is_empty());
    }

    #[test]
    fn numbering_getter_round_trips() {
        let mut doc = Document::new();
        doc.add_bullet_list_item("bullet item", 0);
        doc.add_numbered_list_item("numbered item", 0);
        doc.add_paragraph("plain");

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();
        let paras = doc2.paragraphs();

        let (bullet_id, bullet_lvl) = paras[0].numbering().expect("bullet numbering");
        assert_eq!(bullet_lvl, 0);
        assert_eq!(doc2.numbering_is_bullet(bullet_id), Some(true));

        let (num_id, _) = paras[1].numbering().expect("numbered numbering");
        assert_eq!(doc2.numbering_is_bullet(num_id), Some(false));

        assert!(paras[2].numbering().is_none());
    }

    #[test]
    fn highlight_getter_round_trips() {
        let mut doc = Document::new();
        {
            let mut p = doc.add_paragraph("");
            let mut r = p.add_run("glowing");
            r = r.highlight("yellow");
            let _ = r;
        }

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();
        let paras = doc2.paragraphs();
        let run = paras[0].runs().next().expect("run");
        assert_eq!(run.highlight().as_deref(), Some("yellow"));
    }

    #[test]
    fn run_style_id_getter_round_trips() {
        let mut doc = Document::new();
        {
            let mut p = doc.add_paragraph("");
            let mut r = p.add_run("code text");
            r = r.style("SourceText");
            let _ = r;
        }

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();
        let paras = doc2.paragraphs();
        let run = paras[0].runs().next().expect("run");
        assert_eq!(run.style_id(), Some("SourceText"));
    }

    #[test]
    fn append_hyperlink_round_trips() {
        let mut doc = Document::new();
        doc.add_paragraph("visit ");
        doc.append_hyperlink("GNOME", "https://gnome.org");

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();

        let links = doc2.links();
        assert_eq!(links.len(), 1);
        assert_eq!(links[0].text, "GNOME");
        assert_eq!(links[0].url.as_deref(), Some("https://gnome.org"));
        assert_eq!(doc2.paragraphs()[0].text(), "visit GNOME");

        let paras = doc2.paragraphs();
        let spans = paras[0].hyperlink_spans();
        assert_eq!(spans.len(), 1);
        let (start, end, rel_id) = (spans[0].0, spans[0].1, spans[0].2);
        assert_eq!(end - start, 1);
        let url = doc2.hyperlink_url(rel_id.expect("rel id"));
        assert_eq!(url.as_deref(), Some("https://gnome.org"));
    }

    #[test]
    fn paragraph_hard_break_and_table_cell_hyperlink_round_trip() {
        let mut doc = Document::new();
        let relationship_id = doc.add_hyperlink_relationship("https://example.com/table");

        let mut paragraph = doc.add_paragraph("");
        paragraph.add_run("before");
        paragraph.add_line_break();
        paragraph.add_run("after");

        let mut table = doc.add_table(1, 1);
        let mut cell = table.cell(0, 0).expect("cell");
        cell.remove_first_empty_paragraph();
        cell.add_paragraph("")
            .add_hyperlink("table link", &relationship_id)
            .bold(true);

        let bytes = doc.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();

        assert_eq!(reopened.paragraphs()[0].text(), "before\nafter");
        let tables = reopened.tables();
        let cell = tables[0].cell(0, 0).expect("cell");
        let paragraph = cell.paragraphs().next().expect("paragraph");
        assert_eq!(paragraph.text(), "table link");
        assert!(paragraph.runs().next().expect("run").is_bold());
        let spans = paragraph.hyperlink_spans();
        assert_eq!(spans.len(), 1);
        assert_eq!(
            reopened.hyperlink_url(spans[0].2.expect("relationship id")),
            Some("https://example.com/table".to_string())
        );
    }

    #[test]
    fn writer_hyperlink_tooltip_and_table_indent_round_trip() {
        let mut doc = Document::new();
        let relationship_id = doc.add_hyperlink_relationship("https://example.com");
        doc.add_paragraph("").add_hyperlink_with_tooltip(
            "linked",
            &relationship_id,
            Some("Example site"),
        );
        doc.add_table(1, 1).set_indent(Length::twips(720));

        let bytes = doc.to_bytes().expect("document writes");
        let reopened = Document::from_bytes(&bytes).expect("document reopens");

        let BodyContent::Paragraph(paragraph) = &reopened.document.body.content[0] else {
            panic!("expected paragraph");
        };
        assert_eq!(
            paragraph.hyperlinks[0].tooltip.as_deref(),
            Some("Example site")
        );
        assert!(paragraph.hyperlinks[0].extra_attributes.is_empty());
        let BodyContent::Table(table) = &reopened.document.body.content[1] else {
            panic!("expected table");
        };
        assert_eq!(
            table
                .properties
                .as_ref()
                .and_then(|properties| properties.indent.as_ref()),
            Some(&rdocx_oxml::table::CT_TblWidth::dxa(720))
        );
    }

    #[test]
    fn authored_hyphenation_setting_and_run_language_round_trip_into_layout() {
        let mut document = Document::new();
        document.set_auto_hyphenation(true).unwrap();
        document
            .add_paragraph("")
            .add_run("representation")
            .language("en-US");

        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert!(reopened.build_layout_input().automatic_hyphenation);
        assert_eq!(
            reopened.paragraphs()[0].run(0).unwrap().language(),
            Some("en-US")
        );
        let settings_part = reopened.settings_part_name.as_deref().unwrap();
        let settings_xml =
            std::str::from_utf8(reopened.package.get_part(settings_part).unwrap()).unwrap();
        assert!(settings_xml.contains("<w:autoHyphenation/>"));
    }

    #[test]
    fn authored_settings_do_not_overwrite_an_unrelated_conventional_part() {
        let mut document =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let unrelated = br#"<producer:metadata xmlns:producer="urn:producer"/>"#.to_vec();
        document
            .package
            .set_part(DEFAULT_SETTINGS_PART, unrelated.clone());
        document
            .package
            .content_types
            .add_override(DEFAULT_SETTINGS_PART, "application/example+xml");

        document.set_auto_hyphenation(true).unwrap();
        let bytes = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();

        assert_eq!(
            package.get_part(DEFAULT_SETTINGS_PART),
            Some(unrelated.as_slice())
        );
        assert_eq!(
            package
                .content_types
                .content_type_for(DEFAULT_SETTINGS_PART),
            Some("application/example+xml")
        );
        let relationship = package
            .get_part_rels("/word/document.xml")
            .and_then(|relationships| relationships.get_by_type(rel_types::SETTINGS))
            .expect("authored settings relationship");
        let settings_part =
            OpcPackage::resolve_rel_target("/word/document.xml", &relationship.target);
        assert_ne!(settings_part, DEFAULT_SETTINGS_PART);
        assert!(
            std::str::from_utf8(package.get_part(&settings_part).unwrap())
                .unwrap()
                .contains("<w:autoHyphenation/>")
        );
    }

    #[test]
    fn pending_settings_part_name_is_reserved_against_rich_fragment_imports() {
        let mut document =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        document.set_auto_hyphenation(true).unwrap();
        let settings_part = document.settings_part_name.clone().unwrap();

        let imported = document
            .identifiers
            .reserve_fragment_part_name(&settings_part)
            .unwrap();

        assert_ne!(imported, settings_part);
        assert!(imported.contains("-merge-"));
    }

    #[test]
    fn settings_mutations_roll_back_relationship_exhaustion() {
        fn exhausted_document() -> Document {
            let mut source = Document::new_with_profile(WordCreationProfile::Minimal(
                WordPackageClass::Document,
            ));
            let bytes = source.to_bytes().unwrap();
            let mut package = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
            let relationships = package.get_or_create_part_rels("/word/document.xml");
            relationships.add_with_id(
                &format!("rId{}", u32::MAX),
                "urn:exhaustion",
                "unchanged.bin",
            );
            let mut bytes = Cursor::new(Vec::new());
            package.write_to(&mut bytes).unwrap();
            Document::from_bytes(bytes.get_ref()).unwrap()
        }

        let package_bytes = |document: &Document| {
            let mut bytes = Cursor::new(Vec::new());
            document.package.write_to(&mut bytes).unwrap();
            bytes.into_inner()
        };

        let mut hyphenation = exhausted_document();
        let before = package_bytes(&hyphenation);
        let error = hyphenation.set_auto_hyphenation(true).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("settings relationship allocation failed")
        );
        assert_eq!(package_bytes(&hyphenation), before);
        assert!(hyphenation.settings_part_name.is_none());

        let mut math = exhausted_document();
        let before = package_bytes(&math);
        let error = math
            .set_math_properties(MathProperties::default())
            .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("settings relationship allocation failed")
        );
        assert_eq!(package_bytes(&math), before);
        assert!(math.settings_part_name.is_none());
    }

    #[test]
    fn serialization_reports_checked_styles_and_comments_relationship_exhaustion() {
        let mut style_source =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let source_bytes = style_source.to_bytes().unwrap();
        let mut style_package = OpcPackage::from_reader(Cursor::new(source_bytes)).unwrap();
        let relationships = style_package.get_or_create_part_rels("/word/document.xml");
        relationships
            .items
            .retain(|relationship| relationship.rel_type != rel_types::STYLES);
        relationships.add_with_id(
            &format!("rId{}", u32::MAX),
            "urn:exhaustion",
            "unchanged.bin",
        );
        style_package.parts.remove(DEFAULT_STYLES_PART);
        style_package
            .content_types
            .overrides
            .remove(DEFAULT_STYLES_PART);
        let mut bytes = Cursor::new(Vec::new());
        style_package.write_to(&mut bytes).unwrap();
        let mut style_less = Document::from_bytes(bytes.get_ref()).unwrap();
        let error = style_less.to_bytes().unwrap_err();
        assert!(
            error
                .to_string()
                .contains("styles relationship allocation failed")
        );

        let mut comments_source =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let source_bytes = comments_source.to_bytes().unwrap();
        let mut comments_package = OpcPackage::from_reader(Cursor::new(source_bytes)).unwrap();
        comments_package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id(
                &format!("rId{}", u32::MAX),
                "urn:exhaustion",
                "unchanged.bin",
            );
        let mut bytes = Cursor::new(Vec::new());
        comments_package.write_to(&mut bytes).unwrap();
        let mut comments = Document::from_bytes(bytes.get_ref()).unwrap();
        comments.comments = Some(rdocx_oxml::comments::CT_Comments::new());
        comments.comments_part_name = Some("/word/comments.xml".to_owned());
        let error = comments.to_bytes().unwrap_err();
        assert!(
            error
                .to_string()
                .contains("comments relationship allocation failed")
        );
    }

    #[test]
    fn all_public_numbering_level_properties_survive_reopen() {
        let formats = [
            ListNumberFormat::Decimal,
            ListNumberFormat::UpperRoman,
            ListNumberFormat::LowerRoman,
            ListNumberFormat::UpperLetter,
            ListNumberFormat::LowerLetter,
            ListNumberFormat::Ordinal,
            ListNumberFormat::CardinalText,
            ListNumberFormat::OrdinalText,
            ListNumberFormat::Hex,
            ListNumberFormat::Chicago,
            ListNumberFormat::IdeographDigital,
            ListNumberFormat::JapaneseCounting,
            ListNumberFormat::Aiueo,
            ListNumberFormat::Iroha,
            ListNumberFormat::DecimalFullWidth,
            ListNumberFormat::DecimalHalfWidth,
            ListNumberFormat::JapaneseLegal,
            ListNumberFormat::JapaneseDigitalTenThousand,
            ListNumberFormat::DecimalEnclosedCircle,
            ListNumberFormat::DecimalFullWidth2,
            ListNumberFormat::AiueoFullWidth,
            ListNumberFormat::IrohaFullWidth,
            ListNumberFormat::DecimalZero,
            ListNumberFormat::Bullet,
            ListNumberFormat::Ganada,
            ListNumberFormat::Chosung,
            ListNumberFormat::DecimalEnclosedFullstop,
            ListNumberFormat::DecimalEnclosedParen,
            ListNumberFormat::DecimalEnclosedCircleChinese,
            ListNumberFormat::IdeographEnclosedCircle,
            ListNumberFormat::IdeographTraditional,
            ListNumberFormat::IdeographZodiac,
            ListNumberFormat::IdeographZodiacTraditional,
            ListNumberFormat::TaiwaneseCounting,
            ListNumberFormat::IdeographLegalTraditional,
            ListNumberFormat::TaiwaneseCountingThousand,
            ListNumberFormat::TaiwaneseDigital,
            ListNumberFormat::ChineseCounting,
            ListNumberFormat::ChineseLegalSimplified,
            ListNumberFormat::ChineseCountingThousand,
            ListNumberFormat::KoreanDigital,
            ListNumberFormat::KoreanCounting,
            ListNumberFormat::KoreanLegal,
            ListNumberFormat::KoreanDigital2,
            ListNumberFormat::Hebrew1,
            ListNumberFormat::ArabicAlpha,
            ListNumberFormat::Hebrew2,
            ListNumberFormat::ArabicAbjad,
            ListNumberFormat::HindiVowels,
            ListNumberFormat::HindiConsonants,
            ListNumberFormat::HindiNumbers,
            ListNumberFormat::HindiCounting,
            ListNumberFormat::ThaiLetters,
            ListNumberFormat::ThaiNumbers,
            ListNumberFormat::ThaiCounting,
            ListNumberFormat::VietnameseCounting,
            ListNumberFormat::NumberInDash,
            ListNumberFormat::RussianLower,
            ListNumberFormat::RussianUpper,
            ListNumberFormat::None,
        ];
        let mut document = Document::new();
        for chunk in formats.chunks(9) {
            let levels = chunk
                .iter()
                .cloned()
                .map(ListLevel::new)
                .collect::<Vec<_>>();
            document.add_numbering_definition(&levels).unwrap();
        }
        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        let reopened_formats = reopened
            .numbering_definitions()
            .into_iter()
            .flat_map(|definition| definition.levels)
            .map(|level| level.properties.format)
            .collect::<Vec<_>>();
        assert_eq!(reopened_formats, formats);

        let mut document = Document::new();
        let levels = [
            ListLevel::decimal(),
            ListLevel::new(ListNumberFormat::LowerRoman)
                .start(4)
                .level_text("%1.%2)")
                .suffix(ListLevelSuffix::Space)
                .alignment(Alignment::Center)
                .indentation(Some(Twips(1440)), Some(Twips(240)), None)
                .marker_properties(CT_RPr {
                    bold: Some(true),
                    color: Some("2B6FE3".to_owned()),
                    ..CT_RPr::default()
                })
                .legal_numbering(true)
                .restart(ListLevelRestart::After(0))
                .template_code("A1B2C3D4")
                .tentative(false),
        ];
        let definition_id = document.add_numbering_definition(&levels).unwrap();
        let instance_id = document.add_numbering_instance(definition_id, &[]).unwrap();
        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        let level = reopened.numbering_level(instance_id, 1).unwrap();
        assert_eq!(level.format, NumberingFormat::LowerRoman);
        assert_eq!(level.start, 4);
        assert_eq!(level.level_text, Some("%1.%2)"));
        assert_eq!(level.suffix, ListLevelSuffix::Space);
        assert_eq!(level.alignment, Some(Alignment::Center));
        assert_eq!(level.indent_left, Some(Twips(1440)));
        assert_eq!(level.indent_hanging, Some(Twips(240)));
        assert_eq!(level.legal_numbering, Some(true));
        assert_eq!(level.restart, Some(ListLevelRestart::After(0)));
        assert_eq!(level.template_code, Some("A1B2C3D4"));
        assert_eq!(level.tentative, Some(false));
        assert_eq!(level.marker_properties.unwrap().bold, Some(true));
        assert_eq!(
            level.marker_properties.unwrap().color.as_deref(),
            Some("2B6FE3")
        );

        let mut imported = Document::new();
        let definition_id = imported
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance_id = imported.add_numbering_instance(definition_id, &[]).unwrap();
        let xml = format!(
            r#"<n:numbering xmlns:n="{WORD_NAMESPACE}"><n:abstractNum n:abstractNumId="{definition_id}"><n:lvl n:ilvl="0"><n:start n:val="1"/><n:numFmt n:val="decimal"/><n:pStyle n:val="Normal"></n:pStyle><n:lvlText n:val="%1."/></n:lvl></n:abstractNum><n:num n:numId="{instance_id}"><n:abstractNumId n:val="{definition_id}"/></n:num></n:numbering>"#
        );
        let bytes = replace_numbering_xml(&mut imported, xml.into_bytes());
        let imported = Document::from_bytes(&bytes).unwrap();
        let definition = imported.numbering_definition(definition_id).unwrap();
        assert_eq!(
            definition.paragraph_style_links,
            vec![Some("Normal".to_owned())]
        );
        assert!(!definition.has_unmodeled_properties);
        let level = imported.numbering_level(instance_id, 0).unwrap();
        assert_eq!(level.paragraph_style, Some("Normal"));
        assert!(!level.has_unmodeled_properties);
    }

    #[test]
    fn list_level_semantic_equality_ignores_source_provenance() {
        let authored = ListLevel::new(ListNumberFormat::LowerRoman)
            .start(4)
            .level_text("%1.%2)")
            .suffix(ListLevelSuffix::Space)
            .alignment(Alignment::Center)
            .indentation(Some(Twips(1440)), Some(Twips(240)), None)
            .marker_properties(CT_RPr {
                bold: Some(true),
                color: Some("2B6FE3".to_owned()),
                ..CT_RPr::default()
            })
            .legal_numbering(true)
            .restart(ListLevelRestart::After(0))
            .template_code("A1B2C3D4")
            .tentative(false);
        let source = ct_level_from_list(1, &authored).unwrap();
        let inspected = list_level_from_ct(&source);
        let reconstructed = list_level_values_from_ct(&source);

        assert_eq!(inspected, reconstructed);
        assert_ne!(inspected, reconstructed.clone().start(9));
    }

    #[test]
    fn numbering_instances_and_overrides_round_trip_in_schema_order() {
        let mut document = Document::new();
        let definition_id = document
            .add_numbering_definition(&[ListLevel::decimal(), ListLevel::decimal()])
            .unwrap();
        let instance_id = document
            .add_numbering_instance(
                definition_id,
                &[NumberingLevelOverride::new(1)
                    .start(5)
                    .replacement(ListLevel::new(ListNumberFormat::UpperRoman).level_text("%2)"))],
            )
            .unwrap();
        let bytes = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(Cursor::new(&bytes)).unwrap();
        let xml =
            String::from_utf8(package.get_part("/word/numbering.xml").unwrap().to_vec()).unwrap();
        let instance = xml
            .find(&format!("<w:num w:numId=\"{instance_id}\""))
            .unwrap();
        let abstract_id = xml[instance..].find("<w:abstractNumId ").unwrap();
        let level_override = xml[instance..].find("<w:lvlOverride ").unwrap();
        let start_override = xml[instance + level_override..]
            .find("<w:startOverride ")
            .unwrap();
        let replacement = xml[instance + level_override..]
            .find("<w:lvl w:ilvl=\"1\"")
            .unwrap();
        assert!(abstract_id < level_override);
        assert!(start_override < replacement);

        let reopened = Document::from_bytes(&bytes).unwrap();
        let instance = reopened.numbering_instance(instance_id).unwrap();
        assert_eq!(instance.definition_id, definition_id);
        assert_eq!(instance.level_overrides.len(), 1);
        assert_eq!(instance.level_overrides[0].level, 1);
        assert_eq!(instance.level_overrides[0].start, Some(5));
        assert_eq!(
            instance.level_overrides[0]
                .replacement
                .as_ref()
                .unwrap()
                .format,
            ListNumberFormat::UpperRoman
        );
    }

    #[test]
    fn public_authored_numbering_reports_no_unmodeled_properties() {
        let mut document = Document::new();
        let definition_id = document
            .add_numbering_definition(&[
                ListLevel::decimal().level_text("%1."),
                ListLevel::new(ListNumberFormat::LowerLetter)
                    .level_text("%1.%2.")
                    .legal_numbering(false),
            ])
            .unwrap();
        let instance_id = document
            .add_numbering_instance(definition_id, &[NumberingLevelOverride::new(1).start(3)])
            .unwrap();
        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert!(
            !reopened
                .numbering_definition(definition_id)
                .unwrap()
                .has_unmodeled_properties
        );
        assert!(
            !reopened
                .numbering_instance(instance_id)
                .unwrap()
                .has_unmodeled_properties
        );
        assert!(
            !reopened
                .numbering_level(instance_id, 1)
                .unwrap()
                .has_unmodeled_properties
        );
    }

    #[test]
    fn imported_numbering_projection_update_preserves_unknown_format_and_omissions() {
        let mut source = Document::new();
        let definition_id = source
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance_id = source.add_numbering_instance(definition_id, &[]).unwrap();
        let xml = format!(
            r#"<w:numbering xmlns:w="{WORD_NAMESPACE}"><w:abstractNum w:abstractNumId="{definition_id}"><w:lvl w:ilvl="0"><w:numFmt w:val="producerFormat"/></w:lvl></w:abstractNum><w:num w:numId="{instance_id}"><w:abstractNumId w:val="{definition_id}"/></w:num></w:numbering>"#
        );
        let imported = replace_numbering_xml(&mut source, xml.into_bytes());
        let mut document = Document::from_bytes(&imported).unwrap();
        let canonical = document.to_bytes().unwrap();
        let before = OpcPackage::from_reader(Cursor::new(&canonical))
            .unwrap()
            .get_part(DEFAULT_NUMBERING_PART)
            .unwrap()
            .to_vec();

        let definition = document.numbering_definition(definition_id).unwrap();
        assert_eq!(
            definition.levels[0].properties.format,
            ListNumberFormat::Other("producerFormat".to_owned())
        );
        assert_eq!(definition.levels[0].properties.start, None);
        assert_eq!(
            definition.levels[0].properties.indentation_value(),
            (None, None, None)
        );
        document
            .update_numbering_definition(definition_id, &definition.levels)
            .unwrap();
        let after_bytes = document.to_bytes().unwrap();
        let after = OpcPackage::from_reader(Cursor::new(after_bytes))
            .unwrap()
            .get_part(DEFAULT_NUMBERING_PART)
            .unwrap()
            .to_vec();
        assert_eq!(after, before);
    }

    #[test]
    fn numbering_crud_updates_and_removes_unreferenced_values() {
        let mut document = Document::new();
        let first_definition = document
            .add_numbering_definition(&[ListLevel::decimal(), ListLevel::decimal()])
            .unwrap();
        let second_definition = document
            .add_numbering_definition(&[ListLevel::decimal(), ListLevel::decimal()])
            .unwrap();
        assert_ne!(first_definition, second_definition);
        let instance_id = document
            .add_numbering_instance(first_definition, &[NumberingLevelOverride::new(1).start(3)])
            .unwrap();

        let mut definition = document.numbering_definition(first_definition).unwrap();
        definition.levels[1].properties = definition.levels[1]
            .properties
            .clone()
            .start(7)
            .level_text("%1.%2)")
            .suffix(ListLevelSuffix::Space);
        document
            .update_numbering_definition(first_definition, &definition.levels)
            .unwrap();
        let mut instance = document.numbering_instance(instance_id).unwrap();
        instance.level_overrides[0].start = Some(9);
        document
            .update_numbering_instance(instance_id, second_definition, &instance.level_overrides)
            .unwrap();

        let bytes = document.to_bytes().unwrap();
        let mut reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(
            reopened
                .numbering_definition(first_definition)
                .unwrap()
                .levels[1]
                .properties
                .start,
            Some(7)
        );
        let instance = reopened.numbering_instance(instance_id).unwrap();
        assert_eq!(instance.definition_id, second_definition);
        assert_eq!(instance.level_overrides[0].start, Some(9));

        assert!(reopened.remove_numbering_instance(instance_id).unwrap());
        assert!(
            reopened
                .remove_numbering_definition(first_definition)
                .unwrap()
        );
        assert!(
            reopened
                .remove_numbering_definition(second_definition)
                .unwrap()
        );
        assert!(!reopened.remove_numbering_instance(instance_id).unwrap());
        let reopened = Document::from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert!(reopened.numbering_definitions().is_empty());
        assert!(reopened.numbering_instances().is_empty());
    }

    #[test]
    fn moving_a_live_numbering_instance_cannot_orphan_its_paragraph_level() {
        let mut document = Document::new();
        let complete = document
            .add_numbering_definition(&[
                ListLevel::decimal(),
                ListLevel::decimal(),
                ListLevel::decimal(),
            ])
            .unwrap();
        let incomplete = document
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance = document.add_numbering_instance(complete, &[]).unwrap();
        assert!(
            document
                .add_paragraph("deep item")
                .set_numbering(instance, 2)
        );
        let baseline = document.to_bytes().unwrap();

        assert!(
            document
                .update_numbering_instance(instance, incomplete, &[])
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);
        assert_eq!(
            document.numbering_instance(instance).unwrap().definition_id,
            complete
        );
    }

    #[test]
    fn moving_an_instance_cannot_orphan_a_style_inherited_paragraph_level() {
        let mut document = Document::new();
        let complete = document
            .add_numbering_definition(&[
                ListLevel::decimal(),
                ListLevel::decimal(),
                ListLevel::decimal(),
            ])
            .unwrap();
        let incomplete = document
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance = document.add_numbering_instance(complete, &[]).unwrap();
        document
            .add_style(StyleBuilder::paragraph("DeepListBase", "Deep List Base"))
            .unwrap();
        document
            .link_style_to_numbering("DeepListBase", instance, 2)
            .unwrap();
        document
            .add_style(
                StyleBuilder::paragraph("DeepListChild", "Deep List Child")
                    .based_on("DeepListBase"),
            )
            .unwrap();
        document.add_paragraph("deep item").style("DeepListChild");
        let baseline = document.to_bytes().unwrap();

        assert!(
            document
                .update_numbering_instance(instance, incomplete, &[])
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);
        assert_eq!(
            document.numbering_instance(instance).unwrap().definition_id,
            complete
        );
    }

    #[test]
    fn moving_an_instance_cannot_orphan_a_related_story_paragraph_level() {
        let mut document = Document::new();
        let complete = document
            .add_numbering_definition(&[
                ListLevel::decimal(),
                ListLevel::decimal(),
                ListLevel::decimal(),
            ])
            .unwrap();
        let incomplete = document
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance = document.add_numbering_instance(complete, &[]).unwrap();
        document.set_header("header");
        let (relationship_id, _) = document.header_footer_rel_ids().into_iter().next().unwrap();
        let part_name = {
            let relationship = document
                .package
                .get_part_rels(&document.doc_part_name)
                .unwrap()
                .get_by_id(&relationship_id)
                .unwrap();
            OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target)
        };
        document.package.set_part(
            &part_name,
            format!(
                r#"<n:hdr xmlns:n="{WORD_NAMESPACE}"><n:p><n:pPr><n:numPr><n:ilvl n:val="2"/><n:numId n:val="{instance}"/></n:numPr></n:pPr><n:r><n:t>header</n:t></n:r></n:p></n:hdr>"#
            )
            .into_bytes(),
        );
        let baseline = document.to_bytes().unwrap();

        assert!(
            document
                .update_numbering_instance(instance, incomplete, &[])
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);
        assert_eq!(
            document.numbering_instance(instance).unwrap().definition_id,
            complete
        );
    }

    #[test]
    fn paragraph_numbering_scan_does_not_pair_different_paragraphs() {
        let xml = format!(
            r#"<n:root xmlns:n="{WORD_NAMESPACE}"><n:p><n:pPr><n:numPr><n:ilvl n:val="2"/></n:numPr></n:pPr></n:p><n:p><n:pPr><n:numPr><n:numId n:val="7"/></n:numPr></n:pPr></n:p></n:root>"#
        );

        let properties = xml_paragraph_numbering_properties(xml.as_bytes()).unwrap();
        assert_eq!(properties.len(), 2);
        assert_eq!(properties[0].num_id, None);
        assert_eq!(properties[0].num_ilvl, Some(2));
        assert_eq!(properties[1].num_id, Some(7));
        assert_eq!(properties[1].num_ilvl, None);
    }

    #[test]
    fn opaque_custom_xml_numbering_data_is_not_a_live_story_reference() {
        const CUSTOM_XML_RELATIONSHIP: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
        const CUSTOM_XML_PART: &str = "/customXml/item1.xml";

        let mut document = Document::new();
        let complete = document
            .add_numbering_definition(&[
                ListLevel::decimal(),
                ListLevel::decimal(),
                ListLevel::decimal(),
            ])
            .unwrap();
        let incomplete = document
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance = document.add_numbering_instance(complete, &[]).unwrap();
        document.package.set_part(
            CUSTOM_XML_PART,
            format!(
                r#"<w:p xmlns:w="{WORD_NAMESPACE}"><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="{instance}"/></w:numPr></w:pPr></w:p>"#
            )
            .into_bytes(),
        );
        document
            .package
            .content_types
            .add_override(CUSTOM_XML_PART, "application/xml");
        document
            .package
            .get_or_create_part_rels(&document.doc_part_name)
            .add_with_id(
                "producerCustomXml",
                CUSTOM_XML_RELATIONSHIP,
                "../customXml/item1.xml",
            );

        document
            .update_numbering_instance(instance, incomplete, &[])
            .unwrap();
        assert_eq!(
            document.numbering_instance(instance).unwrap().definition_id,
            incomplete
        );
        assert!(document.remove_numbering_instance(instance).unwrap());
        let saved = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        assert!(saved.get_part(CUSTOM_XML_PART).is_some());
    }

    #[test]
    fn numbering_instance_removal_respects_style_references() {
        let mut document = Document::new();
        let definition = document
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance = document.add_numbering_instance(definition, &[]).unwrap();
        document
            .add_style(StyleBuilder::paragraph(
                "UnusedListStyle",
                "Unused List Style",
            ))
            .unwrap();
        document
            .link_style_to_numbering("UnusedListStyle", instance, 0)
            .unwrap();

        assert!(document.remove_numbering_instance(instance).is_err());
        assert!(document.numbering_instance(instance).is_some());
    }

    #[test]
    fn style_numbering_link_is_atomic_in_both_directions() {
        let mut document = Document::new();
        document
            .add_style(StyleBuilder::paragraph(
                "NumberedHeading",
                "Numbered Heading",
            ))
            .unwrap();
        let definition = document
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance = document.add_numbering_instance(definition, &[]).unwrap();
        let baseline = document.to_bytes().unwrap();

        assert!(
            document
                .add_style(
                    StyleBuilder::paragraph("OneSidedAdd", "One-sided add").paragraph_properties(
                        CT_PPr {
                            num_id: Some(instance),
                            num_ilvl: Some(0),
                            ..CT_PPr::default()
                        }
                    ),
                )
                .is_err()
        );
        assert!(document.style("OneSidedAdd").is_none());
        assert!(
            document
                .set_style(
                    StyleBuilder::paragraph("NumberedHeading", "Numbered Heading")
                        .paragraph_properties(CT_PPr {
                            num_id: Some(instance),
                            num_ilvl: Some(0),
                            ..CT_PPr::default()
                        }),
                )
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);

        assert!(
            document
                .link_style_to_numbering("MissingStyle", instance, 0)
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);
        assert!(
            document
                .link_style_to_numbering("NumberedHeading", instance, 8)
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);

        document
            .link_style_to_numbering("NumberedHeading", instance, 0)
            .unwrap();
        let style = document.style("NumberedHeading").unwrap();
        let properties = style.paragraph_properties().unwrap();
        assert_eq!(properties.num_id, Some(instance));
        assert_eq!(properties.num_ilvl, Some(0));
        assert_eq!(
            document
                .numbering_definition(definition)
                .unwrap()
                .paragraph_style_links,
            vec![Some("NumberedHeading".to_owned())]
        );

        let linked = document.to_bytes().unwrap();
        assert!(
            document
                .unlink_style_from_numbering("NumberedHeading", instance, 1)
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), linked);
        document
            .unlink_style_from_numbering("NumberedHeading", instance, 0)
            .unwrap();
        let style = document.style("NumberedHeading").unwrap();
        assert_eq!(style.paragraph_properties().unwrap().num_id, None);
        assert_eq!(style.paragraph_properties().unwrap().num_ilvl, None);
        assert_eq!(
            document
                .numbering_definition(definition)
                .unwrap()
                .paragraph_style_links,
            vec![None]
        );
    }

    #[test]
    fn numbering_graph_rejects_each_one_sided_style_link() {
        let mut document = Document::new();
        let normal = document
            .styles
            .styles
            .iter_mut()
            .find(|style| style.style_id == "Normal")
            .unwrap();
        normal.ppr = Some(CT_PPr {
            num_id: Some(99),
            ..CT_PPr::default()
        });
        assert!(document.validate_numbering_graph().is_err());
        document
            .styles
            .styles
            .iter_mut()
            .find(|style| style.style_id == "Normal")
            .unwrap()
            .ppr = None;

        let definition = document
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance = document.add_numbering_instance(definition, &[]).unwrap();
        document.numbering.as_mut().unwrap().abstract_nums[0].levels[0].p_style =
            Some("Normal".to_owned());
        assert!(document.validate_numbering_graph().is_err());
        let baseline = document.to_bytes().unwrap();
        assert!(
            document
                .link_style_to_numbering("Heading1", instance, 0)
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn style_numbering_link_overlays_extended_num_pr_payload() {
        let mut source = Document::new();
        source
            .add_style(StyleBuilder::paragraph("ExtLinked", "Extended Linked"))
            .unwrap();
        let definition = source
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance = source.add_numbering_instance(definition, &[]).unwrap();
        let mut package = OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
        let mut styles =
            String::from_utf8(package.get_part(DEFAULT_STYLES_PART).unwrap().to_vec()).unwrap();
        let style_start = styles.find(r#"w:styleId="ExtLinked""#).unwrap();
        let style_end = styles[style_start..].find("</w:style>").unwrap() + style_start;
        styles.insert_str(
            style_end,
            r#"<w:pPr><w:numPr xmlns:ext="urn:producer" ext:root="a&#x20;b"><ext:before/><w:ilvl ext:leaf="level"><ext:level-child/></w:ilvl><w:numId ext:leaf="id"><ext:id-child/></w:numId><ext:after/></w:numPr></w:pPr>"#,
        );
        package.set_part(DEFAULT_STYLES_PART, styles.into_bytes());
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let mut document = Document::from_bytes(bytes.get_ref()).unwrap();

        document
            .link_style_to_numbering("ExtLinked", instance, 0)
            .unwrap();
        let linked = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let linked =
            String::from_utf8(linked.get_part(DEFAULT_STYLES_PART).unwrap().to_vec()).unwrap();
        for retained in [
            r#"ext:root="a&#x20;b""#,
            "<ext:before",
            "<ext:level-child/>",
            "<ext:id-child/>",
            "<ext:after",
        ] {
            assert!(linked.contains(retained), "{linked}");
        }
        assert!(linked.contains(r#"w:val="0""#), "{linked}");
        assert!(
            linked.contains(&format!(r#"w:val="{instance}""#)),
            "{linked}"
        );

        document
            .unlink_style_from_numbering("ExtLinked", instance, 0)
            .unwrap();
        let unlinked = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let unlinked =
            String::from_utf8(unlinked.get_part(DEFAULT_STYLES_PART).unwrap().to_vec()).unwrap();
        let num_pr_start = unlinked.find("<w:numPr").unwrap();
        let num_pr_end = unlinked[num_pr_start..].find("</w:numPr>").unwrap()
            + num_pr_start
            + "</w:numPr>".len();
        assert_eq!(
            &unlinked[num_pr_start..num_pr_end],
            "<w:numPr xmlns:ext=\"urn:producer\" ext:root=\"a&#x20;b\"><ext:before/><w:ilvl ext:leaf=\"level\"><ext:level-child/></w:ilvl><w:numId ext:leaf=\"id\"><ext:id-child/></w:numId><ext:after/>\n      </w:numPr>"
        );
    }

    #[test]
    fn style_linked_numbering_survives_reopen_and_rebuild() {
        let mut document = Document::new();
        document
            .add_style(StyleBuilder::paragraph(
                "NumberedHeading",
                "Numbered Heading",
            ))
            .unwrap();
        let definition = document
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance = document.add_numbering_instance(definition, &[]).unwrap();
        document
            .link_style_to_numbering("NumberedHeading", instance, 0)
            .unwrap();
        document.add_paragraph("First").style("NumberedHeading");
        document
            .add_bookmark(
                "numbered_target",
                crate::comments::RunRange {
                    start: crate::comments::RunPosition {
                        body_index: 0,
                        run_index: 0,
                    },
                    end: crate::comments::RunPosition {
                        body_index: 0,
                        run_index: 1,
                    },
                },
            )
            .unwrap();
        let mut reference = CT_P::new();
        reference.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field(Field::new(
                r"REF numbered_target \w \t",
                "stale REF",
            ))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(reference));
        let mut package =
            OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let document_part = document.doc_part_name.clone();
        let xml = std::str::from_utf8(package.get_part(&document_part).unwrap()).unwrap();
        let toc = r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \t "Numbered Heading,1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p><w:r><w:t>stale TOC</w:t></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#;
        let xml = xml.replacen("<w:body>", &format!("<w:body>{toc}"), 1);
        package.set_part(&document_part, xml.into_bytes());
        let mut authored = Cursor::new(Vec::new());
        package.write_to(&mut authored).unwrap();
        document = Document::from_bytes(authored.get_ref()).unwrap();
        assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
        document
            .update_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert!(
            document
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()
                .iter()
                .any(|field| field.outcome == crate::FieldOutcome::Resolved("1".to_owned()))
        );

        let bytes = document.to_bytes().unwrap();
        let mut reopened = Document::from_bytes(&bytes).unwrap();
        let reopened_bytes = reopened.to_bytes().unwrap();
        if reopened_bytes != bytes {
            let original = OpcPackage::from_reader(Cursor::new(&bytes)).unwrap();
            let rebuilt = OpcPackage::from_reader(Cursor::new(&reopened_bytes)).unwrap();
            let mut changed = original
                .parts
                .iter()
                .filter_map(|(name, contents)| {
                    (rebuilt.parts.get(name) != Some(contents)).then_some(name.as_str())
                })
                .collect::<Vec<_>>();
            changed.sort_unstable();
            let extract_style = |package: &OpcPackage| {
                let styles =
                    std::str::from_utf8(package.get_part(DEFAULT_STYLES_PART).unwrap()).unwrap();
                let start = styles.find(r#"w:styleId="NumberedHeading""#).unwrap();
                let start = styles[..start].rfind("<w:style").unwrap();
                let end = styles[start..].find("</w:style>").unwrap() + start + "</w:style>".len();
                styles[start..end].to_owned()
            };
            panic!(
                "reopened package changed parts: {changed:?}\noriginal: {}\nrebuilt: {}",
                extract_style(&original),
                extract_style(&rebuilt)
            );
        }
        let style = reopened.style("NumberedHeading").unwrap();
        let properties = style.paragraph_properties().unwrap();
        assert_eq!(
            (properties.num_id, properties.num_ilvl),
            (Some(instance), Some(0))
        );
        assert_eq!(
            reopened
                .numbering_definition(definition)
                .unwrap()
                .paragraph_style_links,
            vec![Some("NumberedHeading".to_owned())]
        );
        assert_eq!(reopened.rebuild_toc().unwrap().entry_count, 1);
        reopened
            .update_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn style_linking_accepts_three_distinct_levels() {
        let mut document = Document::new();
        for (style_id, name) in [
            ("NumberedHeading1", "Numbered Heading 1"),
            ("NumberedHeading2", "Numbered Heading 2"),
            ("NumberedHeading3", "Numbered Heading 3"),
        ] {
            document
                .add_style(StyleBuilder::paragraph(style_id, name))
                .unwrap();
        }
        let definition = document
            .add_numbering_definition(&[
                ListLevel::decimal().level_text("%1."),
                ListLevel::decimal().level_text("%1.%2."),
                ListLevel::decimal().level_text("%1.%2.%3."),
            ])
            .unwrap();
        let first = document.add_numbering_instance(definition, &[]).unwrap();
        let second = document.add_numbering_instance(definition, &[]).unwrap();
        for (style_id, level) in [
            ("NumberedHeading1", 0),
            ("NumberedHeading2", 1),
            ("NumberedHeading3", 2),
        ] {
            document
                .link_style_to_numbering(style_id, first, level)
                .unwrap();
        }

        assert_ne!(first, second);
        assert_eq!(
            document
                .numbering_definition(definition)
                .unwrap()
                .paragraph_style_links,
            vec![
                Some("NumberedHeading1".to_owned()),
                Some("NumberedHeading2".to_owned()),
                Some("NumberedHeading3".to_owned()),
            ]
        );
    }

    #[test]
    fn fresh_level_attributes_conflicting_with_retained_aliases_are_atomic() {
        let mut source = Document::new();
        let definition_id = source
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let xml = format!(
            r#"<n:numbering xmlns:n="{WORD_NAMESPACE}"><n:abstractNum n:abstractNumId="{definition_id}"><n:lvl n:ilvl="0" n:tplc="bad" n:tentative="maybe"><n:numFmt n:val="decimal"/></n:lvl></n:abstractNum></n:numbering>"#
        );
        let imported = replace_numbering_xml(&mut source, xml.into_bytes());
        let mut document = Document::from_bytes(&imported).unwrap();
        let baseline = document.to_bytes().unwrap();
        let mut definition = document.numbering_definition(definition_id).unwrap();

        definition.levels[0].properties = ListLevel::decimal().template_code("A1B2C3D4");
        assert!(
            document
                .update_numbering_definition(definition_id, &definition.levels)
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);

        definition.levels[0].properties = ListLevel::decimal().tentative(true);
        assert!(
            document
                .update_numbering_definition(definition_id, &definition.levels)
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);
        let unchanged = document.numbering_definition(definition_id).unwrap();
        assert_eq!(unchanged.levels[0].properties.template_code_value(), None);
        assert_eq!(unchanged.levels[0].properties.tentative_value(), None);
    }

    #[test]
    fn sparse_numbering_definition_levels_keep_identifiers_through_update() {
        let mut source = Document::new();
        let definition_id = source
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let xml = format!(
            r#"<w:numbering xmlns:w="{WORD_NAMESPACE}"><w:abstractNum w:abstractNumId="{definition_id}"><w:lvl w:ilvl="2"><w:numFmt w:val="decimal"/><w:lvlRestart w:val="1"/><w:lvlText w:val="%3."/></w:lvl></w:abstractNum></w:numbering>"#
        );
        let imported = replace_numbering_xml(&mut source, xml.into_bytes());
        let mut document = Document::from_bytes(&imported).unwrap();
        let canonical = document.to_bytes().unwrap();
        let definition = document.numbering_definition(definition_id).unwrap();
        assert_eq!(definition.levels[0].level, 2);
        assert_eq!(
            definition.levels[0].properties.level_text_value(),
            Some("%3.")
        );
        document
            .update_numbering_definition(definition_id, &definition.levels)
            .unwrap();
        assert_eq!(document.to_bytes().unwrap(), canonical);

        let replacement = [NumberingDefinitionLevel {
            level: 2,
            properties: ListLevel::decimal()
                .level_text("%3)")
                .restart(ListLevelRestart::After(0)),
        }];
        document
            .update_numbering_definition(definition_id, &replacement)
            .unwrap();
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(Cursor::new(&saved)).unwrap();
        let numbering =
            String::from_utf8(package.get_part(DEFAULT_NUMBERING_PART).unwrap().to_vec()).unwrap();
        assert!(numbering.contains(r#"<w:lvl w:ilvl="2""#), "{numbering}");
        assert!(numbering.contains(r#"w:val="%3)""#), "{numbering}");
    }

    #[test]
    fn invalid_numbering_mutation_is_atomic() {
        let mut document = Document::new();
        let definition_id = document
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance_id = document.add_numbering_instance(definition_id, &[]).unwrap();
        assert!(document.add_paragraph("item").set_numbering(instance_id, 0));
        let baseline = document.to_bytes().unwrap();

        for invalid in [
            ListLevel::decimal().level_text("%2."),
            ListLevel::decimal().restart(ListLevelRestart::After(0)),
            ListLevel::decimal().template_code("xyz"),
        ] {
            assert!(document.add_numbering_definition(&[invalid]).is_err());
            assert_eq!(document.to_bytes().unwrap(), baseline);
        }
        assert!(document.add_numbering_instance(u32::MAX, &[]).is_err());
        assert_eq!(document.to_bytes().unwrap(), baseline);
        assert!(
            document
                .add_numbering_instance(definition_id, &[NumberingLevelOverride::new(9)])
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);
        assert!(
            document
                .add_numbering_instance(definition_id, &[NumberingLevelOverride::new(1)])
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);
        assert!(
            document
                .add_numbering_instance(
                    definition_id,
                    &[
                        NumberingLevelOverride::new(0),
                        NumberingLevelOverride::new(0),
                    ],
                )
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), baseline);
        assert!(document.remove_numbering_instance(instance_id).is_err());
        assert_eq!(document.to_bytes().unwrap(), baseline);
        assert!(document.remove_numbering_definition(definition_id).is_err());
        assert_eq!(document.to_bytes().unwrap(), baseline);

        document.numbering.as_mut().unwrap().abstract_nums[0].levels[0].p_style =
            Some("MissingStyle".to_owned());
        let invalid_graph = document.to_bytes().unwrap();
        assert!(
            document
                .add_numbering_definition(&[ListLevel::decimal()])
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), invalid_graph);
        document.numbering.as_mut().unwrap().abstract_nums[0].levels[0].p_style = None;

        let next_definition = document
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        assert_eq!(next_definition, definition_id + 1);
        let next_instance = document
            .add_numbering_instance(next_definition, &[])
            .unwrap();
        assert_eq!(next_instance, instance_id + 1);
    }

    #[test]
    fn imported_numbering_extensions_remain_byte_identical() {
        let mut document = Document::new();
        let definition_id = document
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance_id = document
            .add_numbering_instance(definition_id, &[NumberingLevelOverride::new(0).start(4)])
            .unwrap();
        document
            .styles
            .styles
            .iter_mut()
            .find(|style| style.style_id == "Normal")
            .unwrap()
            .ppr = Some(CT_PPr {
            num_id: Some(instance_id),
            num_ilvl: Some(0),
            ..CT_PPr::default()
        });
        let numbering = document.numbering.as_mut().unwrap();
        numbering
            .root_attributes
            .push(("xmlns:ext".to_owned(), "urn:producer".to_owned()));
        let definition = numbering
            .abstract_nums
            .iter_mut()
            .find(|value| value.abstract_num_id == definition_id)
            .unwrap();
        definition
            .extra_xml
            .push((7, b"<ext:definition/>".to_vec()));
        definition.levels[0].p_style = Some("Normal".to_owned());
        definition.levels[0]
            .extra_xml
            .push((7, b"<ext:level value=\"kept\"/>".to_vec()));
        let instance = numbering
            .nums
            .iter_mut()
            .find(|value| value.num_id == instance_id)
            .unwrap();
        instance
            .extra_xml
            .push((1, b"<ext:instance value=\"kept\"/>".to_vec()));
        instance.level_overrides[0]
            .extra_xml
            .push((1, b"<ext:override value=\"kept\"/>".to_vec()));
        let before = numbering.to_xml().unwrap();

        let definition = document.numbering_definition(definition_id).unwrap();
        document
            .update_numbering_definition(definition_id, &definition.levels)
            .unwrap();
        let instance = document.numbering_instance(instance_id).unwrap();
        document
            .update_numbering_instance(
                instance_id,
                instance.definition_id,
                &instance.level_overrides,
            )
            .unwrap();
        let after = document.numbering.as_ref().unwrap().to_xml().unwrap();
        assert_eq!(after, before);
        assert_eq!(
            document
                .numbering_definition(definition_id)
                .unwrap()
                .paragraph_style_links,
            vec![Some("Normal".to_owned())]
        );

        let mut source = Document::new();
        let definition_id = source
            .add_numbering_definition(&[ListLevel::decimal()])
            .unwrap();
        let instance_id = source.add_numbering_instance(definition_id, &[]).unwrap();
        source
            .styles
            .styles
            .iter_mut()
            .find(|style| style.style_id == "Normal")
            .unwrap()
            .ppr = Some(CT_PPr {
            num_id: Some(instance_id),
            num_ilvl: Some(0),
            ..CT_PPr::default()
        });
        let xml = format!(
            r#"<n:numbering xmlns:n="{WORD_NAMESPACE}" xmlns:ext="urn:producer"><ext:root value="kept"/><n:abstractNum n:abstractNumId="{definition_id}" ext:definition="kept"><ext:definition-child/><n:multiLevelType n:val="hybridMultilevel" ext:leaf="type"><ext:type-child/></n:multiLevelType><n:lvl n:ilvl="0"><n:start n:val="1" ext:leaf="start"><ext:start-child/></n:start><n:numFmt n:val="producerFormat" ext:leaf="format"><ext:format-child/></n:numFmt><n:lvlRestart n:val="0" ext:restart="kept"></n:lvlRestart><n:pStyle n:val="Normal"></n:pStyle><n:isLgl n:val="false"/><n:suff n:val="space" ext:leaf="suffix"><ext:suffix-child/></n:suff><ext:level-child/><n:lvlText n:val="%1." ext:leaf="text"><ext:text-child/></n:lvlText><n:lvlJc n:val="left" ext:leaf="alignment"><ext:alignment-child/></n:lvlJc></n:lvl></n:abstractNum><n:num n:numId="{instance_id}" ext:instance="kept"><n:abstractNumId n:val="{definition_id}" ext:reference="kept"><ext:reference-child/></n:abstractNumId><ext:instance-child/><n:lvlOverride n:ilvl="0" ext:override="kept"><n:startOverride n:val="4" ext:start="kept"></n:startOverride><ext:override-child/><n:lvl n:ilvl="0"><n:numFmt n:val="decimal"/><n:pStyle n:val="Normal"/><n:lvlText n:val="%1)"/><ext:replacement-child/></n:lvl></n:lvlOverride></n:num></n:numbering>"#
        );
        let imported = replace_numbering_xml(&mut source, xml.into_bytes());
        let mut imported = Document::from_bytes(&imported).unwrap();
        let canonical_bytes = imported.to_bytes().unwrap();
        let canonical = OpcPackage::from_reader(Cursor::new(&canonical_bytes))
            .unwrap()
            .get_part(DEFAULT_NUMBERING_PART)
            .unwrap()
            .to_vec();
        let canonical_xml = String::from_utf8(canonical.clone()).unwrap();
        assert!(canonical_xml.contains("<ext:type-child/>"));
        assert!(canonical_xml.contains("<n:lvlOverride"));
        assert!(canonical_xml.contains("<n:startOverride"));
        assert!(canonical_xml.contains("<ext:reference-child/>"));
        for payload in [
            "<ext:start-child/>",
            "<ext:format-child/>",
            "<ext:suffix-child/>",
            "<ext:text-child/>",
            "<ext:alignment-child/>",
        ] {
            assert!(canonical_xml.contains(payload), "{canonical_xml}");
        }
        let abstract_reference = canonical_xml.find("<n:abstractNumId ").unwrap();
        let override_position = canonical_xml.find("<n:lvlOverride ").unwrap();
        let start_position = canonical_xml.find("<n:startOverride ").unwrap();
        let replacement_position = canonical_xml[override_position..]
            .find("<n:lvl n:ilvl=\"0\"")
            .unwrap()
            + override_position;
        assert!(abstract_reference < override_position);
        assert!(start_position < replacement_position);

        let definition = imported.numbering_definition(definition_id).unwrap();
        assert_eq!(
            definition.paragraph_style_links,
            vec![Some("Normal".to_owned())]
        );
        assert!(definition.has_unmodeled_properties);
        imported
            .update_numbering_definition(definition_id, &definition.levels)
            .unwrap();
        let instance = imported.numbering_instance(instance_id).unwrap();
        assert!(instance.has_unmodeled_properties);
        imported
            .update_numbering_instance(
                instance_id,
                instance.definition_id,
                &instance.level_overrides,
            )
            .unwrap();
        let after = OpcPackage::from_reader(Cursor::new(imported.to_bytes().unwrap()))
            .unwrap()
            .get_part(DEFAULT_NUMBERING_PART)
            .unwrap()
            .to_vec();
        assert_eq!(after, canonical);
    }

    #[test]
    fn rejected_list_level_update_does_not_materialize_numbering() {
        let mut doc = Document::new();
        assert!(doc.numbering.is_none());

        assert!(!doc.set_list_level(999, 1, ListLevel::decimal()));

        assert!(
            doc.numbering.is_none(),
            "a rejected setter must not add an empty numbering part"
        );
    }

    #[test]
    fn custom_list_and_paragraph_numbering_enforce_the_nine_level_contract() {
        let mut doc = Document::new();
        let levels = vec![ListLevel::decimal(); 10];
        let num_id = doc.add_list_definition(&levels);
        assert_eq!(
            doc.numbering.as_ref().unwrap().abstract_nums[0]
                .levels
                .len(),
            9
        );

        let mut paragraph = doc.add_paragraph("item");
        assert!(!paragraph.set_numbering(num_id, 9));
        assert_eq!(
            paragraph.inner.properties.as_ref().and_then(|p| p.num_id),
            None
        );
        assert!(paragraph.set_numbering(num_id, 8));
        assert_eq!(
            paragraph.inner.properties.as_ref().unwrap().num_ilvl,
            Some(8)
        );
    }

    #[test]
    fn reader_exposes_complete_numbering_level_facts() {
        let mut doc = Document::new();
        let num_id = doc.add_list_definition(&[ListLevel::decimal()]);
        let none_id = doc.add_list_definition(&[ListLevel::decimal()]);
        let level = &mut doc.numbering.as_mut().unwrap().abstract_nums[0].levels[0];
        level.num_fmt = Some(ST_NumberFormat::Other("producerFormat".to_owned()));
        level.start = Some(4);
        level.suffix = Some(ST_LvlSuffix::Space);
        level.lvl_jc = Some(ST_Jc::Center);
        level.p_style = Some("Normal".to_owned());

        let level = doc.numbering_level(num_id, 0).expect("numbering level");
        assert_eq!(level.format, NumberingFormat::Other("producerFormat"));
        assert_eq!(level.format_name, "producerFormat");
        assert_eq!(level.start, 4);
        assert_eq!(level.suffix, ListLevelSuffix::Space);
        assert_eq!(level.alignment, Some(Alignment::Center));
        assert_eq!(level.paragraph_style, Some("Normal"));
        assert!(!level.has_unmodeled_properties);
        assert!(!level.has_paragraph_presentation);
        assert!(!level.has_marker_presentation);

        let level = &mut doc.numbering.as_mut().unwrap().abstract_nums[0].levels[0];
        level.ppr = Some(CT_PPr {
            ind_left: Some(Twips(360)),
            ..Default::default()
        });
        level.rpr = Some(CT_RPr {
            bold: Some(true),
            ..Default::default()
        });
        level.extra_xml.push((0, b"<w:unknown/>".to_vec()));

        let level = doc.numbering_level(num_id, 0).expect("numbering level");
        assert!(level.has_unmodeled_properties);
        assert!(level.has_paragraph_presentation);
        assert!(level.has_marker_presentation);

        doc.numbering.as_mut().unwrap().abstract_nums[1].levels[0].num_fmt =
            Some(ST_NumberFormat::None);
        assert_eq!(
            doc.numbering_level(none_id, 0)
                .expect("no-numbering level")
                .format,
            NumberingFormat::None
        );
    }

    #[test]
    fn numbering_level_distinguishes_modeled_metadata_from_retained_raw_facts() {
        let mut doc = Document::new();
        let definition_metadata_id = doc.add_list_definition(&[ListLevel::decimal()]);
        let raw_level_properties_id = doc.add_list_definition(&[ListLevel::decimal()]);
        let raw_leaf_metadata_id = doc.add_list_definition(&[ListLevel::decimal()]);

        let definition = &mut doc.numbering.as_mut().unwrap().abstract_nums[0];
        definition.nsid = Some("12345678".to_owned());
        definition.tmpl = Some("87654321".to_owned());
        definition.multi_level_type = Some("hybridMultilevel".to_owned());

        let level = &mut doc.numbering.as_mut().unwrap().abstract_nums[1].levels[0];
        level.ppr_raw = Some((
            CT_PPr::default(),
            b"<w:pPr><producer:property/></w:pPr>".to_vec(),
            vec!["producer".to_owned()],
        ));
        level.rpr_raw = Some((
            CT_RPr::default(),
            b"<w:rPr><producer:property/></w:rPr>".to_vec(),
            vec!["producer".to_owned()],
        ));

        let definition = &mut doc.numbering.as_mut().unwrap().abstract_nums[2];
        definition.levels[0].start = Some(1);
        definition.levels[0].start_raw = Some((
            Some(1),
            b"<w:start w:val=\"1\"><producer:fact/></w:start>".to_vec(),
            vec!["w".to_owned(), "producer".to_owned()],
        ));

        assert!(
            !doc.numbering_level(definition_metadata_id, 0)
                .expect("definition metadata level")
                .has_unmodeled_properties
        );
        assert!(
            doc.numbering_level(raw_level_properties_id, 0)
                .expect("raw properties level")
                .has_unmodeled_properties
        );
        assert!(
            doc.numbering_level(raw_leaf_metadata_id, 0)
                .expect("raw leaf metadata level")
                .has_unmodeled_properties
        );
        assert!(
            doc.numbering_definition(
                doc.numbering_instance(raw_leaf_metadata_id)
                    .expect("raw leaf metadata instance")
                    .definition_id,
            )
            .expect("raw leaf metadata definition")
            .has_unmodeled_properties
        );
    }

    #[test]
    fn extreme_numbering_level_reader_fact_does_not_overflow() {
        let mut doc = Document::new();
        let num_id = doc.add_list_definition(&[ListLevel::decimal()]);
        let level = &mut doc.numbering.as_mut().unwrap().abstract_nums[0].levels[0];
        level.ilvl = u32::MAX;
        level.ppr = Some(CT_PPr::default());

        let level = doc
            .numbering_level(num_id, u32::MAX)
            .expect("extreme level remains readable");
        assert!(level.has_paragraph_presentation);
    }

    #[test]
    fn absent_or_unstyled_paragraph_properties_use_default_style_numbering() {
        let mut doc = Document::new();
        let num_id = doc.add_list_definition(&[ListLevel::decimal()]);
        let level = &mut doc.numbering.as_mut().unwrap().abstract_nums[0].levels[0];
        level.p_style = Some("Normal".to_owned());
        level.ppr = Some(CT_PPr {
            keep_next: Some(true),
            ..Default::default()
        });
        let default_style = doc
            .styles
            .styles
            .iter_mut()
            .find(|style| style.style_id == "Normal")
            .expect("default paragraph style");
        default_style.ppr = Some(CT_PPr {
            num_id: Some(num_id),
            ..Default::default()
        });

        doc.add_paragraph("absent properties");
        doc.add_paragraph("unstyled properties");
        let BodyContent::Paragraph(paragraph) = &mut doc.document.body.content[1] else {
            panic!("expected paragraph");
        };
        paragraph.properties = Some(CT_PPr {
            keep_lines: Some(true),
            ..Default::default()
        });

        for index in 0..2 {
            let paragraph = doc.paragraph(index).expect("paragraph");
            let effective = doc.effective_paragraph_properties(&paragraph);
            assert_eq!(effective.num_id, Some(num_id));
            assert_eq!(effective.num_ilvl, Some(0));
            assert_eq!(effective.keep_next, Some(true));
        }
    }

    #[test]
    fn reader_reports_list_paragraph_presentation_separately_from_unknown_xml() {
        let mut doc = Document::new();
        let num_id = doc.add_list_definition(&[ListLevel::decimal()]);
        let level = &mut doc.numbering.as_mut().unwrap().abstract_nums[0].levels[0];
        level.ppr = Some(CT_PPr {
            ind_left: Some(Twips(360)),
            ..CT_PPr::default()
        });

        let level = doc.numbering_level(num_id, 0).expect("numbering level");
        assert!(level.has_paragraph_presentation);
        assert!(!level.has_unmodeled_properties);
    }

    #[test]
    fn reader_exposes_background_and_section_completeness_facts() {
        let mut doc = Document::new();
        assert!(!doc.has_document_background());
        assert!(doc.has_section_layout_formatting());
        assert!(!doc.has_unmodeled_section_properties());

        doc.document.background_xml = Some(b"<w:background/>".to_vec());
        doc.document
            .body
            .sect_pr
            .as_mut()
            .unwrap()
            .extra_xml
            .push(b"<w:printerSettings r:id=\"rId1\"/>".to_vec());

        assert!(doc.has_document_background());
        assert!(doc.has_unmodeled_section_properties());
    }

    #[test]
    fn reader_resolves_concrete_paragraph_and_run_properties() {
        let mut doc = Document::new();
        let num_id = doc.add_list_definition(&[ListLevel::decimal()]);
        let level = &mut doc.numbering.as_mut().unwrap().abstract_nums[0].levels[0];
        level.ppr = Some(CT_PPr {
            keep_next: Some(true),
            ..Default::default()
        });
        level.rpr = Some(CT_RPr {
            bold: Some(true),
            ..Default::default()
        });

        doc.add_style(
            StyleBuilder::paragraph("ListBase", "List Base").run_properties(CT_RPr {
                italic: Some(true),
                ..Default::default()
            }),
        )
        .unwrap();
        doc.link_style_to_numbering("ListBase", num_id, 0).unwrap();
        doc.add_style(StyleBuilder::paragraph("ListChild", "List Child").based_on("ListBase"))
            .unwrap();
        doc.add_paragraph("text").style("ListChild");

        let BodyContent::Paragraph(paragraph) = &mut doc.document.body.content[0] else {
            panic!("expected paragraph");
        };
        let properties = paragraph.properties.as_mut().expect("paragraph style");
        properties.ind_left = Some(Twips(720));
        properties.rpr = Some(CT_RPr {
            strike: Some(true),
            ..Default::default()
        });
        paragraph.runs[0].properties = Some(CT_RPr {
            vanish: Some(true),
            ..Default::default()
        });

        let paragraph = doc.paragraph(0).expect("paragraph");
        let run = paragraph.runs().next().expect("run");
        let paragraph_properties = doc.effective_paragraph_properties(&paragraph);
        let run_properties = doc.effective_run_properties(&paragraph, &run);

        assert_eq!(paragraph_properties.num_id, Some(num_id));
        assert_eq!(paragraph_properties.num_ilvl, Some(0));
        assert_eq!(paragraph_properties.keep_next, Some(true));
        assert_eq!(paragraph_properties.ind_left, Some(Twips(720)));
        assert_eq!(run_properties.italic, Some(true));
        assert_eq!(run_properties.bold, None);
        assert_eq!(run_properties.strike, Some(true));
        assert_eq!(run_properties.vanish, Some(true));
        assert!(
            doc.numbering_level(num_id, 0)
                .expect("numbering level")
                .has_marker_presentation
        );
    }

    #[test]
    fn effective_paragraph_properties_use_the_final_numbering_identity() {
        let mut doc = Document::new();
        let inherited_num_id = doc.add_list_definition(&[ListLevel::decimal()]);
        let direct_num_id = doc.add_list_definition(&[ListLevel::decimal()]);
        doc.numbering.as_mut().unwrap().abstract_nums[0].levels[0].ppr = Some(CT_PPr {
            keep_next: Some(true),
            ..Default::default()
        });
        doc.numbering.as_mut().unwrap().abstract_nums[1].levels[0].ppr = Some(CT_PPr {
            keep_lines: Some(true),
            ..Default::default()
        });
        doc.add_style(StyleBuilder::paragraph("InheritedList", "Inherited List"))
            .unwrap();
        doc.link_style_to_numbering("InheritedList", inherited_num_id, 0)
            .unwrap();

        doc.add_paragraph("direct").numbering(direct_num_id, 0);
        doc.add_paragraph("override")
            .style("InheritedList")
            .numbering(direct_num_id, 0);
        doc.add_paragraph("disabled")
            .style("InheritedList")
            .numbering(0, 0);

        let direct =
            doc.effective_paragraph_properties(&doc.paragraph(0).expect("direct paragraph"));
        assert_eq!(direct.num_id, Some(direct_num_id));
        assert_eq!(direct.keep_lines, Some(true));

        let overridden =
            doc.effective_paragraph_properties(&doc.paragraph(1).expect("overridden paragraph"));
        assert_eq!(overridden.num_id, Some(direct_num_id));
        assert_eq!(overridden.keep_next, None);
        assert_eq!(overridden.keep_lines, Some(true));

        let disabled =
            doc.effective_paragraph_properties(&doc.paragraph(2).expect("disabled paragraph"));
        assert_eq!(disabled.num_id, Some(0));
        assert_eq!(disabled.keep_next, None);
        assert_eq!(disabled.keep_lines, None);
    }

    #[test]
    fn picture_round_trips() {
        // 1x1 red PNG
        let png: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x00, 0x03, 0x00, 0x01, 0x9E, 0xDD, 0x22,
            0x71, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];
        let mut doc = Document::new();
        doc.add_paragraph("before");
        doc.add_picture(png, "dot.png", Length::inches(1.0), Length::inches(1.0));

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();
        let paras = doc2.paragraphs();
        let mut found = None;
        for p in &paras {
            for r in p.runs() {
                if let Some((rel, _alt)) = r.inline_image() {
                    found = Some(rel.to_string());
                }
            }
        }
        let rel = found.expect("no inline image found on read");
        let data = doc2.image_data(&rel).expect("image bytes missing");
        assert_eq!(data, png);
    }

    #[test]
    fn layout_resolves_relationship_images_to_shared_media() {
        let png: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x00, 0x03, 0x00, 0x01, 0x9E, 0xDD, 0x22,
            0x71, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];
        let mut document = Document::new();
        document.add_picture(png, "first.png", Length::inches(1.0), Length::inches(1.0));
        document.add_picture(png, "second.png", Length::inches(1.0), Length::inches(1.0));

        let page = document
            .layout_page(0)
            .expect("layout should succeed")
            .expect("document should have a first page");
        let images = compatibility_page_elements(&page.elements)
            .into_iter()
            .filter_map(|element| match element {
                oxml_layout::PositionedElement::Image {
                    data,
                    content_type,
                    media_id,
                    ..
                } => Some((data, content_type, media_id)),
                _ => None,
            })
            .collect::<Vec<_>>();

        assert_eq!(images.len(), 2);
        assert!(images.iter().all(|(data, _, _)| data.as_slice() == png));
        assert!(
            images
                .iter()
                .all(|(_, content_type, _)| *content_type == "image/png")
        );
        assert_eq!(*images[0].2, oxml_layout::MediaId::from_bytes(png));
        assert_eq!(images[0].2, images[1].2);
    }
}

#[cfg(test)]
mod hyperlink_span_tests {
    use super::*;
    use rdocx_oxml::text::HyperlinkSpan;

    /// `HyperlinkSpan`'s bounds are public, so a caller building the OXML model
    /// by hand can hand us a range past the end of `runs`. `links()` used to
    /// slice with it and panic.
    #[test]
    fn links_clamps_out_of_range_spans() {
        let mut doc = Document::new();
        {
            let mut para = doc.add_paragraph("");
            para.add_run("one");
            para.add_run("two");
        }

        let BodyContent::Paragraph(p) = &mut doc.document.body.content[0] else {
            unreachable!("just added a paragraph")
        };
        p.hyperlinks.push(HyperlinkSpan {
            rel_id: None,
            anchor: Some("bookmark".to_string()),
            tooltip: None,
            doc_location: None,
            run_start: 1,
            run_end: 99,
            extra_attributes: Vec::new(),
            extra_xml: Vec::new(),
            preserved_raw_before: None,
        });
        p.hyperlinks.push(HyperlinkSpan {
            rel_id: None,
            anchor: Some("inverted".to_string()),
            tooltip: None,
            doc_location: None,
            run_start: 5,
            run_end: 1,
            extra_attributes: Vec::new(),
            extra_xml: Vec::new(),
            preserved_raw_before: None,
        });

        let links = doc.links();

        assert_eq!(links.len(), 2);
        assert_eq!(links[0].text, "two");
        assert_eq!(links[1].text, "");
    }
}

#[cfg(test)]
mod watermark_tests {
    use super::*;
    use std::io::Cursor;

    pub(super) const PNG: &[u8] = &[
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
        0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];

    fn text_vml_header(text: &str, color: &str) -> Vec<u8> {
        format!(
            r#"<w:hdr xmlns:w="{}" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:pict><v:shape style="width:468pt;height:117pt;rotation:315" fillcolor="{color}"><v:fill opacity=".5"/><v:textpath string="{text}" style="font-family:&quot;Calibri&quot;"/></v:shape></w:pict></w:r></w:p></w:hdr>"#,
            rdocx_oxml::namespace::W_NS
        )
        .into_bytes()
    }

    fn page_text(layout: &oxml_layout::LayoutResult, index: usize) -> String {
        let mut text = String::new();
        oxml_layout::walk(&layout.pages[index].elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Text(run) = element {
                text.push_str(&run.text);
            }
        });
        text
    }

    fn compatibility_page_elements(
        elements: &[oxml_layout::PositionedElement],
    ) -> Vec<&oxml_layout::PositionedElement> {
        fn collect<'a>(
            elements: &'a [oxml_layout::PositionedElement],
            output: &mut Vec<&'a oxml_layout::PositionedElement>,
        ) {
            for element in elements {
                match element {
                    oxml_layout::PositionedElement::MarkedContent { children, .. } => {
                        collect(children, output);
                    }
                    other => output.push(other),
                }
            }
        }

        let mut output = Vec::new();
        collect(elements, &mut output);
        output
    }

    fn enable_even_headers(document: &mut Document) {
        set_even_headers_value(document, None);
    }

    fn set_even_headers_value(document: &mut Document, value: Option<&str>) {
        let part_name = "/word/settings.xml";
        let setting = value.map_or_else(
            || "<w:evenAndOddHeaders/>".to_owned(),
            |value| format!(r#"<w:evenAndOddHeaders w:val="{value}"/>"#),
        );
        let xml = format!(
            r#"<w:settings xmlns:w="{}">{setting}</w:settings>"#,
            rdocx_oxml::namespace::W_NS
        )
        .into_bytes();
        document.settings = Some(CT_Settings::from_xml(&xml).unwrap());
        document.settings_part_name = Some(part_name.to_owned());
        document.package.set_part(part_name, xml);
        document
            .ensure_part_relationship_checked(
                part_name,
                rel_types::SETTINGS,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
            )
            .unwrap();
    }

    fn assert_watermark_mutation_preserves_reusable_engine(
        mut document: Document,
        mutate: impl FnOnce(&mut Document),
    ) {
        let layout_state = |document: &Document| {
            (
                document
                    .layout_cache
                    .lock()
                    .unwrap_or_else(std::sync::PoisonError::into_inner)
                    .is_some(),
                document
                    .normal_layout_engine
                    .lock()
                    .unwrap_or_else(std::sync::PoisonError::into_inner)
                    .is_some(),
                document
                    .deterministic_layout_cache
                    .lock()
                    .unwrap_or_else(std::sync::PoisonError::into_inner)
                    .is_some(),
            )
        };
        document.add_paragraph("unchanged cacheable body paragraph");

        LAYOUT_INVOCATIONS.set(0);
        let accepted_before = document.layout().expect("populate normal layout cache");
        document
            .to_pdf_deterministic()
            .expect("populate deterministic layout cache");
        assert_eq!(LAYOUT_INVOCATIONS.get(), 2);
        assert_eq!(layout_state(&document), (true, true, true));

        mutate(&mut document);

        assert_eq!(
            layout_state(&document),
            (false, true, false),
            "completed results are invalidated while reusable work survives"
        );

        let accepted_after = document
            .layout()
            .expect("relayout after watermark mutation");
        document
            .to_pdf_deterministic()
            .expect("deterministic relayout after watermark mutation");
        assert!(!Arc::ptr_eq(&accepted_before, &accepted_after));
        assert_eq!(LAYOUT_INVOCATIONS.get(), 4);
    }

    #[test]
    fn watermark_mutations_preserve_reusable_engine_and_invalidate_completed_layouts() {
        assert_watermark_mutation_preserves_reusable_engine(Document::new(), |document| {
            document.set_text_watermark("DRAFT").unwrap();
        });
        assert_watermark_mutation_preserves_reusable_engine(Document::new(), |document| {
            document
                .set_image_watermark(PNG, "watermark.png", Length::pt(72.0), Length::pt(36.0))
                .unwrap();
        });
    }

    #[test]
    fn text_and_image_watermarks_round_trip_through_header_relationships() {
        let mut text = Document::new();
        text.set_header("default header");
        text.set_first_page_header("first header");
        text.set_raw_header_with_images(
            format!(
                r#"<w:hdr xmlns:w="{}"><w:p><w:r><w:t>even header</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
            &[],
            HdrFtrType::Even,
        );
        enable_even_headers(&mut text);
        text.set_text_watermark("DRAFT").unwrap();
        text.set_text_watermark("FINAL").unwrap();
        let reopened = Document::from_bytes(&text.to_bytes().unwrap()).unwrap();
        let input = reopened.build_layout_input();
        assert_eq!(input.headers.len(), 3);
        assert!(input.headers.values().all(|header| {
            matches!(
                header.watermarks(),
                [rdocx_oxml::header_footer::VmlWatermark::Text { text, .. }]
                    if text == "FINAL"
            )
        }));

        let mut image = Document::new();
        image
            .set_image_watermark(PNG, "watermark.png", Length::pt(72.0), Length::pt(36.0))
            .unwrap();
        let reopened = Document::from_bytes(&image.to_bytes().unwrap()).unwrap();
        assert!(!reopened.build_layout_input().headers.is_empty());

        let mut invalid = Document::new();
        invalid.set_header("unchanged");
        let before = invalid.to_bytes().unwrap();
        assert!(
            invalid
                .set_image_watermark(PNG, "bad.png", Length::emu(0), Length::pt(36.0))
                .is_err()
        );
        assert_eq!(invalid.to_bytes().unwrap(), before);
    }

    #[test]
    fn watermark_replacement_removes_stale_authored_images() {
        let mut document = Document::new();
        document.set_header("header");
        document
            .set_image_watermark(PNG, "first.png", Length::pt(72.0), Length::pt(36.0))
            .unwrap();
        let first_parts = document
            .package
            .parts
            .iter()
            .filter(|(_, bytes)| bytes.as_slice() == PNG)
            .map(|(name, _)| name.clone())
            .collect::<Vec<_>>();
        assert_eq!(first_parts.len(), 1);

        let mut replacement = PNG.to_vec();
        *replacement.last_mut().unwrap() ^= 1;
        document
            .set_image_watermark(
                &replacement,
                "second.png",
                Length::pt(72.0),
                Length::pt(36.0),
            )
            .unwrap();
        assert!(
            first_parts
                .iter()
                .all(|part| document.package.get_part(part).is_none())
        );
        assert_eq!(
            document
                .package
                .get_part_rels("/word/header1.xml")
                .unwrap()
                .items
                .iter()
                .filter(|relationship| relationship.rel_type == rel_types::IMAGE)
                .count(),
            1
        );

        document.set_text_watermark("FINAL").unwrap();
        assert!(
            document
                .package
                .parts
                .iter()
                .all(|(_, bytes)| bytes.as_slice() != replacement)
        );
        assert!(
            document
                .package
                .get_part_rels("/word/header1.xml")
                .is_none_or(|relationships| relationships
                    .items
                    .iter()
                    .all(|relationship| relationship.rel_type != rel_types::IMAGE))
        );
    }

    #[test]
    fn imported_header_watermark_order_is_stable_and_preserves_producer_relationships() {
        fn reopened_source() -> Document {
            let mut source = Document::new();
            source.set_header("producer header");
            source
                .package
                .set_part("/word/producer.bin", b"keep".to_vec());
            source
                .package
                .content_types
                .add_override("/word/producer.bin", "application/octet-stream");
            source
                .package
                .get_or_create_part_rels("/word/header1.xml")
                .add_with_id("producerRel", "urn:producer", "producer.bin");
            Document::from_bytes(&source.to_bytes().unwrap()).unwrap()
        }

        fn build(watermark_first: bool) -> Vec<u8> {
            let mut document = reopened_source();
            if watermark_first {
                document
                    .set_image_watermark(PNG, "watermark.png", Length::pt(72.0), Length::pt(36.0))
                    .unwrap();
                document.add_picture(PNG, "body.png", Length::pt(36.0), Length::pt(18.0));
            } else {
                document.add_picture(PNG, "body.png", Length::pt(36.0), Length::pt(18.0));
                document
                    .set_image_watermark(PNG, "watermark.png", Length::pt(72.0), Length::pt(36.0))
                    .unwrap();
            }
            document.set_text_watermark("FINAL").unwrap();
            let relationships = document.package.get_part_rels("/word/header1.xml").unwrap();
            assert_eq!(
                relationships.get_by_id("producerRel").unwrap().target,
                "producer.bin"
            );
            document.to_bytes().unwrap()
        }

        assert_eq!(build(false), build(true));
    }

    #[test]
    fn imported_header_watermark_preserves_producer_drawing_bytes_ids_and_relationship_order() {
        let mut source = Document::new();
        source.set_header_image(PNG, "producer.png", Length::pt(72.0), Length::pt(36.0));
        let source_bytes = source.to_bytes().unwrap();
        let mut package = OpcPackage::from_reader(Cursor::new(source_bytes)).unwrap();
        let mut header =
            String::from_utf8(package.get_part("/word/header1.xml").unwrap().to_vec()).unwrap();
        assert!(header.contains(r#"wp:docPr id="1""#));
        header = header.replacen(r#"wp:docPr id="1""#, r#"wp:docPr id="77""#, 1);
        let drawing_start = header.find("<w:drawing").unwrap();
        let drawing_end = drawing_start
            + header[drawing_start..].find("</w:drawing>").unwrap()
            + "</w:drawing>".len();
        let producer_drawing = header.as_bytes()[drawing_start..drawing_end].to_vec();
        package.set_part("/word/header1.xml", header.into_bytes());
        package.set_part("/word/z-producer.bin", b"z".to_vec());
        package.set_part("/word/a-producer.bin", b"a".to_vec());
        package
            .content_types
            .add_override("/word/z-producer.bin", "application/octet-stream");
        package
            .content_types
            .add_override("/word/a-producer.bin", "application/octet-stream");
        let relationships = package.get_or_create_part_rels("/word/header1.xml");
        let image_id = relationships
            .get_by_type(rel_types::IMAGE)
            .unwrap()
            .id
            .clone();
        relationships.items.insert(
            0,
            oxml_opc::relationship::Relationship {
                id: "zProducer".to_owned(),
                rel_type: "urn:producer:z".to_owned(),
                target: "z-producer.bin".to_owned(),
                target_mode: None,
            },
        );
        relationships.add_with_id("aProducer", "urn:producer:a", "a-producer.bin");
        let producer_order = ["zProducer".to_owned(), image_id, "aProducer".to_owned()];
        let mut serialized = Cursor::new(Vec::new());
        package.write_to(&mut serialized).unwrap();

        let mut document = Document::from_bytes(serialized.get_ref()).unwrap();
        document
            .set_image_watermark(PNG, "watermark.png", Length::pt(36.0), Length::pt(18.0))
            .unwrap();
        let output = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let output_header = output.get_part("/word/header1.xml").unwrap();
        assert!(
            output_header
                .windows(producer_drawing.len())
                .any(|window| window == producer_drawing)
        );
        assert!(String::from_utf8_lossy(output_header).contains(r#"wp:docPr id="77""#));
        let output_order = output
            .get_part_rels("/word/header1.xml")
            .unwrap()
            .items
            .iter()
            .take(3)
            .map(|relationship| relationship.id.clone())
            .collect::<Vec<_>>();
        assert_eq!(output_order, producer_order);
    }

    #[test]
    fn legacy_header_setters_and_watermark_roll_back_relationship_exhaustion() {
        let package_bytes = |document: &Document| {
            let mut output = Cursor::new(Vec::new());
            document.package.write_to(&mut output).unwrap();
            output.into_inner()
        };

        let mut text_seed = Document::new();
        text_seed
            .package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id(
                &format!("rId{}", u32::MAX),
                "urn:exhaustion",
                "unchanged.bin",
            );
        let mut seed_bytes = Cursor::new(Vec::new());
        text_seed.package.write_to(&mut seed_bytes).unwrap();
        let mut text = Document::from_bytes(seed_bytes.get_ref()).unwrap();
        let before = package_bytes(&text);
        assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                text.set_header("header");
            }))
            .is_err()
        );
        assert_eq!(package_bytes(&text), before);

        let mut image_seed = Document::new();
        image_seed.set_header("producer header");
        let mut image_package =
            OpcPackage::from_reader(Cursor::new(image_seed.to_bytes().unwrap())).unwrap();
        image_package
            .get_or_create_part_rels("/word/header1.xml")
            .add_with_id(
                &format!("rId{}", u32::MAX),
                "urn:exhaustion",
                "unchanged.bin",
            );
        let mut seed_bytes = Cursor::new(Vec::new());
        image_package.write_to(&mut seed_bytes).unwrap();
        let source = seed_bytes.into_inner();

        let mut image = Document::from_bytes(&source).unwrap();
        let before = package_bytes(&image);
        assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                image.set_header_image(PNG, "image.png", Length::pt(36.0), Length::pt(18.0));
            }))
            .is_err()
        );
        assert_eq!(package_bytes(&image), before);

        let mut watermark = Document::from_bytes(&source).unwrap();
        let before = package_bytes(&watermark);
        assert!(
            watermark
                .set_image_watermark(PNG, "image.png", Length::pt(36.0), Length::pt(18.0))
                .is_err()
        );
        assert_eq!(package_bytes(&watermark), before);
    }

    #[test]
    fn header_image_relationship_ids_are_scoped_per_part() {
        let mut document = Document::new();
        let image_watermark = |text: &str| {
            format!(
                r#"<w:hdr xmlns:w="{}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:r><w:pict><v:shape style="width:72pt;height:36pt"><v:fill opacity=".5"/><v:imagedata r:id="rId1"/></v:shape></w:pict><w:t>{text}</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes()
        };
        let mut alternate = PNG.to_vec();
        let last = alternate.len() - 1;
        alternate[last] ^= 1;
        document.set_raw_header_with_images(
            image_watermark("default"),
            &[("rId1", PNG, "default.png")],
            HdrFtrType::Default,
        );
        document.set_raw_header_with_images(
            image_watermark("first"),
            &[("rId1", &alternate, "first.png")],
            HdrFtrType::First,
        );
        let input = document.build_layout_input();
        let scoped = input
            .images
            .iter()
            .filter(|(id, _)| id.ends_with("\0rId1"))
            .map(|(_, image)| image.data.as_slice())
            .collect::<Vec<_>>();
        assert_eq!(scoped.len(), 2);
        assert!(scoped.contains(&PNG));
        assert!(scoped.contains(&alternate.as_slice()));
    }

    #[test]
    fn header_image_setter_order_does_not_change_package_bytes() {
        fn build(reverse: bool) -> Vec<u8> {
            let mut document = Document::new();
            let mut alternate = PNG.to_vec();
            *alternate.last_mut().unwrap() ^= 1;
            if reverse {
                document.set_first_page_header_image(
                    &alternate,
                    "first.png",
                    Length::pt(72.0),
                    Length::pt(36.0),
                );
                document.set_header_image(PNG, "default.png", Length::pt(72.0), Length::pt(36.0));
            } else {
                document.set_header_image(PNG, "default.png", Length::pt(72.0), Length::pt(36.0));
                document.set_first_page_header_image(
                    &alternate,
                    "first.png",
                    Length::pt(72.0),
                    Length::pt(36.0),
                );
            }
            document.to_bytes().unwrap()
        }
        assert_eq!(build(false), build(true));
    }

    #[test]
    fn replacing_authored_image_header_with_text_removes_orphan_media() {
        let mut document = Document::new();
        document.set_header_image(PNG, "old.png", Length::pt(72.0), Length::pt(36.0));
        let relationship = document
            .package
            .get_part_rels("/word/header1.xml")
            .and_then(|relationships| relationships.get_by_type(rel_types::IMAGE))
            .cloned()
            .unwrap();
        let image_part = OpcPackage::resolve_rel_target("/word/header1.xml", &relationship.target);
        assert!(document.package.get_part(&image_part).is_some());

        document.set_header("replacement");

        assert!(
            document
                .package
                .get_part_rels("/word/header1.xml")
                .is_none_or(|relationships| relationships.items.is_empty())
        );
        assert!(document.package.get_part(&image_part).is_none());
        let package = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        assert!(package.get_part(&image_part).is_none());
    }

    #[test]
    fn replacing_authored_images_with_text_restores_direct_package_bytes() {
        let mut header_history = Document::new();
        header_history.set_header_image(PNG, "old.png", Length::pt(72.0), Length::pt(36.0));
        header_history.set_header("replacement");
        let mut direct_header = Document::new();
        direct_header.set_header("replacement");
        assert_eq!(
            header_history.to_bytes().unwrap(),
            direct_header.to_bytes().unwrap()
        );

        let mut watermark_history = Document::new();
        watermark_history
            .set_image_watermark(PNG, "old.png", Length::pt(72.0), Length::pt(36.0))
            .unwrap();
        watermark_history.set_text_watermark("FINAL").unwrap();
        let mut direct_watermark = Document::new();
        direct_watermark.set_text_watermark("FINAL").unwrap();
        let history_bytes = watermark_history.to_bytes().unwrap();
        let direct_bytes = direct_watermark.to_bytes().unwrap();
        let history_package = OpcPackage::from_reader(Cursor::new(&history_bytes)).unwrap();
        let direct_package = OpcPackage::from_reader(Cursor::new(&direct_bytes)).unwrap();
        assert_eq!(history_package.content_types, direct_package.content_types);
        assert_eq!(
            history_package.package_rels.to_xml().unwrap(),
            direct_package.package_rels.to_xml().unwrap()
        );
        let relationship_xml = |package: &OpcPackage| {
            let mut relationships = package
                .part_rels
                .iter()
                .map(|(owner, relationships)| (owner.clone(), relationships.to_xml().unwrap()))
                .collect::<Vec<_>>();
            relationships.sort_by(|left, right| left.0.cmp(&right.0));
            relationships
        };
        assert_eq!(
            relationship_xml(&history_package),
            relationship_xml(&direct_package)
        );
        let mut history_parts = history_package.parts.keys().cloned().collect::<Vec<_>>();
        history_parts.sort();
        let mut direct_parts = direct_package.parts.keys().cloned().collect::<Vec<_>>();
        direct_parts.sort();
        assert_eq!(history_parts, direct_parts);
        for part_name in history_parts {
            let history_part = history_package.get_part(&part_name);
            let direct_part = direct_package.get_part(&part_name);
            if history_part != direct_part {
                panic!(
                    "part {part_name} differs\nhistory: {}\ndirect: {}",
                    String::from_utf8_lossy(history_part.unwrap()),
                    String::from_utf8_lossy(direct_part.unwrap())
                );
            }
        }
        assert_eq!(history_bytes, direct_bytes);
    }

    #[test]
    fn orphan_image_default_ignores_override_covered_producer_parts() {
        fn source() -> Vec<u8> {
            let mut document = Document::new_with_profile(WordCreationProfile::Minimal(
                WordPackageClass::Document,
            ));
            let mut package =
                OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
            package.set_part("/word/media/producer.png", b"producer".to_vec());
            package
                .content_types
                .add_override("/word/media/producer.png", "application/x-producer-image");
            let mut bytes = Cursor::new(Vec::new());
            package.write_to(&mut bytes).unwrap();
            bytes.into_inner()
        }

        let source = source();
        let mut history = Document::from_bytes(&source).unwrap();
        history.set_header_image(PNG, "old.png", Length::pt(72.0), Length::pt(36.0));
        history.set_header("replacement");
        let mut direct = Document::from_bytes(&source).unwrap();
        direct.set_header("replacement");

        let history_bytes = history.to_bytes().unwrap();
        let direct_bytes = direct.to_bytes().unwrap();
        let history_package = OpcPackage::from_reader(Cursor::new(&history_bytes)).unwrap();
        let direct_package = OpcPackage::from_reader(Cursor::new(&direct_bytes)).unwrap();
        assert!(!history_package.content_types.defaults.contains_key("png"));
        assert_eq!(history_package.content_types, direct_package.content_types);
        assert_eq!(history_bytes, direct_bytes);
    }

    #[test]
    fn pruned_maximum_header_relationship_id_can_be_reused() {
        let mut document = Document::new();
        let maximum = format!("rId{}", u32::MAX);
        let xml = format!(
            r#"<w:hdr xmlns:w="{}" xmlns:r="{}"><w:p><w:r><w:drawing><a:blip xmlns:a="{}" r:embed="{maximum}"/></w:drawing></w:r></w:p></w:hdr>"#,
            rdocx_oxml::namespace::W_NS,
            drawing_ns::R,
            drawing_ns::A,
        )
        .into_bytes();
        document.set_raw_header_with_images(
            xml,
            &[(maximum.as_str(), PNG, "old.png")],
            HdrFtrType::Default,
        );
        assert!(
            document
                .package
                .get_part_rels("/word/header1.xml")
                .unwrap()
                .get_by_id(&maximum)
                .is_some()
        );

        document.set_header("replacement");
        document.set_header_image(PNG, "new.png", Length::pt(72.0), Length::pt(36.0));
        let relationship = document
            .package
            .get_part_rels("/word/header1.xml")
            .and_then(|relationships| relationships.get_by_type(rel_types::IMAGE))
            .unwrap();
        assert_ne!(relationship.id, maximum);
    }

    #[test]
    fn direct_raw_header_replacement_reuses_its_retired_maximum_relationship_id() {
        fn raw_header(relationship_id: &str, text: &str) -> Vec<u8> {
            format!(
                r#"<w:hdr xmlns:w="{}" xmlns:r="{}"><w:p><w:r><w:drawing><a:blip xmlns:a="{}" r:embed="{relationship_id}"/></w:drawing><w:t>{text}</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS,
                drawing_ns::R,
                drawing_ns::A,
            )
            .into_bytes()
        }

        let maximum = format!("rId{}", u32::MAX);
        let mut final_image = PNG.to_vec();
        *final_image.last_mut().unwrap() ^= 1;
        let mut history = Document::new();
        history.set_raw_header_with_images(
            raw_header(&maximum, "old"),
            &[(maximum.as_str(), PNG, "old.png")],
            HdrFtrType::Default,
        );
        history.set_raw_header_with_images(
            raw_header(&maximum, "final"),
            &[(maximum.as_str(), &final_image, "final.png")],
            HdrFtrType::Default,
        );
        let mut direct = Document::new();
        direct.set_raw_header_with_images(
            raw_header(&maximum, "final"),
            &[(maximum.as_str(), &final_image, "final.png")],
            HdrFtrType::Default,
        );

        assert_eq!(history.to_bytes().unwrap(), direct.to_bytes().unwrap());
    }

    #[test]
    fn legacy_header_setters_preserve_imported_relationship_order() {
        fn source() -> Vec<u8> {
            let mut source = Document::new();
            source.set_header("producer header");
            let bytes = source.to_bytes().unwrap();
            let mut package = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
            let relationships = package.get_or_create_part_rels("/word/header1.xml");
            relationships.add_with_id("rId9", "urn:producer:first", "first.bin");
            relationships.add_with_id("rId2", "urn:producer:second", "second.bin");
            let mut bytes = Cursor::new(Vec::new());
            package.write_to(&mut bytes).unwrap();
            bytes.into_inner()
        }

        for image in [false, true] {
            let mut document = Document::from_bytes(&source()).unwrap();
            if image {
                document.set_header_image(PNG, "new.png", Length::pt(72.0), Length::pt(36.0));
            } else {
                document.set_header("replacement");
            }
            let package =
                OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
            let producer_order = package
                .get_part_rels("/word/header1.xml")
                .unwrap()
                .items
                .iter()
                .filter(|relationship| relationship.rel_type.starts_with("urn:producer:"))
                .map(|relationship| relationship.id.as_str())
                .collect::<Vec<_>>();
            assert_eq!(producer_order, ["rId9", "rId2"]);
        }
    }

    #[test]
    fn set_header_preserves_mixed_case_relationship_owner_spelling() {
        const MIXED_HEADER: &str = "/WORD/HEADER1.XML";
        let mut source = Document::new();
        source.set_header("producer");
        let mut package = OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
        package.set_part("/word/producer.bin", b"producer".to_vec());
        package
            .content_types
            .add_default("bin", "application/octet-stream");
        package
            .get_or_create_part_rels("/word/header1.xml")
            .add_with_id("rId9", "urn:producer", "producer.bin");
        let header = package.remove_part("/word/header1.xml").unwrap();
        package.parts.insert(MIXED_HEADER.to_owned(), header);
        let relationships = package.remove_part_rels("/word/header1.xml").unwrap();
        package
            .part_rels
            .insert(MIXED_HEADER.to_owned(), relationships);
        package
            .get_part_rels_mut("/word/document.xml")
            .unwrap()
            .items
            .iter_mut()
            .find(|relationship| relationship.rel_type == rel_types::HEADER)
            .unwrap()
            .target = "HEADER1.XML".to_owned();
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();

        let mut document = Document::from_bytes(bytes.get_ref()).unwrap();
        document.set_header("replacement");
        let saved = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        assert!(saved.part_rels.contains_key(MIXED_HEADER));
        assert!(!saved.part_rels.contains_key("/word/header1.xml"));
        assert_eq!(
            saved
                .get_part_rels(MIXED_HEADER)
                .unwrap()
                .get_by_id("rId9")
                .unwrap()
                .rel_type,
            "urn:producer"
        );
    }

    #[test]
    fn raw_header_relationship_collisions_are_remapped_without_replacing_preserved_state() {
        let mut source = Document::new();
        source.set_header("preserved");
        source
            .package
            .get_or_create_part_rels("/word/header1.xml")
            .add_with_id("rId1", "urn:preserved", "preserved.bin");
        let mut document = Document::from_bytes(&source.to_bytes().unwrap()).unwrap();
        let xml = format!(
            r#"<w:hdr xmlns:w="{}" xmlns:r="{}"><w:p><w:r><w:drawing><a:blip xmlns:a="{}" r:embed="rId&#49;"/></w:drawing></w:r></w:p></w:hdr>"#,
            rdocx_oxml::namespace::W_NS,
            drawing_ns::R,
            drawing_ns::A,
        )
        .into_bytes();
        document.set_raw_header_with_images(
            xml,
            &[("rId1", PNG, "replacement.png")],
            HdrFtrType::Default,
        );
        let relationships = document.package.get_part_rels("/word/header1.xml").unwrap();
        assert_eq!(
            relationships.get_by_id("rId1").unwrap().rel_type,
            "urn:preserved"
        );
        let image = relationships
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::IMAGE)
            .unwrap();
        assert_ne!(image.id, "rId1");
        let image_id = image.id.clone();
        let header =
            std::str::from_utf8(document.package.get_part("/word/header1.xml").unwrap()).unwrap();
        assert!(
            header.contains(&format!(r#"r:embed="{image_id}""#)),
            "{header}"
        );
        assert!(!header.contains("rId&#49;"), "{header}");
        let saved = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&saved).unwrap();
        let header_part = reopened.active_header_footer_parts().remove(0);
        let header_xml = reopened.package.get_part(&header_part).unwrap();
        let referenced = xml_relationship_ids_in_order(header_xml).unwrap();
        assert!(referenced.contains(&image_id));
        assert!(
            reopened
                .package
                .get_part_rels(&header_part)
                .unwrap()
                .get_by_id(&image_id)
                .is_some()
        );
    }

    #[test]
    fn authored_raw_header_identifier_rewrites_preserve_unrelated_bytes() {
        let mut document = Document::new();
        let xml = format!(
            r#"<w:hdr xmlns:w="{}" xmlns:r="{}" xmlns:wp="{}" xmlns:x="urn:producer"><w:p><w:r><w:drawing><wp:inline><wp:docPr id="99"/><a:blip xmlns:a="{}" r:embed="rId9"/></wp:inline></w:drawing><x:unknown x:a = 'v'><!-- keep  spacing --></x:unknown></w:r></w:p></w:hdr>"#,
            rdocx_oxml::namespace::W_NS,
            drawing_ns::R,
            drawing_ns::WP,
            drawing_ns::A,
        );
        document.set_raw_header_with_images(
            xml.as_bytes().to_vec(),
            &[("rId9", PNG, "raw.png")],
            HdrFtrType::Default,
        );

        let package = OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let actual = package.get_part("/word/header1.xml").unwrap();
        let expected = xml
            .replace(r#"id="99""#, r#"id="1""#)
            .replace(r#"r:embed="rId9""#, r#"r:embed="rId0""#);
        assert_eq!(actual, expected.as_bytes());
    }

    #[test]
    fn rejected_raw_header_input_is_atomic() {
        let mut document = Document::new();
        document.set_header("unchanged");
        let before = document.to_bytes().unwrap();
        let xml = format!(
            r#"<w:hdr xmlns:w="{}" xmlns:r="{}"><w:p><w:r><w:drawing><a:blip xmlns:a="{}" r:embed="rId1"/></w:drawing></w:r></w:p></w:hdr>"#,
            rdocx_oxml::namespace::W_NS,
            drawing_ns::R,
            drawing_ns::A,
        )
        .into_bytes();
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            document.set_raw_header_with_images(
                xml,
                &[("rId1", PNG, "one.png"), ("rId1", PNG, "two.png")],
                HdrFtrType::Default,
            );
        }));
        assert!(result.is_err());
        assert_eq!(document.to_bytes().unwrap(), before);
    }

    #[test]
    fn occupied_orphan_header_part_is_not_overwritten() {
        let mut source = Document::new();
        let mut package = OpcPackage::from_reader(Cursor::new(source.to_bytes().unwrap())).unwrap();
        package.set_part("/word/header1.xml", b"preserved orphan".to_vec());
        package.content_types.add_override(
            "/word/header1.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        );
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let mut document = Document::from_bytes(&bytes.into_inner()).unwrap();

        document.set_header("new header");

        assert_eq!(
            document.package.get_part("/word/header1.xml"),
            Some(b"preserved orphan".as_slice())
        );
        let active = document
            .active_header_footer_parts()
            .into_iter()
            .next()
            .unwrap();
        assert_ne!(active, "/word/header1.xml");
    }

    #[test]
    fn watermark_covers_title_page_fallback_and_every_section() {
        let mut document = Document::new();
        document.add_paragraph("first section");
        let BodyContent::Paragraph(first) = document.document.body.content.last_mut().unwrap()
        else {
            panic!("expected first section paragraph");
        };
        first.properties = Some(CT_PPr {
            sect_pr: Some(CT_SectPr::default_letter()),
            ..Default::default()
        });
        document.add_paragraph("final section");
        document.set_different_first_page(true);

        document.set_text_watermark("DRAFT").unwrap();
        let sections = document
            .document
            .body
            .content
            .iter()
            .filter_map(|content| match content {
                BodyContent::Paragraph(paragraph) => paragraph
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.sect_pr.as_ref()),
                _ => None,
            })
            .chain(document.document.body.sect_pr.iter())
            .collect::<Vec<_>>();
        assert_eq!(sections.len(), 2);
        assert!(
            sections[0]
                .header_refs
                .iter()
                .any(|reference| reference.hdr_ftr_type == HdrFtrType::Default)
        );
        assert!(
            !sections[1]
                .header_refs
                .iter()
                .any(|reference| reference.hdr_ftr_type == HdrFtrType::Default)
        );
        assert!(
            sections[1]
                .header_refs
                .iter()
                .any(|reference| reference.hdr_ftr_type == HdrFtrType::First)
        );
        let active_parts = document.active_header_footer_parts();
        assert_eq!(active_parts.len(), 2);
        assert!(
            active_parts
                .iter()
                .all(|part| document.identifiers.authored_story_parts.contains(part))
        );
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.active_header_footer_parts().len(), 2);

        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert_eq!(layout.pages.len(), 2);
        assert!(layout.pages.iter().all(|page| {
            matches!(
                compatibility_page_elements(&page.elements).first(),
                Some(oxml_layout::PositionedElement::Group(_))
            )
        }));
    }

    #[test]
    fn saved_watermark_materializes_first_and_even_fallback_headers() {
        let mut document = Document::new();
        document.set_different_first_page(true);
        enable_even_headers(&mut document);
        document.set_text_watermark("DRAFT").unwrap();

        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let section = reopened.section_properties().unwrap();
        for hdr_type in [HdrFtrType::Default, HdrFtrType::First, HdrFtrType::Even] {
            let reference = section
                .header_refs
                .iter()
                .find(|reference| reference.hdr_ftr_type == hdr_type)
                .unwrap_or_else(|| panic!("missing saved {hdr_type:?} header"));
            let header = reopened
                .load_header_footer(&reference.rel_id, true)
                .unwrap();
            assert!(matches!(
                header.watermarks(),
                [rdocx_oxml::header_footer::VmlWatermark::Text { text, .. }]
                    if text == "DRAFT"
            ));
        }
    }

    #[test]
    fn inherited_ordinary_header_remains_inherited_after_watermarking() {
        let mut document = Document::new();
        document.set_header("Company header");
        document.add_paragraph("first section");
        let company = document
            .section_properties_mut()
            .header_refs
            .pop()
            .expect("company header reference");
        let BodyContent::Paragraph(first) = document.document.body.content.last_mut().unwrap()
        else {
            panic!("expected first section paragraph");
        };
        first.properties = Some(CT_PPr {
            sect_pr: Some(CT_SectPr {
                header_refs: vec![company],
                ..CT_SectPr::default_letter()
            }),
            ..Default::default()
        });
        document.add_paragraph("second section");

        document.set_text_watermark("DRAFT").unwrap();
        assert!(
            document
                .section_properties()
                .unwrap()
                .header_refs
                .is_empty()
        );
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        assert!(
            reopened
                .section_properties()
                .unwrap()
                .header_refs
                .is_empty()
        );
        let layout = reopened
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert_eq!(layout.pages.len(), 2);
        assert!(page_text(&layout, 1).contains("Company header"));
        assert!(matches!(
            compatibility_page_elements(&layout.pages[1].elements).first(),
            Some(oxml_layout::PositionedElement::Group(_))
        ));
    }

    #[test]
    fn section_page_number_restart_controls_even_header_parity() {
        let mut document = Document::new();
        document.add_paragraph("first section");
        let BodyContent::Paragraph(first) = document.document.body.content.last_mut().unwrap()
        else {
            panic!("expected first section paragraph");
        };
        first.properties = Some(CT_PPr {
            sect_pr: Some(CT_SectPr::default_letter()),
            ..Default::default()
        });
        document.add_paragraph("second section");
        document.set_header("default header");
        document.set_raw_header_with_images(
            format!(
                r#"<w:hdr xmlns:w="{}"><w:p><w:r><w:t>even header</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
            &[],
            HdrFtrType::Even,
        );
        enable_even_headers(&mut document);
        document
            .section_properties_mut()
            .extra_xml
            .push(br#"<w:pgNumType w:start="1"/>"#.to_vec());
        document.set_text_watermark("DRAFT").unwrap();

        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let layout = reopened
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert_eq!(layout.pages.len(), 2);
        assert!(page_text(&layout, 1).contains("default header"));
        assert!(!page_text(&layout, 1).contains("even header"));
    }

    #[test]
    fn blank_first_and_even_variants_do_not_borrow_default_content() {
        let mut first = Document::new();
        first.set_raw_header_with_images(
            format!(
                r#"<w:hdr xmlns:w="{}" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:pict><v:shape style="width:468pt;height:117pt;rotation:315" fillcolor="D9D9D9"><v:fill opacity=".5"/><v:textpath string="DRAFT"/></v:shape></w:pict><w:t>company header</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
            &[],
            HdrFtrType::Default,
        );
        first.set_footer("company footer");
        first.set_first_page_header("");
        first.set_first_page_footer("");
        first.add_paragraph("body");
        let first_layout = first
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert!(
            !compatibility_page_elements(&first_layout.pages[0].elements)
                .iter()
                .any(|element| matches!(element, oxml_layout::PositionedElement::Group(_)))
        );
        assert!(!page_text(&first_layout, 0).contains("company header"));
        assert!(!page_text(&first_layout, 0).contains("company footer"));

        let mut even = Document::new();
        even.set_header("company header");
        even.set_footer("company footer");
        even.set_raw_footer_with_images(
            format!(r#"<w:ftr xmlns:w="{}"/>"#, rdocx_oxml::namespace::W_NS).into_bytes(),
            &[],
            HdrFtrType::Even,
        );
        enable_even_headers(&mut even);
        even.set_text_watermark("DRAFT").unwrap();
        even.add_paragraph(&"body ".repeat(4_000));
        let even_layout = even
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert!(matches!(
            compatibility_page_elements(&even_layout.pages[1].elements).first(),
            Some(oxml_layout::PositionedElement::Group(_))
        ));
        assert!(!page_text(&even_layout, 1).contains("company header"));
        assert!(!page_text(&even_layout, 1).contains("company footer"));
    }

    #[test]
    fn image_watermark_target_is_relative_to_a_custom_header_part() {
        let mut document = Document::new();
        let header_part = "/custom/headers/header.xml";
        document.package.set_part(
            header_part,
            format!(
                r#"<w:hdr xmlns:w="{}"><w:p/></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
        );
        document.package.content_types.add_override(
            header_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        );
        let target = relative_target(&document.doc_part_name, header_part);
        let rel_id = document
            .package
            .get_or_create_part_rels(&document.doc_part_name)
            .add(rel_types::HEADER, &target);
        document.section_properties_mut().header_refs = vec![HdrFtrRef {
            hdr_ftr_type: HdrFtrType::Default,
            rel_id,
        }];

        document
            .set_image_watermark(PNG, "custom.png", Length::pt(72.0), Length::pt(36.0))
            .unwrap();
        let image_relationship = document
            .package
            .get_part_rels(header_part)
            .unwrap()
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::IMAGE)
            .unwrap();
        let image_part = OpcPackage::resolve_rel_target(header_part, &image_relationship.target);
        assert_eq!(document.package.get_part(&image_part), Some(PNG));
    }

    #[test]
    fn named_vml_colour_renders_its_defined_rgb_value() {
        let mut document = Document::new();
        document.set_raw_header_with_images(
            text_vml_header("DRAFT", "silver"),
            &[],
            HdrFtrType::Default,
        );
        document.add_paragraph("body");
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        let elements = compatibility_page_elements(&layout.pages[0].elements);
        let oxml_layout::PositionedElement::Group(group) = elements[0] else {
            panic!("expected watermark group");
        };
        let oxml_layout::PositionedElement::Text(run) = &group.children[0] else {
            panic!("expected watermark text");
        };
        assert_eq!(
            run.color,
            oxml_layout::Color {
                r: 192.0 / 255.0,
                g: 192.0 / 255.0,
                b: 192.0 / 255.0,
                a: 1.0,
            }
        );
    }

    #[test]
    fn invalid_multibyte_vml_colour_is_suppressed_without_panicking() {
        let mut document = Document::new();
        document.set_raw_header_with_images(
            text_vml_header("DRAFT", "€€"),
            &[],
            HdrFtrType::Default,
        );
        document.add_paragraph("body");
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert!(
            !compatibility_page_elements(&layout.pages[0].elements)
                .iter()
                .any(|element| matches!(element, oxml_layout::PositionedElement::Group(_)))
        );
        assert!(
            layout
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("colour")
                    && diagnostic.message.contains("unsupported"))
        );
    }

    #[test]
    fn even_header_selection_follows_the_document_setting() {
        let render = |setting: Option<Option<&str>>| {
            let mut document = Document::new();
            document.set_header("default header");
            document.set_raw_header_with_images(
                format!(
                    r#"<w:hdr xmlns:w="{}"><w:p><w:r><w:t>even header</w:t></w:r></w:p></w:hdr>"#,
                    rdocx_oxml::namespace::W_NS
                )
                .into_bytes(),
                &[],
                HdrFtrType::Even,
            );
            if let Some(value) = setting {
                set_even_headers_value(&mut document, value);
            }
            document.add_paragraph(&"body ".repeat(4_000));
            document
                .layout_for_options(RenderOptions::default(), true)
                .unwrap()
                .layout
                .clone()
        };
        let disabled = render(None);
        let entity_false = render(Some(Some("&#48;")));
        let enabled = render(Some(None));
        assert!(page_text(&disabled, 1).contains("default header"));
        assert!(!page_text(&disabled, 1).contains("even header"));
        assert!(page_text(&entity_false, 1).contains("default header"));
        assert!(!page_text(&entity_false, 1).contains("even header"));
        assert!(page_text(&enabled, 1).contains("even header"));
    }

    #[test]
    fn unresolved_image_watermark_is_suppressed_with_a_diagnostic() {
        let mut document = Document::new();
        let xml = format!(
            r#"<w:hdr xmlns:w="{}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:r><w:pict><v:shape style="width:72pt;height:36pt"><v:imagedata r:id="rIdMissing"/></v:shape></w:pict></w:r></w:p></w:hdr>"#,
            rdocx_oxml::namespace::W_NS
        );
        document.set_raw_header_with_images(xml.into_bytes(), &[], HdrFtrType::Default);
        document.add_paragraph("body");
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert!(
            !compatibility_page_elements(&layout.pages[0].elements)
                .iter()
                .any(|element| matches!(element, oxml_layout::PositionedElement::Group(_)))
        );
        assert!(layout.diagnostics.iter().any(|diagnostic| {
            diagnostic.message.contains("rIdMissing") && diagnostic.message.contains("not resolved")
        }));
    }

    #[test]
    fn watermark_is_centered_in_the_margin_rectangle() {
        let mut document = Document::new();
        document.set_margins(
            Length::pt(36.0),
            Length::pt(144.0),
            Length::pt(108.0),
            Length::pt(72.0),
        );
        document.set_text_watermark("DRAFT").unwrap();
        document.add_paragraph("body");
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        let elements = compatibility_page_elements(&layout.pages[0].elements);
        let oxml_layout::PositionedElement::Group(group) = elements[0] else {
            panic!("expected watermark group");
        };
        let center = group.transform.apply(oxml_layout::Point {
            x: 468.0 / 2.0,
            y: 117.0 / 2.0,
        });
        assert!((center.x - (72.0 + (612.0 - 72.0 - 144.0) / 2.0)).abs() < 1.0e-9);
        assert!((center.y - (36.0 + (792.0 - 36.0 - 108.0) / 2.0)).abs() < 1.0e-9);
    }

    #[test]
    fn watermark_group_precedes_body_elements_on_every_page() {
        let mut document = Document::new();
        document.set_header("default header");
        document.set_first_page_header("first header");
        document.set_raw_header_with_images(
            format!(
                r#"<w:hdr xmlns:w="{}"><w:p><w:r><w:t>even header</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
            &[],
            HdrFtrType::Even,
        );
        enable_even_headers(&mut document);
        document.set_text_watermark("DRAFT").unwrap();
        document.add_paragraph(&"body ".repeat(4_000));
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert!(layout.pages.len() > 2);
        assert!(layout.pages.iter().all(|page| {
            matches!(
                compatibility_page_elements(&page.elements).first(),
                Some(oxml_layout::PositionedElement::Group(_))
            )
        }));
        assert!(
            page_text(&layout, 0).contains("first header"),
            "{:?}",
            page_text(&layout, 0)
        );
        assert!(
            page_text(&layout, 1).contains("even header"),
            "{:?}",
            page_text(&layout, 1)
        );
        assert!(
            page_text(&layout, 2).contains("default header"),
            "{:?}",
            page_text(&layout, 2)
        );
    }

    #[test]
    fn watermark_renders_behind_body_text_on_every_page() {
        let mut document = Document::new();
        document.set_text_watermark("DRAFT").unwrap();
        document.add_paragraph(&"body ".repeat(4_000));
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        let pngs = oxml_pdf::render_all_pages(&layout, 72.0);
        assert!(pngs.len() > 1);
        assert!(pngs.iter().all(|png| png.starts_with(b"\x89PNG\r\n\x1a\n")));
        let digests = pngs
            .iter()
            .map(|png| {
                png.iter().fold(0xcbf2_9ce4_8422_2325_u64, |hash, byte| {
                    (hash ^ u64::from(*byte)).wrapping_mul(0x0000_0100_0000_01b3)
                })
            })
            .collect::<Vec<_>>();
        assert_eq!(
            digests,
            [
                740_018_920_125_384_146,
                740_018_920_125_384_146,
                740_018_920_125_384_146,
                740_018_920_125_384_146,
                1_020_215_290_976_271_429,
            ]
        );
    }
}

#[cfg(test)]
mod reader_fact_tests {
    use super::*;

    #[test]
    fn revision_reader_exposes_tracked_insertion_paragraph_items() {
        let xml = br#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:ins w:id="7" w:author="Ada"><w:r><w:t>before</w:t></w:r><w:hyperlink r:id="rId9"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:ins></w:p></w:body></w:document>"#;
        let mut document = Document::new();
        document.document = CT_Document::from_xml(xml).expect("document parses");
        let crate::BodyItemRef::Paragraph(paragraph) =
            document.body_items().next().expect("paragraph body item")
        else {
            panic!("expected paragraph");
        };
        let crate::ParagraphItemRef::Revision(revision) =
            paragraph.items().next().expect("revision paragraph item")
        else {
            panic!("expected revision");
        };
        let insertion = revision
            .insertion_paragraph()
            .expect("insertion paragraph projection");
        let items = insertion.items().collect::<Vec<_>>();

        assert!(matches!(items[0], crate::ParagraphItemRef::Run(_)));
        let crate::ParagraphItemRef::Hyperlink(link) = &items[1] else {
            panic!("expected hyperlink");
        };
        assert_eq!(link.relationship_id(), Some("rId9"));
    }

    #[test]
    fn unsupported_xml_reader_inspects_preserved_item_bytes() {
        let raw = br#"<w:proofErr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="spellStart"/>"#;
        let fact = UnsupportedXmlRef::from_bytes(raw);

        assert_eq!(fact.raw_xml(), Some(raw.as_slice()));
        assert_eq!(fact.qualified_name(), Some("w:proofErr"));
        assert_eq!(fact.local_name(), "proofErr");
        assert_eq!(
            fact.namespace_uri(),
            Some("http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        );
        assert!(!fact.has_child_content());
    }
}

#[cfg(test)]
mod odttf_tests {
    use super::*;

    /// Build a TrueType header with a two-entry table directory.
    fn fake_font() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend(b"\x00\x01\x00\x00"); // sfntVersion
        data.extend(2u16.to_be_bytes()); // numTables
        data.extend(32u16.to_be_bytes()); // searchRange
        data.extend(1u16.to_be_bytes()); // entrySelector
        data.extend(0u16.to_be_bytes()); // rangeShift
        for (tag, offset, length) in [(b"cmap", 96u32, 40u32), (b"head", 136, 54)] {
            data.extend(tag); // tag
            data.extend(0u32.to_be_bytes()); // checksum
            data.extend(offset.to_be_bytes());
            data.extend(length.to_be_bytes());
        }
        data.extend((0u8..64).map(|i| i.wrapping_mul(7)));
        data
    }

    fn obfuscate(font: &[u8], key: &[u8; 16]) -> Vec<u8> {
        let mut out = font.to_vec();
        for (i, byte) in out.iter_mut().take(32).enumerate() {
            *byte ^= key[i % 16];
        }
        out
    }

    const GUID_HEX: &str = "00112233445566778899AABBCCDDEEFF";

    fn guid_bytes() -> [u8; 16] {
        let mut g = [0u8; 16];
        for (i, b) in g.iter_mut().enumerate() {
            *b = u8::from_str_radix(&GUID_HEX[i * 2..i * 2 + 2], 16).unwrap();
        }
        g
    }

    #[test]
    fn recovers_font_under_either_key_convention() {
        let font = fake_font();
        let name = format!("{GUID_HEX}.odttf");
        for key in odttf_key_candidates(&guid_bytes()) {
            let obfuscated = obfuscate(&font, &key);
            assert_eq!(
                deobfuscate_odttf(&obfuscated, &name).as_deref(),
                Some(font.as_slice()),
                "failed to recover font for key {key:02x?}",
            );
        }
    }

    #[test]
    fn rejects_data_that_does_not_decode_to_a_font() {
        // A GUID that matches nothing in the payload must not yield garbage.
        let junk = vec![0xAB; 64];
        let name = format!("{GUID_HEX}.odttf");
        assert_eq!(deobfuscate_odttf(&junk, &name), None);
    }

    #[test]
    fn rejects_short_or_malformed_input() {
        assert_eq!(deobfuscate_odttf(&[0u8; 8], "abc.odttf"), None);
        assert_eq!(deobfuscate_odttf(&[0u8; 64], "not-a-guid.odttf"), None);
    }

    #[test]
    fn accepts_braced_and_hyphenated_names() {
        let font = fake_font();
        let key = odttf_key_candidates(&guid_bytes())[0];
        let obfuscated = obfuscate(&font, &key);
        for name in [
            "{00112233-4455-6677-8899-AABBCCDDEEFF}.odttf",
            "00112233-4455-6677-8899-AABBCCDDEEFF.odttf",
        ] {
            assert_eq!(
                deobfuscate_odttf(&obfuscated, name).as_deref(),
                Some(font.as_slice()),
                "failed for {name}"
            );
        }
    }
}
