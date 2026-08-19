//! Numbering definitions: `CT_Numbering`, `CT_AbstractNum`, `CT_Num`, `CT_Lvl`.
//!
//! These types represent the content of `numbering.xml`, which defines
//! abstract numbering formats and numbering instances that paragraphs reference.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::error::{OxmlError, Result};
use crate::namespace::{MC_NS, W_NS, matches_local_name};
use crate::properties::{CT_PPr, CT_RPr};
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::shared::ST_Jc;

const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PRESERVED_PROPERTY_NS: &str = "urn:rdocx:preserved-property";
const PRESERVED_PROPERTY_LOCAL: &str = "preservedProperty";
const MAX_PROPERTY_XML_DEPTH: usize = 64;
const PPR_CHILDREN: &[&[u8]] = &[
    b"pStyle",
    b"keepNext",
    b"keepLines",
    b"pageBreakBefore",
    b"framePr",
    b"widowControl",
    b"numPr",
    b"suppressLineNumbers",
    b"pBdr",
    b"shd",
    b"tabs",
    b"suppressAutoHyphens",
    b"kinsoku",
    b"wordWrap",
    b"overflowPunct",
    b"topLinePunct",
    b"autoSpaceDE",
    b"autoSpaceDN",
    b"bidi",
    b"adjustRightInd",
    b"snapToGrid",
    b"spacing",
    b"ind",
    b"contextualSpacing",
    b"mirrorIndents",
    b"suppressOverlap",
    b"jc",
    b"textDirection",
    b"textAlignment",
    b"textboxTightWrap",
    b"outlineLvl",
    b"divId",
    b"cnfStyle",
    b"rPr",
    b"sectPr",
    b"pPrChange",
];

const RPR_CHILDREN: &[&[u8]] = &[
    b"rStyle",
    b"rFonts",
    b"b",
    b"bCs",
    b"i",
    b"iCs",
    b"caps",
    b"smallCaps",
    b"strike",
    b"dstrike",
    b"outline",
    b"shadow",
    b"emboss",
    b"imprint",
    b"noProof",
    b"snapToGrid",
    b"vanish",
    b"webHidden",
    b"color",
    b"spacing",
    b"w",
    b"kern",
    b"position",
    b"sz",
    b"szCs",
    b"highlight",
    b"u",
    b"effect",
    b"bdr",
    b"shd",
    b"fitText",
    b"vertAlign",
    b"rtl",
    b"cs",
    b"em",
    b"lang",
    b"eastAsianLayout",
    b"specVanish",
    b"oMath",
    b"rPrChange",
];

const PPR_MODELLED_CHILDREN: &[&[u8]] = &[
    b"pStyle",
    b"keepNext",
    b"keepLines",
    b"pageBreakBefore",
    b"widowControl",
    b"numPr",
    b"pBdr",
    b"shd",
    b"tabs",
    b"suppressAutoHyphens",
    b"spacing",
    b"ind",
    b"jc",
    b"outlineLvl",
    b"rPr",
    b"sectPr",
    b"pPrChange",
];

const RPR_MODELLED_CHILDREN: &[&[u8]] = &[
    b"rStyle",
    b"rFonts",
    b"b",
    b"bCs",
    b"i",
    b"iCs",
    b"caps",
    b"smallCaps",
    b"strike",
    b"dstrike",
    b"vanish",
    b"color",
    b"spacing",
    b"w",
    b"position",
    b"sz",
    b"szCs",
    b"highlight",
    b"u",
    b"shd",
    b"vertAlign",
    b"ins",
    b"del",
    b"rPrChange",
];

const NUM_PR_CHILDREN: &[&[u8]] = &[b"ilvl", b"numId", b"numberingChange", b"ins"];
const P_BDR_CHILDREN: &[&[u8]] = &[
    b"top", b"left", b"start", b"bottom", b"right", b"end", b"between", b"bar",
];
const TABS_CHILDREN: &[&[u8]] = &[b"tab"];
const SECT_PR_CHILDREN: &[&[u8]] = &[
    b"headerReference",
    b"footerReference",
    b"footnotePr",
    b"endnotePr",
    b"type",
    b"pgSz",
    b"pgMar",
    b"paperSrc",
    b"pgBorders",
    b"lnNumType",
    b"pgNumType",
    b"cols",
    b"formProt",
    b"vAlign",
    b"noEndnote",
    b"titlePg",
    b"textDirection",
    b"bidi",
    b"rtlGutter",
    b"docGrid",
    b"printerSettings",
    b"sectPrChange",
];

#[derive(Clone, Copy)]
enum PropertyKind {
    Paragraph,
    Run,
}

impl PropertyKind {
    fn children(self) -> &'static [&'static [u8]] {
        match self {
            Self::Paragraph => PPR_CHILDREN,
            Self::Run => RPR_CHILDREN,
        }
    }

    fn modelled_children(self) -> &'static [&'static [u8]] {
        match self {
            Self::Paragraph => PPR_MODELLED_CHILDREN,
            Self::Run => RPR_MODELLED_CHILDREN,
        }
    }

    fn local_name(self) -> &'static str {
        match self {
            Self::Paragraph => "pPr",
            Self::Run => "rPr",
        }
    }
}

#[derive(Clone, Copy)]
enum LevelProperty<'a> {
    Paragraph(&'a CT_PPr),
    Run(&'a CT_RPr),
}

impl LevelProperty<'_> {
    fn kind(&self) -> PropertyKind {
        match self {
            Self::Paragraph(_) => PropertyKind::Paragraph,
            Self::Run(_) => PropertyKind::Run,
        }
    }

    fn canonical_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        match self {
            Self::Paragraph(value) if **value == CT_PPr::default() => {
                writer.write_event(Event::Empty(BytesStart::new("w:pPr")))?
            }
            Self::Paragraph(value) => value.to_xml(&mut writer)?,
            Self::Run(value) if **value == CT_RPr::default() => {
                writer.write_event(Event::Empty(BytesStart::new("w:rPr")))?
            }
            Self::Run(value) => value.to_xml(&mut writer)?,
        }
        Ok(writer.into_inner())
    }

    fn tab_sources(&self) -> Option<Vec<Option<usize>>> {
        match self {
            Self::Paragraph(value) => value
                .tabs
                .as_ref()
                .map(|tabs| tabs.tabs.iter().map(|tab| tab.source_occurrence).collect()),
            Self::Run(_) => None,
        }
    }
}

fn paragraph_preservation_eq(current: &CT_PPr, original: &CT_PPr) -> bool {
    current == original
        && LevelProperty::Paragraph(current).tab_sources()
            == LevelProperty::Paragraph(original).tab_sources()
}

fn qualified(prefix: &str, local: &str) -> String {
    format!("{prefix}:{local}")
}

fn generated_prefix(declarations: &[(String, String)], preferred: &str, namespace: &str) -> String {
    if !declarations
        .iter()
        .any(|(prefix, value)| prefix == preferred && value != namespace)
    {
        return preferred.to_string();
    }

    for suffix in 1usize.. {
        let candidate = format!("{preferred}{suffix}");
        if !declarations
            .iter()
            .any(|(prefix, value)| prefix == &candidate && value != namespace)
        {
            return candidate;
        }
    }
    unreachable!("the finite root attribute list cannot occupy every prefix")
}

fn unused_prefix(declarations: &[(String, String)], preferred: &str) -> String {
    for suffix in 0usize.. {
        let candidate = if suffix == 0 {
            preferred.to_string()
        } else {
            format!("{preferred}{suffix}")
        };
        if !declarations.iter().any(|(prefix, _)| prefix == &candidate) {
            return candidate;
        }
    }
    unreachable!("the finite preserved namespace set cannot occupy every prefix")
}

struct PreservationPrefixes {
    carrier: String,
    markup_compatibility: String,
}

impl PreservationPrefixes {
    fn new(declarations: &[(String, String)]) -> Self {
        Self {
            carrier: unused_prefix(declarations, "rdocxPreserve"),
            markup_compatibility: unused_prefix(declarations, "mc"),
        }
    }

    fn namespace_for<'a>(
        &'a self,
        prefix: &str,
        local_declarations: &'a [(String, String)],
        ancestor_scope: &'a [(String, String)],
    ) -> Option<&'a str> {
        local_declarations
            .iter()
            .rev()
            .find(|(candidate, _)| candidate == prefix)
            .or_else(|| {
                ancestor_scope
                    .iter()
                    .rev()
                    .find(|(candidate, _)| candidate == prefix)
            })
            .map(|(_, namespace)| namespace.as_str())
    }

    fn mark_ignorable(
        &self,
        start: BytesStart<'_>,
        ancestor_scope: &[(String, String)],
    ) -> Result<BytesStart<'static>> {
        let name = std::str::from_utf8(start.name().as_ref())?.to_string();
        let mut attributes = Vec::new();
        let mut local_declarations = Vec::new();
        for attribute in start.attributes() {
            let attribute = attribute?;
            let name = std::str::from_utf8(attribute.key.as_ref())?.to_string();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
                .into_owned();
            if let Some(prefix) = name.strip_prefix("xmlns:") {
                local_declarations.push((prefix.to_string(), value.clone()));
            }
            attributes.push((name, value));
        }

        let mut existing_ignorable = None;
        let mut merged_tokens = Vec::new();
        for (index, (name, value)) in attributes.iter().enumerate() {
            let Some((prefix, local)) = name.split_once(':') else {
                continue;
            };
            if local == "Ignorable"
                && self.namespace_for(prefix, &local_declarations, ancestor_scope) == Some(MC_NS)
            {
                existing_ignorable.get_or_insert((index, name.clone()));
                for token in value.split_whitespace() {
                    if !merged_tokens.iter().any(|existing| existing == token) {
                        merged_tokens.push(token.to_string());
                    }
                }
            }
        }
        if !merged_tokens.iter().any(|token| token == &self.carrier) {
            merged_tokens.push(self.carrier.clone());
        }
        let merged_ignorable = merged_tokens.join(" ");

        let mut marked = BytesStart::new(name);
        for (index, (name, value)) in attributes.iter().enumerate() {
            let expanded_ignorable = name.split_once(':').is_some_and(|(prefix, local)| {
                local == "Ignorable"
                    && self.namespace_for(prefix, &local_declarations, ancestor_scope)
                        == Some(MC_NS)
            });
            if expanded_ignorable {
                if existing_ignorable
                    .as_ref()
                    .is_some_and(|(first, _)| *first == index)
                {
                    marked.push_attribute((name.as_str(), merged_ignorable.as_str()));
                }
            } else {
                marked.push_attribute((name.as_str(), value.as_str()));
            }
        }

        let carrier_declaration = format!("xmlns:{}", self.carrier);
        if !attributes
            .iter()
            .any(|(name, _)| name == &carrier_declaration)
        {
            marked.push_attribute((carrier_declaration.as_str(), PRESERVED_PROPERTY_NS));
        }
        if existing_ignorable.is_none() {
            let mc_declaration = format!("xmlns:{}", self.markup_compatibility);
            if !attributes.iter().any(|(name, _)| name == &mc_declaration) {
                marked.push_attribute((mc_declaration.as_str(), MC_NS));
            }
            let ignorable = qualified(&self.markup_compatibility, "Ignorable");
            marked.push_attribute((ignorable.as_str(), merged_ignorable.as_str()));
        }
        Ok(marked)
    }
}

fn append_namespace_declarations(
    attributes: &[(String, String)],
    declarations: &mut Vec<(String, String)>,
) {
    declarations.extend(attributes.iter().filter_map(|(name, value)| {
        name.strip_prefix("xmlns:")
            .map(|prefix| (prefix.to_string(), value.clone()))
    }));
}

fn append_raw_namespace_declarations(
    raw: &[u8],
    declarations: &mut Vec<(String, String)>,
) -> Result<()> {
    let mut reader = Reader::from_reader(raw);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                for attribute in element.attributes() {
                    let attribute = attribute?;
                    let name = std::str::from_utf8(attribute.key.as_ref())?;
                    let Some(prefix) = name.strip_prefix("xmlns:") else {
                        continue;
                    };
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
                        .into_owned();
                    declarations.push((prefix.to_string(), value));
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(error.into()),
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn has_namespace_declaration(
    root_attributes: &[(String, String)],
    prefix: &str,
    namespace: &str,
) -> bool {
    let declaration = format!("xmlns:{prefix}");
    root_attributes
        .iter()
        .any(|(name, value)| name == &declaration && value == namespace)
}

fn prefixed_name(name: &[u8], prefix: &str) -> Result<String> {
    let name = std::str::from_utf8(name)?;
    Ok(match name.strip_prefix("w:") {
        Some(local) => qualified(prefix, local),
        None => name.to_string(),
    })
}

fn prefixed_start(element: &BytesStart<'_>, prefix: &str) -> Result<BytesStart<'static>> {
    let name = prefixed_name(element.name().as_ref(), prefix)?;
    let mut rewritten = BytesStart::new(name);
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = prefixed_name(attribute.key.as_ref(), prefix)?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
            .into_owned();
        rewritten.push_attribute((key.as_str(), value.as_str()));
    }
    Ok(rewritten.into_owned())
}

fn rewrite_generated_prefix(raw: &[u8], prefix: &str) -> Result<Vec<u8>> {
    if prefix == "w" {
        return Ok(raw.to_vec());
    }

    let mut reader = Reader::from_reader(raw);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => {
                writer.write_event(Event::Start(prefixed_start(&element, prefix)?))?;
            }
            Ok(Event::Empty(element)) => {
                writer.write_event(Event::Empty(prefixed_start(&element, prefix)?))?;
            }
            Ok(Event::End(element)) => {
                let name = prefixed_name(element.name().as_ref(), prefix)?;
                writer.write_event(Event::End(BytesEnd::new(name)))?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned())?,
            Err(error) => return Err(error.into()),
        }
        buf.clear();
    }
    Ok(writer.into_inner())
}

fn write_xml_events<W: std::io::Write>(writer: &mut Writer<W>, raw: &[u8]) -> Result<()> {
    let mut reader = Reader::from_reader(raw);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned())?,
            Err(error) => return Err(error.into()),
        }
        buf.clear();
    }
    Ok(())
}

fn write_extras_at<W: std::io::Write>(
    writer: &mut Writer<W>,
    extra_xml: &[(usize, Vec<u8>)],
    position: usize,
) -> Result<()> {
    for (at, raw) in extra_xml {
        if *at == position {
            writer.get_mut().write_all(raw)?;
        }
    }
    Ok(())
}

fn shift_extras_from(extra_xml: &mut [(usize, Vec<u8>)], position: usize) {
    for (at, _) in extra_xml {
        if *at >= position {
            *at += 1;
        }
    }
}

fn capture_root_attributes(start: &BytesStart<'_>) -> Result<Vec<(String, String)>> {
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = std::str::from_utf8(attribute.key.as_ref())?.to_string();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
            .into_owned();
        if (name == "xmlns:w" && value == W_NS) || (name == "xmlns:r" && value == R_NS) {
            continue;
        }
        attributes.push((name, value));
    }
    Ok(attributes)
}

fn capture_extra_attributes(
    start: &BytesStart<'_>,
    modelled: &[&[u8]],
    word_prefixes: &[String],
) -> Result<Vec<(String, String)>> {
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute?;
        if modelled
            .iter()
            .any(|local| is_word_attribute(attribute.key.as_ref(), local, word_prefixes))
        {
            continue;
        }
        let name = std::str::from_utf8(attribute.key.as_ref())?.to_string();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
            .into_owned();
        attributes.push((name, value));
    }
    Ok(attributes)
}

fn namespace_binding(prefix: &str, namespace: &str) -> String {
    format!("\0{prefix}\0{namespace}")
}

fn split_namespace_binding(value: &str) -> Option<(&str, &str)> {
    value.strip_prefix('\0')?.split_once('\0')
}

fn namespace_bindings(scope: &[String]) -> Vec<(String, String)> {
    scope
        .iter()
        .filter_map(|value| split_namespace_binding(value))
        .map(|(prefix, namespace)| (prefix.to_string(), namespace.to_string()))
        .collect()
}

pub(crate) fn word_prefixes_at(
    start: &BytesStart<'_>,
    inherited: &[String],
) -> Result<Vec<String>> {
    let mut prefixes = inherited.to_vec();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        let prefix = if name == b"xmlns" {
            b"".as_slice()
        } else if let Some(prefix) = name.strip_prefix(b"xmlns:") {
            prefix
        } else {
            continue;
        };
        let prefix = std::str::from_utf8(prefix)?.to_string();
        let value =
            attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?;
        prefixes.retain(|candidate| {
            candidate != &prefix
                && split_namespace_binding(candidate)
                    .is_none_or(|(candidate, _)| candidate != prefix)
        });
        prefixes.push(namespace_binding(&prefix, &value));
        if value.as_bytes() == W_NS.as_bytes() {
            prefixes.push(prefix);
        }
    }
    Ok(prefixes)
}

fn is_word_attribute(key: &[u8], local: &[u8], word_prefixes: &[String]) -> bool {
    let Some(separator) = key.iter().position(|byte| *byte == b':') else {
        return false;
    };
    key.get(separator + 1..) == Some(local)
        && word_prefixes
            .iter()
            .any(|prefix| prefix.as_bytes() == &key[..separator])
}

fn is_word_element(name: &[u8], local: &[u8], word_prefixes: &[String]) -> bool {
    let qualified_local = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
    qualified_local == local && is_word_name(name, word_prefixes)
}

fn word_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    word_prefixes: &[String],
) -> Result<Option<String>> {
    let prefixes = word_prefixes_at(element, word_prefixes)?;
    for attribute in element.attributes() {
        let attribute = attribute?;
        if is_word_attribute(attribute.key.as_ref(), name, &prefixes) {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn abstract_raw_boundary(name: &[u8], current: usize, word_prefixes: &[String]) -> (usize, usize) {
    if is_word_element(name, b"nsid", word_prefixes) {
        (0, 1)
    } else if is_word_element(name, b"tmpl", word_prefixes) {
        (2, 3)
    } else if is_word_element(name, b"name", word_prefixes) {
        (3, 4)
    } else if is_word_element(name, b"styleLink", word_prefixes) {
        (4, 5)
    } else if is_word_element(name, b"numStyleLink", word_prefixes) {
        (5, 6)
    } else {
        (current, current)
    }
}

fn root_raw_boundary(name: &[u8], modelled_count: usize, word_prefixes: &[String]) -> usize {
    if is_word_element(name, b"numPicBullet", word_prefixes) {
        0
    } else {
        1 + modelled_count
    }
}

fn push_extra_attributes(start: &mut BytesStart<'_>, attributes: &[(String, String)]) {
    for (name, value) in attributes {
        start.push_attribute((name.as_str(), value.as_str()));
    }
}

fn level_raw_boundary(name: &[u8], current: usize, word_prefixes: &[String]) -> (usize, usize) {
    if is_word_element(name, b"lvlRestart", word_prefixes) {
        (2, 3)
    } else if is_word_element(name, b"pStyle", word_prefixes) {
        (3, 4)
    } else if is_word_element(name, b"isLgl", word_prefixes) {
        (4, 5)
    } else if is_word_element(name, b"suff", word_prefixes) {
        (5, 6)
    } else if is_word_element(name, b"lvlPicBulletId", word_prefixes) {
        (7, 8)
    } else if is_word_element(name, b"legacy", word_prefixes) {
        (8, 9)
    } else {
        (current, current)
    }
}

fn u32_attribute(element: &BytesStart<'_>, name: &[u8], word_prefixes: &[String]) -> Result<u32> {
    Ok(word_attribute_value(element, name, word_prefixes)?
        .map(|value| value.parse())
        .transpose()?
        .unwrap_or(0))
}

fn is_word_name(name: &[u8], word_prefixes: &[String]) -> bool {
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return word_prefixes.iter().any(String::is_empty);
    };
    word_prefixes
        .iter()
        .any(|prefix| prefix.as_bytes() == &name[..separator])
}

fn is_word_property_attribute(name: &[u8], word_prefixes: &[String]) -> bool {
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return false;
    };
    word_prefixes
        .iter()
        .any(|prefix| prefix.as_bytes() == &name[..separator])
}

fn is_relationship_property_attribute(name: &[u8], scope: &[String]) -> bool {
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return false;
    };
    scope.iter().any(|binding| {
        split_namespace_binding(binding).is_some_and(|(prefix, namespace)| {
            prefix.as_bytes() == &name[..separator] && namespace == R_NS
        })
    })
}

fn property_attributes(kind: PropertyKind, local: &[u8]) -> &'static [&'static [u8]] {
    match local {
        b"pStyle"
        | b"keepNext"
        | b"keepLines"
        | b"pageBreakBefore"
        | b"widowControl"
        | b"suppressAutoHyphens"
        | b"jc"
        | b"outlineLvl"
        | b"rStyle"
        | b"b"
        | b"bCs"
        | b"i"
        | b"iCs"
        | b"caps"
        | b"smallCaps"
        | b"strike"
        | b"dstrike"
        | b"vanish"
        | b"w"
        | b"position"
        | b"sz"
        | b"szCs"
        | b"highlight"
        | b"u"
        | b"vertAlign"
        | b"ilvl"
        | b"numId" => &[b"val"],
        b"rFonts" => &[
            b"ascii",
            b"hAnsi",
            b"eastAsia",
            b"cs",
            b"asciiTheme",
            b"hAnsiTheme",
        ],
        b"spacing" if matches!(kind, PropertyKind::Paragraph) => &[
            b"before",
            b"after",
            b"line",
            b"lineRule",
            b"beforeAutospacing",
            b"afterAutospacing",
        ],
        b"spacing" => &[b"val"],
        b"ind" => &[
            b"left",
            b"start",
            b"right",
            b"end",
            b"firstLine",
            b"hanging",
        ],
        b"color" => &[b"val", b"themeColor"],
        b"shd" => &[b"val", b"color", b"fill"],
        b"top" | b"left" | b"start" | b"bottom" | b"right" | b"end" | b"between" | b"bar" => {
            &[b"val", b"sz", b"space", b"color"]
        }
        b"tab" => &[b"val", b"pos", b"leader"],
        b"type" => &[b"val"],
        b"pgSz" => &[b"w", b"h", b"orient"],
        b"pgMar" => &[
            b"top", b"right", b"end", b"bottom", b"left", b"start", b"gutter", b"header", b"footer",
        ],
        b"cols" => &[b"num", b"space", b"equalWidth", b"sep"],
        b"col" => &[b"w", b"space"],
        b"headerReference" | b"footerReference" => &[b"type", b"id"],
        _ => &[],
    }
}

fn nested_property_children(parent: &[u8]) -> &'static [&'static [u8]] {
    match parent {
        b"numPr" => &[b"ilvl", b"numId", b"ins"],
        b"pBdr" => &[
            b"top", b"left", b"start", b"bottom", b"right", b"end", b"between", b"bar",
        ],
        b"tabs" => &[b"tab"],
        b"rPr" => RPR_MODELLED_CHILDREN,
        b"sectPr" => &[
            b"headerReference",
            b"footerReference",
            b"type",
            b"pgSz",
            b"pgMar",
            b"cols",
            b"titlePg",
        ],
        b"cols" => &[b"col"],
        _ => &[],
    }
}

fn nested_property_schema(parent: &[u8]) -> &'static [&'static [u8]] {
    match parent {
        b"numPr" => NUM_PR_CHILDREN,
        b"pBdr" => P_BDR_CHILDREN,
        b"tabs" => TABS_CHILDREN,
        b"rPr" => RPR_CHILDREN,
        b"sectPr" => SECT_PR_CHILDREN,
        b"cols" => &[b"col"],
        _ => &[],
    }
}

fn projected_word_start(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
    kind: PropertyKind,
    root: bool,
    has_producer: &mut bool,
) -> Result<BytesStart<'static>> {
    let local_name = element.local_name();
    let local = std::str::from_utf8(local_name.as_ref())?;
    let name = qualified("w", local);
    let mut projected = BytesStart::new(name);
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        if key.starts_with(b"xmlns") {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?;
            if value.as_bytes() != W_NS.as_bytes() {
                *has_producer = true;
            }
            continue;
        }
        let supported = property_attributes(kind, local_name.as_ref());
        let attribute_local = key.rsplit(|byte| *byte == b':').next().unwrap_or(key);
        let modelled_namespace = is_word_property_attribute(key, word_prefixes)
            || ((local_name.as_ref() == b"headerReference"
                || local_name.as_ref() == b"footerReference")
                && attribute_local == b"id"
                && is_relationship_property_attribute(key, word_prefixes));
        if root || !modelled_namespace || !supported.contains(&attribute_local) {
            *has_producer = true;
            continue;
        }
        let attribute_local = std::str::from_utf8(attribute_local)?;
        let name = qualified("w", attribute_local);
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
            .into_owned();
        projected.push_attribute((name.as_str(), value.as_str()));
    }
    Ok(projected.into_owned())
}

fn project_word_element(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart<'_>,
    inherited_prefixes: &[String],
    writer: &mut Writer<Vec<u8>>,
    kind: PropertyKind,
    depth: usize,
    has_producer: &mut bool,
) -> Result<()> {
    if depth >= MAX_PROPERTY_XML_DEPTH {
        return Err(OxmlError::InvalidValue(format!(
            "property XML depth exceeds {MAX_PROPERTY_XML_DEPTH}"
        )));
    }
    let prefixes = word_prefixes_at(start, inherited_prefixes)?;
    writer.write_event(Event::Start(projected_word_start(
        start,
        &prefixes,
        kind,
        depth == 0,
        has_producer,
    )?))?;
    let parent_local = start.local_name();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) => {
                let child_prefixes = word_prefixes_at(element, &prefixes)?;
                let local = element.local_name();
                let word_name = is_word_name(element.name().as_ref(), &child_prefixes);
                let modelled = if depth == 0 {
                    kind.modelled_children().contains(&local.as_ref())
                } else {
                    nested_property_children(parent_local.as_ref()).contains(&local.as_ref())
                };
                if word_name {
                    if !modelled {
                        *has_producer = true;
                    }
                    project_word_element(
                        reader,
                        element,
                        &prefixes,
                        writer,
                        kind,
                        depth + 1,
                        has_producer,
                    )?;
                } else {
                    *has_producer = true;
                    reader.read_to_end_into(element.name(), &mut Vec::new())?;
                }
            }
            Ok(Event::Empty(ref element)) => {
                let child_prefixes = word_prefixes_at(element, &prefixes)?;
                let local = element.local_name();
                let word_name = is_word_name(element.name().as_ref(), &child_prefixes);
                let modelled = if depth == 0 {
                    kind.modelled_children().contains(&local.as_ref())
                } else {
                    nested_property_children(parent_local.as_ref()).contains(&local.as_ref())
                };
                if word_name {
                    if !modelled {
                        *has_producer = true;
                    }
                    writer.write_event(Event::Empty(projected_word_start(
                        element,
                        &child_prefixes,
                        kind,
                        false,
                        has_producer,
                    )?))?;
                } else {
                    *has_producer = true;
                }
            }
            Ok(Event::End(_)) => {
                let local_name = start.local_name();
                let name = qualified("w", std::str::from_utf8(local_name.as_ref())?);
                writer.write_event(Event::End(BytesEnd::new(name)))?;
                return Ok(());
            }
            Ok(Event::Text(ref text)) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {}
            Ok(Event::Eof) => return Ok(()),
            Ok(_) => *has_producer = true,
            Err(error) => return Err(error.into()),
        }
        buf.clear();
    }
}

fn property_projection(
    raw: &[u8],
    kind: PropertyKind,
    word_prefixes: &[String],
) -> Result<(Vec<u8>, bool)> {
    let mut reader = Reader::from_reader(raw);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut has_producer = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) => {
                project_word_element(
                    &mut reader,
                    element,
                    word_prefixes,
                    &mut writer,
                    kind,
                    0,
                    &mut has_producer,
                )?;
                return Ok((writer.into_inner(), has_producer));
            }
            Ok(Event::Empty(ref element)) => {
                let prefixes = word_prefixes_at(element, word_prefixes)?;
                writer.write_event(Event::Empty(projected_word_start(
                    element,
                    &prefixes,
                    kind,
                    true,
                    &mut has_producer,
                )?))?;
                return Ok((writer.into_inner(), has_producer));
            }
            Ok(Event::Eof) => return Ok((Vec::new(), false)),
            Err(error) => return Err(error.into()),
            _ => {}
        }
        buf.clear();
    }
}

fn ppr_from_raw(raw: &[u8], word_prefixes: &[String]) -> Result<(CT_PPr, bool)> {
    let (_, has_producer) = property_projection(raw, PropertyKind::Paragraph, word_prefixes)?;
    let mut ppr = parse_raw_ppr(raw, word_prefixes)?;
    for (index, tab) in ppr
        .tabs
        .iter_mut()
        .flat_map(|tabs| tabs.tabs.iter_mut())
        .enumerate()
    {
        tab.source_occurrence = Some(index);
    }
    Ok((ppr, has_producer))
}

pub(crate) fn parse_scoped_ppr(raw: &[u8], word_prefixes: &[String]) -> Result<CT_PPr> {
    parse_raw_ppr(raw, word_prefixes)
}

fn parse_raw_ppr(raw: &[u8], word_prefixes: &[String]) -> Result<CT_PPr> {
    let mut reader = Reader::from_reader(raw);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref start)) => {
                let owner_bindings = local_namespace_overrides(start, word_prefixes)?;
                let prefixes = word_prefixes_at(start, word_prefixes)?;
                return CT_PPr::from_xml_with_prefixes_and_owner_bindings(
                    &mut reader,
                    &prefixes,
                    &owner_bindings,
                );
            }
            Ok(Event::Empty(_)) | Ok(Event::Eof) => return Ok(CT_PPr::default()),
            Err(error) => return Err(error.into()),
            _ => {}
        }
        buf.clear();
    }
}

fn rpr_from_raw(raw: &[u8], word_prefixes: &[String]) -> Result<(CT_RPr, bool)> {
    let (_, has_producer) = property_projection(raw, PropertyKind::Run, word_prefixes)?;
    let mut rpr = parse_raw_rpr(raw, word_prefixes)?;
    // Numbering's raw property overlay is the sole preservation source.
    // Keeping the same unmodelled children in the typed projection would
    // duplicate them when canonical properties are merged back into it.
    rpr.revision_xml.clear();
    rpr.revision_xml_positions.clear();
    Ok((rpr, has_producer))
}

fn parse_raw_rpr(raw: &[u8], word_prefixes: &[String]) -> Result<CT_RPr> {
    parse_raw_rpr_with_owner_bindings(raw, word_prefixes, &[])
}

fn parse_raw_rpr_with_owner_bindings(
    raw: &[u8],
    word_prefixes: &[String],
    inherited_owner_bindings: &[(String, String)],
) -> Result<CT_RPr> {
    let mut reader = Reader::from_reader(raw);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref start)) => {
                let local_bindings = local_namespace_overrides(start, word_prefixes)?;
                let owner_bindings =
                    merged_owner_bindings(inherited_owner_bindings, &local_bindings);
                let prefixes = word_prefixes_at(start, word_prefixes)?;
                return CT_RPr::from_xml_with_prefixes_and_owner_bindings(
                    &mut reader,
                    &prefixes,
                    &owner_bindings,
                );
            }
            Ok(Event::Empty(_)) | Ok(Event::Eof) => return Ok(CT_RPr::default()),
            Err(error) => return Err(error.into()),
            _ => {}
        }
        buf.clear();
    }
}

pub(crate) fn local_namespace_overrides(
    start: &BytesStart<'_>,
    inherited: &[String],
) -> Result<Vec<(String, String)>> {
    let mut bindings = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        let prefix = if name == b"xmlns" {
            ""
        } else if let Some(prefix) = name.strip_prefix(b"xmlns:") {
            std::str::from_utf8(prefix)?
        } else {
            continue;
        };
        let namespace = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
            .into_owned();
        let inherited_namespace = inherited.iter().find_map(|binding| {
            split_namespace_binding(binding)
                .filter(|(candidate, _)| *candidate == prefix)
                .map(|(_, namespace)| namespace)
                .or_else(|| (binding == prefix).then_some(W_NS))
        });
        if inherited_namespace != Some(namespace.as_str()) {
            bindings.push((prefix.to_owned(), namespace));
        }
    }
    Ok(bindings)
}

pub(crate) fn parse_scoped_rpr(raw: &[u8], word_prefixes: &[String]) -> Result<CT_RPr> {
    parse_raw_rpr(raw, word_prefixes)
}

pub(crate) fn parse_scoped_rpr_with_owner_bindings(
    raw: &[u8],
    word_prefixes: &[String],
    owner_bindings: &[(String, String)],
) -> Result<CT_RPr> {
    parse_raw_rpr_with_owner_bindings(raw, word_prefixes, owner_bindings)
}

pub(crate) fn merged_owner_bindings(
    inherited: &[(String, String)],
    local: &[(String, String)],
) -> Vec<(String, String)> {
    let mut bindings = inherited.to_vec();
    for (prefix, namespace) in local {
        bindings.retain(|(candidate, _)| candidate != prefix);
        bindings.push((prefix.clone(), namespace.clone()));
    }
    bindings
}

fn generated_property_children(raw: &[u8], kind: PropertyKind) -> Result<Vec<(usize, Vec<u8>)>> {
    generated_children(raw, kind.children())
}

fn generated_children(raw: &[u8], schema_children: &[&[u8]]) -> Result<Vec<(usize, Vec<u8>)>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(true);
    let mut children = Vec::new();
    let mut buf = Vec::new();
    let mut inside = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(_)) if !inside => inside = true,
            Ok(Event::Empty(_)) if !inside => return Ok(children),
            Ok(Event::Start(ref element)) if inside => {
                let local = element.local_name();
                let position = schema_children
                    .iter()
                    .position(|candidate| *candidate == local.as_ref())
                    .expect("typed property writers emit only modelled children");
                children.push((position, capture_element(&mut reader, element)?));
            }
            Ok(Event::Empty(ref element)) if inside => {
                let local = element.local_name();
                let position = schema_children
                    .iter()
                    .position(|candidate| *candidate == local.as_ref())
                    .expect("typed property writers emit only modelled children");
                children.push((position, capture_empty_element(element)?));
            }
            Ok(Event::End(_)) | Ok(Event::Eof) => return Ok(children),
            Err(error) => return Err(error.into()),
            _ => {}
        }
        buf.clear();
    }
}

struct PropertyOverlay {
    start: BytesStart<'static>,
    end_name: String,
    extras: Vec<(usize, Vec<u8>)>,
    modelled: Vec<(usize, Vec<u8>)>,
}

fn property_overlay(
    raw: &[u8],
    kind: PropertyKind,
    word_prefixes: &[String],
) -> Result<PropertyOverlay> {
    property_overlay_with_children(
        raw,
        kind.local_name(),
        kind.children(),
        kind.modelled_children(),
        word_prefixes,
    )
}

fn property_overlay_with_children(
    raw: &[u8],
    local_name: &str,
    schema_children: &[&[u8]],
    modelled_children: &[&[u8]],
    word_prefixes: &[String],
) -> Result<PropertyOverlay> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let (start, end_name, empty) = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => {
                let end_name = std::str::from_utf8(element.name().as_ref())?.to_string();
                break (element.into_owned(), end_name, false);
            }
            Ok(Event::Empty(element)) => {
                let end_name = std::str::from_utf8(element.name().as_ref())?.to_string();
                break (element.into_owned(), end_name, true);
            }
            Ok(Event::Eof) => {
                let name = qualified("w", local_name);
                break (BytesStart::new(name.clone()), name, true);
            }
            Err(error) => return Err(error.into()),
            _ => {}
        }
        buf.clear();
    };

    let root_prefixes = word_prefixes_at(&start, word_prefixes)?;
    let mut extras = Vec::new();
    let mut modelled = Vec::new();
    if !empty {
        let mut pending = Vec::new();
        let mut last_modelled_position = None;
        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) => {
                    let prefixes = word_prefixes_at(element, &root_prefixes)?;
                    if is_word_name(element.name().as_ref(), &prefixes)
                        && let Some(position) = modelled_children.iter().find_map(|local| {
                            (*local == element.local_name().as_ref()).then(|| {
                                schema_children
                                    .iter()
                                    .position(|schema| schema == local)
                                    .unwrap()
                            })
                        })
                    {
                        extras.extend(pending.drain(..).map(|raw| (position, raw)));
                        modelled.push((position, capture_element(&mut reader, element)?));
                        last_modelled_position = Some(position);
                    } else if is_word_name(element.name().as_ref(), &prefixes)
                        && let Some(position) = schema_children
                            .iter()
                            .position(|local| *local == element.local_name().as_ref())
                    {
                        extras.extend(pending.drain(..).map(|raw| (position, raw)));
                        extras.push((position, capture_element(&mut reader, element)?));
                        last_modelled_position = Some(position);
                    } else {
                        pending.push(capture_element(&mut reader, element)?);
                    }
                }
                Ok(Event::Empty(ref element)) => {
                    let prefixes = word_prefixes_at(element, &root_prefixes)?;
                    if is_word_name(element.name().as_ref(), &prefixes)
                        && let Some(position) = modelled_children.iter().find_map(|local| {
                            (*local == element.local_name().as_ref()).then(|| {
                                schema_children
                                    .iter()
                                    .position(|schema| schema == local)
                                    .unwrap()
                            })
                        })
                    {
                        extras.extend(pending.drain(..).map(|raw| (position, raw)));
                        modelled.push((position, capture_empty_element(element)?));
                        last_modelled_position = Some(position);
                    } else if is_word_name(element.name().as_ref(), &prefixes)
                        && let Some(position) = schema_children
                            .iter()
                            .position(|local| *local == element.local_name().as_ref())
                    {
                        extras.extend(pending.drain(..).map(|raw| (position, raw)));
                        extras.push((position, capture_empty_element(element)?));
                        last_modelled_position = Some(position);
                    } else {
                        pending.push(capture_empty_element(element)?);
                    }
                }
                Ok(Event::End(_)) | Ok(Event::Eof) => {
                    let position = last_modelled_position.map_or(0, |position| position + 1);
                    extras.extend(pending.drain(..).map(|raw| (position, raw)));
                    break;
                }
                Err(error) => return Err(error.into()),
                _ => {}
            }
        }
    }

    Ok(PropertyOverlay {
        start,
        end_name,
        extras,
        modelled,
    })
}

fn merge_property_child_attributes(
    original: &[u8],
    generated: &[u8],
    word_prefixes: &[String],
    kind: PropertyKind,
) -> Result<Vec<u8>> {
    let mut original_reader = Reader::from_reader(original);
    let mut original_buf = Vec::new();
    let original_start = loop {
        match original_reader.read_event_into(&mut original_buf) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => break element.into_owned(),
            Ok(Event::Eof) => return Ok(generated.to_vec()),
            Err(error) => return Err(error.into()),
            _ => {}
        }
        original_buf.clear();
    };
    let prefixes = word_prefixes_at(&original_start, word_prefixes)?;
    let local = original_start.local_name();
    let supported = property_attributes(kind, local.as_ref());
    let mut preserved_attributes = Vec::new();
    for attribute in original_start.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let attribute_local = key.rsplit(|byte| *byte == b':').next().unwrap_or(key);
        if key.starts_with(b"xmlns")
            || !is_word_property_attribute(key, &prefixes)
            || !supported.contains(&attribute_local)
        {
            let name = std::str::from_utf8(key)?.to_string();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, original_start.decoder())?
                .into_owned();
            preserved_attributes.push((name, value));
        }
    }

    let mut reader = Reader::from_reader(generated);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut first = true;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(mut element)) if first => {
                push_extra_attributes(&mut element, &preserved_attributes);
                writer.write_event(Event::Start(element.into_owned()))?;
                first = false;
            }
            Ok(Event::Empty(mut element)) if first => {
                push_extra_attributes(&mut element, &preserved_attributes);
                writer.write_event(Event::Empty(element.into_owned()))?;
                first = false;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned())?,
            Err(error) => return Err(error.into()),
        }
        buf.clear();
    }
    Ok(writer.into_inner())
}

fn first_element_start(raw: &[u8]) -> Result<(BytesStart<'static>, bool)> {
    let mut reader = Reader::from_reader(raw);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => return Ok((element.into_owned(), false)),
            Ok(Event::Empty(element)) => return Ok((element.into_owned(), true)),
            Ok(Event::Eof) => {
                return Err(OxmlError::InvalidValue(
                    "property child contains no element".to_string(),
                ));
            }
            Err(error) => return Err(error.into()),
            _ => {}
        }
        buf.clear();
    }
}

fn element_prefix(start: &BytesStart<'_>) -> Result<String> {
    let name = start.name();
    let name = name.as_ref();
    let separator = name.iter().position(|byte| *byte == b':').ok_or_else(|| {
        OxmlError::InvalidValue("generated property QName has no prefix".to_string())
    })?;
    Ok(std::str::from_utf8(&name[..separator])?.to_string())
}

fn producer_leaf_payload(
    raw: &[u8],
    word_prefixes: &[String],
    kind: PropertyKind,
    preservation_prefixes: &PreservationPrefixes,
) -> Result<Option<Vec<u8>>> {
    let (original_start, _) = first_element_start(raw)?;
    let prefixes = word_prefixes_at(&original_start, word_prefixes)?;
    let local = original_start.local_name();
    let local_name = std::str::from_utf8(local.as_ref())?;
    let supported = property_attributes(kind, local.as_ref());
    let mut declarations = Vec::new();
    let mut producer_attributes = Vec::new();
    for attribute in original_start.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let name = std::str::from_utf8(key)?.to_string();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, original_start.decoder())?
            .into_owned();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            declarations.push((name, value));
            continue;
        }
        let attribute_local = key.rsplit(|byte| *byte == b':').next().unwrap_or(key);
        if !is_word_property_attribute(key, &prefixes) || !supported.contains(&attribute_local) {
            producer_attributes.push((name, value));
        }
    }
    let overlay = property_overlay_with_children(raw, local_name, &[], &[], word_prefixes)?;
    let nested = overlay
        .extras
        .iter()
        .map(|(_, child)| child.as_slice())
        .collect::<Vec<_>>();
    if producer_attributes.is_empty() && nested.is_empty() {
        return Ok(None);
    }
    let carrier_name = qualified(&preservation_prefixes.carrier, PRESERVED_PROPERTY_LOCAL);
    let mut start = BytesStart::new(carrier_name.clone());
    push_extra_attributes(&mut start, &declarations);
    push_extra_attributes(&mut start, &producer_attributes);
    let mut writer = Writer::new(Vec::new());
    if nested.is_empty() {
        writer.write_event(Event::Empty(start))?;
    } else {
        writer.write_event(Event::Start(start))?;
        for child in nested {
            writer.get_mut().extend_from_slice(child);
        }
        writer.write_event(Event::End(BytesEnd::new(carrier_name)))?;
    }
    Ok(Some(writer.into_inner()))
}

fn producer_only_property(
    raw: &[u8],
    word_prefixes: &[String],
    kind: PropertyKind,
    output_word_prefix: &str,
    depth: usize,
    preservation_prefixes: &PreservationPrefixes,
) -> Result<Option<Vec<u8>>> {
    if depth >= MAX_PROPERTY_XML_DEPTH {
        return Err(OxmlError::InvalidValue(format!(
            "property XML depth exceeds {MAX_PROPERTY_XML_DEPTH}"
        )));
    }
    let (original_start, _) = first_element_start(raw)?;
    let prefixes = word_prefixes_at(&original_start, word_prefixes)?;
    let local = original_start.local_name();
    let local_name = std::str::from_utf8(local.as_ref())?;
    if local.as_ref() == b"tabs" {
        return producer_only_occurrence_property(
            raw,
            word_prefixes,
            kind,
            output_word_prefix,
            depth,
            preservation_prefixes,
        );
    }
    let supported = property_attributes(kind, local.as_ref());
    let mut start = BytesStart::new(qualified(output_word_prefix, local_name));
    let mut has_producer = false;
    for attribute in original_start.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let attribute_local = key.rsplit(|byte| *byte == b':').next().unwrap_or(key);
        let namespace = key == b"xmlns" || key.starts_with(b"xmlns:");
        let producer = !namespace
            && (!is_word_property_attribute(key, &prefixes)
                || !supported.contains(&attribute_local));
        if namespace || producer {
            let name = std::str::from_utf8(key)?.to_string();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, original_start.decoder())?
                .into_owned();
            start.push_attribute((name.as_str(), value.as_str()));
            has_producer |= producer;
        }
    }

    let (schema_children, modelled_children) = if local.as_ref() == kind.local_name().as_bytes() {
        (kind.children(), kind.modelled_children())
    } else {
        (
            nested_property_schema(local.as_ref()),
            nested_property_children(local.as_ref()),
        )
    };
    if schema_children.is_empty() && local.as_ref() != kind.local_name().as_bytes() {
        return producer_leaf_payload(raw, word_prefixes, kind, preservation_prefixes);
    }
    let overlay = property_overlay_with_children(
        raw,
        local_name,
        schema_children,
        modelled_children,
        word_prefixes,
    )?;
    let mut projected_children = Vec::new();
    for (position, child) in &overlay.modelled {
        if let Some(projected) = producer_only_property(
            child,
            &prefixes,
            kind,
            output_word_prefix,
            depth + 1,
            preservation_prefixes,
        )? {
            projected_children.push((*position, projected));
        }
    }
    has_producer |= !overlay.extras.is_empty() || !projected_children.is_empty();
    if !has_producer {
        return Ok(None);
    }

    let end_name = qualified(output_word_prefix, local_name);
    let mut writer = Writer::new(Vec::new());
    if !projected_children.is_empty() {
        start = preservation_prefixes.mark_ignorable(start, &namespace_bindings(word_prefixes))?;
    }
    if overlay.extras.is_empty() && projected_children.is_empty() {
        writer.write_event(Event::Empty(start))?;
        return Ok(Some(writer.into_inner()));
    }
    writer.write_event(Event::Start(start))?;
    for boundary in 0..=schema_children.len() {
        for (_, extra) in overlay
            .extras
            .iter()
            .filter(|(position, _)| *position == boundary)
        {
            writer.get_mut().extend_from_slice(extra);
        }
        for (_, projected) in projected_children
            .iter()
            .filter(|(position, _)| *position == boundary)
        {
            writer.get_mut().extend_from_slice(projected);
        }
    }
    writer.write_event(Event::End(BytesEnd::new(end_name)))?;
    Ok(Some(writer.into_inner()))
}

struct OccurrenceOverlay {
    word_prefixes: Vec<String>,
    modelled: Vec<Vec<u8>>,
    extras: Vec<(usize, Vec<u8>)>,
}

fn occurrence_overlay(raw: &[u8], word_prefixes: &[String]) -> Result<OccurrenceOverlay> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let start = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => break element.into_owned(),
            Ok(Event::Eof) => {
                return Err(OxmlError::InvalidValue(
                    "repeated property contains no element".to_string(),
                ));
            }
            Err(error) => return Err(error.into()),
            _ => {}
        }
        buf.clear();
    };
    let prefixes = word_prefixes_at(&start, word_prefixes)?;
    let mut modelled = Vec::new();
    let mut extras = Vec::new();
    let mut pending = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) => {
                let child_prefixes = word_prefixes_at(element, &prefixes)?;
                if is_word_element(element.name().as_ref(), b"tab", &child_prefixes) {
                    extras.extend(pending.drain(..).map(|raw| (modelled.len(), raw)));
                    modelled.push(capture_element(&mut reader, element)?);
                } else {
                    pending.push(capture_element(&mut reader, element)?);
                }
            }
            Ok(Event::Empty(ref element)) => {
                let child_prefixes = word_prefixes_at(element, &prefixes)?;
                if is_word_element(element.name().as_ref(), b"tab", &child_prefixes) {
                    extras.extend(pending.drain(..).map(|raw| (modelled.len(), raw)));
                    modelled.push(capture_empty_element(element)?);
                } else {
                    pending.push(capture_empty_element(element)?);
                }
            }
            Ok(Event::End(_)) | Ok(Event::Eof) => {
                extras.extend(pending.drain(..).map(|raw| (modelled.len(), raw)));
                break;
            }
            Err(error) => return Err(error.into()),
            _ => {}
        }
    }
    Ok(OccurrenceOverlay {
        word_prefixes: prefixes,
        modelled,
        extras,
    })
}

fn producer_only_occurrence_property(
    raw: &[u8],
    word_prefixes: &[String],
    kind: PropertyKind,
    output_word_prefix: &str,
    depth: usize,
    preservation_prefixes: &PreservationPrefixes,
) -> Result<Option<Vec<u8>>> {
    if depth >= MAX_PROPERTY_XML_DEPTH {
        return Err(OxmlError::InvalidValue(format!(
            "property XML depth exceeds {MAX_PROPERTY_XML_DEPTH}"
        )));
    }
    let (original_start, _) = first_element_start(raw)?;
    let local = original_start.local_name();
    let local_name = std::str::from_utf8(local.as_ref())?;
    let prefixes = word_prefixes_at(&original_start, word_prefixes)?;
    let supported = property_attributes(kind, local.as_ref());
    let mut start = BytesStart::new(qualified(output_word_prefix, local_name));
    let mut has_producer = false;
    for attribute in original_start.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let attribute_local = key.rsplit(|byte| *byte == b':').next().unwrap_or(key);
        let namespace = key == b"xmlns" || key.starts_with(b"xmlns:");
        let producer = !namespace
            && (!is_word_property_attribute(key, &prefixes)
                || !supported.contains(&attribute_local));
        if namespace || producer {
            let name = std::str::from_utf8(key)?.to_string();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, original_start.decoder())?
                .into_owned();
            start.push_attribute((name.as_str(), value.as_str()));
            has_producer |= producer;
        }
    }
    let overlay = occurrence_overlay(raw, word_prefixes)?;
    let mut projected = Vec::with_capacity(overlay.modelled.len());
    for child in &overlay.modelled {
        projected.push(producer_leaf_payload(
            child,
            &overlay.word_prefixes,
            kind,
            preservation_prefixes,
        )?);
    }
    has_producer |= !overlay.extras.is_empty() || projected.iter().any(Option::is_some);
    if !has_producer {
        return Ok(None);
    }
    let end_name = qualified(output_word_prefix, local_name);
    let mut writer = Writer::new(Vec::new());
    if projected.iter().any(Option::is_some) {
        start = preservation_prefixes.mark_ignorable(start, &namespace_bindings(word_prefixes))?;
    }
    writer.write_event(Event::Start(start))?;
    let mut extras_by_boundary = vec![Vec::new(); overlay.modelled.len() + 1];
    for (boundary, raw) in &overlay.extras {
        extras_by_boundary[*boundary].push(raw.as_slice());
    }
    for (boundary, extras) in extras_by_boundary.iter().enumerate() {
        for raw in extras {
            writer.get_mut().extend_from_slice(raw);
        }
        if let Some(Some(payload)) = projected.get(boundary) {
            writer.get_mut().extend_from_slice(payload);
        }
    }
    writer.write_event(Event::End(BytesEnd::new(end_name)))?;
    Ok(Some(writer.into_inner()))
}

fn match_occurrences(
    original_len: usize,
    generated_sources: &[Option<usize>],
) -> (Vec<Option<usize>>, usize) {
    let mut used = vec![false; original_len];
    let matches = generated_sources
        .iter()
        .map(|source| {
            let source = source.filter(|index| *index < original_len && !used[*index]);
            if let Some(index) = source {
                used[index] = true;
            }
            source
        })
        .collect();
    let work = original_len + generated_sources.len();
    (matches, work)
}

fn merge_repeated_property(
    original: &[u8],
    generated: &[u8],
    word_prefixes: &[String],
    kind: PropertyKind,
    depth: usize,
    generated_sources: Option<&[Option<usize>]>,
    preservation_prefixes: &PreservationPrefixes,
) -> Result<Vec<u8>> {
    Ok(merge_repeated_property_with_work(
        original,
        generated,
        word_prefixes,
        kind,
        depth,
        generated_sources,
        preservation_prefixes,
    )?
    .0)
}

fn merge_repeated_property_with_work(
    original: &[u8],
    generated: &[u8],
    word_prefixes: &[String],
    kind: PropertyKind,
    depth: usize,
    generated_sources: Option<&[Option<usize>]>,
    preservation_prefixes: &PreservationPrefixes,
) -> Result<(Vec<u8>, usize)> {
    let overlay = occurrence_overlay(original, word_prefixes)?;
    let merged_generated =
        merge_property_child_attributes(original, generated, word_prefixes, kind)?;
    let (merged_start, _) = first_element_start(&merged_generated)?;
    let generated_tabs = generated_children(generated, TABS_CHILDREN)?;
    let default_sources = vec![None; generated_tabs.len()];
    let generated_sources = generated_sources.unwrap_or(&default_sources);
    let (matches, mut work) = match_occurrences(overlay.modelled.len(), generated_sources);
    let mut extras_by_boundary = vec![Vec::new(); overlay.modelled.len() + 1];
    for (boundary, raw) in &overlay.extras {
        extras_by_boundary[*boundary].push(raw.as_slice());
        work += 1;
    }
    let mut writer = Writer::new(Vec::new());
    let end_name = std::str::from_utf8(merged_start.name().as_ref())?.to_string();
    writer.write_event(Event::Start(merged_start))?;
    let mut next_boundary = 0;
    for ((_, generated_tab), original_index) in generated_tabs.iter().zip(matches) {
        work += 1;
        if let Some(index) = original_index {
            while next_boundary <= index {
                for raw in &extras_by_boundary[next_boundary] {
                    writer.get_mut().extend_from_slice(raw);
                    work += 1;
                }
                next_boundary += 1;
                work += 1;
            }
            let merged = merge_property_child(
                &overlay.modelled[index],
                generated_tab,
                &overlay.word_prefixes,
                kind,
                depth + 1,
                None,
                preservation_prefixes,
            )?;
            write_xml_events(&mut writer, &merged)?;
        } else {
            write_xml_events(&mut writer, generated_tab)?;
        }
    }
    while next_boundary < extras_by_boundary.len() {
        for raw in &extras_by_boundary[next_boundary] {
            writer.get_mut().extend_from_slice(raw);
            work += 1;
        }
        next_boundary += 1;
        work += 1;
    }
    writer.write_event(Event::End(BytesEnd::new(end_name)))?;
    debug_assert!(
        work <= 4 * (overlay.modelled.len() + generated_tabs.len() + overlay.extras.len() + 1)
    );
    Ok((writer.into_inner(), work))
}

fn merge_property_child(
    original: &[u8],
    generated: &[u8],
    word_prefixes: &[String],
    kind: PropertyKind,
    depth: usize,
    tab_sources: Option<&[Option<usize>]>,
    preservation_prefixes: &PreservationPrefixes,
) -> Result<Vec<u8>> {
    if depth >= MAX_PROPERTY_XML_DEPTH {
        return Err(OxmlError::InvalidValue(format!(
            "property XML depth exceeds {MAX_PROPERTY_XML_DEPTH}"
        )));
    }

    let (original_start, _) = first_element_start(original)?;
    let original_prefixes = word_prefixes_at(&original_start, word_prefixes)?;
    let local = original_start.local_name();
    if local.as_ref() == b"tabs" {
        return merge_repeated_property(
            original,
            generated,
            word_prefixes,
            kind,
            depth,
            tab_sources,
            preservation_prefixes,
        );
    }
    let schema_children = nested_property_schema(local.as_ref());
    if schema_children.is_empty() {
        return merge_property_child_attributes(original, generated, word_prefixes, kind);
    }

    let modelled_children = nested_property_children(local.as_ref());
    let local_name = std::str::from_utf8(local.as_ref())?;
    let overlay = property_overlay_with_children(
        original,
        local_name,
        schema_children,
        modelled_children,
        word_prefixes,
    )?;
    let merged_generated =
        merge_property_child_attributes(original, generated, word_prefixes, kind)?;
    let (merged_start, _) = first_element_start(&merged_generated)?;
    let output_word_prefix = element_prefix(&merged_start)?;
    let end_name = std::str::from_utf8(merged_start.name().as_ref())?.to_string();
    let children = generated_children(generated, schema_children)?;
    let mut used_original = vec![false; overlay.modelled.len()];
    let matches = children
        .iter()
        .map(|(position, _)| {
            let index = overlay
                .modelled
                .iter()
                .enumerate()
                .find(|(index, (original_position, _))| {
                    !used_original[*index] && original_position == position
                })
                .map(|(index, _)| index);
            if let Some(index) = index {
                used_original[index] = true;
            }
            index
        })
        .collect::<Vec<_>>();
    let mut retained = Vec::new();
    for (index, (position, raw)) in overlay.modelled.iter().enumerate() {
        if !used_original[index]
            && let Some(projected) = producer_only_property(
                raw,
                &original_prefixes,
                kind,
                &output_word_prefix,
                depth + 1,
                preservation_prefixes,
            )?
        {
            retained.push((*position, projected));
        }
    }
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Start(merged_start))?;
    let mut next_boundary = 0;
    for ((position, child), original_index) in children.iter().zip(matches) {
        for boundary in next_boundary..=*position {
            write_extras_at(&mut writer, &overlay.extras, boundary)?;
            for (_, projected) in retained
                .iter()
                .filter(|(retained_position, _)| *retained_position == boundary)
            {
                writer.get_mut().extend_from_slice(projected);
            }
        }
        if let Some(index) = original_index {
            let merged = merge_property_child(
                &overlay.modelled[index].1,
                child,
                &original_prefixes,
                kind,
                depth + 1,
                tab_sources,
                preservation_prefixes,
            )?;
            write_xml_events(&mut writer, &merged)?;
        } else {
            write_xml_events(&mut writer, child)?;
        }
        next_boundary = position + 1;
    }
    for boundary in next_boundary..=schema_children.len() {
        write_extras_at(&mut writer, &overlay.extras, boundary)?;
        for (_, projected) in retained
            .iter()
            .filter(|(retained_position, _)| *retained_position == boundary)
        {
            writer.get_mut().extend_from_slice(projected);
        }
    }
    writer.write_event(Event::End(BytesEnd::new(end_name)))?;
    Ok(writer.into_inner())
}

fn write_level_property<W: std::io::Write>(
    writer: &mut Writer<W>,
    property: LevelProperty<'_>,
    preservation: Option<(LevelProperty<'_>, &[u8], &[String])>,
    word_prefix: &str,
    preservation_prefixes: &PreservationPrefixes,
) -> Result<()> {
    let tab_sources = property.tab_sources();
    let Some((original, raw, original_prefixes)) = preservation else {
        let generated = rewrite_generated_prefix(&property.canonical_xml()?, word_prefix)?;
        write_xml_events(writer, &generated)?;
        return Ok(());
    };

    let overlay = property_overlay(raw, property.kind(), original_prefixes)?;
    let generated = rewrite_generated_prefix(&property.canonical_xml()?, word_prefix)?;
    let children = generated_property_children(&generated, property.kind())?;
    let original_generated = rewrite_generated_prefix(&original.canonical_xml()?, word_prefix)?;
    let original_children = generated_property_children(&original_generated, property.kind())?;
    let tab_provenance_unchanged = tab_sources == original.tab_sources();
    let mut used_original = vec![false; overlay.modelled.len()];
    let matches = children
        .iter()
        .map(|(position, _)| {
            let index = overlay
                .modelled
                .iter()
                .enumerate()
                .find(|(index, (original_position, _))| {
                    !used_original[*index] && original_position == position
                })
                .map(|(index, _)| index);
            if let Some(index) = index {
                used_original[index] = true;
            }
            index
        })
        .collect::<Vec<_>>();
    let mut retained = Vec::new();
    for (index, (position, raw)) in overlay.modelled.iter().enumerate() {
        if !used_original[index]
            && let Some(projected) = producer_only_property(
                raw,
                original_prefixes,
                property.kind(),
                word_prefix,
                0,
                preservation_prefixes,
            )?
        {
            retained.push((*position, projected));
        }
    }
    writer.write_event(Event::Start(overlay.start))?;
    let mut next_boundary = 0;
    for ((position, child), original_index) in children.iter().zip(matches) {
        for boundary in next_boundary..=*position {
            write_extras_at(writer, &overlay.extras, boundary)?;
            for (_, projected) in retained
                .iter()
                .filter(|(retained_position, _)| *retained_position == boundary)
            {
                write_xml_events(writer, projected)?;
            }
        }
        let original_raw = original_index.map(|index| overlay.modelled[index].1.as_slice());
        let original_canonical = original_children
            .iter()
            .find(|(original_position, _)| original_position == position)
            .map(|(_, raw)| raw.as_slice());
        let child_is_tabs = first_element_start(child)?.0.local_name().as_ref() == b"tabs";
        if let Some(original_raw) = original_raw
            && original_canonical == Some(child.as_slice())
            && (!child_is_tabs || tab_provenance_unchanged)
        {
            writer.get_mut().write_all(original_raw)?;
        } else if let Some(original_raw) = original_raw {
            let merged = merge_property_child(
                original_raw,
                child,
                original_prefixes,
                property.kind(),
                0,
                tab_sources.as_deref(),
                preservation_prefixes,
            )?;
            write_xml_events(writer, &merged)?;
        } else {
            write_xml_events(writer, child)?;
        }
        next_boundary = position + 1;
    }
    for boundary in next_boundary..=property.kind().children().len() {
        write_extras_at(writer, &overlay.extras, boundary)?;
        for (_, projected) in retained
            .iter()
            .filter(|(retained_position, _)| *retained_position == boundary)
        {
            write_xml_events(writer, projected)?;
        }
    }
    writer.write_event(Event::End(BytesEnd::new(overlay.end_name)))?;
    Ok(())
}

/// `ST_NumberFormat` — Numbering format type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ST_NumberFormat {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
    Ordinal,
    Bullet,
    None,
    /// A producer-defined format value that rdocx does not model.
    Other(String),
}

/// The character or layout item that follows a numbering level marker.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ST_LvlSuffix {
    Tab,
    Space,
    Nothing,
}

impl ST_LvlSuffix {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "tab" => Some(Self::Tab),
            "space" => Some(Self::Space),
            "nothing" => Some(Self::Nothing),
            _ => None,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Tab => "tab",
            Self::Space => "space",
            Self::Nothing => "nothing",
        }
    }
}

impl ST_NumberFormat {
    pub fn from_str(s: &str) -> Self {
        match s {
            "decimal" => Self::Decimal,
            "upperRoman" => Self::UpperRoman,
            "lowerRoman" => Self::LowerRoman,
            "upperLetter" => Self::UpperLetter,
            "lowerLetter" => Self::LowerLetter,
            "ordinal" => Self::Ordinal,
            "bullet" => Self::Bullet,
            "none" => Self::None,
            _ => Self::Other(s.to_owned()),
        }
    }

    pub fn to_str(&self) -> &str {
        match self {
            Self::Decimal => "decimal",
            Self::UpperRoman => "upperRoman",
            Self::LowerRoman => "lowerRoman",
            Self::UpperLetter => "upperLetter",
            Self::LowerLetter => "lowerLetter",
            Self::Ordinal => "ordinal",
            Self::Bullet => "bullet",
            Self::None => "none",
            Self::Other(value) => value,
        }
    }
}

/// `CT_Lvl` — A single level (0–8) in an abstract numbering definition.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Lvl {
    /// Level index (0–8)
    pub ilvl: u32,
    /// Starting number
    pub start: Option<u32>,
    /// Number format
    pub num_fmt: Option<ST_NumberFormat>,
    /// Item emitted between the marker and the paragraph content.
    pub suffix: Option<ST_LvlSuffix>,
    /// Level text (e.g., "%1.", "%1.%2.", bullet char)
    pub lvl_text: Option<String>,
    /// Level justification
    pub lvl_jc: Option<ST_Jc>,
    /// Paragraph properties for this level (typically indentation)
    pub ppr: Option<CT_PPr>,
    /// Run properties for the numbering symbol
    pub rpr: Option<CT_RPr>,
    /// Unmodelled children retained at their modelled-child boundaries.
    pub extra_xml: Vec<(usize, Vec<u8>)>,
    /// Unmodelled attributes and namespace declarations from `w:lvl`.
    pub extra_attributes: Vec<(String, String)>,
    /// Original typed value, raw XML, and namespace scope for an extended `w:pPr`.
    pub ppr_raw: Option<(CT_PPr, Vec<u8>, Vec<String>)>,
    /// Original typed value, raw XML, and namespace scope for an extended `w:rPr`.
    pub rpr_raw: Option<(CT_RPr, Vec<u8>, Vec<String>)>,
}

#[allow(non_snake_case)]
impl CT_Lvl {
    pub fn new(ilvl: u32) -> Self {
        CT_Lvl {
            ilvl,
            start: None,
            num_fmt: None,
            suffix: None,
            lvl_text: None,
            lvl_jc: None,
            ppr: None,
            rpr: None,
            extra_xml: Vec::new(),
            extra_attributes: Vec::new(),
            ppr_raw: None,
            rpr_raw: None,
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>, ilvl: u32) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, ilvl, &["w".to_string()])
    }

    fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        ilvl: u32,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut lvl = CT_Lvl::new(ilvl);
        let mut buf = Vec::new();
        let mut boundary = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"start", &prefixes) {
                        if let Some(value) = word_attribute_value(e, b"val", &prefixes)? {
                            lvl.start = Some(value.parse()?);
                        }
                        reader.read_to_end_into(name, &mut Vec::new())?;
                        boundary = 1;
                    } else if is_word_element(name.as_ref(), b"numFmt", &prefixes) {
                        if let Some(value) = word_attribute_value(e, b"val", &prefixes)? {
                            lvl.num_fmt = Some(ST_NumberFormat::from_str(&value));
                        }
                        reader.read_to_end_into(name, &mut Vec::new())?;
                        boundary = 2;
                    } else if is_word_element(name.as_ref(), b"suff", &prefixes) {
                        if let Some(value) = word_attribute_value(e, b"val", &prefixes)?
                            && let Some(suffix) = ST_LvlSuffix::from_str(&value)
                        {
                            lvl.suffix = Some(suffix);
                        } else {
                            let (at, next) = level_raw_boundary(name.as_ref(), boundary, &prefixes);
                            lvl.extra_xml.push((at, capture_element(reader, e)?));
                            boundary = next;
                            continue;
                        }
                        reader.read_to_end_into(name, &mut Vec::new())?;
                        boundary = 6;
                    } else if is_word_element(name.as_ref(), b"lvlText", &prefixes) {
                        lvl.lvl_text = word_attribute_value(e, b"val", &prefixes)?;
                        reader.read_to_end_into(name, &mut Vec::new())?;
                        boundary = 7;
                    } else if is_word_element(name.as_ref(), b"lvlJc", &prefixes) {
                        if let Some(value) = word_attribute_value(e, b"val", &prefixes)? {
                            lvl.lvl_jc = ST_Jc::from_str(&value).ok();
                        }
                        reader.read_to_end_into(name, &mut Vec::new())?;
                        boundary = 10;
                    } else if is_word_element(name.as_ref(), b"pPr", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        let (ppr, has_producer) = ppr_from_raw(&raw, &prefixes)?;
                        lvl.ppr = Some(ppr.clone());
                        if has_producer {
                            lvl.ppr_raw = Some((ppr, raw, prefixes));
                        }
                        boundary = 11;
                    } else if is_word_element(name.as_ref(), b"rPr", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        let (rpr, has_producer) = rpr_from_raw(&raw, &prefixes)?;
                        lvl.rpr = Some(rpr.clone());
                        if has_producer {
                            lvl.rpr_raw = Some((rpr, raw, prefixes));
                        }
                        boundary = 12;
                    } else {
                        let (at, next) = level_raw_boundary(name.as_ref(), boundary, &prefixes);
                        lvl.extra_xml.push((at, capture_element(reader, e)?));
                        boundary = next;
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"start", &prefixes) {
                        if let Some(val) = word_attribute_value(e, b"val", &prefixes)? {
                            lvl.start = Some(val.parse()?);
                        }
                        boundary = 1;
                    } else if is_word_element(name.as_ref(), b"numFmt", &prefixes) {
                        if let Some(val) = word_attribute_value(e, b"val", &prefixes)? {
                            lvl.num_fmt = Some(ST_NumberFormat::from_str(&val));
                        }
                        boundary = 2;
                    } else if is_word_element(name.as_ref(), b"suff", &prefixes) {
                        if let Some(value) = word_attribute_value(e, b"val", &prefixes)?
                            && let Some(suffix) = ST_LvlSuffix::from_str(&value)
                        {
                            lvl.suffix = Some(suffix);
                            boundary = 6;
                        } else {
                            let (at, next) = level_raw_boundary(name.as_ref(), boundary, &prefixes);
                            lvl.extra_xml.push((at, capture_empty_element(e)?));
                            boundary = next;
                        }
                    } else if is_word_element(name.as_ref(), b"lvlText", &prefixes) {
                        lvl.lvl_text = word_attribute_value(e, b"val", &prefixes)?;
                        boundary = 7;
                    } else if is_word_element(name.as_ref(), b"lvlJc", &prefixes)
                        && let Some(val) = word_attribute_value(e, b"val", &prefixes)?
                    {
                        lvl.lvl_jc = ST_Jc::from_str(&val).ok();
                        boundary = 10;
                    } else if is_word_element(name.as_ref(), b"pPr", &prefixes) {
                        let raw = capture_empty_element(e)?;
                        let (ppr, has_producer) = ppr_from_raw(&raw, &prefixes)?;
                        lvl.ppr = Some(ppr.clone());
                        if has_producer {
                            lvl.ppr_raw = Some((ppr, raw, prefixes));
                        }
                        boundary = 11;
                    } else if is_word_element(name.as_ref(), b"rPr", &prefixes) {
                        let raw = capture_empty_element(e)?;
                        let (rpr, has_producer) = rpr_from_raw(&raw, &prefixes)?;
                        lvl.rpr = Some(rpr.clone());
                        if has_producer {
                            lvl.rpr_raw = Some((rpr, raw, prefixes));
                        }
                        boundary = 12;
                    } else {
                        let (at, next) = level_raw_boundary(name.as_ref(), boundary, &prefixes);
                        lvl.extra_xml.push((at, capture_empty_element(e)?));
                        boundary = next;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"lvl") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(lvl)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let preservation_prefixes = PreservationPrefixes::new(&[]);
        self.to_xml_with_prefix(writer, "w", &preservation_prefixes)
    }

    fn to_xml_with_prefix<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        word_prefix: &str,
        preservation_prefixes: &PreservationPrefixes,
    ) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let level_name = qualified(word_prefix, "lvl");
        let ilvl_name = qualified(word_prefix, "ilvl");
        let mut start = BytesStart::new(level_name.as_str());
        start.push_attribute((ilvl_name.as_str(), buf.format(self.ilvl)));
        push_extra_attributes(&mut start, &self.extra_attributes);
        writer.write_event(Event::Start(start))?;

        write_extras_at(writer, &self.extra_xml, 0)?;
        if let Some(s) = self.start {
            let name = qualified(word_prefix, "start");
            let val_name = qualified(word_prefix, "val");
            let mut e = BytesStart::new(name.as_str());
            e.push_attribute((val_name.as_str(), buf.format(s)));
            writer.write_event(Event::Empty(e))?;
        }

        write_extras_at(writer, &self.extra_xml, 1)?;
        if let Some(fmt) = &self.num_fmt {
            let name = qualified(word_prefix, "numFmt");
            let val_name = qualified(word_prefix, "val");
            let mut e = BytesStart::new(name.as_str());
            e.push_attribute((val_name.as_str(), fmt.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        for boundary in 2..=5 {
            write_extras_at(writer, &self.extra_xml, boundary)?;
        }
        if let Some(suffix) = self.suffix {
            let name = qualified(word_prefix, "suff");
            let val_name = qualified(word_prefix, "val");
            let mut e = BytesStart::new(name.as_str());
            e.push_attribute((val_name.as_str(), suffix.to_str()));
            writer.write_event(Event::Empty(e))?;
        }
        write_extras_at(writer, &self.extra_xml, 6)?;
        if let Some(ref text) = self.lvl_text {
            let name = qualified(word_prefix, "lvlText");
            let val_name = qualified(word_prefix, "val");
            let mut e = BytesStart::new(name.as_str());
            e.push_attribute((val_name.as_str(), text.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        for boundary in 7..=9 {
            write_extras_at(writer, &self.extra_xml, boundary)?;
        }
        if let Some(jc) = self.lvl_jc {
            let name = qualified(word_prefix, "lvlJc");
            let val_name = qualified(word_prefix, "val");
            let mut e = BytesStart::new(name.as_str());
            e.push_attribute((val_name.as_str(), jc.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        write_extras_at(writer, &self.extra_xml, 10)?;
        if let Some(ref ppr) = self.ppr {
            if let Some((original, raw, _)) = &self.ppr_raw
                && paragraph_preservation_eq(ppr, original)
            {
                writer.get_mut().write_all(raw)?;
            } else {
                write_level_property(
                    writer,
                    LevelProperty::Paragraph(ppr),
                    self.ppr_raw.as_ref().map(|(original, raw, prefixes)| {
                        (
                            LevelProperty::Paragraph(original),
                            raw.as_slice(),
                            prefixes.as_slice(),
                        )
                    }),
                    word_prefix,
                    preservation_prefixes,
                )?;
            }
        } else if let Some((_, raw, prefixes)) = &self.ppr_raw
            && let Some(projected) = producer_only_property(
                raw,
                prefixes,
                PropertyKind::Paragraph,
                word_prefix,
                0,
                preservation_prefixes,
            )?
        {
            write_xml_events(writer, &projected)?;
        }

        write_extras_at(writer, &self.extra_xml, 11)?;
        if let Some(ref rpr) = self.rpr {
            if let Some((original, raw, _)) = &self.rpr_raw
                && rpr == original
            {
                writer.get_mut().write_all(raw)?;
            } else {
                write_level_property(
                    writer,
                    LevelProperty::Run(rpr),
                    self.rpr_raw.as_ref().map(|(original, raw, prefixes)| {
                        (
                            LevelProperty::Run(original),
                            raw.as_slice(),
                            prefixes.as_slice(),
                        )
                    }),
                    word_prefix,
                    preservation_prefixes,
                )?;
            }
        } else if let Some((_, raw, prefixes)) = &self.rpr_raw
            && let Some(projected) = producer_only_property(
                raw,
                prefixes,
                PropertyKind::Run,
                word_prefix,
                0,
                preservation_prefixes,
            )?
        {
            write_xml_events(writer, &projected)?;
        }

        write_extras_at(writer, &self.extra_xml, 12)?;

        writer.write_event(Event::End(BytesEnd::new(level_name)))?;
        Ok(())
    }
}

/// `CT_AbstractNum` — An abstract numbering definition with up to 9 levels.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_AbstractNum {
    pub abstract_num_id: u32,
    pub levels: Vec<CT_Lvl>,
    /// Optional multi-level type hint
    pub multi_level_type: Option<String>,
    /// Unmodelled children retained at their modelled-child boundaries.
    pub extra_xml: Vec<(usize, Vec<u8>)>,
    /// Unmodelled attributes and namespace declarations from `w:abstractNum`.
    pub extra_attributes: Vec<(String, String)>,
}

#[allow(non_snake_case)]
impl CT_AbstractNum {
    pub fn new(id: u32) -> Self {
        CT_AbstractNum {
            abstract_num_id: id,
            levels: Vec::new(),
            multi_level_type: None,
            extra_xml: Vec::new(),
            extra_attributes: Vec::new(),
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>, abstract_num_id: u32) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, abstract_num_id, &["w".to_string()])
    }

    fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        abstract_num_id: u32,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut abs = CT_AbstractNum::new(abstract_num_id);
        let mut buf = Vec::new();
        let mut boundary = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"multiLevelType", &prefixes) {
                        abs.multi_level_type = word_attribute_value(e, b"val", &prefixes)?;
                        reader.read_to_end_into(name, &mut Vec::new())?;
                        boundary = 2;
                    } else if is_word_element(name.as_ref(), b"lvl", &prefixes) {
                        let ilvl = u32_attribute(e, b"ilvl", &prefixes)?;
                        let mut level = CT_Lvl::from_xml_with_prefixes(reader, ilvl, &prefixes)?;
                        level.extra_attributes =
                            capture_extra_attributes(e, &[b"ilvl"], &prefixes)?;
                        abs.levels.push(level);
                        boundary = 7 + abs.levels.len();
                    } else {
                        let (at, next) = abstract_raw_boundary(name.as_ref(), boundary, &prefixes);
                        abs.extra_xml.push((at, capture_element(reader, e)?));
                        boundary = next;
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"multiLevelType", &prefixes) {
                        abs.multi_level_type = word_attribute_value(e, b"val", &prefixes)?;
                        boundary = 2;
                    } else if is_word_element(name.as_ref(), b"lvl", &prefixes) {
                        let mut level = CT_Lvl::new(u32_attribute(e, b"ilvl", &prefixes)?);
                        level.extra_attributes =
                            capture_extra_attributes(e, &[b"ilvl"], &prefixes)?;
                        abs.levels.push(level);
                        boundary = 7 + abs.levels.len();
                    } else {
                        let (at, next) = abstract_raw_boundary(name.as_ref(), boundary, &prefixes);
                        abs.extra_xml.push((at, capture_empty_element(e)?));
                        boundary = next;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"abstractNum") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(abs)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let preservation_prefixes = PreservationPrefixes::new(&[]);
        self.to_xml_with_prefix(writer, "w", &preservation_prefixes)
    }

    fn to_xml_with_prefix<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        word_prefix: &str,
        preservation_prefixes: &PreservationPrefixes,
    ) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let abstract_name = qualified(word_prefix, "abstractNum");
        let id_name = qualified(word_prefix, "abstractNumId");
        let mut start = BytesStart::new(abstract_name.as_str());
        start.push_attribute((id_name.as_str(), buf.format(self.abstract_num_id)));
        push_extra_attributes(&mut start, &self.extra_attributes);
        writer.write_event(Event::Start(start))?;

        write_extras_at(writer, &self.extra_xml, 0)?;
        write_extras_at(writer, &self.extra_xml, 1)?;
        if let Some(ref mlt) = self.multi_level_type {
            let name = qualified(word_prefix, "multiLevelType");
            let val_name = qualified(word_prefix, "val");
            let mut e = BytesStart::new(name.as_str());
            e.push_attribute((val_name.as_str(), mlt.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        for boundary in 2..=6 {
            write_extras_at(writer, &self.extra_xml, boundary)?;
        }
        for (index, lvl) in self.levels.iter().enumerate() {
            write_extras_at(writer, &self.extra_xml, 7 + index)?;
            lvl.to_xml_with_prefix(writer, word_prefix, preservation_prefixes)?;
        }
        write_extras_at(writer, &self.extra_xml, 7 + self.levels.len())?;

        writer.write_event(Event::End(BytesEnd::new(abstract_name)))?;
        Ok(())
    }
}

/// `CT_Num` — A numbering instance that references an abstract numbering definition.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Num {
    pub num_id: u32,
    pub abstract_num_id: u32,
    /// Unmodelled children such as level overrides.
    pub extra_xml: Vec<(usize, Vec<u8>)>,
    /// Unmodelled attributes and namespace declarations from `w:num`.
    pub extra_attributes: Vec<(String, String)>,
}

#[allow(non_snake_case)]
impl CT_Num {
    pub fn from_xml(reader: &mut Reader<&[u8]>, num_id: u32) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, num_id, &["w".to_string()])
    }

    fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        num_id: u32,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut abstract_num_id = 0;
        let mut extra_xml = Vec::new();
        let mut buf = Vec::new();
        let mut position = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(e.name().as_ref(), b"abstractNumId", &prefixes)
                        && let Some(val) = word_attribute_value(e, b"val", &prefixes)?
                    {
                        abstract_num_id = val.parse()?;
                        position += 1;
                    } else {
                        extra_xml.push((position, capture_empty_element(e)?));
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(e.name().as_ref(), b"abstractNumId", &prefixes) {
                        if let Some(value) = word_attribute_value(e, b"val", &prefixes)? {
                            abstract_num_id = value.parse()?;
                        }
                        reader.read_to_end_into(e.name(), &mut Vec::new())?;
                        position = 1;
                    } else {
                        let raw = capture_element(reader, e)?;
                        extra_xml.push((position, raw));
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"num") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Num {
            num_id,
            abstract_num_id,
            extra_xml,
            extra_attributes: Vec::new(),
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.to_xml_with_prefix(writer, "w")
    }

    fn to_xml_with_prefix<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        word_prefix: &str,
    ) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let num_name = qualified(word_prefix, "num");
        let num_id_name = qualified(word_prefix, "numId");
        let mut start = BytesStart::new(num_name.as_str());
        start.push_attribute((num_id_name.as_str(), buf.format(self.num_id)));
        push_extra_attributes(&mut start, &self.extra_attributes);
        writer.write_event(Event::Start(start))?;

        write_extras_at(writer, &self.extra_xml, 0)?;
        let abs_name = qualified(word_prefix, "abstractNumId");
        let val_name = qualified(word_prefix, "val");
        let mut abs_ref = BytesStart::new(abs_name.as_str());
        abs_ref.push_attribute((val_name.as_str(), buf.format(self.abstract_num_id)));
        writer.write_event(Event::Empty(abs_ref))?;
        write_extras_at(writer, &self.extra_xml, 1)?;

        writer.write_event(Event::End(BytesEnd::new(num_name)))?;
        Ok(())
    }
}

/// `CT_Numbering` — Root element of the numbering definitions part.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Numbering {
    pub abstract_nums: Vec<CT_AbstractNum>,
    pub nums: Vec<CT_Num>,
    /// Namespace declarations and compatibility attributes from the root.
    pub root_attributes: Vec<(String, String)>,
    /// Unmodelled root children retained at their modelled-child boundaries.
    pub extra_xml: Vec<(usize, Vec<u8>)>,
}

#[allow(non_snake_case)]
impl CT_Numbering {
    pub fn new() -> Self {
        CT_Numbering {
            abstract_nums: Vec::new(),
            nums: Vec::new(),
            root_attributes: Vec::new(),
            extra_xml: Vec::new(),
        }
    }

    fn preserved_namespace_declarations(&self) -> Result<Vec<(String, String)>> {
        let mut declarations = Vec::new();
        append_namespace_declarations(&self.root_attributes, &mut declarations);
        for (_, raw) in &self.extra_xml {
            append_raw_namespace_declarations(raw, &mut declarations)?;
        }
        for abstract_num in &self.abstract_nums {
            append_namespace_declarations(&abstract_num.extra_attributes, &mut declarations);
            for (_, raw) in &abstract_num.extra_xml {
                append_raw_namespace_declarations(raw, &mut declarations)?;
            }
            for level in &abstract_num.levels {
                append_namespace_declarations(&level.extra_attributes, &mut declarations);
                for (_, raw) in &level.extra_xml {
                    append_raw_namespace_declarations(raw, &mut declarations)?;
                }
                if let Some((_, raw, _)) = &level.ppr_raw {
                    append_raw_namespace_declarations(raw, &mut declarations)?;
                }
                if let Some((_, raw, _)) = &level.rpr_raw {
                    append_raw_namespace_declarations(raw, &mut declarations)?;
                }
            }
        }
        for num in &self.nums {
            append_namespace_declarations(&num.extra_attributes, &mut declarations);
            for (_, raw) in &num.extra_xml {
                append_raw_namespace_declarations(raw, &mut declarations)?;
            }
        }
        Ok(declarations)
    }

    /// Parse from XML bytes (the content of numbering.xml).
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut abstract_nums = Vec::new();
        let mut nums = Vec::new();
        let mut root_attributes = Vec::new();
        let mut extra_xml = Vec::new();
        let mut buf = Vec::new();
        let mut position = 0;
        let mut word_prefixes = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, &word_prefixes)?;
                    if is_word_element(name.as_ref(), b"abstractNum", &prefixes) {
                        let id = u32_attribute(e, b"abstractNumId", &prefixes)?;
                        let mut abstract_num =
                            CT_AbstractNum::from_xml_with_prefixes(&mut reader, id, &prefixes)?;
                        abstract_num.extra_attributes =
                            capture_extra_attributes(e, &[b"abstractNumId"], &prefixes)?;
                        abstract_nums.push(abstract_num);
                        position += 1;
                    } else if is_word_element(name.as_ref(), b"num", &prefixes) {
                        let id = u32_attribute(e, b"numId", &prefixes)?;
                        let mut num = CT_Num::from_xml_with_prefixes(&mut reader, id, &prefixes)?;
                        num.extra_attributes = capture_extra_attributes(e, &[b"numId"], &prefixes)?;
                        nums.push(num);
                        position += 1;
                    } else if is_word_element(name.as_ref(), b"numbering", &prefixes) {
                        root_attributes = capture_root_attributes(e)?;
                        word_prefixes = prefixes;
                    } else {
                        extra_xml.push((
                            root_raw_boundary(name.as_ref(), position, &prefixes),
                            capture_element(&mut reader, e)?,
                        ));
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, &word_prefixes)?;
                    if is_word_element(name.as_ref(), b"numbering", &prefixes) {
                        root_attributes = capture_root_attributes(e)?;
                        word_prefixes = prefixes;
                    } else if is_word_element(name.as_ref(), b"abstractNum", &prefixes) {
                        let mut abstract_num =
                            CT_AbstractNum::new(u32_attribute(e, b"abstractNumId", &prefixes)?);
                        abstract_num.extra_attributes =
                            capture_extra_attributes(e, &[b"abstractNumId"], &prefixes)?;
                        abstract_nums.push(abstract_num);
                        position += 1;
                    } else if is_word_element(name.as_ref(), b"num", &prefixes) {
                        nums.push(CT_Num {
                            num_id: u32_attribute(e, b"numId", &prefixes)?,
                            abstract_num_id: 0,
                            extra_xml: Vec::new(),
                            extra_attributes: capture_extra_attributes(e, &[b"numId"], &prefixes)?,
                        });
                        position += 1;
                    } else {
                        extra_xml.push((
                            root_raw_boundary(name.as_ref(), position, &prefixes),
                            capture_empty_element(e)?,
                        ));
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Numbering {
            abstract_nums,
            nums,
            root_attributes,
            extra_xml,
        })
    }

    /// Serialize to XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
        let declarations = self.preserved_namespace_declarations()?;
        let word_prefix = generated_prefix(&declarations, "w", W_NS);
        let relationship_prefix = generated_prefix(&declarations, "r", R_NS);
        let preservation_prefixes = PreservationPrefixes::new(&declarations);

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let root_name = qualified(&word_prefix, "numbering");
        let mut start = BytesStart::new(root_name.as_str());
        let word_declaration = format!("xmlns:{word_prefix}");
        if !has_namespace_declaration(&self.root_attributes, &word_prefix, W_NS) {
            start.push_attribute((word_declaration.as_str(), W_NS));
        }
        let relationship_declaration = format!("xmlns:{relationship_prefix}");
        if !has_namespace_declaration(&self.root_attributes, &relationship_prefix, R_NS) {
            start.push_attribute((relationship_declaration.as_str(), R_NS));
        }
        for (name, value) in &self.root_attributes {
            start.push_attribute((name.as_str(), value.as_str()));
        }
        writer.write_event(Event::Start(start))?;

        write_extras_at(&mut writer, &self.extra_xml, 0)?;
        let mut position = 0;
        for abs in &self.abstract_nums {
            write_extras_at(&mut writer, &self.extra_xml, 1 + position)?;
            abs.to_xml_with_prefix(&mut writer, &word_prefix, &preservation_prefixes)?;
            position += 1;
        }

        for num in &self.nums {
            write_extras_at(&mut writer, &self.extra_xml, 1 + position)?;
            num.to_xml_with_prefix(&mut writer, &word_prefix)?;
            position += 1;
        }

        write_extras_at(&mut writer, &self.extra_xml, 1 + position)?;

        writer.write_event(Event::End(BytesEnd::new(root_name)))?;

        Ok(writer.into_inner())
    }

    /// Get the next available abstract numbering ID.
    pub fn next_abstract_num_id(&self) -> u32 {
        let maximum = self
            .abstract_nums
            .iter()
            .map(|a| a.abstract_num_id)
            .max()
            .unwrap_or(0);
        if self.abstract_nums.is_empty() {
            return 0;
        }
        if let Some(next) = maximum.checked_add(1) {
            return next;
        }
        let used: std::collections::HashSet<_> = self
            .abstract_nums
            .iter()
            .map(|item| item.abstract_num_id)
            .collect();
        (0..=u32::MAX)
            .find(|candidate| !used.contains(candidate))
            .expect("an in-memory numbering collection cannot occupy every u32 identifier")
    }

    /// Get the next available numbering instance ID.
    pub fn next_num_id(&self) -> u32 {
        let maximum = self.nums.iter().map(|n| n.num_id).max().unwrap_or(0);
        if self.nums.is_empty() {
            return 1;
        }
        if let Some(next) = maximum.checked_add(1) {
            return next;
        }
        let used: std::collections::HashSet<_> = self.nums.iter().map(|item| item.num_id).collect();
        (1..=u32::MAX)
            .find(|candidate| !used.contains(candidate))
            .expect("an in-memory numbering collection cannot occupy every nonzero u32 identifier")
    }

    /// Create a bullet list definition and return its numId.
    pub fn add_bullet_list(&mut self) -> u32 {
        self.add_list(&[(ST_NumberFormat::Bullet, Some(1))])
    }

    /// Create a numbered (decimal) list definition and return its numId.
    pub fn add_numbered_list(&mut self) -> u32 {
        self.add_list(&[(ST_NumberFormat::Decimal, Some(1))])
    }

    /// Create a list definition with explicit per-level formats and return
    /// its numId.
    ///
    /// `levels[i]` specifies level `i` as `(format, start)`; a `start` of
    /// `None` defaults to 1 (`start` has no meaning for bullet levels). All
    /// nine levels are always defined so a paragraph referencing a deeper
    /// level than was specified still renders: unspecified levels continue
    /// from the last specified format's family — the bullet-glyph rotation
    /// for bullets, the decimal/letter/roman rotation otherwise — matching
    /// the [`Self::add_bullet_list`] / [`Self::add_numbered_list`] templates.
    ///
    /// An empty `levels` behaves like [`Self::add_numbered_list`]. Only the
    /// first nine entries are representable and any later entries are ignored.
    pub fn add_list(&mut self, levels: &[(ST_NumberFormat, Option<u32>)]) -> u32 {
        let abs_id = self.next_abstract_num_id();
        let num_id = self.next_num_id();

        let mut abs = CT_AbstractNum::new(abs_id);
        abs.multi_level_type = Some("hybridMultilevel".to_string());

        let mut last_specified = ST_NumberFormat::Decimal;
        for i in 0..9u32 {
            let (num_fmt, start) = match levels.get(i as usize) {
                Some((fmt, start)) => {
                    last_specified = fmt.clone();
                    (fmt.clone(), start.unwrap_or(1))
                }
                None => (level_fill_format(&last_specified, i), 1),
            };

            abs.levels.push(build_level(i, num_fmt, start));
        }

        let abstract_position = 1 + self.abstract_nums.len();
        shift_extras_from(&mut self.extra_xml, abstract_position);
        self.abstract_nums.push(abs);
        let num_position = 1 + self.abstract_nums.len() + self.nums.len();
        shift_extras_from(&mut self.extra_xml, num_position);
        self.nums.push(CT_Num {
            num_id,
            abstract_num_id: abs_id,
            extra_xml: Vec::new(),
            extra_attributes: Vec::new(),
        });

        num_id
    }

    /// Redefine one level of an existing list definition, for callers that
    /// only learn a deeper level's format when content first reaches it.
    ///
    /// Returns `false` when `num_id` is unknown or `ilvl` is out of range
    /// (levels are 0–8).
    pub fn set_list_level(
        &mut self,
        num_id: u32,
        ilvl: u32,
        num_fmt: ST_NumberFormat,
        start: Option<u32>,
    ) -> bool {
        if ilvl > 8 {
            return false;
        }

        let Some(num) = self.nums.iter().find(|n| n.num_id == num_id) else {
            return false;
        };
        let abstract_num_id = num.abstract_num_id;
        let Some(abs) = self
            .abstract_nums
            .iter_mut()
            .find(|a| a.abstract_num_id == abstract_num_id)
        else {
            return false;
        };

        let level = build_level(ilvl, num_fmt, start.unwrap_or(1));
        match abs.levels.iter_mut().find(|l| l.ilvl == ilvl) {
            Some(existing) => {
                existing.start = level.start;
                existing.num_fmt = level.num_fmt;
                existing.lvl_text = level.lvl_text;
            }
            None => {
                let index = abs
                    .levels
                    .iter()
                    .position(|existing| existing.ilvl > ilvl)
                    .unwrap_or(abs.levels.len());
                let boundary = 7 + index;
                shift_extras_from(&mut abs.extra_xml, boundary);
                abs.levels.insert(index, level);
            }
        }

        true
    }

    /// Look up the abstract numbering definition for a given numId.
    pub fn get_abstract_num_for(&self, num_id: u32) -> Option<&CT_AbstractNum> {
        let num = self.nums.iter().find(|n| n.num_id == num_id)?;
        self.abstract_nums
            .iter()
            .find(|a| a.abstract_num_id == num.abstract_num_id)
    }
}

impl Default for CT_Numbering {
    fn default() -> Self {
        Self::new()
    }
}

/// Bullet glyph rotation shared by the list templates: • ◦ ▪ repeating.
const BULLET_CHARS: [&str; 9] = [
    "\u{2022}", // bullet •
    "\u{25E6}", // white bullet ◦
    "\u{25AA}", // black small square ▪
    "\u{2022}", // repeat pattern
    "\u{25E6}", "\u{25AA}", "\u{2022}", "\u{25E6}", "\u{25AA}",
];

/// Numeric format rotation shared by the list templates:
/// decimal, lowerLetter, lowerRoman repeating.
const NUMBERED_FORMATS: [ST_NumberFormat; 9] = [
    ST_NumberFormat::Decimal,
    ST_NumberFormat::LowerLetter,
    ST_NumberFormat::LowerRoman,
    ST_NumberFormat::Decimal,
    ST_NumberFormat::LowerLetter,
    ST_NumberFormat::LowerRoman,
    ST_NumberFormat::Decimal,
    ST_NumberFormat::LowerLetter,
    ST_NumberFormat::LowerRoman,
];

/// Template format for an unspecified level, keyed on the last format the
/// caller did specify: bullets stay bullets; anything numeric continues the
/// numbered rotation.
fn level_fill_format(last_specified: &ST_NumberFormat, ilvl: u32) -> ST_NumberFormat {
    match last_specified {
        ST_NumberFormat::Bullet => ST_NumberFormat::Bullet,
        _ => NUMBERED_FORMATS[ilvl as usize % NUMBERED_FORMATS.len()].clone(),
    }
}

/// One level in the shared template shape: bullet glyph or `%N.` text,
/// left-justified, indented 720tw per depth with a 360tw hanging indent.
fn build_level(ilvl: u32, num_fmt: ST_NumberFormat, start: u32) -> CT_Lvl {
    let mut lvl = CT_Lvl::new(ilvl);
    lvl.start = Some(start);
    lvl.lvl_text = Some(match num_fmt {
        ST_NumberFormat::Bullet => BULLET_CHARS[ilvl as usize % BULLET_CHARS.len()].to_string(),
        _ => format!("%{}.", ilvl + 1),
    });
    lvl.num_fmt = Some(num_fmt);
    lvl.lvl_jc = Some(ST_Jc::Left);

    // Standard indentation: 720tw per level
    let indent = (ilvl + 1) as i32 * 720;
    lvl.ppr = Some(CT_PPr {
        ind_left: Some(crate::units::Twips(indent)),
        ind_hanging: Some(crate::units::Twips(360)),
        ..Default::default()
    });

    lvl
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::borders::CT_TabStop;
    use crate::shared::ST_TabJc;
    use crate::units::Twips;

    #[test]
    fn round_trip_numbering() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_bullet_list();
        assert_eq!(num_id, 1);

        let xml = numbering.to_xml().unwrap();
        let parsed = CT_Numbering::from_xml(&xml).unwrap();

        assert_eq!(parsed.abstract_nums.len(), 1);
        assert_eq!(parsed.nums.len(), 1);
        assert_eq!(parsed.nums[0].num_id, 1);
        assert_eq!(parsed.nums[0].abstract_num_id, 0);

        let abs = &parsed.abstract_nums[0];
        assert_eq!(abs.levels.len(), 9);
        assert_eq!(abs.levels[0].num_fmt, Some(ST_NumberFormat::Bullet));
        assert_eq!(abs.levels[0].lvl_text, Some("\u{2022}".to_string()));
    }

    #[test]
    fn round_trip_numbered_list() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        assert_eq!(num_id, 1);

        let xml = numbering.to_xml().unwrap();
        let parsed = CT_Numbering::from_xml(&xml).unwrap();

        let abs = &parsed.abstract_nums[0];
        assert_eq!(abs.levels[0].num_fmt, Some(ST_NumberFormat::Decimal));
        assert_eq!(abs.levels[0].lvl_text, Some("%1.".to_string()));
        assert_eq!(abs.levels[1].num_fmt, Some(ST_NumberFormat::LowerLetter));
    }

    #[test]
    fn producer_defined_number_format_round_trips_without_substitution() {
        let xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="chicago"/></w:lvl></w:abstractNum></w:numbering>"#;

        let numbering = CT_Numbering::from_xml(xml).unwrap();
        assert_eq!(
            numbering.abstract_nums[0].levels[0].num_fmt,
            Some(ST_NumberFormat::Other("chicago".to_owned()))
        );

        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        assert!(
            output.contains(r#"<w:numFmt w:val="chicago"/>"#),
            "{output}"
        );
    }

    #[test]
    fn level_suffix_round_trips_in_its_schema_slot() {
        let xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="space"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>"#;
        let numbering = CT_Numbering::from_xml(xml).expect("numbering parses");
        let level = &numbering.abstract_nums[0].levels[0];

        assert_eq!(level.suffix, Some(ST_LvlSuffix::Space));

        let output =
            String::from_utf8(numbering.to_xml().expect("numbering writes")).expect("XML is UTF-8");
        let suffix = output.find("<w:suff").expect("suffix writes");
        let text = output.find("<w:lvlText").expect("level text writes");
        assert!(suffix < text, "suffix must precede level text: {output}");
    }

    #[test]
    fn level_raw_is_lgl_stays_before_suffix() {
        let xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:isLgl/><w:suff w:val="space"/></w:lvl></w:abstractNum></w:numbering>"#;
        let numbering = CT_Numbering::from_xml(xml).expect("numbering parses");
        let output =
            String::from_utf8(numbering.to_xml().expect("numbering writes")).expect("XML is UTF-8");

        let is_lgl = output.find("<w:isLgl/>").expect("raw isLgl writes");
        let suffix = output.find("<w:suff").expect("suffix writes");
        assert!(is_lgl < suffix, "isLgl must precede suffix: {output}");
    }

    #[test]
    fn multiple_lists() {
        let mut numbering = CT_Numbering::new();
        let bullet_id = numbering.add_bullet_list();
        let num_id = numbering.add_numbered_list();

        assert_eq!(bullet_id, 1);
        assert_eq!(num_id, 2);

        let xml = numbering.to_xml().unwrap();
        let parsed = CT_Numbering::from_xml(&xml).unwrap();

        assert_eq!(parsed.abstract_nums.len(), 2);
        assert_eq!(parsed.nums.len(), 2);
    }

    #[test]
    fn level_indentation() {
        let mut numbering = CT_Numbering::new();
        numbering.add_bullet_list();

        let abs = &numbering.abstract_nums[0];
        // Level 0: 720tw indent, 360tw hanging
        assert_eq!(
            abs.levels[0].ppr.as_ref().unwrap().ind_left,
            Some(Twips(720))
        );
        assert_eq!(
            abs.levels[0].ppr.as_ref().unwrap().ind_hanging,
            Some(Twips(360))
        );
        // Level 2: 2160tw indent
        assert_eq!(
            abs.levels[2].ppr.as_ref().unwrap().ind_left,
            Some(Twips(2160))
        );
    }

    #[test]
    fn parse_numbering_xml() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"#;

        let numbering = CT_Numbering::from_xml(xml).unwrap();
        assert_eq!(numbering.abstract_nums.len(), 1);
        assert_eq!(numbering.nums.len(), 1);

        let abs = &numbering.abstract_nums[0];
        assert_eq!(abs.abstract_num_id, 0);
        assert_eq!(abs.multi_level_type, Some("hybridMultilevel".to_string()));
        assert_eq!(abs.levels.len(), 2);
        assert_eq!(abs.levels[0].start, Some(1));
        assert_eq!(abs.levels[0].num_fmt, Some(ST_NumberFormat::Decimal));
        assert_eq!(abs.levels[0].lvl_text, Some("%1.".to_string()));
        assert_eq!(
            abs.levels[0].ppr.as_ref().unwrap().ind_left,
            Some(Twips(720))
        );
        assert_eq!(abs.levels[1].num_fmt, Some(ST_NumberFormat::LowerLetter));

        let num = &numbering.nums[0];
        assert_eq!(num.num_id, 1);
        assert_eq!(num.abstract_num_id, 0);
    }

    #[test]
    fn get_abstract_num_for_lookup() {
        let mut numbering = CT_Numbering::new();
        numbering.add_bullet_list();
        numbering.add_numbered_list();

        let abs = numbering.get_abstract_num_for(2).unwrap();
        assert_eq!(abs.levels[0].num_fmt, Some(ST_NumberFormat::Decimal));

        assert!(numbering.get_abstract_num_for(99).is_none());
    }

    #[test]
    fn add_list_mixed_levels_round_trip() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_list(&[
            (ST_NumberFormat::Bullet, None),
            (ST_NumberFormat::Decimal, Some(3)),
        ]);
        assert_eq!(num_id, 1);

        let xml = numbering.to_xml().unwrap();
        let parsed = CT_Numbering::from_xml(&xml).unwrap();

        let abs = parsed.get_abstract_num_for(num_id).unwrap();
        assert_eq!(abs.levels.len(), 9);
        assert_eq!(abs.levels[0].num_fmt, Some(ST_NumberFormat::Bullet));
        assert_eq!(abs.levels[0].lvl_text, Some("\u{2022}".to_string()));
        assert_eq!(abs.levels[1].num_fmt, Some(ST_NumberFormat::Decimal));
        assert_eq!(abs.levels[1].lvl_text, Some("%2.".to_string()));
        assert_eq!(abs.levels[1].start, Some(3));
        // Unspecified levels continue the last specified family (numeric).
        assert_eq!(abs.levels[2].num_fmt, Some(ST_NumberFormat::LowerRoman));
        assert_eq!(abs.levels[2].start, Some(1));
    }

    #[test]
    fn add_list_fill_keeps_bullets_for_bullet_lists() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_list(&[(ST_NumberFormat::Bullet, None)]);

        let abs = numbering.get_abstract_num_for(num_id).unwrap();
        for level in &abs.levels {
            assert_eq!(level.num_fmt, Some(ST_NumberFormat::Bullet));
        }
    }

    #[test]
    fn add_list_delegation_matches_legacy_templates() {
        let mut via_helpers = CT_Numbering::new();
        via_helpers.add_bullet_list();
        via_helpers.add_numbered_list();

        let mut via_add_list = CT_Numbering::new();
        via_add_list.add_list(&[(ST_NumberFormat::Bullet, Some(1))]);
        via_add_list.add_list(&[(ST_NumberFormat::Decimal, Some(1))]);

        assert_eq!(via_helpers.abstract_nums, via_add_list.abstract_nums);
        assert_eq!(via_helpers.nums, via_add_list.nums);
    }

    #[test]
    fn set_list_level_redefines_one_level() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_list(&[(ST_NumberFormat::Bullet, None)]);

        assert!(numbering.set_list_level(num_id, 1, ST_NumberFormat::Decimal, Some(3)));

        let abs = numbering.get_abstract_num_for(num_id).unwrap();
        assert_eq!(abs.levels[1].num_fmt, Some(ST_NumberFormat::Decimal));
        assert_eq!(abs.levels[1].lvl_text, Some("%2.".to_string()));
        assert_eq!(abs.levels[1].start, Some(3));
        // Neighbors untouched.
        assert_eq!(abs.levels[0].num_fmt, Some(ST_NumberFormat::Bullet));
        assert_eq!(abs.levels[2].num_fmt, Some(ST_NumberFormat::Bullet));
    }

    #[test]
    fn set_list_level_rejects_unknown_targets() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_list(&[(ST_NumberFormat::Bullet, None)]);

        assert!(!numbering.set_list_level(99, 0, ST_NumberFormat::Decimal, None));
        assert!(!numbering.set_list_level(num_id, 9, ST_NumberFormat::Decimal, None));
    }

    #[test]
    fn list_mutations_preserve_unmodelled_numbering_xml() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
             xmlns:ext="urn:producer:extension"
             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
             mc:Ignorable="w15">
  <w:abstractNum ext:abstractNumId="producer-abstract" q:abstractNumId="0">
    <w:nsid w:val="12345678"/>
    <w:tmpl w:val="87654321"/>
    <w:lvl xmlns:q="urn:producer:shadow" q:ilvl="shadow-level"
           ext:ilvl="producer-level" w:ilvl="0"
           xmlns:ilvl="urn:producer:level"
           xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlRestart w:val="0"/>
      <w:lvlText w:val="%1."/>
      <ilvl:extension/>
      <w15:extension w15:value="bound-prefix"/>
      <w16:extension w16:value="level-bound-prefix"/>
    </w:lvl>
    <ext:tmpl ext:value="after-level"/>
  </w:abstractNum>
  <w:num ext:numId="producer-instance" q:numId="1">
    <w:abstractNumId w:val="0"/>
    <w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride>
  </w:num>
  <ext:numPicBullet ext:value="after-instance"/>
  <w:extLst><w:ext w:uri="preserve-me"/></w:extLst>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        assert!(numbering.set_list_level(1, 0, ST_NumberFormat::UpperRoman, Some(3)));
        numbering.add_list(&[(ST_NumberFormat::Bullet, None)]);

        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        for raw in [
            r#"<w:nsid w:val="12345678"/>"#,
            r#"<w:tmpl w:val="87654321"/>"#,
            r#"<w:lvlRestart w:val="0"/>"#,
            r#"<ilvl:extension/>"#,
            r#"<w15:extension w15:value="bound-prefix"/>"#,
            r#"<w16:extension w16:value="level-bound-prefix"/>"#,
            r#"<w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride>"#,
            r#"<ext:tmpl ext:value="after-level"/>"#,
            r#"<ext:numPicBullet ext:value="after-instance"/>"#,
            r#"<w:extLst><w:ext w:uri="preserve-me"/></w:extLst>"#,
        ] {
            assert!(output.contains(raw), "missing preserved XML: {raw}");
        }
        assert!(
            output.contains(r#"xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml""#),
            "{output}"
        );
        assert!(
            output.contains(
                r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#
            )
        );
        assert!(output.contains(r#"mc:Ignorable="w15""#));
        assert!(
            output.contains(r#"xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml""#)
        );
        for attribute in [
            r#"ext:abstractNumId="producer-abstract""#,
            r#"ext:ilvl="producer-level""#,
            r#"xmlns:ilvl="urn:producer:level""#,
            r#"ext:numId="producer-instance""#,
        ] {
            assert!(output.contains(attribute), "missing attribute: {attribute}");
        }
        for rewritten_attribute in ["q:abstractNumId=", "q:numId="] {
            assert!(
                !output.contains(rewritten_attribute),
                "aliased modelled attribute was emitted twice: {rewritten_attribute}"
            );
        }
        assert!(output.contains(r#"q:ilvl="shadow-level""#));
        let first_level = output.find(r#"<w:lvl w:ilvl="0""#).unwrap();
        let foreign_template = output.find("<ext:tmpl ").unwrap();
        let instance = output.find("<w:num w:numId=").unwrap();
        let foreign_picture = output.find("<ext:numPicBullet ").unwrap();
        let final_extension = output.find("<w:extLst>").unwrap();
        assert!(first_level < foreign_template && foreign_template < instance);
        assert!(instance < foreign_picture && foreign_picture < final_extension);
    }

    #[test]
    fn sparse_level_materialization_keeps_raw_children_in_schema_order() {
        let xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:lvlRestart w:val="0"/>
      <w:lvlText w:val="%1."/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"#;
        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        assert!(numbering.set_list_level(1, 0, ST_NumberFormat::UpperRoman, Some(2)));

        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        let start = output.find("<w:start ").unwrap();
        let format = output.find("<w:numFmt ").unwrap();
        let restart = output.find("<w:lvlRestart ").unwrap();
        let text = output.find("<w:lvlText ").unwrap();
        assert!(start < format && format < restart && restart < text);
    }

    #[test]
    fn inserting_a_missing_level_keeps_abstract_predecessors_in_schema_order() {
        let xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:nsid w:val="12345678"/>
    <w:tmpl w:val="87654321"/>
    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>
    <w:lvl w:ilvl="2"><w:numFmt w:val="lowerRoman"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"#;
        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        assert!(numbering.set_list_level(1, 1, ST_NumberFormat::UpperLetter, Some(1)));

        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        let nsid = output.find("<w:nsid ").unwrap();
        let template = output.find("<w:tmpl ").unwrap();
        let level_zero = output.find(r#"<w:lvl w:ilvl="0""#).unwrap();
        let level_one = output.find(r#"<w:lvl w:ilvl="1""#).unwrap();
        let level_two = output.find(r#"<w:lvl w:ilvl="2""#).unwrap();
        assert!(nsid < template && template < level_zero);
        assert!(level_zero < level_one && level_one < level_two);
    }

    #[test]
    fn self_closing_numbering_elements_are_not_captured_as_unknown_children() {
        let empty = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        let parsed = CT_Numbering::from_xml(empty).unwrap();
        let output = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert_eq!(output.matches("<w:numbering").count(), 1);
        assert!(parsed.extra_xml.is_empty());

        let with_abstract = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="7"/></w:numbering>"#;
        let parsed = CT_Numbering::from_xml(with_abstract).unwrap();
        assert_eq!(parsed.abstract_nums[0].abstract_num_id, 7);
        assert_eq!(parsed.next_abstract_num_id(), 8);

        let nested =
            br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                                     xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"/>
    <w:lvl w:ilvl="1">
      <w:pPr ext:marker="p"><ext:pPrData/></w:pPr>
      <w:rPr ext:marker="r"><ext:rPrData/></w:rPr>
    </w:lvl>
  </w:abstractNum>
</w:numbering>"#;
        let parsed = CT_Numbering::from_xml(nested).unwrap();
        assert!(parsed.abstract_nums[0].levels[0].extra_xml.is_empty());
        assert!(parsed.abstract_nums[0].levels[1].extra_xml.is_empty());
        assert!(parsed.abstract_nums[0].levels[1].ppr_raw.is_some());
        assert!(parsed.abstract_nums[0].levels[1].rpr_raw.is_some());
        let output = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert_eq!(output.matches(r#"<w:lvl w:ilvl="0""#).count(), 1);
        assert_eq!(output.matches("<w:pPr").count(), 1);
        assert_eq!(output.matches("<w:rPr").count(), 1);
        assert!(output.contains(r#"<w:pPr ext:marker="p"><ext:pPrData/></w:pPr>"#));
        assert!(output.contains(r#"<w:rPr ext:marker="r"><ext:rPrData/></w:rPr>"#));
    }

    #[test]
    fn expanded_scalar_numbering_elements_are_modelled_once() {
        let xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"></w:multiLevelType>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"></w:start>
      <w:numFmt w:val="decimal"></w:numFmt>
      <w:lvlText w:val="%1."></w:lvlText>
      <w:lvlJc w:val="left"></w:lvlJc>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"></w:abstractNumId></w:num>
</w:numbering>"#;
        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        assert_eq!(numbering.nums[0].abstract_num_id, 0);
        assert!(numbering.set_list_level(1, 0, ST_NumberFormat::UpperRoman, Some(4)));

        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        for name in ["multiLevelType", "start", "numFmt", "lvlText", "lvlJc"] {
            assert_eq!(
                output.matches(&format!("<w:{name}")).count(),
                1,
                "expanded scalar duplicated: {name}"
            );
        }
        assert_eq!(output.matches("<w:abstractNumId ").count(), 1);
        assert!(output.contains(r#"<w:start w:val="4"/>"#));
        assert!(output.contains(r#"<w:numFmt w:val="upperRoman"/>"#));
    }

    #[test]
    fn adding_the_first_list_keeps_schema_final_root_content_last() {
        let xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:numIdMac w:val="7"/>
</w:numbering>"#;
        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.add_list(&[(ST_NumberFormat::Decimal, None)]);

        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        let abstract_num = output.find("<w:abstractNum ").unwrap();
        let num = output.find("<w:num w:numId=").unwrap();
        let final_content = output.find("<w:numIdMac ").unwrap();
        assert!(abstract_num < num && num < final_content);
    }

    #[test]
    fn list_identifier_allocation_uses_a_gap_after_the_maximum_value() {
        let mut numbering = CT_Numbering {
            abstract_nums: vec![CT_AbstractNum::new(u32::MAX)],
            nums: vec![CT_Num {
                num_id: u32::MAX,
                abstract_num_id: u32::MAX,
                extra_xml: Vec::new(),
                extra_attributes: Vec::new(),
            }],
            root_attributes: Vec::new(),
            extra_xml: Vec::new(),
        };

        let num_id = numbering.add_list(&[(ST_NumberFormat::Decimal, None)]);
        assert_ne!(num_id, u32::MAX);
        assert_eq!(numbering.abstract_nums.last().unwrap().abstract_num_id, 0);
        assert_eq!(num_id, 1);
    }

    #[test]
    fn list_definition_accepts_only_the_first_nine_supplied_levels() {
        let mut numbering = CT_Numbering::new();
        let levels = vec![(ST_NumberFormat::Decimal, None); 10];
        let num_id = numbering.add_list(&levels);

        assert_eq!(
            numbering.get_abstract_num_for(num_id).unwrap().levels.len(),
            9
        );
    }

    #[test]
    fn foreign_canonical_prefix_spellings_keep_their_expanded_names() {
        let xml = br#"<q:numbering
  xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w="urn:producer:word"
  xmlns:r="urn:producer:relationships"
  xmlns:w1="urn:occupied:word"
  xmlns:r1="urn:occupied:relationships">
  <q:abstractNum q:abstractNumId="0"><q:lvl q:ilvl="0"/></q:abstractNum>
  <w:extension r:token="producer"/>
</q:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.abstract_nums[0].levels[0].start = Some(2);
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();

        assert!(output.contains(r#"xmlns:w="urn:producer:word""#));
        assert!(output.contains(r#"xmlns:r="urn:producer:relationships""#));
        assert!(output.contains(
            r#"xmlns:w2="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#
        ));
        assert!(output.contains(
            r#"xmlns:r2="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
        ));
        assert!(output.contains(r#"<w:extension r:token="producer"/>"#));
        assert!(output.contains(r#"<w2:start w2:val="2"/>"#));

        let reparsed = CT_Numbering::from_xml(output.as_bytes()).unwrap();
        assert_eq!(reparsed.abstract_nums[0].levels[0].start, Some(2));
    }

    #[test]
    fn descendant_declarations_cannot_shadow_generated_prefixes() {
        let xml = br#"<q:numbering
  xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w="urn:producer:w" xmlns:w1="urn:producer:w1"
  xmlns:r="urn:producer:r">
  <q:abstractNum xmlns:w2="urn:producer:w2" xmlns:r1="urn:producer:r1"
                 q:abstractNumId="0">
    <q:lvl xmlns:w3="urn:producer:w3" q:ilvl="0">
      <q:pPr xmlns:w4="urn:producer:w4" xmlns:ext="urn:producer:extension"
             ext:marker="keep"><q:ind q:left="720"/></q:pPr>
    </q:lvl>
  </q:abstractNum>
  <q:num xmlns:w5="urn:producer:w5" q:numId="1">
    <q:abstractNumId q:val="0"/>
  </q:num>
</q:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .ind_left = Some(Twips(1440));
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();

        assert!(output.contains("<w6:numbering "), "{output}");
        assert!(output.contains(
            r#"xmlns:w6="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#
        ));
        assert!(output.contains(
            r#"xmlns:r2="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
        ));
        assert!(output.contains("<w6:abstractNum w6:abstractNumId="));
        assert!(output.contains(r#"xmlns:w2="urn:producer:w2""#));
        assert!(output.contains("<w6:lvl w6:ilvl="));
        assert!(output.contains(r#"xmlns:w3="urn:producer:w3""#));
        assert!(output.contains(r#"<q:pPr xmlns:w4="urn:producer:w4""#));
        assert!(output.contains("<w6:num w6:numId="));
        assert!(output.contains(r#"xmlns:w5="urn:producer:w5""#));

        let reparsed = CT_Numbering::from_xml(output.as_bytes()).unwrap();
        assert_eq!(
            reparsed.abstract_nums[0].levels[0]
                .ppr
                .as_ref()
                .unwrap()
                .ind_left,
            Some(Twips(1440))
        );
    }

    #[test]
    fn typed_level_property_edits_preserve_producer_xml() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:pPr ext:marker="p"><ext:pPrData/><w:ind w:left="720"/></w:pPr>
      <w:rPr ext:marker="r"><ext:rPrData/><w:b/></w:rPr>
    </w:lvl>
  </w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        let level = &mut numbering.abstract_nums[0].levels[0];
        level.ppr.as_mut().unwrap().ind_left = Some(Twips(1440));
        level.rpr.as_mut().unwrap().bold = Some(false);
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();

        for preserved in [
            r#"ext:marker="p""#,
            r#"<ext:pPrData/>"#,
            r#"ext:marker="r""#,
            r#"<ext:rPrData/>"#,
        ] {
            assert_eq!(output.matches(preserved).count(), 1, "{output}");
        }
        assert!(output.contains(r#"<w:ind w:left="1440"/>"#), "{output}");
        assert!(output.contains(r#"<w:b w:val="false"/>"#), "{output}");

        let reparsed = CT_Numbering::from_xml(output.as_bytes()).unwrap();
        let level = &reparsed.abstract_nums[0].levels[0];
        assert_eq!(level.ppr.as_ref().unwrap().ind_left, Some(Twips(1440)));
        assert_eq!(level.rpr.as_ref().unwrap().bold, Some(false));
    }

    #[test]
    fn property_local_name_collisions_remain_foreign_during_typed_edits() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:pPr ext:marker="p"><ext:ind ext:left="999"/><w:ind w:left="720"/></w:pPr>
      <w:rPr ext:marker="r"><ext:b/><w:b/></w:rPr>
    </w:lvl>
  </w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        let level = &mut numbering.abstract_nums[0].levels[0];
        assert_eq!(level.ppr.as_ref().unwrap().ind_left, Some(Twips(720)));
        assert_eq!(level.rpr.as_ref().unwrap().bold, Some(true));
        level.ppr.as_mut().unwrap().ind_left = Some(Twips(1440));
        level.rpr.as_mut().unwrap().bold = Some(false);
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();

        assert_eq!(output.matches(r#"<ext:ind ext:left="999"/>"#).count(), 1);
        assert_eq!(output.matches("<ext:b/>").count(), 1);
        assert_eq!(output.matches(r#"<w:ind w:left="1440"/>"#).count(), 1);
        assert_eq!(output.matches(r#"<w:b w:val="false"/>"#).count(), 1);
    }

    #[test]
    fn unsupported_word_property_xml_survives_typed_edits() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:producer w:token="nested"/></w:numPr>
      <w:ind w:left="720" w:producer="attribute"/>
    </w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        let ppr = numbering.abstract_nums[0].levels[0].ppr.as_mut().unwrap();
        ppr.ind_left = Some(Twips(1440));
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();

        assert!(
            output.contains(r#"<w:producer w:token="nested"/>"#),
            "{output}"
        );
        assert!(output.contains(r#"w:producer="attribute""#), "{output}");
        assert!(output.contains(r#"<w:ind w:left="1440""#), "{output}");
    }

    #[test]
    fn default_foreign_elements_and_no_namespace_attributes_remain_unmodelled() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr>
      <ind xmlns="urn:producer" left="999"/>
      <w:ind w:left="720" producer="keep"/>
    </w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        let ppr = numbering.abstract_nums[0].levels[0].ppr.as_mut().unwrap();
        assert_eq!(ppr.ind_left, Some(Twips(720)));
        ppr.ind_left = Some(Twips(1440));
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();

        assert!(output.contains(r#"<ind xmlns="urn:producer" left="999"/>"#));
        assert!(output.contains(r#"producer="keep""#));
        assert!(output.contains(r#"<w:ind w:left="1440""#));
    }

    #[test]
    fn property_overlay_uses_wordprocessingml_schema_order() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr ext:marker="p">
      <w:numPr><w:numId w:val="1"/></w:numPr><ext:p/><w:pBdr><w:top w:val="single"/></w:pBdr>
      <w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:suppressAutoHyphens/>
    </w:pPr>
    <w:rPr ext:marker="r">
      <w:strike/><w:dstrike/><ext:r/><w:vanish/>
      <w:highlight w:val="yellow"/><ext:u/><w:u w:val="single"/>
    </w:rPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        let level = &mut numbering.abstract_nums[0].levels[0];
        level.ppr.as_mut().unwrap().num_id = Some(2);
        level.rpr.as_mut().unwrap().strike = Some(false);
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();

        let num_pr = output.find("<w:numPr>").unwrap();
        let p_extra = output.find("<ext:p/>").unwrap();
        let borders = output.find("<w:pBdr>").unwrap();
        let tabs = output.find("<w:tabs>").unwrap();
        let hyphens = output.find("<w:suppressAutoHyphens").unwrap();
        assert!(num_pr < p_extra && p_extra < borders && borders < tabs && tabs < hyphens);
        let strike = output.find("<w:strike ").unwrap();
        let dstrike = output.find("<w:dstrike").unwrap();
        let r_extra = output.find("<ext:r/>").unwrap();
        let vanish = output.find("<w:vanish").unwrap();
        let highlight = output.find("<w:highlight").unwrap();
        let u_extra = output.find("<ext:u/>").unwrap();
        let underline = output.find("<w:u ").unwrap();
        assert!(strike < dstrike && dstrike < r_extra && r_extra < vanish);
        assert!(highlight < u_extra && u_extra < underline);
    }

    #[test]
    fn property_projection_rejects_excessive_depth_normally() {
        let mut nested = String::new();
        for _ in 0..70 {
            nested.push_str("<w:producer>");
        }
        for _ in 0..70 {
            nested.push_str("</w:producer>");
        }
        let xml = format!(
            r#"<w:numbering xmlns:w="{W_NS}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:pPr><w:numPr>{nested}</w:numPr></w:pPr></w:lvl></w:abstractNum></w:numbering>"#
        );

        let error = CT_Numbering::from_xml(xml.as_bytes()).unwrap_err();
        assert!(
            matches!(error, crate::OxmlError::InvalidValue(ref message) if message.contains("property XML depth")),
            "{error}"
        );
    }

    #[test]
    fn changed_composite_properties_keep_nested_producer_xml() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:numPr ext:marker="keep"><w:numId w:val="1"/><w:producer w:token="word"><ext:deep/></w:producer><ext:data/></w:numPr></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .num_id = Some(2);
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();

        assert!(output.contains(r#"<w:numId w:val="2"/>"#), "{output}");
        assert!(output.contains(r#"ext:marker="keep""#), "{output}");
        let producer = output.find(r#"<w:producer w:token="word">"#).unwrap();
        let deep = output.find("<ext:deep/>").unwrap();
        let producer_end = output.find("</w:producer>").unwrap();
        assert!(producer < deep && deep < producer_end, "{output}");
        assert!(output.contains("<ext:data/>"), "{output}");
    }

    #[test]
    fn numbering_preservation_does_not_duplicate_typed_changes() {
        let xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:pPrChange w:id="6" w:author="Ada"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr>
    <w:rPr><w:rPrChange w:id="7" w:author="Ben"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.abstract_nums[0].levels[0]
            .rpr
            .as_mut()
            .unwrap()
            .bold = Some(true);
        numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .jc = Some(ST_Jc::Center);
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();

        let bold = output.find("<w:b").unwrap();
        let change = output.find("<w:rPrChange").unwrap();
        assert!(bold < change, "{output}");
        assert!(output.contains(r#"<w:rPr><w:i/></w:rPr>"#), "{output}");
        assert_eq!(output.matches("<w:rPrChange").count(), 1, "{output}");
        assert_eq!(output.matches("<w:pPrChange").count(), 1, "{output}");
        assert!(
            output.find("<w:jc w:val=\"center\"").unwrap() < output.find("<w:pPrChange").unwrap()
        );
    }

    #[test]
    fn numbering_model_identity_uses_element_and_attribute_namespaces() {
        let xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <abstractNum xmlns="urn:producer" abstractNumId="99"/>
  <w:abstractNum abstractNumId="8" w:abstractNumId="0">
    <lvl xmlns="urn:producer" ilvl="7"/>
    <w:lvl ilvl="6" w:ilvl="0">
      <start xmlns="urn:producer" val="99"/>
      <w:start val="88" w:val="1"/>
      <w:numFmt val="upperRoman" w:val="decimal"/>
    </w:lvl>
  </w:abstractNum>
  <num xmlns="urn:producer" numId="77"/>
  <w:num numId="5" w:numId="1"><w:abstractNumId val="9" w:val="0"/></w:num>
</w:numbering>"#;

        let numbering = CT_Numbering::from_xml(xml).unwrap();
        assert_eq!(numbering.abstract_nums.len(), 1);
        assert_eq!(numbering.abstract_nums[0].abstract_num_id, 0);
        assert_eq!(numbering.abstract_nums[0].levels.len(), 1);
        assert_eq!(numbering.abstract_nums[0].levels[0].ilvl, 0);
        assert_eq!(numbering.abstract_nums[0].levels[0].start, Some(1));
        assert_eq!(
            numbering.abstract_nums[0].levels[0].num_fmt,
            Some(ST_NumberFormat::Decimal)
        );
        assert_eq!(numbering.nums.len(), 1);
        assert_eq!(numbering.nums[0].num_id, 1);
        assert_eq!(numbering.nums[0].abstract_num_id, 0);

        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        assert!(
            output.contains(r#"<abstractNum xmlns="urn:producer" abstractNumId="99"/>"#),
            "{output}"
        );
        assert!(output.contains(r#"abstractNumId="8""#), "{output}");
        assert!(output.contains(r#"ilvl="6""#), "{output}");
        assert!(
            output.contains(r#"<start xmlns="urn:producer" val="99"/>"#),
            "{output}"
        );
        assert!(
            output.contains(r#"<num xmlns="urn:producer" numId="77"/>"#),
            "{output}"
        );
        assert!(output.contains(r#"numId="5""#), "{output}");
    }

    #[test]
    fn producer_only_tabs_survive_unrelated_edit_and_explicit_clear() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:tabs><ext:data/></w:tabs><w:ind w:left="720"/></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut unrelated = CT_Numbering::from_xml(xml).unwrap();
        unrelated.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .ind_left = Some(Twips(1440));
        let output = String::from_utf8(unrelated.to_xml().unwrap()).unwrap();
        let tabs = output.find("<w:tabs>").unwrap();
        let producer = output.find("<ext:data/>").unwrap();
        let tabs_end = output.find("</w:tabs>").unwrap();
        assert!(tabs < producer && producer < tabs_end, "{output}");
        assert!(output.contains(r#"w:left="1440""#), "{output}");

        let mut cleared = CT_Numbering::from_xml(xml).unwrap();
        cleared.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs = None;
        let output = String::from_utf8(cleared.to_xml().unwrap()).unwrap();
        let tabs = output.find("<w:tabs>").unwrap();
        let producer = output.find("<ext:data/>").unwrap();
        let tabs_end = output.find("</w:tabs>").unwrap();
        assert!(tabs < producer && producer < tabs_end, "{output}");
    }

    #[test]
    fn repeated_tabs_keep_occurrence_identity_and_between_nodes() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:tabs>
      <w:tab w:val="left" w:pos="720" ext:id="a"/><ext:between/>
      <w:tab w:val="left" w:pos="1440" ext:id="b"/>
    </w:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut removed = CT_Numbering::from_xml(xml).unwrap();
        removed.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs
            .remove(0);
        let output = String::from_utf8(removed.to_xml().unwrap()).unwrap();
        assert!(!output.contains(r#"ext:id="a""#), "{output}");
        let between = output.find("<ext:between/>").unwrap();
        let surviving = output.find(r#"w:pos="1440" ext:id="b""#).unwrap();
        assert!(between < surviving, "{output}");

        let mut inserted = CT_Numbering::from_xml(xml).unwrap();
        inserted.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs
            .insert(0, CT_TabStop::new(ST_TabJc::Left, Twips(360)));
        let output = String::from_utf8(inserted.to_xml().unwrap()).unwrap();
        let inserted_tab = output.find(r#"w:pos="360""#).unwrap();
        let first = output.find(r#"w:pos="720" ext:id="a""#).unwrap();
        let between = output.find("<ext:between/>").unwrap();
        let second = output.find(r#"w:pos="1440" ext:id="b""#).unwrap();
        assert!(
            inserted_tab < first && first < between && between < second,
            "{output}"
        );
    }

    #[test]
    fn explicit_property_clears_keep_only_producer_projection() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr ext:container="keep"><w:tabs><w:tab w:val="left" w:pos="720" ext:id="a"/><ext:data/></w:tabs><w:ind w:left="1440"/></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut tabs_cleared = CT_Numbering::from_xml(xml).unwrap();
        tabs_cleared.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs = None;
        let output = String::from_utf8(tabs_cleared.to_xml().unwrap()).unwrap();
        assert!(
            output.contains(r#"<rdocxPreserve:preservedProperty ext:id="a"/>"#),
            "{output}"
        );
        assert!(output.contains("<ext:data/>"), "{output}");
        assert!(!output.contains(r#"w:pos="720""#), "{output}");
        let reparsed = CT_Numbering::from_xml(output.as_bytes()).unwrap();
        assert!(
            reparsed.abstract_nums[0].levels[0]
                .ppr
                .as_ref()
                .unwrap()
                .tabs
                .as_ref()
                .unwrap()
                .tabs
                .is_empty()
        );

        let mut container_cleared = CT_Numbering::from_xml(xml).unwrap();
        container_cleared.abstract_nums[0].levels[0].ppr = None;
        let output = String::from_utf8(container_cleared.to_xml().unwrap()).unwrap();
        assert!(
            output.contains(r#"<w:pPr ext:container="keep""#),
            "{output}"
        );
        assert!(
            output.contains(r#"<rdocxPreserve:preservedProperty ext:id="a"/>"#),
            "{output}"
        );
        assert!(output.contains("<ext:data/>"), "{output}");
        assert!(!output.contains(r#"w:left="1440""#), "{output}");
        assert!(!output.contains(r#"w:pos="720""#), "{output}");
    }

    #[test]
    fn cleared_tabs_keep_word_and_no_namespace_producer_attributes() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:rdocxPreserve="urn:producer:occupied"
  xmlns:mc="urn:producer:mc">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:tabs><w:tab w:val="left" w:pos="720" producer="plain" w:producer="word" token="rdocxPreserve:value"/></w:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs = None;
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        assert!(
            output.contains(r#"xmlns:rdocxPreserve="urn:producer:occupied""#),
            "{output}"
        );
        assert!(
            output.contains(r#"<w:tabs xmlns:rdocxPreserve1="urn:rdocx:preserved-property" xmlns:mc1="http://schemas.openxmlformats.org/markup-compatibility/2006" mc1:Ignorable="rdocxPreserve1">"#),
            "{output}"
        );
        assert!(
            output.contains(r#"<rdocxPreserve1:preservedProperty producer="plain" w:producer="word" token="rdocxPreserve:value"/>"#),
            "{output}"
        );
        assert!(output.contains(r#"producer="plain""#), "{output}");
        assert!(output.contains(r#"w:producer="word""#), "{output}");
        assert!(!output.contains("<w:tab "), "{output}");
        let mut reader = Reader::from_str(&output);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pPr" => {
                    break;
                }
                Ok(Event::Eof) => panic!("missing pPr"),
                _ => {}
            }
            buf.clear();
        }
        let direct = CT_PPr::from_xml(&mut reader).unwrap();
        assert!(direct.tabs.unwrap().tabs.is_empty());
        let reparsed = CT_Numbering::from_xml(output.as_bytes()).unwrap();
        assert!(
            reparsed.abstract_nums[0].levels[0]
                .ppr
                .as_ref()
                .unwrap()
                .tabs
                .as_ref()
                .unwrap()
                .tabs
                .is_empty()
        );
    }

    #[test]
    fn provenance_only_tab_replacement_does_not_reuse_raw_ppr() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:tabs><w:tab w:val="left" w:pos="720" ext:id="source"/></w:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        let parsed = &mut numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs[0];
        assert_eq!(parsed.source_occurrence, Some(0));
        *parsed = CT_TabStop::new(ST_TabJc::Left, Twips(720));
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"w:pos="720""#), "{output}");
        assert!(!output.contains(r#"ext:id="source""#), "{output}");
    }

    #[test]
    fn carrier_extends_existing_expanded_ignorable_attribute_once() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:mc="urn:producer:occupied"
  xmlns:compat="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:w15="urn:word:fifteen"
  xmlns:ext="urn:producer">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:tabs compat:Ignorable="w15"><w:tab w:val="left" w:pos="720" ext:id="source"/></w:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs = None;
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        assert_eq!(output.matches(":Ignorable=").count(), 1, "{output}");
        assert!(
            output.contains(r#"compat:Ignorable="w15 rdocxPreserve""#),
            "{output}"
        );
        assert!(
            output.contains(r#"xmlns:rdocxPreserve="urn:rdocx:preserved-property""#),
            "{output}"
        );
    }

    #[test]
    fn carrier_resolves_ignorable_against_property_ancestor_scope() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:compat="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:ext="urn:producer">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0" xmlns:compat="urn:producer:near">
    <w:pPr><w:tabs compat:Ignorable="producer"><w:tab w:val="left" w:pos="720" ext:id="source"/></w:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs = None;
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        assert!(
            output.contains(r#"compat:Ignorable="producer""#),
            "{output}"
        );
        assert!(output.contains(r#":Ignorable="rdocxPreserve""#), "{output}");
        assert_eq!(output.matches(":Ignorable=").count(), 2, "{output}");
        let reparsed = CT_Numbering::from_xml(output.as_bytes()).unwrap();
        assert!(
            reparsed.abstract_nums[0].levels[0]
                .ppr
                .as_ref()
                .unwrap()
                .tabs
                .as_ref()
                .unwrap()
                .tabs
                .is_empty()
        );
    }

    #[test]
    fn numbering_projection_ignores_foreign_nested_run_properties() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><ext:rPr><ext:b/></ext:rPr><w:rPr><ext:b/></w:rPr><w:ind w:left="720"/></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        let ppr = numbering.abstract_nums[0].levels[0].ppr.as_mut().unwrap();
        assert_eq!(ppr.rpr.as_ref().and_then(|rpr| rpr.bold), None);
        ppr.ind_left = Some(Twips(1440));
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        assert!(output.contains("<ext:rPr><ext:b/></ext:rPr>"), "{output}");
        assert!(output.contains("<w:rPr>"), "{output}");
        assert_eq!(output.matches("<ext:b/>").count(), 2, "{output}");
        assert!(output.contains(r#"w:left="1440""#), "{output}");
    }

    #[test]
    fn repeated_producer_payload_keeps_clear_boundaries() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:tabs>
      <w:tab w:val="left" w:pos="720" ext:id="a"/><ext:between/>
      <w:tab w:val="left" w:pos="1440" ext:id="b"/>
    </w:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs = None;
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        let first = output
            .find(r#"<rdocxPreserve:preservedProperty ext:id="a"/>"#)
            .unwrap();
        let between = output.find("<ext:between/>").unwrap();
        let second = output
            .find(r#"<rdocxPreserve:preservedProperty ext:id="b"/>"#)
            .unwrap();
        assert!(first < between && between < second, "{output}");
        let reparsed = CT_Numbering::from_xml(output.as_bytes()).unwrap();
        assert!(
            reparsed.abstract_nums[0].levels[0]
                .ppr
                .as_ref()
                .unwrap()
                .tabs
                .as_ref()
                .unwrap()
                .tabs
                .is_empty()
        );
    }

    #[test]
    fn typed_tab_edit_keeps_occurrence_producer_attributes() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:tabs><w:tab w:val="left" w:pos="720" ext:id="a"/></w:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs[0]
            .pos = Twips(1440);
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"w:pos="1440" ext:id="a""#), "{output}");
    }

    #[test]
    fn tab_identity_collisions_and_inserted_edits_keep_occurrence_ownership() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:tabs>
      <w:tab w:val="left" w:pos="720" ext:id="a"/><ext:between/>
      <w:tab w:val="left" w:pos="1440" ext:id="b"/>
    </w:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut collision = CT_Numbering::from_xml(xml).unwrap();
        collision.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs[0]
            .pos = Twips(1440);
        let output = String::from_utf8(collision.to_xml().unwrap()).unwrap();
        let first = output.find(r#"w:pos="1440" ext:id="a""#).unwrap();
        let between = output.find("<ext:between/>").unwrap();
        let second = output[first + 1..]
            .find(r#"w:pos="1440" ext:id="b""#)
            .map(|offset| first + 1 + offset)
            .unwrap();
        assert!(first < between && between < second, "{output}");

        let mut inserted_and_edited = CT_Numbering::from_xml(xml).unwrap();
        let tabs = &mut inserted_and_edited.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs;
        tabs[0].pos = Twips(1080);
        tabs.insert(0, CT_TabStop::new(ST_TabJc::Left, Twips(360)));
        let output = String::from_utf8(inserted_and_edited.to_xml().unwrap()).unwrap();
        let inserted = output.find(r#"w:pos="360""#).unwrap();
        let first = output.find(r#"w:pos="1080" ext:id="a""#).unwrap();
        let between = output.find("<ext:between/>").unwrap();
        let second = output.find(r#"w:pos="1440" ext:id="b""#).unwrap();
        assert!(
            inserted < first && first < between && between < second,
            "{output}"
        );
    }

    #[test]
    fn later_insertions_and_removals_anchor_earlier_tab_edits() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:tabs>
      <w:tab w:val="left" w:pos="720" ext:id="a"/><ext:between-a-b/>
      <w:tab w:val="left" w:pos="1440" ext:id="b"/><ext:between-b-c/>
      <w:tab w:val="left" w:pos="2160" ext:id="c"/>
    </w:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut inserted = CT_Numbering::from_xml(xml).unwrap();
        let tabs = &mut inserted.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs;
        tabs[0].pos = Twips(1080);
        tabs.push(CT_TabStop::new(ST_TabJc::Left, Twips(2880)));
        let output = String::from_utf8(inserted.to_xml().unwrap()).unwrap();
        let first = output.find(r#"w:pos="1080" ext:id="a""#).unwrap();
        let first_boundary = output.find("<ext:between-a-b/>").unwrap();
        let second = output.find(r#"w:pos="1440" ext:id="b""#).unwrap();
        let second_boundary = output.find("<ext:between-b-c/>").unwrap();
        let third = output.find(r#"w:pos="2160" ext:id="c""#).unwrap();
        let inserted = output.find(r#"w:pos="2880""#).unwrap();
        assert!(
            first < first_boundary
                && first_boundary < second
                && second < second_boundary
                && second_boundary < third
                && third < inserted,
            "{output}"
        );

        let mut removed = CT_Numbering::from_xml(xml).unwrap();
        let tabs = &mut removed.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs;
        tabs[0].pos = Twips(1080);
        tabs.pop();
        let output = String::from_utf8(removed.to_xml().unwrap()).unwrap();
        let first = output.find(r#"w:pos="1080" ext:id="a""#).unwrap();
        let first_boundary = output.find("<ext:between-a-b/>").unwrap();
        let second = output.find(r#"w:pos="1440" ext:id="b""#).unwrap();
        assert!(
            first < first_boundary && first_boundary < second,
            "{output}"
        );
        assert!(!output.contains(r#"ext:id="c""#), "{output}");
    }

    #[test]
    fn provenance_keeps_same_segment_edits_and_first_duplicate_claim() {
        let inserted_xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:tabs><w:tab w:val="left" w:pos="720" ext:id="a"/><ext:between/><w:tab w:val="left" w:pos="1440" ext:id="b"/></w:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;
        let mut inserted = CT_Numbering::from_xml(inserted_xml).unwrap();
        let tabs = &mut inserted.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs;
        tabs[0].pos = Twips(1080);
        tabs.insert(1, CT_TabStop::new(ST_TabJc::Left, Twips(1200)));
        let output = String::from_utf8(inserted.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"w:pos="1080" ext:id="a""#), "{output}");
        assert!(output.contains(r#"w:pos="1200"/>"#), "{output}");
        assert!(output.contains(r#"w:pos="1440" ext:id="b""#), "{output}");

        let removed_xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><w:tabs><w:tab w:val="left" w:pos="720" ext:id="a"/><ext:first/><w:tab w:val="left" w:pos="1080" ext:id="x"/><ext:second/><w:tab w:val="left" w:pos="1440" ext:id="b"/></w:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;
        let mut removed = CT_Numbering::from_xml(removed_xml).unwrap();
        let tabs = &mut removed.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs;
        tabs[0].pos = Twips(900);
        tabs.remove(1);
        let output = String::from_utf8(removed.to_xml().unwrap()).unwrap();
        let first = output.find(r#"w:pos="900" ext:id="a""#).unwrap();
        let first_extra = output.find("<ext:first/>").unwrap();
        let second_extra = output.find("<ext:second/>").unwrap();
        let second = output.find(r#"w:pos="1440" ext:id="b""#).unwrap();
        assert!(
            first < first_extra && first_extra < second_extra && second_extra < second,
            "{output}"
        );
        assert!(!output.contains(r#"ext:id="x""#), "{output}");

        let mut duplicated = CT_Numbering::from_xml(inserted_xml).unwrap();
        let tabs = &mut duplicated.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs;
        let mut duplicate = tabs[0].clone();
        duplicate.pos = Twips(1200);
        tabs.insert(1, duplicate);
        let output = String::from_utf8(duplicated.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"w:pos="720" ext:id="a""#), "{output}");
        assert!(output.contains(r#"w:pos="1200"/>"#), "{output}");
        assert_eq!(output.matches(r#"ext:id="a""#).count(), 1, "{output}");
    }

    #[test]
    fn aliased_tabs_use_generated_container_qname_and_local_child_alias() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr><q:tabs xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:tab q:val="left" q:pos="720" ext:id="a"/></q:tabs></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs
            .insert(0, CT_TabStop::new(ST_TabJc::Left, Twips(360)));
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        assert!(
            output.contains(
                r#"<w:tabs xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#
            ),
            "{output}"
        );
        assert!(output.contains(r#"w:pos="720" ext:id="a""#), "{output}");
        assert!(output.contains("</w:tabs>"), "{output}");
        assert!(!output.contains("</q:tabs>"), "{output}");
    }

    #[test]
    fn high_count_tab_overlay_uses_bounded_occurrence_queues() {
        const TAB_COUNT: usize = 10_000;
        let sources = (0..TAB_COUNT).map(Some).collect::<Vec<_>>();
        let (_, work) = match_occurrences(TAB_COUNT, &sources);
        assert!(work <= TAB_COUNT + sources.len());

        let mut original_tabs = format!(r#"<w:tabs xmlns:w="{W_NS}" xmlns:ext="urn:producer">"#);
        let mut generated_tabs = format!(r#"<w:tabs xmlns:w="{W_NS}">"#);
        for index in 0..TAB_COUNT {
            original_tabs.push_str(&format!(
                r#"<w:tab w:val="left" w:pos="{}" ext:id="{}"/>"#,
                index + 1,
                index
            ));
            generated_tabs.push_str(&format!(
                r#"<w:tab w:val="left" w:pos="{}"/>"#,
                if index == 0 { 20_000 } else { index + 1 }
            ));
            if index + 1 < TAB_COUNT {
                original_tabs.push_str("<ext:between/>");
            }
        }
        original_tabs.push_str("</w:tabs>");
        generated_tabs.push_str("</w:tabs>");
        let preservation_prefixes = PreservationPrefixes::new(&[]);
        let (_, writer_work) = merge_repeated_property_with_work(
            original_tabs.as_bytes(),
            generated_tabs.as_bytes(),
            &["w".to_string()],
            PropertyKind::Paragraph,
            0,
            Some(&sources),
            &preservation_prefixes,
        )
        .unwrap();
        assert!(writer_work <= 4 * (TAB_COUNT * 3 + 1));

        let mut xml = format!(
            r#"<w:numbering xmlns:w="{W_NS}" xmlns:ext="urn:producer"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:pPr><w:tabs>"#
        );
        xml.push_str(
            original_tabs
                .strip_prefix(&format!(
                    r#"<w:tabs xmlns:w="{W_NS}" xmlns:ext="urn:producer">"#
                ))
                .unwrap()
                .strip_suffix("</w:tabs>")
                .unwrap(),
        );
        xml.push_str("</w:tabs></w:pPr></w:lvl></w:abstractNum></w:numbering>");

        let mut numbering = CT_Numbering::from_xml(xml.as_bytes()).unwrap();
        numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .tabs
            .as_mut()
            .unwrap()
            .tabs[0]
            .pos = Twips(20_000);
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"ext:id="9999""#));
        assert!(output.contains("<ext:between/>"));
    }

    #[test]
    fn property_extensions_keep_schema_slots_when_typed_children_change() {
        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:pPr ext:marker="p"><w:pStyle w:val="List"/><ext:data/><w:keepNext/></w:pPr>
    </w:lvl>
  </w:abstractNum>
</w:numbering>"#;

        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        let ppr = numbering.abstract_nums[0].levels[0].ppr.as_mut().unwrap();
        ppr.style_id = None;
        ppr.ind_left = Some(Twips(720));
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();

        let extension = output.find("<ext:data/>").unwrap();
        let keep_next = output.find("<w:keepNext").unwrap();
        let indentation = output.find("<w:ind ").unwrap();
        assert!(extension < keep_next && keep_next < indentation, "{output}");

        let xml = br#"<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ext="urn:producer:extension">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
    <w:pPr ext:marker="p"><ext:data/><w:keepNext/></w:pPr>
  </w:lvl></w:abstractNum>
</w:numbering>"#;
        let mut numbering = CT_Numbering::from_xml(xml).unwrap();
        numbering.abstract_nums[0].levels[0]
            .ppr
            .as_mut()
            .unwrap()
            .style_id = Some("List".to_string());
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();
        let style = output.find("<w:pStyle ").unwrap();
        let extension = output.find("<ext:data/>").unwrap();
        let keep_next = output.find("<w:keepNext").unwrap();
        assert!(style < extension && extension < keep_next, "{output}");
    }

    #[test]
    fn canonical_properties_use_parent_writer_indentation() {
        let mut numbering = CT_Numbering::new();
        numbering.add_numbered_list();
        let output = String::from_utf8(numbering.to_xml().unwrap()).unwrap();

        assert!(
            output.contains("\n      <w:pPr>\n        <w:ind "),
            "{output}"
        );
        assert!(!output.contains("</w:lvlJc><w:pPr>"), "{output}");
    }

    #[test]
    fn level_public_preservation_fields_and_canonical_equality_stay_stable() {
        let level = CT_Lvl {
            ilvl: 0,
            start: Some(1),
            num_fmt: Some(ST_NumberFormat::Decimal),
            suffix: None,
            lvl_text: Some("%1.".to_string()),
            lvl_jc: Some(ST_Jc::Left),
            ppr: Some(CT_PPr {
                ind_left: Some(Twips(720)),
                ..Default::default()
            }),
            rpr: Some(CT_RPr {
                bold: Some(true),
                ..Default::default()
            }),
            extra_xml: Vec::new(),
            extra_attributes: Vec::new(),
            ppr_raw: None,
            rpr_raw: None,
        };
        let numbering = CT_Numbering {
            abstract_nums: vec![CT_AbstractNum {
                abstract_num_id: 0,
                levels: vec![level.clone()],
                multi_level_type: None,
                extra_xml: Vec::new(),
                extra_attributes: Vec::new(),
            }],
            nums: Vec::new(),
            root_attributes: Vec::new(),
            extra_xml: Vec::new(),
        };

        let reparsed = CT_Numbering::from_xml(&numbering.to_xml().unwrap()).unwrap();
        assert_eq!(reparsed.abstract_nums[0].levels[0], level);

        let mut with_extra = level.clone();
        with_extra.extra_xml.push((0, b"<ext:data/>".to_vec()));
        assert_ne!(with_extra, level);
    }
}
