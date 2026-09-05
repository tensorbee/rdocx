//! Native inventory and bounded replacement of Word building blocks.

use std::collections::HashSet;

use oxml_opc::OpcPackage;
use oxml_opc::relationship::rel_types;
use rdocx_oxml::document::CT_Body;
use rdocx_oxml::glossary::{CT_DocPart, CT_GlossaryDocument};

use crate::{Document, Error, Result};

/// The supported classification of a glossary entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BuildingBlockKind {
    AutoText,
    BuildingBlock,
}

/// The editable supported projection of one existing glossary entry.
#[derive(Debug, Clone, PartialEq)]
pub struct BuildingBlock {
    pub name: String,
    pub kind: BuildingBlockKind,
    pub category: Option<String>,
    pub description: Option<String>,
    pub guid: Option<String>,
    pub gallery: Option<String>,
    pub behaviors: Vec<String>,
    pub body: CT_Body,
}

/// One building block and its stable package-scoped identity.
#[derive(Debug, Clone, PartialEq)]
pub struct BuildingBlockInfo {
    pub glossary_part: String,
    pub ordinal: usize,
    pub block: BuildingBlock,
}

pub(crate) fn load_glossary(
    package: &OpcPackage,
    document_part: &str,
) -> Result<Option<(String, CT_GlossaryDocument)>> {
    let Some(relationships) = package.get_part_rels(document_part) else {
        return Ok(None);
    };
    let glossary = relationships
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == rel_types::GLOSSARY_DOCUMENT)
        .collect::<Vec<_>>();
    if glossary.is_empty() {
        return Ok(None);
    }
    if glossary.len() != 1 {
        return Err(Error::Other(
            "the main document must own at most one glossary relationship".to_owned(),
        ));
    }
    let mut ids = HashSet::new();
    if relationships
        .items
        .iter()
        .any(|relationship| !ids.insert(relationship.id.as_str()))
    {
        return Err(Error::Other(
            "the main document relationship ids must be unique".to_owned(),
        ));
    }
    let relationship = glossary[0];
    if relationship
        .target_mode
        .as_deref()
        .is_some_and(|mode| mode != "Internal")
    {
        return Err(Error::Other(
            "the glossary relationship must be internal".to_owned(),
        ));
    }
    validate_internal_target(document_part, &relationship.target)?;
    let part_name = OpcPackage::resolve_rel_target(document_part, &relationship.target);
    let xml = package
        .get_part(&part_name)
        .ok_or_else(|| Error::Other("the glossary relationship target is missing".to_owned()))?;
    if package
        .content_types
        .overrides
        .get(&part_name)
        .map(String::as_str)
        != Some(oxml_opc::content_types::WORD_GLOSSARY)
    {
        return Err(Error::Other(
            "the glossary relationship target has the wrong content type".to_owned(),
        ));
    }
    Ok(Some((part_name, CT_GlossaryDocument::from_xml(xml)?)))
}

pub(crate) fn validate_internal_target(source_part: &str, target: &str) -> Result<()> {
    if !relationship_target_is_normalized_pack_uri(target) {
        return Err(Error::Other(
            "the glossary target is not a safe part URI".to_owned(),
        ));
    }
    let mut depth = if target.starts_with('/') {
        0
    } else {
        source_part
            .trim_start_matches('/')
            .rsplit_once('/')
            .map(|(directory, _)| directory.split('/').filter(|part| !part.is_empty()).count())
            .unwrap_or(0)
    };
    for component in target.split('/') {
        match component {
            "" | "." => {}
            ".." => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Other("the glossary target escapes the package root".to_owned())
                })?;
            }
            _ => depth += 1,
        }
    }
    Ok(())
}

fn relationship_target_is_normalized_pack_uri(target: &str) -> bool {
    if target.is_empty()
        || target.ends_with('/')
        || target.contains("//")
        || target.contains(['\\', '?', '#'])
        || !target.is_ascii()
    {
        return false;
    }
    if !target.starts_with('/')
        && target
            .split('/')
            .next()
            .is_some_and(|segment| segment.contains(':'))
    {
        return false;
    }
    let bytes = target.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        let byte = bytes[index];
        if byte == b'%' {
            let (Some(&high), Some(&low)) = (bytes.get(index + 1), bytes.get(index + 2)) else {
                return false;
            };
            if !high.is_ascii_hexdigit()
                || !low.is_ascii_hexdigit()
                || high.is_ascii_lowercase()
                || low.is_ascii_lowercase()
            {
                return false;
            }
            let decoded = (pack_uri_hex_value(high) << 4) | pack_uri_hex_value(low);
            if decoded.is_ascii_alphanumeric()
                || matches!(decoded, b'-' | b'.' | b'_' | b'~' | b'/' | b'\\')
                || decoded.is_ascii_control()
            {
                return false;
            }
            index += 3;
            continue;
        }
        if !(byte.is_ascii_alphanumeric()
            || matches!(
                byte,
                b'-' | b'.'
                    | b'_'
                    | b'~'
                    | b'!'
                    | b'$'
                    | b'&'
                    | b'\''
                    | b'('
                    | b')'
                    | b'*'
                    | b'+'
                    | b','
                    | b';'
                    | b'='
                    | b':'
                    | b'@'
                    | b'/'
            ))
        {
            return false;
        }
        index += 1;
    }
    target
        .split('/')
        .filter(|segment| !segment.is_empty())
        .all(|segment| matches!(segment, "..") || (segment != "." && !segment.ends_with('.')))
}

fn pack_uri_hex_value(byte: u8) -> u8 {
    match byte {
        b'0'..=b'9' => byte - b'0',
        b'A'..=b'F' => byte - b'A' + 10,
        _ => 0,
    }
}

fn facade_block(part: &CT_DocPart) -> BuildingBlock {
    BuildingBlock {
        name: part.properties.name.clone().unwrap_or_default(),
        kind: if part.properties.types.iter().any(|value| value == "autoExp") {
            BuildingBlockKind::AutoText
        } else {
            BuildingBlockKind::BuildingBlock
        },
        category: part.properties.category.clone(),
        description: part.properties.description.clone(),
        guid: part.properties.guid.clone(),
        gallery: part.properties.gallery.clone(),
        behaviors: part.properties.behaviors.clone(),
        body: part.body.clone(),
    }
}

fn apply_block(part: &mut CT_DocPart, block: BuildingBlock) -> Result<()> {
    if block.name.is_empty() {
        return Err(Error::Other(
            "a building block name must not be empty".to_owned(),
        ));
    }
    if block.category.is_some() != block.gallery.is_some() {
        return Err(Error::Other(
            "a building block category requires both name and gallery".to_owned(),
        ));
    }
    if block
        .gallery
        .as_deref()
        .is_some_and(|gallery| !valid_building_block_gallery(gallery))
    {
        return Err(Error::Other(
            "a building block gallery must be a schema enumeration".to_owned(),
        ));
    }
    if block
        .behaviors
        .iter()
        .any(|behavior| !matches!(behavior.as_str(), "content" | "p" | "pg"))
    {
        return Err(Error::Other(
            "a building block behavior must be content, p, or pg".to_owned(),
        ));
    }
    if block.guid.as_deref().is_some_and(|guid| !valid_guid(guid)) {
        return Err(Error::Other(
            "a building block GUID must use braced uppercase lexical form".to_owned(),
        ));
    }
    part.properties.name = Some(block.name);
    part.properties.category = block.category;
    part.properties.description = block.description;
    part.properties.guid = block.guid;
    part.properties.gallery = block.gallery;
    part.properties.behaviors = block.behaviors;
    match block.kind {
        BuildingBlockKind::AutoText => {
            if !part.properties.types.iter().any(|value| value == "autoExp") {
                part.properties.types.push("autoExp".to_owned());
            }
        }
        BuildingBlockKind::BuildingBlock => {
            part.properties.types.retain(|value| value != "autoExp");
        }
    }
    part.body = block.body;
    Ok(())
}

fn valid_building_block_gallery(value: &str) -> bool {
    matches!(
        value,
        "placeholder"
            | "any"
            | "default"
            | "docParts"
            | "coverPg"
            | "eq"
            | "ftrs"
            | "hdrs"
            | "pgNum"
            | "tbls"
            | "watermarks"
            | "autoTxt"
            | "txtBox"
            | "pgNumT"
            | "pgNumB"
            | "pgNumMargins"
            | "tblOfContents"
            | "bib"
            | "custQuickParts"
            | "custCoverPg"
            | "custEq"
            | "custFtrs"
            | "custHdrs"
            | "custPgNum"
            | "custTbls"
            | "custWatermarks"
            | "custAutoTxt"
            | "custTxtBox"
            | "custPgNumT"
            | "custPgNumB"
            | "custPgNumMargins"
            | "custTblOfContents"
            | "custBib"
            | "custom1"
            | "custom2"
            | "custom3"
            | "custom4"
            | "custom5"
    )
}

fn valid_guid(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() == 38
        && bytes.first() == Some(&b'{')
        && bytes.last() == Some(&b'}')
        && [9, 14, 19, 24]
            .into_iter()
            .all(|index| bytes.get(index) == Some(&b'-'))
        && bytes.iter().enumerate().all(|(index, byte)| {
            matches!(index, 0 | 9 | 14 | 19 | 24 | 37) || matches!(byte, b'0'..=b'9' | b'A'..=b'F')
        })
}

impl Document {
    /// Inventory relationship-owned glossary entries in source order.
    pub fn building_blocks(&self) -> Result<Vec<BuildingBlockInfo>> {
        let Some(glossary) = self.glossary.as_ref() else {
            return Ok(Vec::new());
        };
        let part_name = self
            .glossary_part_name
            .as_ref()
            .expect("loaded glossary has a part name");
        Ok(glossary
            .doc_parts
            .iter()
            .enumerate()
            .map(|(ordinal, part)| BuildingBlockInfo {
                glossary_part: part_name.clone(),
                ordinal,
                block: facade_block(part),
            })
            .collect())
    }

    /// Replace one existing glossary entry through a staged save and reopen.
    pub fn replace_building_block(
        &mut self,
        glossary_part: &str,
        ordinal: usize,
        block: BuildingBlock,
    ) -> Result<BuildingBlockInfo> {
        if self.glossary_part_name.as_deref() != Some(glossary_part) {
            return Err(Error::Other("stale glossary part identity".to_owned()));
        }
        let mut candidate = self.clone_for_staging();
        let part = candidate
            .glossary
            .as_mut()
            .and_then(|glossary| glossary.doc_parts.get_mut(ordinal))
            .ok_or_else(|| Error::Other("stale building block ordinal".to_owned()))?;
        apply_block(part, block.clone())?;
        candidate.glossary_dirty = true;
        let bytes = candidate.to_bytes()?;
        let reopened = Document::from_bytes(&bytes)?;
        let result = reopened
            .building_blocks()?
            .into_iter()
            .find(|entry| entry.glossary_part == glossary_part && entry.ordinal == ordinal)
            .ok_or_else(|| {
                Error::Other("building block identity did not survive reopen".to_owned())
            })?;
        if result.block != block {
            return Err(Error::Other(
                "building block replacement did not survive reopen".to_owned(),
            ));
        }
        self.commit_staged_mutation(reopened);
        Ok(result)
    }
}
