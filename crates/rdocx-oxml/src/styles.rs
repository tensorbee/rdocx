//! Style elements: `CT_Styles`, `CT_Style`, `CT_DocDefaults`.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::error::{OxmlError, Result};
use crate::namespace::{W_NS, matches_local_name};
use crate::numbering::{parse_scoped_ppr, parse_scoped_rpr, word_prefixes_at};
use crate::properties::{CT_PPr, CT_RPr, is_word_attribute, is_word_element};
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::table::{CT_TblPr, CT_TcPr, CT_TrPr};

/// The type of a style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
}

impl StyleType {
    pub fn from_str(s: &str) -> Result<Self> {
        match s {
            "paragraph" => Ok(StyleType::Paragraph),
            "character" => Ok(StyleType::Character),
            "table" => Ok(StyleType::Table),
            "numbering" => Ok(StyleType::Numbering),
            _ => Err(OxmlError::InvalidValue(format!("invalid style type: {s}"))),
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            StyleType::Paragraph => "paragraph",
            StyleType::Character => "character",
            StyleType::Table => "table",
            StyleType::Numbering => "numbering",
        }
    }
}

/// `CT_Style` — A single style definition.
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_Style {
    pub style_id: String,
    pub style_type: StyleType,
    pub name: Option<String>,
    pub based_on: Option<String>,
    pub next_style: Option<String>,
    pub linked_style: Option<String>,
    pub auto_redefine: Option<bool>,
    pub hidden: Option<bool>,
    pub ui_priority: Option<u32>,
    pub semi_hidden: Option<bool>,
    pub unhide_when_used: Option<bool>,
    pub quick_format: Option<bool>,
    pub locked: Option<bool>,
    pub is_default: bool,
    pub ppr: Option<CT_PPr>,
    pub rpr: Option<CT_RPr>,
    /// Typed projection of the style's base table properties.
    pub table_properties: Option<CT_TblPr>,
    #[doc(hidden)]
    pub table_properties_original: Option<CT_TblPr>,
    /// Preserved self-contained bytes for the base table properties.
    pub table_properties_xml: Option<Vec<u8>>,
    /// Typed projection of the style's base table row properties.
    ///
    /// Boxed, like the other composite members added to this family, so a
    /// style stays cheap on the stack. Test threads build whole documents by
    /// value against a 2 MiB ceiling.
    pub table_row_properties: Option<Box<CT_TrPr>>,
    /// Typed projection of the style's base table cell properties.
    ///
    /// Boxed for the same stack-budget reason as `table_row_properties`.
    pub table_cell_properties: Option<Box<CT_TcPr>>,
    /// Preserved conditional table-style regions and typed projections.
    pub conditional_table_styles: Vec<CT_TblStylePr>,
    /// Unmodeled root attributes and namespace declarations.
    #[doc(hidden)]
    pub extra_attributes: Vec<(String, String)>,
    /// Preserved explicit forms of modeled scalar children.
    #[doc(hidden)]
    pub modeled_xml: Vec<(u8, Vec<u8>)>,
    /// Preserved style children keyed to their schema-order rank.
    pub extra_xml: Vec<(u8, Vec<u8>)>,
}

/// One conditional table-style region, in Word's increasing-priority order.
///
/// **The declaration order is the resolution order.** Resolution applies
/// regions in ascending order and a later region overwrites an earlier one, so
/// moving a variant changes what Word-authored tables render as. The order is
/// whole table, vertical bands, horizontal bands, column edges, row edges, then
/// the four corners.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum TableStyleRegion {
    WholeTable,
    Band1Vert,
    Band2Vert,
    Band1Horz,
    Band2Horz,
    FirstCol,
    LastCol,
    FirstRow,
    LastRow,
    NwCell,
    NeCell,
    SwCell,
    SeCell,
}

impl TableStyleRegion {
    /// Every region, in increasing-priority order.
    pub const ALL: [TableStyleRegion; 13] = [
        TableStyleRegion::WholeTable,
        TableStyleRegion::Band1Vert,
        TableStyleRegion::Band2Vert,
        TableStyleRegion::Band1Horz,
        TableStyleRegion::Band2Horz,
        TableStyleRegion::FirstCol,
        TableStyleRegion::LastCol,
        TableStyleRegion::FirstRow,
        TableStyleRegion::LastRow,
        TableStyleRegion::NwCell,
        TableStyleRegion::NeCell,
        TableStyleRegion::SwCell,
        TableStyleRegion::SeCell,
    ];

    /// The `w:type` value naming this region.
    pub fn to_str(self) -> &'static str {
        match self {
            TableStyleRegion::WholeTable => "wholeTable",
            TableStyleRegion::Band1Vert => "band1Vert",
            TableStyleRegion::Band2Vert => "band2Vert",
            TableStyleRegion::Band1Horz => "band1Horz",
            TableStyleRegion::Band2Horz => "band2Horz",
            TableStyleRegion::FirstCol => "firstCol",
            TableStyleRegion::LastCol => "lastCol",
            TableStyleRegion::FirstRow => "firstRow",
            TableStyleRegion::LastRow => "lastRow",
            TableStyleRegion::NwCell => "nwCell",
            TableStyleRegion::NeCell => "neCell",
            TableStyleRegion::SwCell => "swCell",
            TableStyleRegion::SeCell => "seCell",
        }
    }

    /// The region a `w:type` value names, or `None` when it names none.
    ///
    /// An unrecognised value is never an error. The region round-trips from
    /// its preserved bytes and takes no part in resolution.
    pub fn from_str(value: &str) -> Option<Self> {
        Self::ALL
            .into_iter()
            .find(|region| region.to_str() == value)
    }
}

/// One conditional table-style region.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_TblStylePr {
    /// The region this layer formats, or `None` when `w:type` names a region
    /// this workspace does not recognise. An unrecognised region serialises
    /// back from [`CT_TblStylePr::raw_xml`] unchanged.
    pub region: Option<TableStyleRegion>,
    pub paragraph_properties: Option<CT_PPr>,
    /// Region run properties (`w:rPr`).
    ///
    /// Boxed so the five-layer region stays cheap on the stack, following the
    /// `CT_PPr::borders` precedent.
    pub run_properties: Option<Box<CT_RPr>>,
    pub table_properties: Option<CT_TblPr>,
    /// Region row properties (`w:trPr`). Modeled and round-tripped here. Its
    /// layout application belongs to F-268a.
    ///
    /// Boxed for the same stack-budget reason as `run_properties`.
    pub row_properties: Option<Box<CT_TrPr>>,
    pub cell_properties: Option<CT_TcPr>,
    /// Unmodeled region attributes and namespace declarations.
    #[doc(hidden)]
    pub extra_attributes: Vec<(String, String)>,
    pub raw_xml: Vec<u8>,
}

#[allow(non_snake_case)]
impl CT_Style {
    fn style_property_bytes(
        &self,
        rank: u8,
        local: &[u8],
        generated: Option<&[u8]>,
    ) -> Result<Vec<u8>> {
        let sources = self
            .modeled_xml
            .iter()
            .filter(|(owner_rank, _)| *owner_rank == rank)
            .map(|(_, raw)| raw)
            .collect::<Vec<_>>();
        let mut selected = None;
        for (index, source) in sources.iter().enumerate() {
            if defaults_owned_child_name(source, &[])?.as_deref() == Some(local) {
                selected = Some(index);
            }
        }
        let Some(selected) = selected else {
            let mut bytes = sources.into_iter().flatten().copied().collect::<Vec<_>>();
            if let Some(generated) = generated {
                bytes.extend_from_slice(generated);
            }
            return Ok(bytes);
        };
        let unchanged = if generated.is_none() {
            false
        } else if local == b"rPr" {
            Some(parse_style_rpr(sources[selected], &[])?) == self.rpr
        } else {
            Some(parse_scoped_ppr(sources[selected], &[])?) == self.ppr
        };
        let mut bytes = Vec::new();
        for (index, source) in sources.iter().enumerate() {
            let owned = defaults_owned_child_name(source, &[])?.as_deref() == Some(local);
            if !owned {
                bytes.extend_from_slice(source);
            } else if let Some(generated) = generated {
                if unchanged || index != selected {
                    bytes.extend_from_slice(source);
                } else {
                    bytes.extend_from_slice(&merge_defaults_property_root(
                        source,
                        generated,
                        &[],
                        local,
                    )?);
                }
            }
        }
        Ok(bytes)
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>, attrs: &BytesStart) -> Result<Self> {
        let prefixes = word_prefixes_at(attrs, &["w".to_string()])?;
        let namespace_bindings = namespace_bindings_at(attrs, &[])?;
        Self::from_xml_with_prefixes(reader, attrs, &prefixes, &namespace_bindings)
    }

    fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        attrs: &BytesStart,
        word_prefixes: &[String],
        namespace_bindings: &[(String, String)],
    ) -> Result<Self> {
        reject_conflicting_style_prefixes(attrs)?;
        reject_conflicting_style_bindings(namespace_bindings)?;
        let mut style_id = String::new();
        let mut style_type = StyleType::Paragraph;
        let mut is_default = false;

        for attr in attrs.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            if is_word_attribute(key, b"styleId", word_prefixes) {
                style_id = std::str::from_utf8(&attr.value)?.to_string();
            } else if is_word_attribute(key, b"type", word_prefixes) {
                style_type = StyleType::from_str(std::str::from_utf8(&attr.value)?)?;
            } else if is_word_attribute(key, b"default", word_prefixes) {
                is_default = std::str::from_utf8(&attr.value)? == "1"
                    || std::str::from_utf8(&attr.value)? == "true";
            }
        }

        let mut name = None;
        let mut based_on = None;
        let mut next_style = None;
        let mut linked_style = None;
        let mut auto_redefine = None;
        let mut hidden = None;
        let mut ui_priority = None;
        let mut semi_hidden = None;
        let mut unhide_when_used = None;
        let mut quick_format = None;
        let mut locked = None;
        let mut ppr = None;
        let mut rpr = None;
        let mut table_properties = None;
        let mut table_properties_xml = None;
        let mut table_row_properties = None;
        let mut table_cell_properties = None;
        let mut conditional_table_styles = Vec::new();
        let mut modeled_xml = Vec::new();
        let mut extra_xml = Vec::new();
        let extra_attributes = raw_style_attributes(
            attrs,
            &[b"styleId", b"type", b"default"],
            word_prefixes,
            namespace_bindings,
        )?;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let ename = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    let bindings = namespace_bindings_at(e, namespace_bindings)?;
                    if is_word_element(ename.as_ref(), b"pPr", &prefixes)
                        || is_word_element(ename.as_ref(), b"rPr", &prefixes)
                    {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            &bindings,
                        )?;
                        if is_word_element(ename.as_ref(), b"pPr", &prefixes) {
                            ppr = Some(parse_scoped_ppr(&raw, &prefixes)?);
                            modeled_xml.push((20, raw));
                        } else {
                            rpr = Some(parse_style_rpr(&raw, &prefixes)?);
                            modeled_xml.push((21, raw));
                        }
                    } else if matches!(ename.local_name().as_ref(), b"pPr" | b"rPr") {
                        let rank = if ename.local_name().as_ref() == b"pPr" {
                            20
                        } else {
                            21
                        };
                        modeled_xml.push((
                            rank,
                            crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                &bindings,
                            )?,
                        ));
                    } else if let Some(rank) = style_scalar_rank(ename.as_ref(), &prefixes) {
                        let string_value = matches!(rank, 0 | 2 | 3 | 4)
                            .then(|| get_val_attr(e, &prefixes))
                            .transpose()?
                            .flatten();
                        let toggle_value = matches!(rank, 5 | 6 | 8 | 9 | 10 | 11)
                            .then(|| parse_style_toggle(e, &prefixes))
                            .transpose()?
                            .flatten();
                        let priority_value = (rank == 7)
                            .then(|| parse_style_priority(e, &prefixes))
                            .transpose()?
                            .flatten();
                        let raw = capture_empty_element(e)?;
                        if string_value.is_some()
                            || toggle_value.is_some()
                            || priority_value.is_some()
                        {
                            apply_style_scalar(
                                rank,
                                string_value,
                                toggle_value,
                                priority_value,
                                &mut name,
                                &mut based_on,
                                &mut next_style,
                                &mut linked_style,
                                &mut auto_redefine,
                                &mut hidden,
                                &mut ui_priority,
                                &mut semi_hidden,
                                &mut unhide_when_used,
                                &mut quick_format,
                                &mut locked,
                            );
                            if style_scalar_empty_requires_snapshot(e, rank)? {
                                modeled_xml
                                    .push((rank, make_style_raw_self_contained(&raw, &bindings)?));
                            }
                        } else {
                            extra_xml.push((rank, make_style_raw_self_contained(&raw, &bindings)?));
                        }
                    } else {
                        extra_xml.push((
                            style_child_rank(ename.as_ref(), &prefixes),
                            make_style_raw_self_contained(&capture_empty_element(e)?, &bindings)?,
                        ));
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let ename = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    let bindings = namespace_bindings_at(e, namespace_bindings)?;
                    let scalar_rank = style_scalar_rank(ename.as_ref(), &prefixes);
                    if let Some(rank) = scalar_rank {
                        let string_value = matches!(rank, 0 | 2 | 3 | 4)
                            .then(|| get_val_attr(e, &prefixes))
                            .transpose()?
                            .flatten();
                        let toggle_value = matches!(rank, 5 | 6 | 8 | 9 | 10 | 11)
                            .then(|| parse_style_toggle(e, &prefixes))
                            .transpose()?
                            .flatten();
                        let priority_value = (rank == 7)
                            .then(|| parse_style_priority(e, &prefixes))
                            .transpose()?
                            .flatten();
                        let raw = capture_element(reader, e)?;
                        if explicit_style_element_has_only_trivia(&raw)?
                            && (string_value.is_some()
                                || toggle_value.is_some()
                                || priority_value.is_some())
                        {
                            apply_style_scalar(
                                rank,
                                string_value,
                                toggle_value,
                                priority_value,
                                &mut name,
                                &mut based_on,
                                &mut next_style,
                                &mut linked_style,
                                &mut auto_redefine,
                                &mut hidden,
                                &mut ui_priority,
                                &mut semi_hidden,
                                &mut unhide_when_used,
                                &mut quick_format,
                                &mut locked,
                            );
                            modeled_xml
                                .push((rank, make_style_raw_self_contained(&raw, &bindings)?));
                        } else {
                            extra_xml.push((rank, make_style_raw_self_contained(&raw, &bindings)?));
                        }
                    } else if is_word_element(ename.as_ref(), b"pPr", &prefixes)
                        || is_word_element(ename.as_ref(), b"rPr", &prefixes)
                    {
                        let trim_start = reader.config().trim_text_start;
                        let trim_end = reader.config().trim_text_end;
                        reader.config_mut().trim_text(false);
                        let captured = capture_element(reader, e);
                        reader.config_mut().trim_text_start = trim_start;
                        reader.config_mut().trim_text_end = trim_end;
                        let raw = crate::text::raw_with_external_bindings(&captured?, &bindings)?;
                        if is_word_element(ename.as_ref(), b"pPr", &prefixes) {
                            ppr = Some(parse_scoped_ppr(&raw, &prefixes)?);
                            modeled_xml.push((20, raw));
                        } else {
                            rpr = Some(parse_style_rpr(&raw, &prefixes)?);
                            modeled_xml.push((21, raw));
                        }
                    } else if matches!(ename.local_name().as_ref(), b"pPr" | b"rPr") {
                        let trim_start = reader.config().trim_text_start;
                        let trim_end = reader.config().trim_text_end;
                        reader.config_mut().trim_text(false);
                        let captured = capture_element(reader, e);
                        reader.config_mut().trim_text_start = trim_start;
                        reader.config_mut().trim_text_end = trim_end;
                        let rank = if ename.local_name().as_ref() == b"pPr" {
                            20
                        } else {
                            21
                        };
                        modeled_xml.push((
                            rank,
                            crate::text::raw_with_external_bindings(&captured?, &bindings)?,
                        ));
                    } else if is_word_element(ename.as_ref(), b"tblPr", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        let preserved = make_style_raw_self_contained(&raw, &bindings)?;
                        table_properties =
                            Some(parse_style_table_properties(&preserved, &prefixes)?);
                        table_properties_xml = Some(preserved);
                    } else if is_word_element(ename.as_ref(), b"trPr", &prefixes) {
                        table_row_properties = Some(Box::new(CT_TrPr::from_xml_with_prefixes(
                            reader, &prefixes,
                        )?));
                    } else if is_word_element(ename.as_ref(), b"tcPr", &prefixes) {
                        table_cell_properties = Some(Box::new(CT_TcPr::from_xml_with_prefixes(
                            reader, &prefixes,
                        )?));
                    } else if is_word_element(ename.as_ref(), b"tblStylePr", &prefixes) {
                        reject_conflicting_style_prefixes(e)?;
                        reject_conflicting_style_bindings(&bindings)?;
                        let region = get_word_attr(e, b"type", &prefixes)?
                            .as_deref()
                            .and_then(TableStyleRegion::from_str);
                        let raw =
                            make_style_raw_self_contained(&capture_element(reader, e)?, &bindings)?;
                        let layers = parse_conditional_style_properties(&raw, &prefixes)?;
                        conditional_table_styles.push(CT_TblStylePr {
                            region,
                            paragraph_properties: layers.paragraph_properties,
                            run_properties: layers.run_properties,
                            table_properties: layers.table_properties,
                            row_properties: layers.row_properties,
                            cell_properties: layers.cell_properties,
                            extra_attributes: raw_style_attributes(
                                e,
                                &[b"type"],
                                &prefixes,
                                &bindings,
                            )?,
                            raw_xml: raw,
                        });
                    } else {
                        extra_xml.push((
                            style_child_rank(ename.as_ref(), &prefixes),
                            make_style_raw_self_contained(&capture_element(reader, e)?, &bindings)?,
                        ));
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"style") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Style {
            style_id,
            style_type,
            name,
            based_on,
            next_style,
            linked_style,
            auto_redefine,
            hidden,
            ui_priority,
            semi_hidden,
            unhide_when_used,
            quick_format,
            locked,
            is_default,
            ppr,
            rpr,
            table_properties: table_properties.clone(),
            table_properties_original: table_properties.clone(),
            table_properties_xml,
            table_row_properties,
            table_cell_properties,
            conditional_table_styles,
            extra_attributes,
            modeled_xml,
            extra_xml,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut e = BytesStart::new("w:style");
        e.push_attribute(("w:type", self.style_type.to_str()));
        e.push_attribute(("w:styleId", self.style_id.as_str()));
        if self.is_default {
            e.push_attribute(("w:default", "1"));
        }
        push_style_attributes(&mut e, &self.extra_attributes);
        let _style_scope = crate::text::emitted_binding_scope(&e, false, false)?;
        writer.write_event(Event::Start(e))?;

        let mut extras = self.extra_xml.iter().collect::<Vec<_>>();
        extras.sort_by_key(|(rank, _)| *rank);
        let mut extra_index = 0;

        write_style_extras(writer, &extras, &mut extra_index, 0)?;
        write_style_string(writer, "w:name", 0, self.name.as_deref(), &self.modeled_xml)?;

        write_style_extras(writer, &extras, &mut extra_index, 2)?;
        write_style_string(
            writer,
            "w:basedOn",
            2,
            self.based_on.as_deref(),
            &self.modeled_xml,
        )?;

        write_style_extras(writer, &extras, &mut extra_index, 3)?;
        write_style_string(
            writer,
            "w:next",
            3,
            self.next_style.as_deref(),
            &self.modeled_xml,
        )?;

        write_style_extras(writer, &extras, &mut extra_index, 4)?;
        write_style_string(
            writer,
            "w:link",
            4,
            self.linked_style.as_deref(),
            &self.modeled_xml,
        )?;
        write_style_extras(writer, &extras, &mut extra_index, 5)?;
        write_style_toggle(
            writer,
            "w:autoRedefine",
            5,
            self.auto_redefine,
            &self.modeled_xml,
        )?;
        write_style_extras(writer, &extras, &mut extra_index, 6)?;
        write_style_toggle(writer, "w:hidden", 6, self.hidden, &self.modeled_xml)?;
        write_style_extras(writer, &extras, &mut extra_index, 7)?;
        write_style_priority(writer, self.ui_priority, &self.modeled_xml)?;
        write_style_extras(writer, &extras, &mut extra_index, 8)?;
        write_style_toggle(
            writer,
            "w:semiHidden",
            8,
            self.semi_hidden,
            &self.modeled_xml,
        )?;
        write_style_extras(writer, &extras, &mut extra_index, 9)?;
        write_style_toggle(
            writer,
            "w:unhideWhenUsed",
            9,
            self.unhide_when_used,
            &self.modeled_xml,
        )?;
        write_style_extras(writer, &extras, &mut extra_index, 10)?;
        write_style_toggle(
            writer,
            "w:qFormat",
            10,
            self.quick_format,
            &self.modeled_xml,
        )?;
        write_style_extras(writer, &extras, &mut extra_index, 11)?;
        write_style_toggle(writer, "w:locked", 11, self.locked, &self.modeled_xml)?;
        write_style_extras(writer, &extras, &mut extra_index, 20)?;
        let mut generated_ppr = Writer::new(Vec::new());
        if let Some(ref ppr) = self.ppr {
            ppr.to_xml(&mut generated_ppr)?;
        }
        let ppr = self.ppr.as_ref().map(|_| {
            if generated_ppr.get_ref().is_empty() {
                b"<w:pPr/>".as_slice()
            } else {
                generated_ppr.get_ref().as_slice()
            }
        });
        writer
            .get_mut()
            .write_all(&self.style_property_bytes(20, b"pPr", ppr)?)?;
        write_style_extras(writer, &extras, &mut extra_index, 21)?;
        let mut generated_rpr = Writer::new(Vec::new());
        if let Some(ref rpr) = self.rpr {
            rpr.to_xml(&mut generated_rpr)?;
        }
        let rpr = self.rpr.as_ref().map(|_| {
            if generated_rpr.get_ref().is_empty() {
                b"<w:rPr/>".as_slice()
            } else {
                generated_rpr.get_ref().as_slice()
            }
        });
        writer
            .get_mut()
            .write_all(&self.style_property_bytes(21, b"rPr", rpr)?)?;
        write_style_extras(writer, &extras, &mut extra_index, 22)?;
        if let Some(ref properties) = self.table_properties {
            let preserved_matches = self.table_properties_original.as_ref() == Some(properties);
            if preserved_matches {
                writer
                    .get_mut()
                    .write_all(self.table_properties_xml.as_deref().unwrap_or_default())?;
            } else {
                properties.to_xml(writer)?;
            }
        }
        write_style_extras(writer, &extras, &mut extra_index, 23)?;
        if let Some(ref properties) = self.table_row_properties {
            properties.to_xml(writer)?;
        }
        write_style_extras(writer, &extras, &mut extra_index, 24)?;
        if let Some(ref properties) = self.table_cell_properties {
            properties.to_xml(writer)?;
        }
        write_style_extras(writer, &extras, &mut extra_index, 25)?;
        for conditional in &self.conditional_table_styles {
            writer
                .get_mut()
                .write_all(&serialize_conditional_table_style(conditional)?)?;
        }
        write_style_extras(writer, &extras, &mut extra_index, u8::MAX)?;

        writer.write_event(Event::End(BytesEnd::new("w:style")))?;
        Ok(())
    }
}

/// `CT_DocDefaults` — Document-level default properties.
#[derive(Debug, Clone, Default, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_DocDefaults {
    pub rpr: Option<CT_RPr>,
    pub ppr: Option<CT_PPr>,
    /// Namespace-closed original wrapper, preserving attributes and child slots.
    pub source_xml: Option<Vec<u8>>,
}

#[allow(non_snake_case)]
impl CT_DocDefaults {
    fn from_source_xml(raw: &[u8], inherited: &[String]) -> Result<Self> {
        let (root, children) = defaults_source_children(raw)?;
        let prefixes = word_prefixes_at(&root, inherited)?;
        let mut defaults = Self {
            source_xml: Some(raw.to_vec()),
            ..Self::default()
        };
        for child in children {
            let Some((start, nested)) = defaults_optional_children(&child)? else {
                continue;
            };
            let scope = word_prefixes_at(&start, &prefixes)?;
            let run = is_word_element(start.name().as_ref(), b"rPrDefault", &scope);
            let paragraph = is_word_element(start.name().as_ref(), b"pPrDefault", &scope);
            if !run && !paragraph {
                continue;
            }
            // Last physical owned wrapper determines the direct property facts.
            if run {
                defaults.rpr = None;
            }
            if paragraph {
                defaults.ppr = None;
            }
            for property in nested {
                let Some((start, _)) = defaults_optional_children(&property)? else {
                    continue;
                };
                let property_scope = word_prefixes_at(&start, &scope)?;
                if run && is_word_element(start.name().as_ref(), b"rPr", &property_scope) {
                    defaults.rpr = Some(parse_style_rpr(&property, &scope)?);
                } else if paragraph
                    && is_word_element(start.name().as_ref(), b"pPr", &property_scope)
                {
                    defaults.ppr = Some(parse_scoped_ppr(&property, &scope)?);
                }
            }
        }
        Ok(defaults)
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, &["w".to_string()])
    }

    fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut defaults = CT_DocDefaults::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if matches_local_name(name.as_ref(), b"rPrDefault") {
                        // Read into rPrDefault, expecting rPr child
                        defaults.rpr = Self::parse_pr_default(reader, b"rPrDefault")?;
                    } else if matches_local_name(name.as_ref(), b"pPrDefault") {
                        defaults.ppr = Self::parse_ppr_default(reader, &prefixes)?;
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"docDefaults") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(defaults)
    }

    fn parse_pr_default(reader: &mut Reader<&[u8]>, end_tag: &[u8]) -> Result<Option<CT_RPr>> {
        let mut rpr = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"rPr") {
                        rpr = Some(CT_RPr::from_xml(reader)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), end_tag) => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(rpr)
    }

    fn parse_ppr_default(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Option<CT_PPr>> {
        let mut ppr = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"pPr", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        ppr = Some(parse_scoped_ppr(&raw, word_prefixes)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"pPrDefault") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(ppr)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if let Some(raw) = &self.source_xml {
            let _defaults_scope = crate::text::raw_root_binding_scope(raw, false)?;
            let original = Self::from_source_xml(raw, &[])?;
            if original.rpr == self.rpr && original.ppr == self.ppr {
                writer.get_mut().write_all(raw)?;
                return Ok(());
            }
            let mut run = Writer::new(Vec::new());
            if let Some(properties) = &self.rpr {
                properties.to_xml(&mut run)?;
            }
            let mut paragraph = Writer::new(Vec::new());
            if let Some(properties) = &self.ppr {
                properties.to_xml(&mut paragraph)?;
            }
            // An explicitly empty final owner suppresses any earlier duplicate.
            let run = self.rpr.as_ref().map(|_| {
                if run.get_ref().is_empty() {
                    b"<w:rPr/>".to_vec()
                } else {
                    run.into_inner()
                }
            });
            let paragraph = self.ppr.as_ref().map(|_| {
                if paragraph.get_ref().is_empty() {
                    b"<w:pPr/>".to_vec()
                } else {
                    paragraph.into_inner()
                }
            });
            let merged = merge_defaults_source(
                raw,
                run.as_deref(),
                paragraph.as_deref(),
                original.rpr != self.rpr,
                original.ppr != self.ppr,
            )?;
            writer.get_mut().write_all(&merged)?;
            return Ok(());
        }
        writer.write_event(Event::Start(BytesStart::new("w:docDefaults")))?;

        if let Some(ref rpr) = self.rpr {
            writer.write_event(Event::Start(BytesStart::new("w:rPrDefault")))?;
            rpr.to_xml(writer)?;
            writer.write_event(Event::End(BytesEnd::new("w:rPrDefault")))?;
        }

        if let Some(ref ppr) = self.ppr {
            writer.write_event(Event::Start(BytesStart::new("w:pPrDefault")))?;
            ppr.to_xml(writer)?;
            writer.write_event(Event::End(BytesEnd::new("w:pPrDefault")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:docDefaults")))?;
        Ok(())
    }
}

fn parse_style_rpr(raw: &[u8], inherited: &[String]) -> Result<CT_RPr> {
    let (root, children) = defaults_source_children(raw)?;
    let scope = word_prefixes_at(&root, inherited)?;
    let end = std::str::from_utf8(root.name().as_ref())?.to_owned();
    let mut projection = Writer::new(Vec::new());
    projection.write_event(Event::Start(root))?;
    for child in children {
        let name = defaults_owned_child_name(&child, &scope)?;
        if name.as_deref().is_some_and(|name| {
            matches!(
                name,
                b"rFonts"
                    | b"sz"
                    | b"szCs"
                    | b"u"
                    | b"b"
                    | b"bCs"
                    | b"i"
                    | b"iCs"
                    | b"caps"
                    | b"smallCaps"
                    | b"strike"
                    | b"dstrike"
                    | b"vanish"
                    | b"rtl"
                    | b"lang"
                    | b"color"
                    | b"spacing"
                    | b"position"
                    | b"vertAlign"
            )
        }) {
            let (leaf, _) = defaults_source_children(&child)?;
            projection.write_event(Event::Empty(leaf))?;
        } else {
            projection.get_mut().extend_from_slice(&child);
        }
    }
    projection.write_event(Event::End(BytesEnd::new(end)))?;
    parse_scoped_rpr(projection.get_ref(), inherited)
}

// These concrete two-level wrappers retain all unowned bytes in their source slot.
fn defaults_optional_children(raw: &[u8]) -> Result<Option<(BytesStart<'static>, Vec<Vec<u8>>)>> {
    let mut reader = Reader::from_reader(raw);
    match reader.read_event()? {
        Event::Start(_) | Event::Empty(_) => defaults_source_children(raw).map(Some),
        _ => Ok(None),
    }
}

fn defaults_source_children(raw: &[u8]) -> Result<(BytesStart<'static>, Vec<Vec<u8>>)> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let (start, empty) = match reader.read_event()? {
        Event::Start(start) => (start.into_owned(), false),
        Event::Empty(start) => (start.into_owned(), true),
        _ => {
            return Err(OxmlError::InvalidValue(
                "default wrapper root missing".to_owned(),
            ));
        }
    };
    let mut children = Vec::new();
    if !empty {
        loop {
            let begin = reader.buffer_position() as usize;
            match reader.read_event()? {
                Event::Start(ref child) => {
                    capture_element(&mut reader, child)?;
                    children.push(raw[begin..reader.buffer_position() as usize].to_vec());
                }
                Event::End(_) => break,
                Event::Eof => {
                    return Err(OxmlError::InvalidValue(
                        "default wrapper end missing".to_owned(),
                    ));
                }
                _ => children.push(raw[begin..reader.buffer_position() as usize].to_vec()),
            }
        }
    }
    Ok((start, children))
}

fn defaults_owned_child_name(raw: &[u8], inherited: &[String]) -> Result<Option<Vec<u8>>> {
    let Some((root, _)) = defaults_optional_children(raw)? else {
        return Ok(None);
    };
    let scope = word_prefixes_at(&root, inherited)?;
    let local = root.name().local_name().as_ref().to_vec();
    Ok(is_word_element(root.name().as_ref(), &local, &scope).then_some(local))
}

// Style and defaults setters use this attribute-level merge. All physical
// duplicates lose changed owned attributes, while their opaque bytes survive.
fn defaults_owned_attributes(
    raw: &[u8],
    inherited: &[String],
) -> Result<std::collections::HashMap<Vec<u8>, Vec<u8>>> {
    let mut values = std::collections::HashMap::new();
    if let Some((root, _)) = defaults_optional_children(raw)? {
        let scope = word_prefixes_at(&root, inherited)?;
        for attr in root.attributes() {
            let attr = attr?;
            let local = attr.key.local_name().as_ref().to_vec();
            if is_word_attribute(attr.key.as_ref(), &local, &scope) {
                values.insert(local, attr.value.to_vec());
            }
        }
    }
    Ok(values)
}

pub(crate) fn merge_defaults_scalar_attributes(
    raw: &[u8],
    before: Option<&Vec<u8>>,
    after: Option<&Vec<u8>>,
    inherited: &[String],
    selected: bool,
) -> Result<Vec<u8>> {
    merge_defaults_scalar_attributes_in(raw, before, after, inherited, selected, None)
}

pub(crate) fn merge_defaults_scalar_attributes_in(
    raw: &[u8],
    before: Option<&Vec<u8>>,
    after: Option<&Vec<u8>>,
    inherited: &[String],
    selected: bool,
    emission: Option<&crate::text::FieldXmlEmission>,
) -> Result<Vec<u8>> {
    let canonical = vec!["w".to_owned()];
    let old = defaults_owned_attributes(before.map_or(&[], Vec::as_slice), &canonical)?;
    let new = defaults_owned_attributes(after.map_or(&[], Vec::as_slice), &canonical)?;
    let changed = old
        .keys()
        .chain(new.keys())
        .filter(|key| old.get(*key) != new.get(*key))
        .cloned()
        .collect::<std::collections::HashSet<_>>();
    if changed.is_empty() {
        if let Some(emission) = emission {
            emission.reserve_inner_xml(raw.len())?;
        }
        return Ok(raw.to_vec());
    }
    let (root, _) = defaults_source_children(raw)?;
    let scope = word_prefixes_at(&root, inherited)?;
    let bindings = namespace_bindings_at(&root, &[])?;
    // Reuse an existing qualified Word binding. A new binding is needed only
    // when the physical owner has no usable attribute prefix.
    let existing_prefix = scope
        .iter()
        .find(|prefix| !prefix.is_empty() && !prefix.starts_with('\0'))
        .map(String::as_str);
    let mut prefix = existing_prefix
        .map(std::borrow::Cow::Borrowed)
        .unwrap_or_else(|| std::borrow::Cow::Owned("defaultsWord".to_owned()));
    while existing_prefix.is_none()
        && (bindings
            .iter()
            .any(|(name, _)| name.as_str() == prefix.as_ref())
            || inherited.iter().any(|known| known == prefix.as_ref())
            || raw.windows(prefix.len() + 1).any(|bytes| {
                bytes[..prefix.len()] == *prefix.as_bytes() && bytes[prefix.len()] == b':'
            }))
    {
        prefix.to_mut().push('_');
    }
    // The actual lexical walk produces borrowed source spans. No replacement
    // XML buffer or generated attribute name exists until the exact plan is
    // validated and reserved below.
    let mut retained = Vec::new();
    let mut cursor = 1;
    while cursor < raw.len()
        && !raw[cursor].is_ascii_whitespace()
        && raw[cursor] != b'/'
        && raw[cursor] != b'>'
    {
        cursor += 1;
    }
    retained.push(0..cursor);
    loop {
        let begin = cursor;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if matches!(raw.get(cursor), Some(b'/') | Some(b'>')) {
            retained.push(begin..cursor);
            break;
        }
        let key_start = cursor;
        while cursor < raw.len() && !raw[cursor].is_ascii_whitespace() && raw[cursor] != b'=' {
            cursor += 1;
        }
        let key = &raw[key_start..cursor];
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if raw.get(cursor) != Some(&b'=') {
            return Err(OxmlError::InvalidValue(
                "default attribute assignment missing".to_owned(),
            ));
        }
        cursor += 1;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *raw
            .get(cursor)
            .ok_or_else(|| OxmlError::InvalidValue("default attribute quote missing".to_owned()))?;
        if quote != b'\'' && quote != b'"' {
            return Err(OxmlError::InvalidValue(
                "default attribute quote invalid".to_owned(),
            ));
        }
        cursor += 1;
        while cursor < raw.len() && raw[cursor] != quote {
            cursor += 1;
        }
        if cursor == raw.len() {
            return Err(OxmlError::InvalidValue(
                "default attribute end missing".to_owned(),
            ));
        }
        cursor += 1;
        let local = key.rsplit(|byte| *byte == b':').next().unwrap_or(key);
        if !changed.contains(local) || !is_word_attribute(key, local, &scope) {
            retained.push(begin..cursor);
        }
    }
    let append = selected && changed.iter().any(|key| new.contains_key(key));
    let mut attributes = if append {
        let mut keys = changed
            .iter()
            .filter_map(|key| new.get(key).map(|value| (key, value)))
            .collect::<Vec<_>>();
        keys.sort_by_key(|(key, _)| *key);
        keys.into_iter()
            .map(|(key, value)| Ok((std::str::from_utf8(key)?, value.as_slice())))
            .collect::<Result<Vec<_>>>()?
    } else {
        Vec::new()
    };
    retained.push(cursor..raw.len());
    let mut length = retained
        .iter()
        .try_fold(0usize, |sum, range| sum.checked_add(range.len()))
        .ok_or_else(|| OxmlError::InvalidValue("default attribute output overflow".to_owned()))?;
    if append && existing_prefix.is_none() {
        length = length
            .checked_add(10)
            .and_then(|n| n.checked_add(prefix.len()))
            .and_then(|n| n.checked_add(W_NS.len()))
            .ok_or_else(|| {
                OxmlError::InvalidValue("default namespace output overflow".to_owned())
            })?;
    }
    for (key, value) in &attributes {
        length = length
            .checked_add(5)
            .and_then(|n| n.checked_add(prefix.len()))
            .and_then(|n| n.checked_add(key.len()))
            .and_then(|n| n.checked_add(value.len()))
            .ok_or_else(|| {
                OxmlError::InvalidValue("default attribute output overflow".to_owned())
            })?;
    }
    if let Some(emission) = emission {
        emission.reserve_inner_xml(length)?;
    }
    let mut result = Vec::with_capacity(length);
    let tail = retained.pop().expect("the actual source tail was recorded");
    for range in retained {
        result.extend_from_slice(&raw[range]);
    }
    if append && existing_prefix.is_none() {
        result.extend_from_slice(b" xmlns:");
        result.extend_from_slice(prefix.as_bytes());
        result.extend_from_slice(b"=\"");
        result.extend_from_slice(W_NS.as_bytes());
        result.push(b'"');
    }
    for (key, value) in attributes.drain(..) {
        result.push(b' ');
        result.extend_from_slice(prefix.as_bytes());
        result.push(b':');
        result.extend_from_slice(key.as_bytes());
        result.extend_from_slice(b"=\"");
        result.extend_from_slice(value);
        result.push(b'"');
    }
    result.extend_from_slice(&raw[tail]);
    debug_assert_eq!(result.len(), length);
    Ok(result)
}

// The typed paragraph border owner has no raw container slot. Merge only the
// changed owned edge attributes while retaining its physical source payload.
fn style_border_semantic_key(name: Vec<u8>) -> Vec<u8> {
    match name.as_slice() {
        b"start" => b"left".to_vec(),
        b"end" => b"right".to_vec(),
        _ => name,
    }
}

pub(crate) fn merge_style_border_source(
    raw: &[u8],
    before: Option<&Vec<u8>>,
    after: &[u8],
    inherited: &[String],
    selected_owner: bool,
) -> Result<Vec<u8>> {
    merge_style_border_source_in(raw, before, after, inherited, selected_owner, None)
}

pub(crate) fn merge_style_border_source_in(
    raw: &[u8],
    before: Option<&Vec<u8>>,
    after: &[u8],
    inherited: &[String],
    selected_owner: bool,
    emission: Option<&crate::text::FieldXmlEmission>,
) -> Result<Vec<u8>> {
    if let Some(emission) = emission {
        let (output, _outer) = emission.new_source_record_output();
        let mut writer = Writer::new(output);
        write_style_border_source(
            &mut writer,
            raw,
            before,
            after,
            inherited,
            selected_owner,
            Some(emission),
        )?;
        Ok(writer.into_inner().into_mapped_parts().0)
    } else {
        let mut writer = Writer::new(Vec::new());
        write_style_border_source(
            &mut writer,
            raw,
            before,
            after,
            inherited,
            selected_owner,
            None,
        )?;
        Ok(writer.into_inner())
    }
}

fn write_style_border_source<W: std::io::Write>(
    writer: &mut Writer<W>,
    raw: &[u8],
    before: Option<&Vec<u8>>,
    after: &[u8],
    inherited: &[String],
    selected_owner: bool,
    emission: Option<&crate::text::FieldXmlEmission>,
) -> Result<()> {
    let (root, source) = defaults_source_children(raw)?;
    let scope = word_prefixes_at(&root, inherited)?;
    let canonical = vec!["w".to_owned()];
    let mut previous = std::collections::HashMap::new();
    if let Some(before) = before {
        let before = crate::text::raw_with_external_bindings_in(
            before,
            &[("w".to_owned(), W_NS.to_owned())],
            emission,
        )?;
        let (before_root, children) = defaults_source_children(&before)?;
        let before_scope = word_prefixes_at(&before_root, &canonical)?;
        let bindings = namespace_bindings_at(&before_root, &[])?;
        for child in children {
            if let Some(name) = defaults_owned_child_name(&child, &before_scope)? {
                previous.insert(
                    style_border_semantic_key(name),
                    crate::text::raw_with_external_bindings_in(&child, &bindings, emission)?,
                );
            }
        }
    }
    let bindings = [("w".to_owned(), W_NS.to_owned())];
    let after = crate::text::raw_with_external_bindings_in(after, &bindings, emission)?;
    let (generated_root, generated) = defaults_source_children(&after)?;
    let generated_scope = word_prefixes_at(&generated_root, &[])?;
    let mut supplied = std::collections::HashMap::new();
    let mut order = Vec::new();
    for child in generated {
        if let Some(name) = defaults_owned_child_name(&child, &generated_scope)? {
            let name = style_border_semantic_key(name);
            order.push(name.clone());
            supplied.insert(
                name,
                crate::text::raw_with_external_bindings_in(
                    &child,
                    &namespace_bindings_at(&generated_root, &[])?,
                    emission,
                )?,
            );
        }
    }
    let mut selected = std::collections::HashMap::new();
    for (index, child) in source.iter().enumerate() {
        if let Some(name) = defaults_owned_child_name(child, &scope)? {
            let name = style_border_semantic_key(name);
            selected.insert(name, index);
        }
    }
    let generated_bindings = namespace_bindings_at(&generated_root, &[])?;
    let root_name = root.name();
    let end = std::str::from_utf8(root_name.as_ref())?;
    writer.write_event(Event::Start(root.borrow()))?;
    let mut inserted = std::collections::HashSet::new();
    for child in &source {
        if let Some(name) = defaults_owned_child_name(child, &scope)? {
            let name = style_border_semantic_key(name);
            if let Some(rank) = order.iter().position(|value| value == &name) {
                for new_name in order
                    .iter()
                    .take(rank)
                    .filter(|value| !selected.contains_key(*value))
                {
                    if selected_owner && inserted.insert(new_name.clone()) {
                        writer
                            .get_mut()
                            .write_all(&crate::text::raw_with_external_bindings_in(
                                &supplied[new_name],
                                &generated_bindings,
                                emission,
                            )?)?;
                    }
                }
            }
            // Explicit deletion removes the owned edge, including every physical
            // alias or duplicate. Opaque siblings and the enclosing owner remain.
            if previous.contains_key(&name) && !supplied.contains_key(&name) {
                continue;
            }
            if previous.get(&name) != supplied.get(&name) && supplied.contains_key(&name) {
                writer
                    .get_mut()
                    .write_all(&merge_defaults_scalar_attributes_in(
                        child,
                        previous.get(&name),
                        supplied.get(&name),
                        &scope,
                        true,
                        emission,
                    )?)?;
                inserted.insert(name);
                continue;
            }
        }
        writer.get_mut().write_all(child)?;
    }
    for name in order
        .iter()
        .filter(|name| selected_owner && !selected.contains_key(*name) && !inserted.contains(*name))
    {
        writer
            .get_mut()
            .write_all(&crate::text::raw_with_external_bindings_in(
                &supplied[name],
                &generated_bindings,
                emission,
            )?)?;
    }
    writer.write_event(Event::End(BytesEnd::new(end)))?;
    Ok(())
}

fn merge_defaults_property_root(
    raw: &[u8],
    generated: &[u8],
    inherited: &[String],
    local: &[u8],
) -> Result<Vec<u8>> {
    let (root, source_children) = defaults_source_children(raw)?;
    let scope = word_prefixes_at(&root, inherited)?;
    let mut before = Writer::new(Vec::new());
    if local == b"rPr" {
        parse_style_rpr(raw, inherited)?.to_xml(&mut before)?;
    } else {
        parse_scoped_ppr(raw, inherited)?.to_xml(&mut before)?;
    }
    let bindings = [("w".to_owned(), W_NS.to_owned())];
    let generated = crate::text::raw_with_external_bindings(generated, &bindings)?;
    let (generated_root, children) = defaults_source_children(&generated)?;
    let generated_scope = word_prefixes_at(&generated_root, &[])?;
    let mut previous = std::collections::HashMap::new();
    if !before.get_ref().is_empty() {
        let raw_before = crate::text::raw_with_external_bindings(before.get_ref(), &bindings)?;
        let (before_root, before_children) = defaults_source_children(&raw_before)?;
        let before_scope = word_prefixes_at(&before_root, &[])?;
        for child in before_children {
            if let Some(name) = defaults_owned_child_name(&child, &before_scope)? {
                previous.insert(name, child);
            }
        }
    }
    let mut supplied = std::collections::HashMap::new();
    let mut order = Vec::new();
    for child in &children {
        if let Some(name) = defaults_owned_child_name(child, &generated_scope)? {
            order.push(name.clone());
            supplied.insert(name, child.clone());
        }
    }
    let mut changed = previous
        .keys()
        .chain(supplied.keys())
        .filter(|name| previous.get(*name) != supplied.get(*name))
        .cloned()
        .collect::<Vec<_>>();
    changed.sort();
    changed.dedup();
    let mut selected = std::collections::HashMap::new();
    for (index, child) in source_children.iter().enumerate() {
        if let Some(name) = defaults_owned_child_name(child, &scope)? {
            selected.insert(name, index);
        }
    }
    let generated_bindings = namespace_bindings_at(&generated_root, &[])?;
    let end = std::str::from_utf8(root.name().as_ref())?.to_owned();
    let source_bindings =
        namespace_bindings_at(&root, &crate::numbering::namespace_bindings(&scope))?;
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Start(root))?;
    let mut inserted = std::collections::HashSet::new();
    for (index, child) in source_children.iter().enumerate() {
        let name = defaults_owned_child_name(child, &scope)?;
        if let Some(ref name) = name {
            let rank = order.iter().position(|value| value == name);
            for new_name in order
                .iter()
                .filter(|value| changed.contains(value) && !selected.contains_key(*value))
            {
                if !inserted.contains(new_name)
                    && rank.is_some_and(|rank| {
                        order.iter().position(|value| value == new_name).unwrap() < rank
                    })
                {
                    writer
                        .get_mut()
                        .extend_from_slice(&crate::text::raw_with_external_bindings(
                            &supplied[new_name],
                            &generated_bindings,
                        )?);
                    inserted.insert(new_name.clone());
                }
            }
            if local == b"pPr" && name == b"pBdr" && changed.contains(name) {
                if let Some(generated) = supplied.get(name) {
                    let generated_border =
                        crate::text::raw_with_external_bindings(generated, &generated_bindings)?;
                    if crate::borders::CT_PBdr::from_source_xml(&generated_border)?
                        == crate::borders::CT_PBdr::default()
                    {
                        let child =
                            crate::text::raw_with_external_bindings(child, &source_bindings)?;
                        writer
                            .get_mut()
                            .extend_from_slice(&crate::borders::CT_PBdr::merge_source(
                                &child, &child, None, false,
                            )?);
                    } else {
                        writer
                            .get_mut()
                            .extend_from_slice(&merge_style_border_source(
                                child,
                                previous.get(name),
                                generated,
                                &scope,
                                selected.get(name) == Some(&index),
                            )?);
                    }
                } else {
                    // Explicit clear removes each owned edge, not its surrounding
                    // source container or opaque siblings.
                    let child = crate::text::raw_with_external_bindings(child, &source_bindings)?;
                    writer
                        .get_mut()
                        .extend_from_slice(&crate::borders::CT_PBdr::merge_source(
                            &child, &child, None, false,
                        )?);
                }
                inserted.insert(name.clone());
                continue;
            }
            if changed.contains(name)
                && matches!(
                    name.as_slice(),
                    b"rFonts"
                        | b"sz"
                        | b"szCs"
                        | b"u"
                        | b"b"
                        | b"bCs"
                        | b"i"
                        | b"iCs"
                        | b"color"
                        | b"lang"
                        | b"spacing"
                        | b"ind"
                        | b"jc"
                        | b"keepNext"
                        | b"keepLines"
                        | b"pageBreakBefore"
                        | b"shd"
                        | b"caps"
                        | b"smallCaps"
                        | b"strike"
                        | b"dstrike"
                        | b"vanish"
                        | b"rtl"
                        | b"position"
                        | b"vertAlign"
                )
            {
                if !supplied.contains_key(name)
                    && !matches!(name.as_slice(), b"rFonts" | b"sz" | b"szCs")
                {
                    inserted.insert(name.clone());
                    continue;
                }
                writer
                    .get_mut()
                    .extend_from_slice(&merge_defaults_scalar_attributes(
                        child,
                        previous.get(name),
                        supplied.get(name),
                        &scope,
                        selected.get(name) == Some(&index),
                    )?);

                inserted.insert(name.clone());
                continue;
            }
            if changed.contains(name) && selected.get(name) == Some(&index) {
                if let Some(generated) = supplied.get(name) {
                    writer
                        .get_mut()
                        .extend_from_slice(&crate::text::raw_with_external_bindings(
                            generated,
                            &generated_bindings,
                        )?);
                }
                inserted.insert(name.clone());
                continue;
            }
        }
        writer.get_mut().extend_from_slice(child);
    }
    for name in order
        .iter()
        .filter(|name| changed.contains(name) && !inserted.contains(*name))
    {
        writer
            .get_mut()
            .extend_from_slice(&crate::text::raw_with_external_bindings(
                &supplied[name],
                &generated_bindings,
            )?);
    }
    writer.write_event(Event::End(BytesEnd::new(end)))?;
    Ok(writer.into_inner())
}

fn merge_defaults_property_wrapper(
    raw: &[u8],
    inherited: &[String],
    local: &[u8],
    generated: Option<&[u8]>,
) -> Result<Vec<u8>> {
    let (root, children) = defaults_source_children(raw)?;
    let scope = word_prefixes_at(&root, inherited)?;
    let mut selected = None;
    for (index, child) in children.iter().enumerate() {
        if let Some((start, _)) = defaults_optional_children(child)? {
            let scope = word_prefixes_at(&start, &scope)?;
            if is_word_element(start.name().as_ref(), local, &scope) {
                selected = Some(index);
            }
        }
    }
    let mut writer = Writer::new(Vec::new());
    let end = std::str::from_utf8(root.name().as_ref())?.to_owned();
    writer.write_event(Event::Start(root))?;
    for (index, child) in children.iter().enumerate() {
        if Some(index) == selected {
            if let Some(generated) = generated {
                writer
                    .get_mut()
                    .extend_from_slice(&merge_defaults_property_root(
                        child, generated, &scope, local,
                    )?);
            }
        } else {
            writer.get_mut().extend_from_slice(child);
        }
    }
    if selected.is_none()
        && let Some(generated) = generated
    {
        writer
            .get_mut()
            .extend_from_slice(&crate::text::raw_with_external_bindings(
                generated,
                &[("w".to_owned(), W_NS.to_owned())],
            )?);
    }
    writer.write_event(Event::End(BytesEnd::new(end)))?;
    Ok(writer.into_inner())
}

fn merge_defaults_source(
    raw: &[u8],
    run: Option<&[u8]>,
    paragraph: Option<&[u8]>,
    run_changed: bool,
    paragraph_changed: bool,
) -> Result<Vec<u8>> {
    let (root, children) = defaults_source_children(raw)?;
    let scope = word_prefixes_at(&root, &[])?;
    let mut run_slot = None;
    let mut paragraph_slot = None;
    for (index, child) in children.iter().enumerate() {
        if let Some((start, _)) = defaults_optional_children(child)? {
            let child_scope = word_prefixes_at(&start, &scope)?;
            if is_word_element(start.name().as_ref(), b"rPrDefault", &child_scope) {
                run_slot = Some(index);
            }
            if is_word_element(start.name().as_ref(), b"pPrDefault", &child_scope) {
                paragraph_slot = Some(index);
            }
        }
    }
    let mut writer = Writer::new(Vec::new());
    let end = std::str::from_utf8(root.name().as_ref())?.to_owned();
    writer.write_event(Event::Start(root))?;
    // Insert a newly supplied rPrDefault before any existing pPrDefault.
    if run_slot.is_none()
        && let Some(run) = run
    {
        writer.get_mut().extend_from_slice(&merge_defaults_property_wrapper(b"<w:rPrDefault xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>", &scope, b"rPr", Some(run))?);
    }
    for (index, child) in children.iter().enumerate() {
        if Some(index) == run_slot && run_changed {
            writer
                .get_mut()
                .extend_from_slice(&merge_defaults_property_wrapper(
                    child, &scope, b"rPr", run,
                )?);
        } else if Some(index) == paragraph_slot && paragraph_changed {
            writer
                .get_mut()
                .extend_from_slice(&merge_defaults_property_wrapper(
                    child, &scope, b"pPr", paragraph,
                )?);
        } else {
            writer.get_mut().extend_from_slice(child);
        }
    }
    if paragraph_slot.is_none()
        && let Some(paragraph) = paragraph
    {
        writer.get_mut().extend_from_slice(&merge_defaults_property_wrapper(b"<w:pPrDefault xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>", &scope, b"pPr", Some(paragraph))?);
    }
    writer.write_event(Event::End(BytesEnd::new(end)))?;
    Ok(writer.into_inner())
}

/// `CT_Styles` — The styles part (word/styles.xml).
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_Styles {
    pub doc_defaults: Option<CT_DocDefaults>,
    pub styles: Vec<CT_Style>,
    /// Root children the model does not read, such as `w:latentStyles`.
    /// They stay in source order and are written after `w:docDefaults` and
    /// before the first `w:style`.
    pub extra_xml: Vec<Vec<u8>>,
}

#[allow(non_snake_case)]
impl CT_Styles {
    pub fn new() -> Self {
        CT_Styles {
            doc_defaults: None,
            styles: Vec::new(),
            extra_xml: Vec::new(),
        }
    }

    /// Parse from XML bytes (the content of word/styles.xml).
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut doc_defaults = None;
        let mut styles = Vec::new();
        let mut extra_xml = Vec::new();
        let mut buf = Vec::new();
        let mut word_prefixes = Vec::new();
        let mut namespace_bindings = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, &word_prefixes)?;
                    if is_word_element(name.as_ref(), b"docDefaults", &prefixes) {
                        let bindings = namespace_bindings_at(e, &namespace_bindings)?;
                        reader.config_mut().trim_text(false);
                        let captured = capture_element(&mut reader, e)?;
                        reader.config_mut().trim_text(true);
                        let raw = crate::text::raw_with_external_bindings(&captured, &bindings)?;
                        doc_defaults = Some(CT_DocDefaults::from_source_xml(&raw, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"style", &prefixes) {
                        let bindings = namespace_bindings_at(e, &namespace_bindings)?;
                        styles.push(CT_Style::from_xml_with_prefixes(
                            &mut reader,
                            e,
                            &prefixes,
                            &bindings,
                        )?);
                    } else if is_word_element(name.as_ref(), b"styles", &prefixes) {
                        // Root element, continue
                        word_prefixes = prefixes;
                        namespace_bindings = namespace_bindings_at(e, &namespace_bindings)?;
                    } else {
                        let bindings = namespace_bindings_at(e, &namespace_bindings)?;
                        reader.config_mut().trim_text(false);
                        let captured = capture_element(&mut reader, e)?;
                        reader.config_mut().trim_text(true);
                        extra_xml.push(make_style_raw_self_contained(&captured, &bindings)?);
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let prefixes = word_prefixes_at(e, &word_prefixes)?;
                    let bindings = namespace_bindings_at(e, &namespace_bindings)?;
                    if is_word_element(e.name().as_ref(), b"docDefaults", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            &bindings,
                        )?;
                        doc_defaults = Some(CT_DocDefaults::from_source_xml(&raw, &prefixes)?);
                    } else if is_word_element(e.name().as_ref(), b"style", &prefixes) {
                        let owned = e.to_owned();
                        let mut empty = Reader::from_reader(&[] as &[u8]);
                        styles.push(CT_Style::from_xml_with_prefixes(
                            &mut empty, &owned, &prefixes, &bindings,
                        )?);
                    } else if !is_word_element(e.name().as_ref(), b"styles", &prefixes) {
                        extra_xml.push(make_style_raw_self_contained(
                            &capture_empty_element(e)?,
                            &bindings,
                        )?);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Styles {
            doc_defaults,
            styles,
            extra_xml,
        })
    }

    /// Serialize to XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut styles_start = BytesStart::new("w:styles");
        styles_start.push_attribute(("xmlns:w", W_NS));
        styles_start.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        let _binding_scope = crate::text::emitted_root_binding_scope(&styles_start, 0, false)?;
        writer.write_event(Event::Start(styles_start))?;

        if let Some(ref defaults) = self.doc_defaults {
            defaults.to_xml(&mut writer)?;
        }

        for raw in &self.extra_xml {
            writer.get_mut().extend_from_slice(raw);
        }

        for style in &self.styles {
            style.to_xml(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:styles")))?;

        Ok(writer.into_inner())
    }

    /// Find a style by its ID.
    pub fn get_by_id(&self, style_id: &str) -> Option<&CT_Style> {
        self.styles.iter().find(|s| s.style_id == style_id)
    }

    /// Find the default style for a given type.
    pub fn get_default(&self, style_type: StyleType) -> Option<&CT_Style> {
        self.styles
            .iter()
            .find(|s| s.style_type == style_type && s.is_default)
    }

    /// Create a minimal default styles part for a new document.
    pub fn new_default() -> Self {
        use crate::units::HalfPoint;

        let normal = CT_Style {
            style_id: "Normal".to_string(),
            style_type: StyleType::Paragraph,
            name: Some("Normal".to_string()),
            based_on: None,
            next_style: None,
            linked_style: None,
            auto_redefine: None,
            hidden: None,
            ui_priority: None,
            semi_hidden: None,
            unhide_when_used: None,
            quick_format: None,
            locked: None,
            is_default: true,
            ppr: None,
            rpr: None,
            table_properties: None,
            table_properties_original: None,
            table_properties_xml: None,
            table_row_properties: None,
            table_cell_properties: None,
            conditional_table_styles: Vec::new(),
            extra_attributes: Vec::new(),
            modeled_xml: Vec::new(),
            extra_xml: Vec::new(),
        };

        let heading1 = CT_Style {
            style_id: "Heading1".to_string(),
            style_type: StyleType::Paragraph,
            name: Some("heading 1".to_string()),
            based_on: Some("Normal".to_string()),
            next_style: Some("Normal".to_string()),
            linked_style: None,
            auto_redefine: None,
            hidden: None,
            ui_priority: None,
            semi_hidden: None,
            unhide_when_used: None,
            quick_format: None,
            locked: None,
            is_default: false,
            ppr: Some(CT_PPr {
                keep_next: Some(true),
                keep_lines: Some(true),
                space_before: Some(crate::units::Twips(240)),
                space_after: Some(crate::units::Twips(0)),
                ..Default::default()
            }),
            rpr: Some(CT_RPr {
                sz: Some(HalfPoint(32)),
                sz_cs: Some(HalfPoint(32)),
                bold: Some(true),
                bold_cs: Some(true),
                color: Some("2F5496".to_string()),
                ..Default::default()
            }),
            table_properties: None,
            table_properties_original: None,
            table_properties_xml: None,
            table_row_properties: None,
            table_cell_properties: None,
            conditional_table_styles: Vec::new(),
            extra_attributes: Vec::new(),
            modeled_xml: Vec::new(),
            extra_xml: Vec::new(),
        };

        let doc_defaults = CT_DocDefaults {
            source_xml: None,
            rpr: Some(CT_RPr {
                font_ascii: Some("Calibri".to_string()),
                font_hansi: Some("Calibri".to_string()),
                font_east_asia: Some("Calibri".to_string()),
                font_cs: Some("Times New Roman".to_string()),
                sz: Some(HalfPoint(22)),
                sz_cs: Some(HalfPoint(22)),
                ..Default::default()
            }),
            ppr: Some(CT_PPr {
                space_after: Some(crate::units::Twips(160)),
                line_spacing: Some(crate::units::Twips(259)),
                line_rule: Some("auto".to_string()),
                ..Default::default()
            }),
        };

        CT_Styles {
            doc_defaults: Some(doc_defaults),
            styles: vec![normal, heading1],
            extra_xml: Vec::new(),
        }
    }
}

impl Default for CT_Styles {
    fn default() -> Self {
        Self::new()
    }
}

fn style_child_rank(name: &[u8], word_prefixes: &[String]) -> u8 {
    let local = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
    if !is_word_element(name, local, word_prefixes) {
        return 26;
    }
    match local {
        b"name" => 0,
        b"aliases" => 1,
        b"basedOn" => 2,
        b"next" => 3,
        b"link" => 4,
        b"autoRedefine" => 5,
        b"hidden" => 6,
        b"uiPriority" => 7,
        b"semiHidden" => 8,
        b"unhideWhenUsed" => 9,
        b"qFormat" => 10,
        b"locked" => 11,
        b"personal" => 12,
        b"personalCompose" => 13,
        b"personalReply" => 14,
        b"rsid" => 19,
        b"pPr" => 20,
        b"rPr" => 21,
        b"tblPr" => 22,
        b"trPr" => 23,
        b"tcPr" => 24,
        b"tblStylePr" => 25,
        _ => 26,
    }
}

fn style_scalar_rank(name: &[u8], word_prefixes: &[String]) -> Option<u8> {
    let rank = style_child_rank(name, word_prefixes);
    matches!(rank, 0 | 2..=11).then_some(rank)
}

fn style_scalar_empty_requires_snapshot(element: &BytesStart<'_>, rank: u8) -> Result<bool> {
    let canonical_name: &[u8] = match rank {
        0 => b"w:name",
        2 => b"w:basedOn",
        3 => b"w:next",
        4 => b"w:link",
        5 => b"w:autoRedefine",
        6 => b"w:hidden",
        7 => b"w:uiPriority",
        8 => b"w:semiHidden",
        9 => b"w:unhideWhenUsed",
        10 => b"w:qFormat",
        11 => b"w:locked",
        _ => unreachable!("style scalar rank"),
    };
    if element.name().as_ref() != canonical_name {
        return Ok(true);
    }

    for attribute in element.attributes() {
        if attribute?.key.as_ref() != b"w:val" {
            return Ok(true);
        }
    }
    Ok(false)
}

#[allow(clippy::too_many_arguments)]
fn apply_style_scalar(
    rank: u8,
    string_value: Option<String>,
    toggle_value: Option<bool>,
    priority_value: Option<u32>,
    name: &mut Option<String>,
    based_on: &mut Option<String>,
    next_style: &mut Option<String>,
    linked_style: &mut Option<String>,
    auto_redefine: &mut Option<bool>,
    hidden: &mut Option<bool>,
    ui_priority: &mut Option<u32>,
    semi_hidden: &mut Option<bool>,
    unhide_when_used: &mut Option<bool>,
    quick_format: &mut Option<bool>,
    locked: &mut Option<bool>,
) {
    match rank {
        0 => *name = string_value,
        2 => *based_on = string_value,
        3 => *next_style = string_value,
        4 => *linked_style = string_value,
        5 => *auto_redefine = toggle_value,
        6 => *hidden = toggle_value,
        7 => *ui_priority = priority_value,
        8 => *semi_hidden = toggle_value,
        9 => *unhide_when_used = toggle_value,
        10 => *quick_format = toggle_value,
        11 => *locked = toggle_value,
        _ => unreachable!("style scalar rank"),
    }
}

fn write_style_extras<W: std::io::Write>(
    writer: &mut Writer<W>,
    extras: &[&(u8, Vec<u8>)],
    index: &mut usize,
    before_rank: u8,
) -> Result<()> {
    while let Some((rank, raw)) = extras.get(*index).copied() {
        if *rank >= before_rank {
            break;
        }
        writer.get_mut().write_all(raw)?;
        *index += 1;
    }
    Ok(())
}

fn namespace_bindings_at(
    element: &BytesStart,
    inherited: &[(String, String)],
) -> Result<Vec<(String, String)>> {
    let mut bindings = inherited.to_vec();
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let prefix = if key == b"xmlns" {
            Some(String::new())
        } else {
            key.strip_prefix(b"xmlns:")
                .map(|prefix| String::from_utf8_lossy(prefix).into_owned())
        };
        let Some(prefix) = prefix else {
            continue;
        };
        let uri = std::str::from_utf8(&attribute.value)?.to_owned();
        if let Some(binding) = bindings
            .iter_mut()
            .find(|(candidate, _)| candidate == &prefix)
        {
            binding.1 = uri;
        } else {
            bindings.push((prefix, uri));
        }
    }
    Ok(bindings)
}

fn reject_conflicting_style_prefixes(element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        let expected = match attribute.key.as_ref() {
            b"xmlns:w" => Some(W_NS),
            b"xmlns:r" => {
                Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            }
            _ => None,
        };
        let Some(expected) = expected else {
            continue;
        };
        let value = std::str::from_utf8(&attribute.value)?;
        if value != expected {
            return Err(OxmlError::InvalidValue(format!(
                "{} conflicts with the fixed style writer namespace",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
    }
    Ok(())
}

fn raw_style_attributes(
    element: &BytesStart<'_>,
    modeled: &[&[u8]],
    word_prefixes: &[String],
    namespace_bindings: &[(String, String)],
) -> Result<Vec<(String, String)>> {
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
            .into_owned();
        let canonical_namespace = (key == b"xmlns:w" && value == W_NS)
            || (key == b"xmlns:r"
                && value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        if !canonical_namespace
            && !modeled
                .iter()
                .any(|local| is_word_attribute(key, local, word_prefixes))
        {
            attributes.push((std::str::from_utf8(key)?.to_owned(), value));
        }
    }
    if !attributes.is_empty() {
        for (prefix, namespace) in namespace_bindings {
            if prefix.is_empty()
                || (prefix == "w" && namespace == W_NS)
                || (prefix == "r"
                    && namespace
                        == "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            {
                continue;
            }
            let declaration = format!("xmlns:{prefix}");
            if !attributes.iter().any(|(name, _)| name == &declaration) {
                attributes.push((declaration, namespace.clone()));
            }
        }
    }
    Ok(attributes)
}

fn reject_conflicting_style_bindings(bindings: &[(String, String)]) -> Result<()> {
    for (prefix, namespace) in bindings {
        let expected = match prefix.as_str() {
            "w" => Some(W_NS),
            "r" => Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
            _ => None,
        };
        if expected.is_some_and(|expected| namespace != expected) {
            return Err(OxmlError::InvalidValue(format!(
                "xmlns:{prefix} conflicts with the fixed style writer namespace"
            )));
        }
    }
    Ok(())
}

fn push_style_attributes(element: &mut BytesStart<'_>, attributes: &[(String, String)]) {
    for (name, value) in attributes {
        element.push_attribute((name.as_str(), value.as_str()));
    }
}

fn make_style_raw_self_contained(
    raw: &[u8],
    namespace_bindings: &[(String, String)],
) -> Result<Vec<u8>> {
    let external_bindings = namespace_bindings
        .iter()
        .filter(|(prefix, namespace)| {
            !((prefix == "w" && namespace == W_NS)
                || (prefix == "r"
                    && namespace
                        == "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))
        })
        .cloned()
        .collect::<Vec<_>>();
    crate::text::raw_with_external_bindings(raw, &external_bindings)
}

fn parse_style_table_properties(raw: &[u8], word_prefixes: &[String]) -> Result<CT_TblPr> {
    let mut reader = Reader::from_reader(raw);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref start)
                if is_word_element(start.name().as_ref(), b"tblPr", word_prefixes) =>
            {
                let prefixes = word_prefixes_at(start, word_prefixes)?;
                return CT_TblPr::from_xml_with_prefixes(&mut reader, &prefixes);
            }
            Event::Eof => return Ok(CT_TblPr::default()),
            _ => {}
        }
        buf.clear();
    }
}

fn explicit_style_element_has_only_trivia(raw: &[u8]) -> Result<bool> {
    let mut reader = Reader::from_reader(raw);
    let mut inside = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(_) if !inside => inside = true,
            Event::End(_) if inside => return Ok(true),
            Event::Text(text) if inside && text.as_ref().iter().all(u8::is_ascii_whitespace) => {}
            Event::Comment(_) | Event::PI(_) if inside => {}
            Event::Eof => return Ok(false),
            _ if inside => return Ok(false),
            _ => {}
        }
        buf.clear();
    }
}

/// The five typed property layers a `w:tblStylePr` region carries.
#[derive(Default)]
struct ConditionalStyleLayers {
    paragraph_properties: Option<CT_PPr>,
    run_properties: Option<Box<CT_RPr>>,
    table_properties: Option<CT_TblPr>,
    row_properties: Option<Box<CT_TrPr>>,
    cell_properties: Option<CT_TcPr>,
}

fn parse_conditional_style_properties(
    raw: &[u8],
    word_prefixes: &[String],
) -> Result<ConditionalStyleLayers> {
    let mut reader = Reader::from_reader(raw);
    let mut layers = ConditionalStyleLayers::default();
    let mut buf = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref start) => {
                depth += 1;
                let prefixes = word_prefixes_at(start, word_prefixes)?;
                if depth != 2 {
                    // Only the region's own children are layers. Anything
                    // deeper belongs to a layer or to a preserved subtree.
                } else if is_word_element(start.name().as_ref(), b"pPr", &prefixes) {
                    let captured = capture_element(&mut reader, start)?;
                    layers.paragraph_properties = Some(parse_scoped_ppr(&captured, &prefixes)?);
                    depth -= 1;
                } else if is_word_element(start.name().as_ref(), b"rPr", &prefixes) {
                    let captured = capture_element(&mut reader, start)?;
                    layers.run_properties = Some(Box::new(parse_style_rpr(&captured, &prefixes)?));
                    depth -= 1;
                } else if is_word_element(start.name().as_ref(), b"tblPr", &prefixes) {
                    layers.table_properties =
                        Some(CT_TblPr::from_xml_with_prefixes(&mut reader, &prefixes)?);
                    depth -= 1;
                } else if is_word_element(start.name().as_ref(), b"trPr", &prefixes) {
                    layers.row_properties = Some(Box::new(CT_TrPr::from_xml_with_prefixes(
                        &mut reader,
                        &prefixes,
                    )?));
                    depth -= 1;
                } else if is_word_element(start.name().as_ref(), b"tcPr", &prefixes) {
                    layers.cell_properties =
                        Some(CT_TcPr::from_xml_with_prefixes(&mut reader, &prefixes)?);
                    depth -= 1;
                }
            }
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(layers)
}

fn serialize_conditional_table_style(conditional: &CT_TblStylePr) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(conditional.raw_xml.as_slice());
    let mut prefixes = vec!["w".to_owned()];
    let mut original_region = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref start) => {
                prefixes = word_prefixes_at(start, &prefixes)?;
                original_region = get_word_attr(start, b"type", &prefixes)?;
                break;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    let Some(region) = conditional.region else {
        // An unrecognised `w:type` was never projected, so its preserved bytes
        // are the only faithful serialisation of it.
        return Ok(conditional.raw_xml.clone());
    };
    let original = parse_conditional_style_properties(&conditional.raw_xml, &prefixes)?;
    if original_region.as_deref() == Some(region.to_str())
        && original.paragraph_properties == conditional.paragraph_properties
        && original.run_properties == conditional.run_properties
        && original.table_properties == conditional.table_properties
        && original.row_properties == conditional.row_properties
        && original.cell_properties == conditional.cell_properties
    {
        return Ok(conditional.raw_xml.clone());
    }

    let mut extras = preserved_conditional_style_children(&conditional.raw_xml)?;
    extras.sort_by_key(|(rank, _)| *rank);
    let extra_refs = extras.iter().collect::<Vec<_>>();
    let mut extra_index = 0;
    let mut writer = Writer::new(Vec::new());
    let mut start = BytesStart::new("w:tblStylePr");
    start.push_attribute(("w:type", region.to_str()));
    push_style_attributes(&mut start, &conditional.extra_attributes);
    writer.write_event(Event::Start(start))?;
    write_style_extras(&mut writer, &extra_refs, &mut extra_index, 0)?;
    if let Some(properties) = &conditional.paragraph_properties {
        properties.to_xml(&mut writer)?;
    }
    write_style_extras(&mut writer, &extra_refs, &mut extra_index, 1)?;
    if let Some(properties) = &conditional.run_properties {
        properties.to_xml(&mut writer)?;
    }
    write_style_extras(&mut writer, &extra_refs, &mut extra_index, 2)?;
    if let Some(properties) = &conditional.table_properties {
        properties.to_xml(&mut writer)?;
    }
    write_style_extras(&mut writer, &extra_refs, &mut extra_index, 3)?;
    if let Some(properties) = &conditional.row_properties {
        properties.to_xml(&mut writer)?;
    }
    write_style_extras(&mut writer, &extra_refs, &mut extra_index, 4)?;
    if let Some(properties) = &conditional.cell_properties {
        properties.to_xml(&mut writer)?;
    }
    write_style_extras(&mut writer, &extra_refs, &mut extra_index, u8::MAX)?;
    writer.write_event(Event::End(BytesEnd::new("w:tblStylePr")))?;
    Ok(writer.into_inner())
}

fn preserved_conditional_style_children(raw: &[u8]) -> Result<Vec<(u8, Vec<u8>)>> {
    let mut reader = Reader::from_reader(raw);
    let mut inside = false;
    let mut word_prefixes = vec!["w".to_owned()];
    let mut extras = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref start) if !inside => {
                word_prefixes = word_prefixes_at(start, &word_prefixes)?;
                inside = true;
            }
            Event::Start(ref start) => {
                let prefixes = word_prefixes_at(start, &word_prefixes)?;
                let local = start.local_name();
                if is_word_element(start.name().as_ref(), local.as_ref(), &prefixes)
                    && matches!(
                        local.as_ref(),
                        b"pPr" | b"rPr" | b"tblPr" | b"trPr" | b"tcPr"
                    )
                {
                    reader.read_to_end_into(start.name(), &mut Vec::new())?;
                } else {
                    extras.push((5, capture_element(&mut reader, start)?));
                }
            }
            Event::Empty(ref empty) => {
                let prefixes = word_prefixes_at(empty, &word_prefixes)?;
                let local = empty.local_name();
                if !(is_word_element(empty.name().as_ref(), local.as_ref(), &prefixes)
                    && matches!(
                        local.as_ref(),
                        b"pPr" | b"rPr" | b"tblPr" | b"trPr" | b"tcPr"
                    ))
                {
                    extras.push((5, capture_empty_element(empty)?));
                }
            }
            Event::End(_) | Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(extras)
}

/// Extract the `w:val` attribute from an element.
fn get_val_attr(e: &BytesStart, word_prefixes: &[String]) -> Result<Option<String>> {
    get_word_attr(e, b"val", word_prefixes)
}

fn get_word_attr(e: &BytesStart, local: &[u8], word_prefixes: &[String]) -> Result<Option<String>> {
    for attr in e.attributes() {
        let attr = attr?;
        if is_word_attribute(attr.key.as_ref(), local, word_prefixes) {
            return Ok(Some(std::str::from_utf8(&attr.value)?.to_string()));
        }
    }
    Ok(None)
}

fn parse_style_toggle(e: &BytesStart, word_prefixes: &[String]) -> Result<Option<bool>> {
    let value = get_word_attr(e, b"val", word_prefixes)?;
    match value.as_deref() {
        None | Some("1" | "true" | "on") => Ok(Some(true)),
        Some("0" | "false" | "off") => Ok(Some(false)),
        Some(value) => Err(OxmlError::InvalidValue(format!(
            "invalid style toggle value: {value}"
        ))),
    }
}

fn parse_style_priority(e: &BytesStart, word_prefixes: &[String]) -> Result<Option<u32>> {
    get_word_attr(e, b"val", word_prefixes)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| OxmlError::InvalidValue(format!("invalid style UI priority: {value}")))
        })
        .transpose()
}

fn preserved_style_attribute(raw: &[u8], local: &[u8]) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(raw);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref start) | Event::Empty(ref start) => {
                let prefixes = word_prefixes_at(start, &["w".to_owned()])?;
                return get_word_attr(start, local, &prefixes);
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buf.clear();
    }
}

fn preserved_modeled_xml(modeled_xml: &[(u8, Vec<u8>)], rank: u8) -> Option<&[u8]> {
    modeled_xml
        .iter()
        .find(|(candidate, _)| *candidate == rank)
        .map(|(_, raw)| raw.as_slice())
}

fn rewritten_style_scalar_start(
    element: &BytesStart<'_>,
    value: &str,
) -> Result<BytesStart<'static>> {
    let prefixes = word_prefixes_at(element, &["w".to_owned()])?;
    let element_name = std::str::from_utf8(element.name().as_ref())?.to_owned();
    let mut rewritten = BytesStart::new(element_name.clone());
    let mut wrote_value = false;
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = std::str::from_utf8(attribute.key.as_ref())?.to_owned();
        if is_word_attribute(attribute.key.as_ref(), b"val", &prefixes) {
            if !wrote_value {
                rewritten.push_attribute((key.as_str(), value));
                wrote_value = true;
            }
        } else {
            let attribute_value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
                .into_owned();
            rewritten.push_attribute((key.as_str(), attribute_value.as_str()));
        }
    }
    if !wrote_value {
        let element_prefix = element_name
            .split_once(':')
            .map(|(prefix, _)| prefix.to_owned());
        let existing_prefix = element_prefix.or_else(|| {
            prefixes
                .iter()
                .find(|prefix| !prefix.is_empty() && !prefix.starts_with('\0'))
                .cloned()
        });
        let prefix = existing_prefix.unwrap_or_else(|| {
            let mut candidate = "rdocxStyle".to_owned();
            let mut suffix = 1usize;
            while element
                .attributes()
                .filter_map(std::result::Result::ok)
                .any(|attribute| attribute.key.as_ref() == format!("xmlns:{candidate}").as_bytes())
            {
                candidate = format!("rdocxStyle{suffix}");
                suffix += 1;
            }
            let declaration = format!("xmlns:{candidate}");
            rewritten.push_attribute((declaration.as_str(), W_NS));
            candidate
        });
        let key = format!("{prefix}:val");
        rewritten.push_attribute((key.as_str(), value));
    }
    Ok(rewritten.into_owned())
}

fn write_rewritten_style_scalar<W: std::io::Write>(
    writer: &mut Writer<W>,
    raw: &[u8],
    value: &str,
) -> Result<()> {
    let mut reader = Reader::from_reader(raw);
    let mut root_pending = true;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(element) if root_pending => {
                root_pending = false;
                writer.write_event(Event::Start(rewritten_style_scalar_start(&element, value)?))?;
            }
            Event::Empty(element) if root_pending => {
                writer.write_event(Event::Empty(rewritten_style_scalar_start(&element, value)?))?;
                break;
            }
            Event::Eof => break,
            event => writer.write_event(event.into_owned())?,
        }
        buf.clear();
    }
    Ok(())
}

fn write_style_string<W: std::io::Write>(
    writer: &mut Writer<W>,
    name: &str,
    rank: u8,
    value: Option<&str>,
    modeled_xml: &[(u8, Vec<u8>)],
) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    if let Some(raw) = preserved_modeled_xml(modeled_xml, rank)
        && preserved_style_attribute(raw, b"val")?.as_deref() == Some(value)
    {
        writer.get_mut().write_all(raw)?;
        return Ok(());
    }
    if let Some(raw) = preserved_modeled_xml(modeled_xml, rank) {
        return write_rewritten_style_scalar(writer, raw, value);
    }
    let mut element = BytesStart::new(name);
    element.push_attribute(("w:val", value));
    writer.write_event(Event::Empty(element))?;
    Ok(())
}

fn write_style_toggle<W: std::io::Write>(
    writer: &mut Writer<W>,
    name: &str,
    rank: u8,
    value: Option<bool>,
    modeled_xml: &[(u8, Vec<u8>)],
) -> Result<()> {
    if let Some(value) = value {
        if let Some(raw) = preserved_modeled_xml(modeled_xml, rank) {
            let prefixes = vec!["w".to_owned()];
            let mut reader = Reader::from_reader(raw);
            let mut buf = Vec::new();
            let preserved = loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(ref start) | Event::Empty(ref start) => {
                        let prefixes = word_prefixes_at(start, &prefixes)?;
                        break parse_style_toggle(start, &prefixes)?;
                    }
                    Event::Eof => break None,
                    _ => {}
                }
                buf.clear();
            };
            if preserved == Some(value) {
                writer.get_mut().write_all(raw)?;
                return Ok(());
            }
            return write_rewritten_style_scalar(writer, raw, if value { "1" } else { "0" });
        }
        let mut element = BytesStart::new(name);
        if !value {
            element.push_attribute(("w:val", "0"));
        }
        writer.write_event(Event::Empty(element))?;
    }
    Ok(())
}

fn write_style_priority<W: std::io::Write>(
    writer: &mut Writer<W>,
    priority: Option<u32>,
    modeled_xml: &[(u8, Vec<u8>)],
) -> Result<()> {
    let Some(priority) = priority else {
        return Ok(());
    };
    if let Some(raw) = preserved_modeled_xml(modeled_xml, 7)
        && preserved_style_attribute(raw, b"val")?
            .as_deref()
            .and_then(|value| value.parse::<u32>().ok())
            == Some(priority)
    {
        writer.get_mut().write_all(raw)?;
        return Ok(());
    }
    let value = priority.to_string();
    if let Some(raw) = preserved_modeled_xml(modeled_xml, 7) {
        return write_rewritten_style_scalar(writer, raw, &value);
    }
    let mut element = BytesStart::new("w:uiPriority");
    element.push_attribute(("w:val", value.as_str()));
    writer.write_event(Event::Empty(element))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unmodelled_root_children_round_trip_before_the_styles() {
        let xml = format!(
            concat!(
                r#"<w:styles xmlns:w="{}" xmlns:x="urn:x">"#,
                r#"<w:docDefaults/>"#,
                r#"<w:latentStyles w:count="1"><w:lsdException w:name="Normal" x:a="1"/></w:latentStyles>"#,
                r#"<x:empty/>"#,
                r#"<w:style w:type="paragraph" w:styleId="Empty"/>"#,
                r#"<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>"#,
                r#"</w:styles>"#,
            ),
            W_NS
        );
        let styles = CT_Styles::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(styles.extra_xml.len(), 2);
        assert_eq!(styles.styles.len(), 2);
        assert_eq!(styles.styles[0].style_id, "Empty");
        let written = String::from_utf8(styles.to_xml().unwrap()).unwrap();
        let latent = written
            .find(concat!(
                r#"<w:latentStyles w:count="1" xmlns:x="urn:x">"#,
                r#"<w:lsdException w:name="Normal" x:a="1"/></w:latentStyles>"#,
            ))
            .unwrap_or_else(|| panic!("{written}"));
        let empty = written.find(r#"<x:empty xmlns:x="urn:x"/>"#).unwrap();
        let empty_style = written
            .find(r#"<w:style w:type="paragraph" w:styleId="Empty">"#)
            .unwrap();
        assert!(written.find("<w:docDefaults").unwrap() < latent);
        assert!(latent < empty);
        assert!(empty < empty_style);
        assert_eq!(CT_Styles::from_xml(written.as_bytes()).unwrap(), styles);
    }

    #[test]
    fn round_trip_styles() {
        let styles = CT_Styles::new_default();
        let xml = styles.to_xml().unwrap();
        let parsed = CT_Styles::from_xml(&xml).unwrap();

        // Namespace-close only the retained owners. Every other byte remains exact.
        let original = String::from_utf8(xml.clone()).unwrap();
        let (prefix, heading) = original
            .split_once(r#"<w:style w:type="paragraph" w:styleId="Heading1">"#)
            .unwrap();
        assert_eq!(prefix.matches("<w:docDefaults>").count(), 1);
        let prefix = prefix.replace(
            "<w:docDefaults>",
            &format!("<w:docDefaults xmlns:w=\"{W_NS}\">"),
        );
        let mut heading = heading.to_owned();
        for owner in ["pPr", "rPr"] {
            let opening = format!("<w:{owner}>");
            assert_eq!(heading.matches(&opening).count(), 1);
            heading = heading.replace(&opening, &format!("<w:{owner} xmlns:w=\"{W_NS}\">"));
        }
        let expected =
            format!(r#"{prefix}<w:style w:type="paragraph" w:styleId="Heading1">{heading}"#);
        let imported = parsed.to_xml().unwrap();
        assert_eq!(imported, expected.as_bytes());
        let reopened = CT_Styles::from_xml(&imported).unwrap();
        assert_eq!(reopened.to_xml().unwrap(), imported);
        assert_eq!(reopened, parsed);

        assert_eq!(parsed.styles.len(), 2);
        assert!(parsed.doc_defaults.is_some());

        let normal = parsed.get_by_id("Normal").unwrap();
        assert_eq!(normal.name, Some("Normal".to_string()));
        assert!(normal.is_default);

        let h1 = parsed.get_by_id("Heading1").unwrap();
        assert_eq!(h1.based_on, Some("Normal".to_string()));
    }

    #[test]
    fn find_default_style() {
        let styles = CT_Styles::new_default();
        let default_para = styles.get_default(StyleType::Paragraph).unwrap();
        assert_eq!(default_para.style_id, "Normal");
    }

    #[test]
    fn linked_style_and_ui_flags_accept_aliases_and_serialize_in_schema_order() {
        let xml = format!(
            r#"<q:styles xmlns:q="{W_NS}" xmlns:x="urn:producer"><q:style q:type="paragraph" q:styleId="Corpus" x:root="&amp;"><q:name q:val="Corpus"> <!--name-trivia--> </q:name><q:next q:val="Normal"><?next kept?></q:next><q:link q:val="CorpusChar"> <!--link-trivia--> </q:link><q:autoRedefine q:val="0" x:keep="yes"/><q:hidden> </q:hidden><q:uiPriority q:val="17"></q:uiPriority><q:semiHidden q:val="false"></q:semiHidden><q:unhideWhenUsed></q:unhideWhenUsed><q:qFormat q:val="1"></q:qFormat><q:locked q:val="true"></q:locked><q:pPr><q:spacing q:after="0"/></q:pPr></q:style></q:styles>"#
        );
        let mut styles = CT_Styles::from_xml(xml.as_bytes()).unwrap();
        let style = styles.get_by_id("Corpus").unwrap();
        assert_eq!(style.next_style.as_deref(), Some("Normal"));
        assert_eq!(style.linked_style.as_deref(), Some("CorpusChar"));
        assert_eq!(style.auto_redefine, Some(false));
        assert_eq!(style.hidden, Some(true));
        assert_eq!(style.ui_priority, Some(17));
        assert_eq!(style.semi_hidden, Some(false));
        assert_eq!(style.unhide_when_used, Some(true));
        assert_eq!(style.quick_format, Some(true));
        assert_eq!(style.locked, Some(true));

        let serialized = String::from_utf8(styles.to_xml().unwrap()).unwrap();
        let positions = [
            "<q:next",
            "<q:link",
            "<q:autoRedefine",
            "<q:hidden",
            "<q:uiPriority",
            "<q:semiHidden",
            "<q:unhideWhenUsed",
            "<q:qFormat",
            "<q:locked",
            "<q:pPr",
        ]
        .map(|token| serialized.find(token).unwrap());
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
        assert!(serialized.contains(&format!(
            r#"<q:pPr xmlns:q="{W_NS}"><q:spacing q:after="0"/></q:pPr>"#
        )));
        assert_eq!(serialized.matches("name-trivia").count(), 1);
        assert_eq!(serialized.matches("next kept").count(), 1);
        assert_eq!(serialized.matches("link-trivia").count(), 1);
        assert_eq!(serialized.matches(r#"x:root="&amp;""#).count(), 1);
        assert_eq!(serialized.matches(r#"x:keep="yes""#).count(), 1);
        let reparsed = CT_Styles::from_xml(serialized.as_bytes()).unwrap();
        assert_eq!(reparsed.styles[0], styles.styles[0]);

        styles.styles[0].next_style = Some("Other".to_owned());
        styles.styles[0].auto_redefine = Some(true);
        styles.styles[0].ui_priority = Some(23);
        let changed = String::from_utf8(styles.to_xml().unwrap()).unwrap();
        assert!(changed.contains("<q:next"));
        assert!(changed.contains(r#"q:val="Other""#));
        assert_eq!(changed.matches("next kept").count(), 1);
        assert!(changed.contains("<q:autoRedefine"));
        assert!(changed.contains(r#"q:val="1""#));
        assert_eq!(changed.matches(r#"x:keep="yes""#).count(), 1);
        assert!(changed.contains("<q:uiPriority"));
        assert!(changed.contains(r#"q:val="23""#));
        let reparsed = CT_Styles::from_xml(changed.as_bytes()).unwrap();
        assert_eq!(reparsed.styles[0].next_style.as_deref(), Some("Other"));
        assert_eq!(reparsed.styles[0].auto_redefine, Some(true));
        assert_eq!(reparsed.styles[0].ui_priority, Some(23));
    }

    #[test]
    fn style_parser_rejects_prefixes_that_conflict_with_fixed_writer_bindings() {
        let xml = format!(
            r#"<q:styles xmlns:q="{W_NS}" xmlns:w="urn:not-word"><q:style q:type="paragraph" q:styleId="Corpus"></q:style></q:styles>"#
        );
        let error = CT_Styles::from_xml(xml.as_bytes()).unwrap_err();
        assert!(error.to_string().contains("xmlns:w"));
    }

    #[test]
    fn changed_table_style_properties_keep_one_unknown_child_in_schema_order() {
        let xml = format!(
            r#"<q:styles xmlns:q="{W_NS}" xmlns:x="urn:producer"><q:style q:type="table" q:styleId="Corpus"><q:tblPr><q:tblStyleRowBandSize q:val="2"/><x:producer x:exact="kept"/><q:shd q:val="clear" q:fill="EEEEEE"/></q:tblPr></q:style></q:styles>"#
        );
        let mut styles = CT_Styles::from_xml(xml.as_bytes()).unwrap();
        styles.styles[0]
            .table_properties
            .as_mut()
            .unwrap()
            .shading
            .as_mut()
            .unwrap()
            .fill = Some("D9EAF7".to_owned());

        let serialized = String::from_utf8(styles.to_xml().unwrap()).unwrap();
        assert_eq!(serialized.matches("tblStyleRowBandSize").count(), 1);
        assert_eq!(serialized.matches("x:producer").count(), 1);
        assert!(
            serialized.find("tblStyleRowBandSize").unwrap() < serialized.find("w:shd").unwrap()
        );
        assert_eq!(
            CT_Styles::from_xml(serialized.as_bytes()).unwrap().styles[0]
                .table_properties
                .as_ref()
                .and_then(|properties| properties.shading.as_ref())
                .and_then(|shading| shading.fill.as_deref()),
            Some("D9EAF7")
        );
    }

    #[test]
    fn aliased_style_paragraph_properties_use_ancestor_namespace_scope() {
        let xml = format!(
            r#"<q:styles xmlns:q="{W_NS}" xmlns:ext="urn:producer"><q:style q:type="paragraph" q:styleId="Alias"><q:pPr><ext:jc ext:val="right"/><q:jc q:val="center"/></q:pPr></q:style></q:styles>"#
        );
        let parsed = CT_Styles::from_xml(xml.as_bytes()).unwrap();
        let ppr = parsed.styles[0].ppr.as_ref().unwrap();
        assert_eq!(ppr.jc, Some(crate::shared::ST_Jc::Center));
    }

    #[test]
    fn direct_style_parser_uses_supplied_start_ancestor_scope() {
        let xml = format!(
            r#"<outer xmlns:ext="urn:producer"><q:style xmlns:q="{W_NS}" q:type="paragraph" q:styleId="Direct"><ext:pPr><ext:jc ext:val="right"/></ext:pPr><q:pPr><ext:jc ext:val="right"/><q:jc q:val="center"/></q:pPr></q:style></outer>"#
        );
        let mut reader = Reader::from_str(&xml);
        let mut buf = Vec::new();
        let parsed = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"style" => {
                    break CT_Style::from_xml(&mut reader, element).unwrap();
                }
                Ok(Event::Eof) => panic!("missing style"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        };
        assert_eq!(
            parsed.ppr.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn direct_style_parser_does_not_promote_foreign_start_prefix() {
        let xml = r#"<outer><ext:style xmlns:ext="urn:producer" ext:type="paragraph" ext:styleId="Foreign"><ext:pPr><ext:jc ext:val="right"/></ext:pPr></ext:style></outer>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();
        let parsed = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"style" => {
                    break CT_Style::from_xml(&mut reader, element).unwrap();
                }
                Ok(Event::Eof) => panic!("missing style"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        };
        assert!(parsed.ppr.is_none());
    }

    #[test]
    fn direct_style_parser_accepts_default_word_namespace() {
        let xml = format!(
            r#"<outer xmlns:ext="urn:producer"><style xmlns="{W_NS}" xmlns:w="{W_NS}" w:type="paragraph" w:styleId="Direct"><ext:pPr><ext:jc ext:val="right"/></ext:pPr><pPr><ext:jc ext:val="right"/><jc w:val="center"/></pPr></style></outer>"#
        );
        let mut reader = Reader::from_str(&xml);
        let mut buf = Vec::new();
        let parsed = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"style" => {
                    break CT_Style::from_xml(&mut reader, element).unwrap();
                }
                Ok(Event::Eof) => panic!("missing style"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        };
        assert_eq!(
            parsed.ppr.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn table_style_properties_are_namespace_aware_schema_ordered_and_preserved() {
        let xml = format!(
            r#"<q:styles xmlns:q="{W_NS}" xmlns:bad="urn:not-word" xmlns:x="urn:region"><q:style q:type="table" q:styleId="Dense"><bad:tblPr><bad:tblBorders><bad:top bad:val="double"/></bad:tblBorders></bad:tblPr><q:pPr><q:spacing q:after="40"/></q:pPr><q:rPr><q:b/></q:rPr><q:tblPr><q:tblBorders><q:top q:val="single" q:sz="8" q:color="112233"/></q:tblBorders><ext:keep xmlns:ext="urn:producer" ext:value="byte-identical"/></q:tblPr><q:tblStylePr q:type="firstRow" x:region="&amp;"><q:pPr><q:spacing q:after="0"/></q:pPr><ext:conditional xmlns:ext="urn:producer" ext:value="preserved"/><q:tcPr><q:shd q:val="clear" q:fill="AABBCC"/></q:tcPr></q:tblStylePr></q:style></q:styles>"#
        );
        let mut styles = CT_Styles::from_xml(xml.as_bytes()).unwrap();
        let style = styles.get_by_id("Dense").unwrap();
        assert_eq!(
            style
                .table_properties
                .as_ref()
                .and_then(|properties| properties.borders.as_ref())
                .and_then(|borders| borders.top.as_ref())
                .and_then(|border| border.sz),
            Some(8)
        );
        assert_eq!(style.conditional_table_styles.len(), 1);
        assert_eq!(
            style.conditional_table_styles[0].region,
            Some(TableStyleRegion::FirstRow)
        );

        let serialized = String::from_utf8(styles.to_xml().unwrap()).unwrap();
        assert_eq!(serialized.matches("<q:tblPr").count(), 1);
        assert!(
            serialized
                .contains(r#"<ext:keep xmlns:ext="urn:producer" ext:value="byte-identical"/>"#)
        );
        let ppr = serialized.find("<q:pPr").unwrap();
        let rpr = serialized.find("<q:rPr").unwrap();
        let table_properties = serialized.find("<q:tblPr").unwrap();
        let conditional = serialized.find("<q:tblStylePr").unwrap();
        assert!(ppr < rpr && rpr < table_properties && table_properties < conditional);
        assert!(serialized.contains(&format!(
            r#"<q:pPr xmlns:q="{W_NS}"><q:spacing q:after="40"/></q:pPr>"#
        )));
        assert!(serialized.contains(&format!(r#"<q:rPr xmlns:q="{W_NS}"><q:b/></q:rPr>"#)));
        assert!(serialized.contains(r#"<bad:tblPr xmlns:bad="urn:not-word">"#));
        CT_Styles::from_xml(serialized.as_bytes()).expect("preserved prefixes remain bound");

        styles.styles[0]
            .table_properties
            .as_mut()
            .unwrap()
            .borders
            .as_mut()
            .unwrap()
            .top
            .as_mut()
            .unwrap()
            .color = Some("445566".to_owned());
        let changed = String::from_utf8(styles.to_xml().unwrap()).unwrap();
        assert_eq!(changed.matches("<w:tblPr>").count(), 1);
        assert_eq!(changed.matches("445566").count(), 1);
        assert!(
            changed.contains(r#"<ext:keep xmlns:ext="urn:producer" ext:value="byte-identical"/>"#)
        );
        styles.styles[0].conditional_table_styles[0]
            .cell_properties
            .as_mut()
            .unwrap()
            .shading
            .as_mut()
            .unwrap()
            .fill = Some("DDEEFF".to_owned());
        styles.styles[0].conditional_table_styles[0].region = Some(TableStyleRegion::LastRow);
        let changed = String::from_utf8(styles.to_xml().unwrap()).unwrap();
        assert_eq!(changed.matches("DDEEFF").count(), 1);
        assert!(!changed.contains("AABBCC"));
        assert!(changed.contains(r#"<w:tblStylePr w:type="lastRow" x:region="&amp;""#));
        assert_eq!(changed.matches(r#"x:region="&amp;""#).count(), 1);
        assert!(!changed.contains(r#"tblStylePr q:type="firstRow""#));
        assert!(
            changed
                .contains(r#"<ext:conditional xmlns:ext="urn:producer" ext:value="preserved"/>"#)
        );
        let reparsed = CT_Styles::from_xml(changed.as_bytes())
            .expect("typed conditional projection remains valid XML");
        assert_eq!(
            reparsed.styles[0].conditional_table_styles[0].region,
            Some(TableStyleRegion::LastRow)
        );
        assert_eq!(
            reparsed.styles[0].conditional_table_styles[0]
                .cell_properties
                .as_ref()
                .and_then(|properties| properties.shading.as_ref())
                .and_then(|shading| shading.fill.as_deref()),
            Some("DDEEFF")
        );
    }
    #[test]
    fn fx200_defaults_source_preserves_unknown_wrappers_and_qualified_owner_after_font_edit() {
        let foreign =
            r#"<f:rPrDefault xmlns:f="urn:foreign" f:flag='&#x41;'><f:rPr/></f:rPrDefault>"#;
        let opaque = r#"<z:keep xmlns:z="urn:opaque" z:value='&#65;'><z:child/></z:keep>"#;
        let source = format!(
            r#"<q:styles xmlns:q="{W_NS}" xmlns:x="urn:attr"><q:docDefaults x:root='&#65;'><!--before-->{foreign}<q:rPrDefault x:wrapper='keep'>  {opaque}<q:rPr x:property='&#65;'><q:rFonts q:ascii="Prior" q:hAnsi="Other"/><q:b q:val='true'/><!--inside-->{opaque}</q:rPr><!--after--></q:rPrDefault><q:pPrDefault x:unchanged='&#x42;'><q:pPr x:property='&#66;'><q:keepNext/></q:pPr>{opaque}</q:pPrDefault></q:docDefaults></q:styles>"#
        );
        let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
        let defaults = styles.doc_defaults.as_ref().unwrap();
        assert_eq!(
            defaults.rpr.as_ref().unwrap().font_ascii.as_deref(),
            Some("Prior")
        );
        assert_eq!(defaults.rpr.as_ref().unwrap().bold, Some(true));
        let original_ppr = defaults.ppr.clone();
        let original_source = defaults.source_xml.clone().unwrap();
        let mut untouched = Writer::new(Vec::new());
        defaults.to_xml(&mut untouched).unwrap();
        assert_eq!(untouched.into_inner(), original_source);
        let properties = styles.doc_defaults.as_mut().unwrap().rpr.as_mut().unwrap();
        properties.font_ascii = Some("Selected".to_owned());
        properties.font_hansi = Some("Selected".to_owned());
        let saved = styles.to_xml().unwrap();
        let text = std::str::from_utf8(&saved).unwrap();
        assert!(text.contains(foreign) && text.contains(opaque));
        assert!(text.contains("x:root='&#65;'") && text.contains("x:wrapper='keep'"));
        assert!(text.contains("<!--before-->") && text.contains("<!--after-->"));
        assert!(text.contains("x:property='&#65;'"));
        assert!(text.contains("<q:b q:val='true'/><!--inside-->"));
        let original_text = std::str::from_utf8(&original_source).unwrap();
        let begin = original_text.find("<q:pPrDefault").unwrap();
        let end = original_text.find("</q:pPrDefault>").unwrap() + "</q:pPrDefault>".len();
        assert!(text.contains(&original_text[begin..end]));
        let reopened = CT_Styles::from_xml(&saved).unwrap();
        let defaults = reopened.doc_defaults.as_ref().unwrap();
        assert_eq!(
            defaults.rpr.as_ref().unwrap().font_ascii.as_deref(),
            Some("Selected")
        );
        assert_eq!(defaults.rpr.as_ref().unwrap().bold, Some(true));
        assert_eq!(defaults.ppr, original_ppr);
    }

    #[test]
    fn fx200_empty_default_wrapper_and_duplicate_wrapper_last_owner_are_preserved() {
        let source = format!(
            r#"<w:styles xmlns:w="{W_NS}" xmlns:f="urn:foreign"><w:docDefaults f:keep='yes'/></w:styles>"#
        );
        let styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
        let defaults = styles.doc_defaults.as_ref().unwrap();
        assert!(defaults.rpr.is_none() && defaults.ppr.is_none());
        assert!(
            std::str::from_utf8(&styles.to_xml().unwrap())
                .unwrap()
                .contains("f:keep='yes'")
        );
        let source = format!(
            r#"<w:styles xmlns:w="{W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault><w:rPrDefault><!--last--></w:rPrDefault></w:docDefaults></w:styles>"#
        );
        let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
        assert!(styles.doc_defaults.as_ref().unwrap().rpr.is_none());
        styles.doc_defaults.as_mut().unwrap().rpr = Some(CT_RPr {
            font_ascii: Some("Selected".to_owned()),
            ..Default::default()
        });
        let saved = styles.to_xml().unwrap();
        assert!(std::str::from_utf8(&saved).unwrap().contains("<w:b/>"));
        let reopened = CT_Styles::from_xml(&saved).unwrap();
        let properties = reopened.doc_defaults.unwrap().rpr.unwrap();
        assert_eq!(properties.font_ascii.as_deref(), Some("Selected"));
        assert!(properties.bold.is_none());
    }
    #[test]
    fn fx200_same_wrapper_duplicate_default_properties_do_not_resurface_on_font_clear() {
        let source = format!(
            r#"<w:styles xmlns:w="{W_NS}" xmlns:f="urn:foreign"><w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr><w:rPr f:last='&#65;'><w:rFonts w:ascii="Selected"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>"#
        );
        let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
        let properties = styles.doc_defaults.as_mut().unwrap().rpr.as_mut().unwrap();
        assert!(properties.bold.is_none());
        assert_eq!(properties.font_ascii.as_deref(), Some("Selected"));
        properties.font_ascii = None;
        let saved = styles.to_xml().unwrap();
        let text = std::str::from_utf8(&saved).unwrap();
        assert!(text.contains("<w:b/>") && text.contains("f:last='&#65;'"));
        let reopened = CT_Styles::from_xml(&saved).unwrap();
        let properties = reopened.doc_defaults.unwrap().rpr.unwrap();
        assert!(properties.bold.is_none() && properties.font_ascii.is_none());
    }
    #[test]
    fn fx200_duplicate_default_font_children_clear_owned_slots_without_reviving_or_erasing_raw() {
        for replacement in [None, Some("B")] {
            let source = format!(
                r#"<w:styles xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:f="urn:foreign"><w:docDefaults><w:rPrDefault><w:rPr><q:rFonts q:ascii='A' f:earlier='&#65;'><f:opaque/><!--font-child--></q:rFonts><!--between--><w:rFonts w:hint='cs' f:last='&#66;'/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>"#
            );
            let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
            let properties = styles.doc_defaults.as_mut().unwrap().rpr.as_mut().unwrap();
            assert_eq!(properties.font_ascii.as_deref(), Some("A"));
            assert_eq!(properties.font_hint.as_deref(), Some("cs"));
            properties.font_ascii = replacement.map(str::to_owned);
            let saved = styles.to_xml().unwrap();
            let text = std::str::from_utf8(&saved).unwrap();
            for raw in [
                "f:earlier='&#65;'",
                "<f:opaque/><!--font-child-->",
                "<!--between-->",
                "w:hint='cs'",
                "f:last='&#66;'",
            ] {
                assert!(text.contains(raw), "missing {raw}: {text}");
            }
            assert!(!text.contains("q:ascii='A'"));
            let reopened = CT_Styles::from_xml(&saved).unwrap();
            let properties = reopened.doc_defaults.unwrap().rpr.unwrap();
            assert_eq!(properties.font_ascii.as_deref(), replacement);
            assert_eq!(properties.font_hint.as_deref(), Some("cs"));
        }
    }

    #[test]
    fn fx200_duplicate_default_size_children_clear_all_owned_values_preserving_foreign_payload() {
        let source = format!(
            r#"<w:styles xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:f="urn:foreign"><w:docDefaults><w:rPrDefault><w:rPr><q:sz q:val='24' f:first='&#65;'/><!--size--><w:sz w:val='28' f:last='&#66;'><f:opaque/></w:sz><q:szCs q:val='20' f:first='CS'/><w:szCs w:val='32' f:last='CS2'/><f:sz f:val='999'/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>"#
        );
        let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
        let properties = styles.doc_defaults.as_mut().unwrap().rpr.as_mut().unwrap();
        assert_eq!(properties.sz, Some(crate::units::HalfPoint(28)));
        assert_eq!(properties.sz_cs, Some(crate::units::HalfPoint(32)));
        properties.sz = None;
        properties.sz_cs = None;
        let saved = styles.to_xml().unwrap();
        let text = std::str::from_utf8(&saved).unwrap();
        for raw in [
            "f:first='&#65;'",
            "<!--size-->",
            "f:last='&#66;'",
            "<f:opaque/>",
            "f:first='CS'",
            "f:last='CS2'",
            "<f:sz f:val='999'/>",
        ] {
            assert!(text.contains(raw), "missing {raw}: {text}");
        }
        for value in ["q:val='24'", "w:val='28'", "q:val='20'", "w:val='32'"] {
            assert!(!text.contains(value));
        }
        let reopened = CT_Styles::from_xml(&saved).unwrap();
        let properties = reopened.doc_defaults.unwrap().rpr.unwrap();
        assert!(properties.sz.is_none() && properties.sz_cs.is_none());
    }
    #[test]
    fn fx200_style_property_root_attributes_and_untouched_owner_bytes_survive_updates() {
        let source = format!(
            r#"<w:styles xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:f="urn:foreign"><w:style w:type="paragraph" w:styleId="Imported"><w:name w:val="Imported"/><q:pPr f:keep='&#65;'><q:keepNext q:val='true'/><!--ppr--><f:opaque/></q:pPr><q:rPr f:keep='&#66;'><q:rFonts q:ascii='A' q:hint='cs'/><!--rpr--><f:opaque/></q:rPr></w:style></w:styles>"#
        );
        let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
        let original = styles.styles[0]
            .modeled_xml
            .iter()
            .find(|(rank, _)| *rank == 20)
            .unwrap()
            .1
            .clone();
        styles.styles[0].rpr.as_mut().unwrap().font_ascii = Some("B".to_owned());
        let saved = styles.to_xml().unwrap();
        assert!(
            saved
                .windows(original.len())
                .any(|window| window == original)
        );
        let text = std::str::from_utf8(&saved).unwrap();
        for raw in [
            "f:keep='&#65;'",
            "f:keep='&#66;'",
            "q:hint='cs'",
            "<!--rpr-->",
            "<f:opaque/>",
        ] {
            assert!(text.contains(raw), "{text}");
        }
        let reopened = CT_Styles::from_xml(&saved).unwrap();
        assert_eq!(
            reopened.styles[0]
                .rpr
                .as_ref()
                .unwrap()
                .font_ascii
                .as_deref(),
            Some("B")
        );
        assert_eq!(
            reopened.styles[0].ppr.as_ref().unwrap().keep_next,
            Some(true)
        );
    }

    #[test]
    fn fx200_duplicate_style_property_owners_keep_empty_final_owner_and_foreign_peers() {
        let source = format!(
            r#"<w:styles xmlns:w="{W_NS}" xmlns:f="urn:foreign"><w:style w:type="paragraph" w:styleId="Imported"><w:name w:val="Imported"/><w:pPr><w:keepNext/></w:pPr><w:pPr f:last='&#65;'/><w:rPr><w:b/></w:rPr><f:rPr f:keep='peer'/><w:rPr f:last='&#66;'><w:rFonts w:ascii='A'/></w:rPr></w:style></w:styles>"#
        );
        let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
        assert!(styles.styles[0].ppr.as_ref().unwrap().keep_next.is_none());
        assert!(styles.styles[0].rpr.as_ref().unwrap().bold.is_none());
        styles.styles[0].rpr.as_mut().unwrap().font_ascii = None;
        let saved = styles.to_xml().unwrap();
        let text = std::str::from_utf8(&saved).unwrap();
        for raw in [
            "<w:b/>",
            "<w:keepNext/>",
            "f:last='&#65;'",
            "f:last='&#66;'",
            "f:keep='peer'",
        ] {
            assert!(text.contains(raw), "{text}");
        }
        let first = text.find("<w:b/>").unwrap();
        let peer = text.find("f:keep='peer'").unwrap();
        let last = text.find("f:last='&#66;'").unwrap();
        assert!(first < peer && peer < last, "{text}");
        let reopened = CT_Styles::from_xml(&saved).unwrap();
        let style = &reopened.styles[0];
        assert!(
            style.rpr.as_ref().unwrap().bold.is_none()
                && style.rpr.as_ref().unwrap().font_ascii.is_none()
        );
        assert!(style.ppr.as_ref().unwrap().keep_next.is_none());
        styles.styles[0].rpr = None;
        let cleared = styles.to_xml().unwrap();
        let reopened = CT_Styles::from_xml(&cleared).unwrap();
        assert!(reopened.styles[0].rpr.is_none());
        assert!(
            std::str::from_utf8(&cleared)
                .unwrap()
                .contains("f:keep='peer'")
        );
    }
    #[test]
    fn fx200_style_border_container_edit_retains_aliased_root_and_opaque_payload() {
        let source = format!(
            r#"<w:styles xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:f="urn:foreign"><w:style w:type="paragraph" w:styleId="Imported"><q:pPr f:outer='&#65;'><q:pBdr f:keep='&#66;'><q:bottom q:val='single' f:edge='&#67;'/><!--border--><f:opaque>  keep  </f:opaque><f:bottom f:val='foreign'/></q:pBdr><q:keepNext/></q:pPr></w:style></w:styles>"#
        );
        let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
        let edge = styles.styles[0]
            .ppr
            .as_mut()
            .unwrap()
            .borders
            .as_mut()
            .unwrap()
            .bottom
            .as_mut()
            .unwrap();
        edge.val = crate::shared::ST_Border::Double;
        edge.sz = Some(8);
        edge.color = Some("123456".to_owned());
        let saved = styles.to_xml().unwrap();
        let text = std::str::from_utf8(&saved).unwrap();
        for raw in [
            "f:outer='&#65;'",
            "f:keep='&#66;'",
            "f:edge='&#67;'",
            "<!--border-->",
            "<f:opaque>  keep  </f:opaque>",
            "<f:bottom f:val='foreign'/>",
        ] {
            assert!(text.contains(raw), "{text}");
        }
        let reopened = CT_Styles::from_xml(&saved).unwrap();
        let properties = reopened.styles[0].ppr.as_ref().unwrap();
        let border = properties
            .borders
            .as_ref()
            .unwrap()
            .bottom
            .as_ref()
            .unwrap();
        assert_eq!(border.val, crate::shared::ST_Border::Double);
        assert_eq!(border.sz, Some(8));
        assert_eq!(border.color.as_deref(), Some("123456"));
        assert_eq!(properties.keep_next, Some(true));
        assert_eq!(reopened.to_xml().unwrap(), saved);
    }

    #[test]
    fn fx200_duplicate_style_border_clear_keeps_payload_but_suppresses_earlier_edges() {
        let source = format!(
            r#"<w:styles xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:f="urn:foreign"><w:style w:type="paragraph" w:styleId="Imported"><w:pPr><q:pBdr f:first='&#65;'><q:top q:val='single'/><f:opaque first='same'/></q:pBdr><f:pBdr f:keep='foreign'/><q:pBdr f:last='&#66;'><q:bottom q:val='double'/><!--last--><f:opaque last='same'/></q:pBdr><w:keepNext/></w:pPr></w:style></w:styles>"#
        );
        let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
        assert!(
            styles.styles[0]
                .ppr
                .as_ref()
                .unwrap()
                .borders
                .as_ref()
                .unwrap()
                .bottom
                .is_some()
        );
        styles.styles[0].ppr.as_mut().unwrap().borders = None;
        let saved = styles.to_xml().unwrap();
        let text = std::str::from_utf8(&saved).unwrap();
        for raw in [
            "f:first='&#65;'",
            "f:last='&#66;'",
            "<f:opaque first='same'/>",
            "<f:opaque last='same'/>",
            "<!--last-->",
            "<f:pBdr f:keep='foreign'/>",
        ] {
            assert!(text.contains(raw), "{text}");
        }
        assert!(text.find("f:first='&#65;'").unwrap() < text.find("f:keep='foreign'").unwrap());
        assert!(text.find("f:keep='foreign'").unwrap() < text.find("f:last='&#66;'").unwrap());
        let reopened = CT_Styles::from_xml(&saved).unwrap();
        let properties = reopened.styles[0].ppr.as_ref().unwrap();
        let border = properties.borders.as_ref().unwrap();
        assert!(
            border.top.is_none()
                && border.bottom.is_none()
                && border.left.is_none()
                && border.right.is_none()
                && border.between.is_none()
                && border.bar.is_none()
        );
        assert_eq!(properties.keep_next, Some(true));
        assert_eq!(reopened.to_xml().unwrap(), saved);
    }

    #[test]
    fn fx200_direct_style_reader_restores_both_trim_flags_and_opaque_sibling_whitespace() {
        for property in [
            "<w:pPr><w:keepNext/></w:pPr>",
            "<f:pPr><f:payload>  inner  </f:payload></f:pPr>",
            "<w:rPr><w:b/></w:rPr>",
            "<f:rPr><f:payload>  inner  </f:payload></f:rPr>",
        ] {
            for (trim_start, trim_end) in [(false, false), (false, true), (true, false)] {
                let source = format!(
                    r#"<w:style xmlns:w="{W_NS}" xmlns:f="urn:foreign" w:type="paragraph" w:styleId="Direct">{property}<f:opaque>  keep  </f:opaque></w:style>"#
                );
                let mut reader = Reader::from_reader(source.as_bytes());
                reader.config_mut().trim_text_start = trim_start;
                reader.config_mut().trim_text_end = trim_end;
                let Event::Start(start) = reader.read_event().unwrap() else {
                    panic!("style root missing")
                };
                let style = CT_Style::from_xml(&mut reader, &start).unwrap();
                assert_eq!(reader.config().trim_text_start, trim_start);
                assert_eq!(reader.config().trim_text_end, trim_end);
                let expected = match (trim_start, trim_end) {
                    (false, false) => ">  keep  </f:opaque>",
                    (false, true) => ">  keep</f:opaque>",
                    _ => ">keep  </f:opaque>",
                };
                assert!(
                    style
                        .extra_xml
                        .iter()
                        .any(|(_, raw)| std::str::from_utf8(raw).unwrap().contains(expected)),
                    "{style:?}"
                );
            }
            let opening = &property[..property.find('>').unwrap() + 1];
            let malformed = format!(
                r#"<w:style xmlns:w="{W_NS}" xmlns:f="urn:foreign" w:type="paragraph" w:styleId="Direct">{opening}</f:wrong></w:style>"#
            );
            let mut reader = Reader::from_reader(malformed.as_bytes());
            reader.config_mut().trim_text(false);
            let Event::Start(start) = reader.read_event().unwrap() else {
                panic!("style root missing")
            };
            assert!(CT_Style::from_xml(&mut reader, &start).is_err());
            assert!(!reader.config().trim_text_start && !reader.config().trim_text_end);
        }
    }
    #[test]
    fn fx200_started_style_border_edges_match_empty_projection_and_keep_source_on_updates() {
        for prefix in ["w", "q"] {
            let source = format!(
                r#"<w:styles xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:f="urn:foreign"><w:style w:type="paragraph" w:styleId="Started"><{prefix}:pPr><{prefix}:pBdr><{prefix}:bottom {prefix}:val='single' f:edge='&#65;'><f:opaque>  keep  </f:opaque><!--edge--></{prefix}:bottom></{prefix}:pBdr></{prefix}:pPr></w:style></w:styles>"#
            );
            let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
            let original = styles.to_xml().unwrap();
            let before = styles.styles[0]
                .ppr
                .as_ref()
                .unwrap()
                .borders
                .as_ref()
                .unwrap()
                .bottom
                .as_ref()
                .unwrap();
            assert_eq!(before.val, crate::shared::ST_Border::Single);
            assert!(before.source_xml.is_some());
            assert_eq!(
                CT_Styles::from_xml(&original).unwrap().to_xml().unwrap(),
                original
            );
            let edge = styles.styles[0]
                .ppr
                .as_mut()
                .unwrap()
                .borders
                .as_mut()
                .unwrap()
                .bottom
                .as_mut()
                .unwrap();
            edge.val = crate::shared::ST_Border::Double;
            edge.sz = Some(8);
            edge.color = Some("123456".to_owned());
            let saved = styles.to_xml().unwrap();
            let text = std::str::from_utf8(&saved).unwrap();
            for raw in [
                "f:edge='&#65;'",
                "<f:opaque>  keep  </f:opaque>",
                "<!--edge-->",
            ] {
                assert!(text.contains(raw), "{text}");
            }
            assert!(text.contains(&format!("</{prefix}:bottom>")), "{text}");
            let reopened = CT_Styles::from_xml(&saved).unwrap();
            let actual = reopened.styles[0]
                .ppr
                .as_ref()
                .unwrap()
                .borders
                .as_ref()
                .unwrap()
                .bottom
                .as_ref()
                .unwrap();
            assert_eq!(actual.val, crate::shared::ST_Border::Double);
            assert_eq!(actual.sz, Some(8));
            assert_eq!(actual.color.as_deref(), Some("123456"));
            assert_eq!(reopened.to_xml().unwrap(), saved);
            let empty = format!(
                r#"<w:pPr xmlns:w="{W_NS}"><w:pBdr><w:bottom w:val="double" w:sz="8" w:color="123456"/></w:pBdr></w:pPr>"#
            );
            let canonical =
                crate::numbering::parse_scoped_ppr(empty.as_bytes(), &["w".to_owned()]).unwrap();
            let expected = canonical.borders.as_ref().unwrap().bottom.as_ref().unwrap();
            assert_eq!(
                (actual.val, actual.sz, actual.color.as_deref()),
                (expected.val, expected.sz, expected.color.as_deref())
            );
        }
    }
    #[test]
    fn fx200_style_border_opaque_descendants_keep_bytes_and_never_become_direct_edges() {
        for opaque in [
            r#"<f:opaque><q:bottom q:val='double'></q:bottom></f:opaque>"#,
            r#"<f:opaque xmlns:q='urn:shadow'><q:bottom q:val='double'></q:bottom><w:left w:val='double'/></f:opaque>"#,
        ] {
            let source = format!(
                r#"<w:styles xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:f="urn:foreign"><w:style w:type="paragraph" w:styleId="Opaque"><q:pPr><q:pBdr>{opaque}<q:right q:val='single' f:edge='&#65;'><f:inside>  real  </f:inside></q:right></q:pBdr></q:pPr></w:style></w:styles>"#
            );
            let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
            let borders = styles.styles[0]
                .ppr
                .as_ref()
                .unwrap()
                .borders
                .as_ref()
                .unwrap();
            assert!(borders.bottom.is_none() && borders.left.is_none());
            assert_eq!(
                borders.right.as_ref().unwrap().val,
                crate::shared::ST_Border::Single
            );
            let noop = styles.to_xml().unwrap();
            assert!(std::str::from_utf8(&noop).unwrap().contains(opaque));
            styles.styles[0]
                .ppr
                .as_mut()
                .unwrap()
                .borders
                .as_mut()
                .unwrap()
                .right
                .as_mut()
                .unwrap()
                .val = crate::shared::ST_Border::Double;
            let saved = styles.to_xml().unwrap();
            let text = std::str::from_utf8(&saved).unwrap();
            assert!(
                text.contains(opaque)
                    && text.contains("<f:inside>  real  </f:inside>")
                    && text.contains("f:edge='&#65;'")
            );
            let reopened = CT_Styles::from_xml(&saved).unwrap();
            let actual = reopened.styles[0]
                .ppr
                .as_ref()
                .unwrap()
                .borders
                .as_ref()
                .unwrap();
            assert!(actual.bottom.is_none() && actual.left.is_none());
            assert_eq!(
                actual.right.as_ref().unwrap().val,
                crate::shared::ST_Border::Double
            );
            assert_eq!(reopened.to_xml().unwrap(), saved);
        }
    }

    #[test]
    fn fx200_started_logical_style_border_edges_update_once_without_payload_duplication() {
        for (physical, logical) in [("start", "left"), ("end", "right")] {
            let source = format!(
                r#"<w:styles xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:f="urn:foreign"><w:style w:type="paragraph" w:styleId="Logical"><q:pPr><q:pBdr f:container='&#65;'><q:{physical} q:val='single' f:edge='&#66;'><f:payload>  unique  </f:payload><!--unique--></q:{physical}></q:pBdr></q:pPr></w:style></w:styles>"#
            );
            let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
            let initial = styles.to_xml().unwrap();
            assert_eq!(
                CT_Styles::from_xml(&initial).unwrap().to_xml().unwrap(),
                initial
            );
            for (value, size, color) in [
                (crate::shared::ST_Border::Double, 8, "123456"),
                (crate::shared::ST_Border::Single, 12, "654321"),
            ] {
                let borders = styles.styles[0]
                    .ppr
                    .as_mut()
                    .unwrap()
                    .borders
                    .as_mut()
                    .unwrap();
                let edge = if logical == "left" {
                    borders.left.as_mut().unwrap()
                } else {
                    borders.right.as_mut().unwrap()
                };
                edge.val = value;
                edge.sz = Some(size);
                edge.color = Some(color.to_owned());
                let saved = styles.to_xml().unwrap();
                let text = std::str::from_utf8(&saved).unwrap();
                assert_eq!(
                    text.matches(&format!("<q:{physical} ")).count(),
                    1,
                    "{text}"
                );
                assert_eq!(
                    text.matches(&format!("</q:{physical}>")).count(),
                    1,
                    "{text}"
                );
                assert_eq!(
                    text.matches("<f:payload>  unique  </f:payload>").count(),
                    1,
                    "{text}"
                );
                assert_eq!(text.matches("<!--unique-->").count(), 1, "{text}");
                assert!(text.contains("f:edge='&#66;'") && text.contains("f:container='&#65;'"));
                styles = CT_Styles::from_xml(&saved).unwrap();
                let borders = styles.styles[0]
                    .ppr
                    .as_ref()
                    .unwrap()
                    .borders
                    .as_ref()
                    .unwrap();
                let actual = if logical == "left" {
                    borders.left.as_ref().unwrap()
                } else {
                    borders.right.as_ref().unwrap()
                };
                assert_eq!(
                    (actual.val, actual.sz, actual.color.as_deref()),
                    (value, Some(size), Some(color))
                );
                assert_eq!(styles.to_xml().unwrap(), saved);
            }
        }
    }

    #[test]
    fn fx200_style_partial_border_clear_removes_all_aliases_without_touching_siblings() {
        let source = format!(
            r#"<w:styles xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:f="urn:f"><w:style w:type="paragraph" w:styleId="Partial"><q:pPr><q:pBdr f:root='&#65;'><q:start q:val='single' f:gone='yes'><f:gone>owned</f:gone></q:start><q:left q:val='double'/><!--keep--><q:end q:val='single' f:stay='&#66;'><f:stay>  body  </f:stay></q:end><f:left f:val='foreign'/></q:pBdr></q:pPr></w:style></w:styles>"#
        );
        let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
        styles.styles[0]
            .ppr
            .as_mut()
            .unwrap()
            .borders
            .as_mut()
            .unwrap()
            .left = None;
        let saved = styles.to_xml().unwrap();
        let text = std::str::from_utf8(&saved).unwrap();
        assert!(!text.contains("<q:start") && !text.contains("<q:left"));
        for exact in [
            "f:root='&#65;'",
            "<!--keep-->",
            "f:stay='&#66;'",
            "<f:stay>  body  </f:stay>",
            "<f:left f:val='foreign'/>",
        ] {
            assert!(text.contains(exact), "{text}");
        }
        let reopened = CT_Styles::from_xml(&saved).unwrap();
        let borders = reopened.styles[0]
            .ppr
            .as_ref()
            .unwrap()
            .borders
            .as_ref()
            .unwrap();
        assert!(borders.left.is_none());
        assert_eq!(
            borders.right.as_ref().unwrap().val,
            crate::shared::ST_Border::Single
        );
        assert_eq!(reopened.to_xml().unwrap(), saved);
    }
    #[test]
    fn fx200_changed_style_scalar_keeps_physical_foreign_attributes_and_started_payload() {
        for started in [false, true] {
            let leaf = if started {
                "<q:u q:val='single' f:keep='&#65;'><f:opaque>  keep  </f:opaque><!--inside--></q:u>"
            } else {
                "<q:u q:val='single' f:keep='&#65;'/>"
            };
            let source = format!(
                r#"<w:styles xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:f="urn:foreign"><w:style w:type="paragraph" w:styleId="Scalar"><q:rPr f:root='&#66;'>{leaf}<q:b/></q:rPr></w:style></w:styles>"#
            );
            let mut styles = CT_Styles::from_xml(source.as_bytes()).unwrap();
            styles.styles[0].rpr.as_mut().unwrap().underline =
                Some(crate::shared::ST_Underline::Double);
            let saved = styles.to_xml().unwrap();
            let text = std::str::from_utf8(&saved).unwrap();
            assert!(
                text.contains("f:keep='&#65;'")
                    && text.contains("f:root='&#66;'")
                    && text.contains("<q:b/>"),
                "{text}"
            );
            if started {
                assert!(
                    text.contains("<f:opaque>  keep  </f:opaque><!--inside-->"),
                    "{text}"
                );
            }
            let reopened = CT_Styles::from_xml(&saved).unwrap();
            assert_eq!(
                reopened.styles[0].rpr.as_ref().unwrap().underline,
                Some(crate::shared::ST_Underline::Double)
            );
            assert_eq!(reopened.to_xml().unwrap(), saved);
        }
    }
}
