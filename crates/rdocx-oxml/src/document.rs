//! Document-level elements: `CT_Document` and `CT_Body`.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::borders::CT_BorderEdge;
use crate::content_control::{CT_Sdt, SdtOwner};
use crate::error::{OxmlError, Result};
use crate::header_footer::{HdrFtrRef, HdrFtrType};
use crate::namespace::{W_NS, matches_local_name};
use crate::numbering::{
    local_namespace_overrides, merged_owner_bindings, namespace_bindings, word_prefixes_at,
};
use crate::properties::{
    get_word_val_attr, is_word_attribute, is_word_element, parse_integer_measurement,
};
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::revision::{CT_Revision, RevisionKind};
use crate::shared::{ST_PageOrientation, ST_SectionType};
use crate::table::{CT_Tbl, ST_VerticalJc};
use crate::text::{
    CT_P, ROOT_MC_BINDING, ROOT_R_BINDING, ROOT_WP_BINDING, capture_root_attribute_record,
    declare_w14_on_part_root, is_root_attribute_record, push_root_attribute_record,
    root_binding_scope,
};
use crate::units::Twips;

/// Content that can appear in a document body (paragraphs and tables).
#[derive(Debug, Clone, PartialEq)]
pub enum BodyContent {
    Paragraph(CT_P),
    Table(CT_Tbl),
    ContentControl(CT_Sdt),
    /// Raw XML for unknown elements and controls that cannot be typed safely.
    RawXml(Vec<u8>),
}

/// Column definition for multi-column layouts.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Column {
    /// Column width in twips
    pub width: Option<Twips>,
    /// Space after this column in twips
    pub space: Option<Twips>,
}

/// `CT_Columns` — Column layout configuration.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Columns {
    /// Number of columns (if equal width)
    pub num: Option<u32>,
    /// Space between columns in twips (when equal width)
    pub space: Option<Twips>,
    /// Whether columns are equal width
    pub equal_width: Option<bool>,
    /// Separator line between columns
    pub sep: Option<bool>,
    /// Individual column definitions (for unequal widths)
    pub columns: Vec<CT_Column>,
}

/// `CT_PageNumberType` -- the M23 page-number restart with retained M24 state.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_PageNumberType {
    /// Displayed page number at which this section starts.
    pub start: Option<u32>,
    /// Original `w:pgNumType` subtree, including unsupported attributes.
    #[doc(hidden)]
    pub raw_xml: Option<Vec<u8>>,
    /// Parsed start value used to detect an authored change.
    #[doc(hidden)]
    pub parsed_start: Option<u32>,
    /// Qualified name of the retained start attribute.
    #[doc(hidden)]
    pub start_attribute_name: Option<Vec<u8>>,
    /// In-scope Word prefix preferred when adding a missing start attribute.
    #[doc(hidden)]
    pub start_insertion_prefix: Option<Vec<u8>>,
    /// Every inherited or local prefix reserved by the retained source scope.
    #[doc(hidden)]
    pub reserved_insertion_prefixes: Vec<Vec<u8>>,
}

impl CT_PageNumberType {
    /// Create an authored page-number restart.
    pub fn new(start: u32) -> Self {
        Self {
            start: Some(start),
            raw_xml: None,
            parsed_start: None,
            start_attribute_name: None,
            start_insertion_prefix: None,
            reserved_insertion_prefixes: Vec::new(),
        }
    }
}

impl Default for CT_Columns {
    fn default() -> Self {
        CT_Columns {
            num: Some(1),
            space: Some(Twips(720)),
            equal_width: Some(true),
            sep: None,
            columns: Vec::new(),
        }
    }
}

/// `ST_PageBorderZOrder` -- whether a page border draws in front of content.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ST_PageBorderZOrder {
    Front,
    Back,
}

impl ST_PageBorderZOrder {
    pub fn from_str(value: &str) -> Option<Self> {
        match value {
            "front" => Some(Self::Front),
            "back" => Some(Self::Back),
            _ => None,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Front => "front",
            Self::Back => "back",
        }
    }
}

/// `ST_PageBorderDisplay` -- which pages of a section carry the page border.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ST_PageBorderDisplay {
    AllPages,
    FirstPage,
    NotFirstPage,
}

impl ST_PageBorderDisplay {
    pub fn from_str(value: &str) -> Option<Self> {
        match value {
            "allPages" => Some(Self::AllPages),
            "firstPage" => Some(Self::FirstPage),
            "notFirstPage" => Some(Self::NotFirstPage),
            _ => None,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::AllPages => "allPages",
            Self::FirstPage => "firstPage",
            Self::NotFirstPage => "notFirstPage",
        }
    }
}

/// `ST_PageBorderOffset` -- what the page-border edge offsets are measured from.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ST_PageBorderOffset {
    Page,
    Text,
}

impl ST_PageBorderOffset {
    pub fn from_str(value: &str) -> Option<Self> {
        match value {
            "page" => Some(Self::Page),
            "text" => Some(Self::Text),
            _ => None,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Page => "page",
            Self::Text => "text",
        }
    }
}

/// `CT_PageBorders` -- the `w:pgBorders` frame drawn around a section's pages.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_PageBorders {
    /// `w:zOrder`, whether the frame draws in front of or behind page content.
    pub z_order: Option<ST_PageBorderZOrder>,
    /// `w:display`, which physical pages of the section receive the frame.
    pub display: Option<ST_PageBorderDisplay>,
    /// `w:offsetFrom`, the rectangle each edge offset is measured from.
    pub offset_from: Option<ST_PageBorderOffset>,
    pub top: Option<CT_BorderEdge>,
    pub left: Option<CT_BorderEdge>,
    pub bottom: Option<CT_BorderEdge>,
    pub right: Option<CT_BorderEdge>,
    /// Attributes this type does not model, in source order, written first.
    ///
    /// Values are serialized verbatim, so a caller storing one owns the
    /// escaping, exactly as `CT_BorderEdge::extra_attributes` does.
    pub extra_attributes: Vec<(String, String)>,
    /// Child elements this type does not model, in source order.
    pub extra_xml: Vec<Vec<u8>>,
}

/// `ST_LineNumberRestart` -- where a section's line numbering starts over.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ST_LineNumberRestart {
    NewPage,
    NewSection,
    Continuous,
}

impl ST_LineNumberRestart {
    pub fn from_str(value: &str) -> Option<Self> {
        match value {
            "newPage" => Some(Self::NewPage),
            "newSection" => Some(Self::NewSection),
            "continuous" => Some(Self::Continuous),
            _ => None,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::NewPage => "newPage",
            Self::NewSection => "newSection",
            Self::Continuous => "continuous",
        }
    }
}

/// `CT_LineNumber` -- the `w:lnNumType` margin line numbering of a section.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_LineNumber {
    /// `w:countBy`, the interval at which a number is printed.
    pub count_by: Option<u32>,
    /// `w:start`, the first line number of the section.
    pub start: Option<u32>,
    /// `w:distance`, the gap between the number and the text it labels.
    pub distance: Option<Twips>,
    /// `w:restart`, where numbering starts over.
    pub restart: Option<ST_LineNumberRestart>,
    /// Attributes this type does not model, in source order, written first.
    pub extra_attributes: Vec<(String, String)>,
}

/// `CT_PaperSource` -- the `w:paperSrc` printer trays a section prints from.
///
/// Tray selection happens in the printer driver, so this has no on-page
/// consequence and pagination is identical with and without it.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_PaperSource {
    /// `w:first`, the tray feeding the section's first page.
    pub first: Option<u32>,
    /// `w:other`, the tray feeding every later page.
    pub other: Option<u32>,
    /// Attributes this type does not model, in source order, written first.
    pub extra_attributes: Vec<(String, String)>,
}

/// `ST_DocGrid` -- how a section's character grid constrains its text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub enum ST_DocGrid {
    /// No grid. Lines and characters keep their ordinary metrics.
    Default,
    /// Line advance snaps to the grid pitch. Characters keep their advance.
    Lines,
    /// Line advance snaps to the pitch and characters snap to the char space.
    LinesAndChars,
    /// Characters snap to the char space, and line advance snaps too.
    SnapToChars,
}

impl ST_DocGrid {
    pub fn from_str(value: &str) -> Option<Self> {
        match value {
            "default" => Some(Self::Default),
            "lines" => Some(Self::Lines),
            "linesAndChars" => Some(Self::LinesAndChars),
            "snapToChars" => Some(Self::SnapToChars),
            _ => None,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::Lines => "lines",
            Self::LinesAndChars => "linesAndChars",
            Self::SnapToChars => "snapToChars",
        }
    }

    /// Whether this grid type puts line advance on the grid pitch.
    ///
    /// `Default` is deliberately false. Keeping the default type off the grid
    /// path is what leaves every ungridded document on the arithmetic it
    /// already had, with no new floating-point step on the existing branch.
    pub fn snaps_lines(self) -> bool {
        matches!(self, Self::Lines | Self::LinesAndChars | Self::SnapToChars)
    }

    /// Whether this grid type puts per-character advance on the grid.
    pub fn snaps_characters(self) -> bool {
        matches!(self, Self::LinesAndChars | Self::SnapToChars)
    }
}

/// `CT_DocGrid` -- the `w:docGrid` character grid of a section.
///
/// `line_pitch` is the twip distance between two grid lines. `char_space` is
/// the twip addition to one East Asian character cell.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub struct CT_DocGrid {
    /// `w:type`, which of the grid constraints apply.
    pub grid_type: Option<ST_DocGrid>,
    /// `w:linePitch`, the grid line spacing in twips.
    pub line_pitch: Option<Twips>,
    /// `w:charSpace`, the grid character advance addition in twips.
    pub char_space: Option<i32>,
    /// Attributes this type does not model, in source order, written first.
    pub extra_attributes: Vec<(String, String)>,
}

/// Section-level note configuration, shared by `w:footnotePr` and `w:endnotePr`.
///
/// F-269 models the values and the authoring surface. Their effect on marker
/// text, placement and restart belongs to F-274.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CT_NoteProperties {
    /// `w:pos/@w:val`, where the notes are placed.
    pub pos: Option<String>,
    /// `w:numFmt/@w:val`, the marker number format.
    pub num_fmt: Option<String>,
    /// `w:numStart/@w:val`, the first marker number.
    pub num_start: Option<u32>,
    /// `w:numRestart/@w:val`, where marker numbering starts over.
    pub num_restart: Option<String>,
    /// IDs of special note records selected by document settings.
    pub special_references: Vec<i32>,
    /// Unmodelled children retained at their schema insertion slots.
    ///
    /// A modelled child carrying an attribute this type does not read is kept
    /// here whole rather than typed, which is the choice `w:pgSz` already makes.
    #[doc(hidden)]
    pub extra_xml: Vec<(usize, Vec<u8>)>,
}

/// Value and occurrence anchor for one repeated section reference.
///
/// Distinguishable values follow their source predecessor or successor. Equal
/// duplicates are indistinguishable through the public `Vec` mutation surface,
/// so they resolve deterministically by their ordinal among equal current
/// values. The following anchor takes precedence when both neighbors resolve.
#[doc(hidden)]
#[derive(Debug, Clone, PartialEq)]
pub struct CT_SectPrReferenceAnchor {
    reference: HdrFtrRef,
    source_occurrence: usize,
}

impl CT_SectPrReferenceAnchor {
    fn at(references: &[HdrFtrRef], index: usize) -> Option<Self> {
        let reference = references.get(index)?.clone();
        let source_occurrence = references[..index]
            .iter()
            .filter(|candidate| **candidate == reference)
            .count();
        Some(Self {
            reference,
            source_occurrence,
        })
    }

    fn resolve(&self, references: &[HdrFtrRef]) -> Option<usize> {
        references
            .iter()
            .enumerate()
            .filter(|(_, reference)| **reference == self.reference)
            .nth(self.source_occurrence)
            .map(|(index, _)| index)
    }
}

/// Retained section-child position, including repeated-child occurrence anchors.
#[doc(hidden)]
#[derive(Debug, Clone, PartialEq)]
pub enum CT_SectPrRawPosition {
    /// A schema slot and modeled occurrence boundary.
    Schema { slot: usize, occurrence: usize },
    /// A boundary tied to neighboring header-reference values and occurrences.
    Header {
        preceding: Option<CT_SectPrReferenceAnchor>,
        following: Option<CT_SectPrReferenceAnchor>,
    },
    /// A boundary tied to neighboring footer-reference values and occurrences.
    Footer {
        preceding: Option<CT_SectPrReferenceAnchor>,
        following: Option<CT_SectPrReferenceAnchor>,
    },
}

/// `CT_SectPr` — Section properties (page size, margins, columns, orientation).
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_SectPr {
    /// Page width in twips
    pub page_width: Option<Twips>,
    /// Page height in twips
    pub page_height: Option<Twips>,
    /// Page orientation
    pub orientation: Option<ST_PageOrientation>,
    /// Top margin in twips
    pub margin_top: Option<Twips>,
    /// Right margin in twips
    pub margin_right: Option<Twips>,
    /// Bottom margin in twips
    pub margin_bottom: Option<Twips>,
    /// Left margin in twips
    pub margin_left: Option<Twips>,
    /// Gutter margin in twips
    pub gutter: Option<Twips>,
    /// Header distance from top edge in twips
    pub header_distance: Option<Twips>,
    /// Footer distance from bottom edge in twips
    pub footer_distance: Option<Twips>,
    /// Section break type
    pub section_type: Option<ST_SectionType>,
    /// Column layout
    pub columns: Option<CT_Columns>,
    /// Page-number restart. Unsupported format and chapter attributes remain raw.
    pub page_number: Option<CT_PageNumberType>,
    /// Section footnote configuration. F-274 owns its effect on note rendering.
    ///
    /// Boxed, with the other large members, so a section stays cheap on the
    /// stack. `CT_SectPr` sits inside `CT_PPr`, which test threads build by
    /// value against a 2 MiB ceiling.
    pub footnote_pr: Option<Box<CT_NoteProperties>>,
    /// Section endnote configuration. F-274 owns its effect on note rendering.
    pub endnote_pr: Option<Box<CT_NoteProperties>>,
    /// Printer tray selection. Round-tripped, with no effect on page geometry.
    pub paper_source: Option<Box<CT_PaperSource>>,
    /// Page border frame.
    pub page_borders: Option<Box<CT_PageBorders>>,
    /// Margin line numbering.
    pub line_numbers: Option<Box<CT_LineNumber>>,
    /// Vertical alignment of the body band within the page.
    pub vertical_alignment: Option<ST_VerticalJc>,
    /// `w:textDirection`. Authored and preserved here, projected by F-266c.
    pub text_direction: Option<String>,
    /// `w:docGrid`, the section character grid.
    ///
    /// Boxed with the other composite members so a section stays cheap on the
    /// stack. `CT_SectPr` sits inside `CT_PPr`, which test threads build by
    /// value against a 2 MiB ceiling.
    pub doc_grid: Option<Box<CT_DocGrid>>,
    /// Title page (different first page header/footer)
    pub title_pg: Option<bool>,
    /// Header references
    pub header_refs: Vec<HdrFtrRef>,
    /// Footer references
    pub footer_refs: Vec<HdrFtrRef>,
    /// Unknown child elements captured as raw XML.
    pub extra_xml: Vec<Vec<u8>>,
    /// Schema slots and modeled-child boundaries for retained child elements.
    #[doc(hidden)]
    pub extra_xml_positions: Vec<CT_SectPrRawPosition>,
    /// Prior section properties from the schema-final `w:sectPrChange`.
    pub change: Option<CT_Revision>,
}

#[allow(non_snake_case)]
impl CT_SectPr {
    fn empty() -> Self {
        Self {
            page_width: None,
            page_height: None,
            orientation: None,
            margin_top: None,
            margin_right: None,
            margin_bottom: None,
            margin_left: None,
            gutter: None,
            header_distance: None,
            footer_distance: None,
            section_type: None,
            columns: None,
            page_number: None,
            footnote_pr: None,
            endnote_pr: None,
            paper_source: None,
            page_borders: None,
            line_numbers: None,
            vertical_alignment: None,
            text_direction: None,
            doc_grid: None,
            title_pg: None,
            header_refs: Vec::new(),
            footer_refs: Vec::new(),
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            change: None,
        }
    }

    /// Report whether one raw section carrier retains root attributes.
    #[doc(hidden)]
    pub fn raw_position_is_root_attributes(position: &CT_SectPrRawPosition, raw: &[u8]) -> bool {
        matches!(
            position,
            CT_SectPrRawPosition::Schema {
                slot: usize::MAX,
                occurrence: usize::MAX
            }
        ) && is_root_attribute_record(raw)
    }

    fn from_empty_root(root: &BytesStart<'_>, word_prefixes: &[String]) -> Result<Self> {
        let mut section = Self::empty();
        if let Some(record) = capture_root_attribute_record(root, word_prefixes)? {
            section.extra_xml.push(record);
            section
                .extra_xml_positions
                .push(CT_SectPrRawPosition::Schema {
                    slot: usize::MAX,
                    occurrence: usize::MAX,
                });
        }
        Ok(section)
    }

    /// Default US Letter page with 1-inch margins.
    pub fn default_letter() -> Self {
        CT_SectPr {
            page_width: Some(Twips(12240)),  // 8.5"
            page_height: Some(Twips(15840)), // 11"
            orientation: Some(ST_PageOrientation::Portrait),
            margin_top: Some(Twips(1440)),    // 1"
            margin_right: Some(Twips(1440)),  // 1"
            margin_bottom: Some(Twips(1440)), // 1"
            margin_left: Some(Twips(1440)),   // 1"
            gutter: Some(Twips(0)),
            header_distance: Some(Twips(720)),
            footer_distance: Some(Twips(720)),
            section_type: None,
            columns: None,
            page_number: None,
            footnote_pr: None,
            endnote_pr: None,
            paper_source: None,
            page_borders: None,
            line_numbers: None,
            vertical_alignment: None,
            text_direction: None,
            doc_grid: None,
            title_pg: None,
            header_refs: Vec::new(),
            footer_refs: Vec::new(),
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            change: None,
        }
    }

    /// Default A4 page with 1-inch margins.
    pub fn default_a4() -> Self {
        CT_SectPr {
            page_width: Some(Twips(11906)),  // 210mm
            page_height: Some(Twips(16838)), // 297mm
            orientation: Some(ST_PageOrientation::Portrait),
            margin_top: Some(Twips(1440)),
            margin_right: Some(Twips(1440)),
            margin_bottom: Some(Twips(1440)),
            margin_left: Some(Twips(1440)),
            gutter: Some(Twips(0)),
            header_distance: Some(Twips(720)),
            footer_distance: Some(Twips(720)),
            section_type: None,
            columns: None,
            page_number: None,
            footnote_pr: None,
            endnote_pr: None,
            paper_source: None,
            page_borders: None,
            line_numbers: None,
            vertical_alignment: None,
            text_direction: None,
            doc_grid: None,
            title_pg: None,
            header_refs: Vec::new(),
            footer_refs: Vec::new(),
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            change: None,
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, &["w".to_owned()])
    }

    pub(crate) fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        Self::from_xml_with_prefixes_and_owner_bindings(reader, word_prefixes, &[])
    }

    pub(crate) fn from_xml_with_prefixes_and_owner_bindings(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
    ) -> Result<Self> {
        Self::from_xml_with_prefixes_owner_bindings_and_root(
            reader,
            word_prefixes,
            owner_bindings,
            None,
        )
    }

    pub(crate) fn from_xml_with_prefixes_owner_bindings_and_root(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
        root: Option<&BytesStart<'_>>,
    ) -> Result<Self> {
        let mut sect = Self::empty();
        let mut change_raw_index = 0usize;
        let mut raw_position = (0usize, 0usize);
        let mut header_count = 0usize;
        let mut footer_count = 0usize;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"pgSz", &prefixes) {
                        Self::parse_page_size_attributes(&mut sect, e, &prefixes)?;
                        raw_position = (4, 0);
                    } else if is_word_element(name.as_ref(), b"pgMar", &prefixes) {
                        Self::parse_page_margin_attributes(&mut sect, e, &prefixes)?;
                        raw_position = (5, 0);
                    } else if is_word_element(name.as_ref(), b"type", &prefixes) {
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            sect.section_type = ST_SectionType::from_str(&val).ok();
                        }
                        raw_position = (3, 0);
                    } else if is_word_element(name.as_ref(), b"pgNumType", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        if sect.page_number.is_none() {
                            sect.page_number =
                                Some(Self::parse_page_number(e, &prefixes, owner_bindings, raw)?);
                        } else {
                            sect.push_extra_xml(raw, 6, 1);
                        }
                        raw_position = (6, 1);
                    } else if is_word_element(name.as_ref(), b"cols", &prefixes) {
                        sect.columns = Some(Self::parse_cols_empty(e, &prefixes)?);
                        raw_position = (7, 0);
                    } else if is_word_element(name.as_ref(), b"footnotePr", &prefixes) {
                        if sect.footnote_pr.is_none() && e.attributes().next().is_none() {
                            sect.footnote_pr = Some(Box::default());
                        } else {
                            let raw = crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?;
                            sect.push_extra_xml(raw, 2, 0);
                        }
                        raw_position = (2, 0);
                    } else if is_word_element(name.as_ref(), b"endnotePr", &prefixes) {
                        if sect.endnote_pr.is_none() && e.attributes().next().is_none() {
                            sect.endnote_pr = Some(Box::default());
                        } else {
                            let raw = crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?;
                            sect.push_extra_xml(raw, 2, 1);
                        }
                        raw_position = (2, 1);
                    } else if is_word_element(name.as_ref(), b"paperSrc", &prefixes) {
                        if sect.paper_source.is_none() {
                            sect.paper_source =
                                Some(Box::new(Self::parse_paper_source(e, &prefixes)?));
                        } else {
                            let raw = crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?;
                            sect.push_extra_xml(raw, 5, 1);
                        }
                        raw_position = (5, 1);
                    } else if is_word_element(name.as_ref(), b"pgBorders", &prefixes) {
                        if sect.page_borders.is_none() {
                            sect.page_borders =
                                Some(Box::new(Self::parse_page_border_attrs(e, &prefixes)?));
                        } else {
                            let raw = crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?;
                            sect.push_extra_xml(raw, 5, 2);
                        }
                        raw_position = (5, 2);
                    } else if is_word_element(name.as_ref(), b"lnNumType", &prefixes) {
                        if sect.line_numbers.is_none() {
                            sect.line_numbers =
                                Some(Box::new(Self::parse_line_number(e, &prefixes)?));
                        } else {
                            let raw = crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?;
                            sect.push_extra_xml(raw, 5, 3);
                        }
                        raw_position = (5, 3);
                    } else if is_word_element(name.as_ref(), b"vAlign", &prefixes) {
                        let typed = sole_val_attribute(e, &prefixes)?
                            .filter(|value| ST_VerticalJc::from_str(value).to_str() == value)
                            .map(|value| ST_VerticalJc::from_str(&value));
                        match typed {
                            Some(value) if sect.vertical_alignment.is_none() => {
                                sect.vertical_alignment = Some(value);
                            }
                            _ => {
                                let raw = crate::text::raw_with_external_bindings(
                                    &capture_empty_element(e)?,
                                    owner_bindings,
                                )?;
                                sect.push_extra_xml(raw, 7, 0);
                            }
                        }
                        raw_position = (7, 1);
                    } else if is_word_element(name.as_ref(), b"textDirection", &prefixes) {
                        match sole_val_attribute(e, &prefixes)? {
                            Some(value) if sect.text_direction.is_none() => {
                                sect.text_direction = Some(value);
                            }
                            _ => {
                                let raw = crate::text::raw_with_external_bindings(
                                    &capture_empty_element(e)?,
                                    owner_bindings,
                                )?;
                                sect.push_extra_xml(raw, 8, 0);
                            }
                        }
                        raw_position = (8, 1);
                    } else if is_word_element(name.as_ref(), b"docGrid", &prefixes) {
                        if sect.doc_grid.is_none() {
                            sect.doc_grid = Some(Box::new(Self::parse_doc_grid(e, &prefixes)?));
                        } else {
                            // The schema allows one. A second is retained
                            // after the typed one rather than before it, the
                            // way a second `w:paperSrc` already is.
                            let raw = crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?;
                            sect.push_extra_xml(raw, 8, 2);
                        }
                        raw_position = (8, 2);
                    } else if is_word_element(name.as_ref(), b"headerReference", &prefixes) {
                        let mut hdr_type = HdrFtrType::Default;
                        let mut rel_id = String::new();
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if is_word_attribute(key, b"type", &prefixes) {
                                hdr_type = HdrFtrType::from_str(val);
                            } else if attribute_in_namespace(
                                key,
                                b"id",
                                crate::namespace::R_NS,
                                &prefixes,
                            ) {
                                rel_id = val.to_string();
                            }
                        }
                        if !rel_id.is_empty() {
                            sect.header_refs.push(HdrFtrRef {
                                hdr_ftr_type: hdr_type,
                                rel_id,
                            });
                            header_count += 1;
                        }
                        raw_position = (0, header_count);
                    } else if is_word_element(name.as_ref(), b"footerReference", &prefixes) {
                        let mut ftr_type = HdrFtrType::Default;
                        let mut rel_id = String::new();
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if is_word_attribute(key, b"type", &prefixes) {
                                ftr_type = HdrFtrType::from_str(val);
                            } else if attribute_in_namespace(
                                key,
                                b"id",
                                crate::namespace::R_NS,
                                &prefixes,
                            ) {
                                rel_id = val.to_string();
                            }
                        }
                        if !rel_id.is_empty() {
                            sect.footer_refs.push(HdrFtrRef {
                                hdr_ftr_type: ftr_type,
                                rel_id,
                            });
                            footer_count += 1;
                        }
                        raw_position = (1, footer_count);
                    } else if is_word_element(name.as_ref(), b"titlePg", &prefixes) {
                        sect.title_pg = Some(
                            get_word_val_attr(e, &prefixes)?
                                .is_none_or(|value| value != "0" && value != "false"),
                        );
                        raw_position = (8, 0);
                    } else if is_word_element(name.as_ref(), b"sectPrChange", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            if let Some(previous) = sect.change.replace(revision) {
                                sect.extra_xml
                                    .insert(change_raw_index, previous.into_raw_xml());
                                sect.extra_xml_positions.insert(
                                    change_raw_index,
                                    CT_SectPrRawPosition::Schema {
                                        slot: 9,
                                        occurrence: 0,
                                    },
                                );
                            }
                            change_raw_index = sect.extra_xml.len();
                        } else {
                            sect.push_extra_xml(raw, 9, 0);
                        }
                        raw_position = (9, 0);
                    } else {
                        // Capture unknown empty elements
                        let position = Self::raw_child_schema_slot(name.as_ref(), &prefixes)
                            .unwrap_or(raw_position);
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        sect.push_extra_xml(raw, position.0, position.1);
                        raw_position = position;
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"pgSz", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if raw_element_has_child_content(&raw)?
                            || !page_size_attributes_are_modeled(e, &prefixes)?
                        {
                            sect.push_extra_xml(raw, 4, 0);
                        } else {
                            Self::parse_page_size_attributes(&mut sect, e, &prefixes)?;
                        }
                        raw_position = (4, 0);
                    } else if is_word_element(name.as_ref(), b"pgMar", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if raw_element_has_child_content(&raw)?
                            || !page_margin_attributes_are_modeled(e, &prefixes)?
                        {
                            sect.push_extra_xml(raw, 5, 0);
                        } else {
                            Self::parse_page_margin_attributes(&mut sect, e, &prefixes)?;
                        }
                        raw_position = (5, 0);
                    } else if is_word_element(name.as_ref(), b"headerReference", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if raw_element_has_child_content(&raw)?
                            || !story_reference_attributes_are_modeled(e, &prefixes)?
                        {
                            sect.push_extra_xml(raw, 0, header_count);
                        } else if let Some(reference) = Self::parse_story_reference(e, &prefixes)? {
                            sect.header_refs.push(reference);
                            header_count += 1;
                        }
                        raw_position = (0, header_count);
                    } else if is_word_element(name.as_ref(), b"footerReference", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if raw_element_has_child_content(&raw)?
                            || !story_reference_attributes_are_modeled(e, &prefixes)?
                        {
                            sect.push_extra_xml(raw, 1, footer_count);
                        } else if let Some(reference) = Self::parse_story_reference(e, &prefixes)? {
                            sect.footer_refs.push(reference);
                            footer_count += 1;
                        }
                        raw_position = (1, footer_count);
                    } else if is_word_element(name.as_ref(), b"cols", &prefixes) {
                        sect.columns = Some(Self::parse_cols_start(reader, e, &prefixes)?);
                        raw_position = (7, 0);
                    } else if is_word_element(name.as_ref(), b"footnotePr", &prefixes) {
                        if sect.footnote_pr.is_none() && e.attributes().next().is_none() {
                            sect.footnote_pr = Some(Box::new(Self::parse_note_properties(
                                reader,
                                &prefixes,
                                owner_bindings,
                                b"footnotePr",
                            )?));
                        } else {
                            let raw = crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                owner_bindings,
                            )?;
                            sect.push_extra_xml(raw, 2, 0);
                        }
                        raw_position = (2, 0);
                    } else if is_word_element(name.as_ref(), b"endnotePr", &prefixes) {
                        if sect.endnote_pr.is_none() && e.attributes().next().is_none() {
                            sect.endnote_pr = Some(Box::new(Self::parse_note_properties(
                                reader,
                                &prefixes,
                                owner_bindings,
                                b"endnotePr",
                            )?));
                        } else {
                            let raw = crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                owner_bindings,
                            )?;
                            sect.push_extra_xml(raw, 2, 1);
                        }
                        raw_position = (2, 1);
                    } else if is_word_element(name.as_ref(), b"pgBorders", &prefixes) {
                        if sect.page_borders.is_none() {
                            sect.page_borders = Some(Box::new(Self::parse_page_borders(
                                reader,
                                e,
                                &prefixes,
                                owner_bindings,
                            )?));
                        } else {
                            let raw = crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                owner_bindings,
                            )?;
                            sect.push_extra_xml(raw, 5, 2);
                        }
                        raw_position = (5, 2);
                    } else if is_word_element(name.as_ref(), b"paperSrc", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        sect.push_extra_xml(raw, 5, 0);
                        raw_position = (5, 1);
                    } else if is_word_element(name.as_ref(), b"lnNumType", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        sect.push_extra_xml(raw, 5, 2);
                        raw_position = (5, 3);
                    } else if is_word_element(name.as_ref(), b"vAlign", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        sect.push_extra_xml(raw, 7, 0);
                        raw_position = (7, 1);
                    } else if is_word_element(name.as_ref(), b"textDirection", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        sect.push_extra_xml(raw, 8, 0);
                        raw_position = (8, 1);
                    } else if is_word_element(name.as_ref(), b"docGrid", &prefixes) {
                        // The schema makes `w:docGrid` empty. One that carries
                        // children is a producer extension, so it stays raw at
                        // the slot the typed element would have written.
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        sect.push_extra_xml(raw, 8, 1);
                        raw_position = (8, 2);
                    } else if is_word_element(name.as_ref(), b"pgNumType", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if sect.page_number.is_none() {
                            sect.page_number =
                                Some(Self::parse_page_number(e, &prefixes, owner_bindings, raw)?);
                        } else {
                            sect.push_extra_xml(raw, 6, 1);
                        }
                        raw_position = (6, 1);
                    } else if is_word_element(name.as_ref(), b"sectPrChange", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            if let Some(previous) = sect.change.replace(revision) {
                                sect.extra_xml
                                    .insert(change_raw_index, previous.into_raw_xml());
                                sect.extra_xml_positions.insert(
                                    change_raw_index,
                                    CT_SectPrRawPosition::Schema {
                                        slot: 9,
                                        occurrence: 0,
                                    },
                                );
                            }
                            change_raw_index = sect.extra_xml.len();
                        } else {
                            sect.push_extra_xml(raw, 9, 0);
                        }
                        raw_position = (9, 0);
                    } else {
                        // Capture unknown start elements as raw XML
                        let position = Self::raw_child_schema_slot(name.as_ref(), &prefixes)
                            .unwrap_or(raw_position);
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        sect.push_extra_xml(raw, position.0, position.1);
                        raw_position = position;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"sectPr") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        if let Some(record) = root
            .map(|root| capture_root_attribute_record(root, word_prefixes))
            .transpose()?
            .flatten()
        {
            sect.extra_xml.push(record);
            sect.extra_xml_positions.push(CT_SectPrRawPosition::Schema {
                slot: usize::MAX,
                occurrence: usize::MAX,
            });
        }
        sect.bind_story_reference_positions();
        Ok(sect)
    }

    fn parse_page_size_attributes(
        sect: &mut Self,
        element: &BytesStart<'_>,
        word_prefixes: &[String],
    ) -> Result<()> {
        for attribute in element.attributes() {
            let attribute = attribute?;
            let key = attribute.key.as_ref();
            let value = std::str::from_utf8(&attribute.value)?;
            let measurement = || parse_integer_measurement(element.name().as_ref(), key, value);
            if is_word_attribute(key, b"w", word_prefixes) {
                sect.page_width = Some(Twips(measurement()?));
            } else if is_word_attribute(key, b"h", word_prefixes) {
                sect.page_height = Some(Twips(measurement()?));
            } else if is_word_attribute(key, b"orient", word_prefixes) {
                sect.orientation = ST_PageOrientation::from_str(value).ok();
            }
        }
        Ok(())
    }

    fn parse_page_margin_attributes(
        sect: &mut Self,
        element: &BytesStart<'_>,
        word_prefixes: &[String],
    ) -> Result<()> {
        for attribute in element.attributes() {
            let attribute = attribute?;
            let key = attribute.key.as_ref();
            let field = if is_word_attribute(key, b"top", word_prefixes) {
                &mut sect.margin_top
            } else if is_word_attribute(key, b"right", word_prefixes)
                || is_word_attribute(key, b"end", word_prefixes)
            {
                &mut sect.margin_right
            } else if is_word_attribute(key, b"bottom", word_prefixes) {
                &mut sect.margin_bottom
            } else if is_word_attribute(key, b"left", word_prefixes)
                || is_word_attribute(key, b"start", word_prefixes)
            {
                &mut sect.margin_left
            } else if is_word_attribute(key, b"gutter", word_prefixes) {
                &mut sect.gutter
            } else if is_word_attribute(key, b"header", word_prefixes) {
                &mut sect.header_distance
            } else if is_word_attribute(key, b"footer", word_prefixes) {
                &mut sect.footer_distance
            } else {
                continue;
            };
            let value = std::str::from_utf8(&attribute.value)?;
            *field = Some(Twips(parse_integer_measurement(
                element.name().as_ref(),
                key,
                value,
            )?));
        }
        Ok(())
    }

    fn parse_story_reference(
        element: &BytesStart<'_>,
        word_prefixes: &[String],
    ) -> Result<Option<HdrFtrRef>> {
        let mut hdr_ftr_type = HdrFtrType::Default;
        let mut relationship_id = None;
        for attribute in element.attributes() {
            let attribute = attribute?;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
                .into_owned();
            if is_word_attribute(attribute.key.as_ref(), b"type", word_prefixes) {
                hdr_ftr_type = HdrFtrType::from_str(&value);
            } else if attribute_in_namespace(
                attribute.key.as_ref(),
                b"id",
                crate::namespace::R_NS,
                word_prefixes,
            ) {
                relationship_id = Some(value);
            }
        }
        Ok(relationship_id.map(|rel_id| HdrFtrRef {
            hdr_ftr_type,
            rel_id,
        }))
    }

    fn parse_page_number(
        e: &BytesStart<'_>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
        raw_xml: Vec<u8>,
    ) -> Result<CT_PageNumberType> {
        let mut start = None;
        let mut start_attribute_name = None;
        for attr in e.attributes() {
            let attr = attr?;
            if is_word_attribute(attr.key.as_ref(), b"start", word_prefixes) {
                start = attr
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, e.decoder())?
                    .parse()
                    .ok();
                start_attribute_name = Some(attr.key.as_ref().to_vec());
            }
        }
        let element_name = e.name();
        let element_prefix = element_name
            .as_ref()
            .iter()
            .position(|byte| *byte == b':')
            .map(|separator| &element_name.as_ref()[..separator]);
        let start_insertion_prefix = element_prefix
            .filter(|prefix| {
                !prefix.is_empty()
                    && word_prefixes
                        .iter()
                        .any(|candidate| candidate.as_bytes() == *prefix)
            })
            .or_else(|| {
                word_prefixes
                    .iter()
                    .find(|prefix| !prefix.is_empty() && !prefix.starts_with('\0'))
                    .map(String::as_bytes)
            })
            .map(<[u8]>::to_vec);
        let mut reserved_insertion_prefixes = owner_bindings
            .iter()
            .map(|(prefix, _)| prefix.as_bytes().to_vec())
            .collect::<Vec<_>>();
        reserved_insertion_prefixes.extend(
            namespace_bindings(word_prefixes)
                .into_iter()
                .map(|(prefix, _)| prefix.into_bytes()),
        );
        reserved_insertion_prefixes.sort();
        reserved_insertion_prefixes.dedup();
        Ok(CT_PageNumberType {
            start,
            raw_xml: Some(raw_xml),
            parsed_start: start,
            start_attribute_name,
            start_insertion_prefix,
            reserved_insertion_prefixes,
        })
    }

    fn push_extra_xml(&mut self, raw: Vec<u8>, slot: usize, occurrence: usize) {
        self.extra_xml.push(raw);
        self.extra_xml_positions
            .push(CT_SectPrRawPosition::Schema { slot, occurrence });
    }

    fn bind_story_reference_positions(&mut self) {
        for position in &mut self.extra_xml_positions {
            let (slot, occurrence) = match position {
                CT_SectPrRawPosition::Schema { slot, occurrence } => (*slot, *occurrence),
                _ => continue,
            };
            let (references, is_header) = match slot {
                0 => (&self.header_refs, true),
                1 => (&self.footer_refs, false),
                _ => continue,
            };
            let preceding = occurrence
                .checked_sub(1)
                .and_then(|index| CT_SectPrReferenceAnchor::at(references, index));
            let following = CT_SectPrReferenceAnchor::at(references, occurrence);
            *position = if is_header {
                CT_SectPrRawPosition::Header {
                    preceding,
                    following,
                }
            } else {
                CT_SectPrRawPosition::Footer {
                    preceding,
                    following,
                }
            };
        }
    }

    /// Schema slot and sub-slot for a `w:sectPr` child this type keeps raw.
    ///
    /// The sub-slot orders a retained child against the modeled siblings that
    /// share its slot, which is what keeps `xsd:sequence` intact now that
    /// `w:vAlign` and `w:textDirection` sit between them.
    fn raw_child_schema_slot(name: &[u8], word_prefixes: &[String]) -> Option<(usize, usize)> {
        let local = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
        if !is_word_element(name, local, word_prefixes) {
            return None;
        }
        match local {
            b"formProt" => Some((7, 0)),
            b"noEndnote" => Some((7, 1)),
            b"bidi" | b"rtlGutter" => Some((8, 1)),
            b"printerSettings" => Some((8, 2)),
            _ => None,
        }
    }

    /// Read `w:docGrid`, keeping every attribute this type does not model.
    fn parse_doc_grid(e: &BytesStart, word_prefixes: &[String]) -> Result<CT_DocGrid> {
        let mut grid = CT_DocGrid::default();
        for attribute in e.attributes() {
            let attribute = attribute?;
            let key = attribute.key.as_ref();
            let value = std::str::from_utf8(&attribute.value)?;
            let modeled = if is_word_attribute(key, b"type", word_prefixes) {
                match ST_DocGrid::from_str(value) {
                    Some(grid_type) => {
                        grid.grid_type = Some(grid_type);
                        true
                    }
                    None => false,
                }
            } else if is_word_attribute(key, b"linePitch", word_prefixes) {
                match value.parse() {
                    Ok(parsed) => {
                        grid.line_pitch = Some(Twips(parsed));
                        true
                    }
                    Err(_) => false,
                }
            } else if is_word_attribute(key, b"charSpace", word_prefixes) {
                match value.parse() {
                    Ok(parsed) => {
                        grid.char_space = Some(parsed);
                        true
                    }
                    Err(_) => false,
                }
            } else {
                false
            };
            if !modeled {
                grid.extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), value.to_owned()));
            }
        }
        Ok(grid)
    }

    fn parse_paper_source(e: &BytesStart, word_prefixes: &[String]) -> Result<CT_PaperSource> {
        let mut source = CT_PaperSource::default();
        for attribute in e.attributes() {
            let attribute = attribute?;
            let key = attribute.key.as_ref();
            let value = std::str::from_utf8(&attribute.value)?;
            if is_word_attribute(key, b"first", word_prefixes) {
                store_unsigned_attribute(
                    value,
                    key,
                    &mut source.first,
                    &mut source.extra_attributes,
                )?;
            } else if is_word_attribute(key, b"other", word_prefixes) {
                store_unsigned_attribute(
                    value,
                    key,
                    &mut source.other,
                    &mut source.extra_attributes,
                )?;
            } else {
                source
                    .extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), value.to_owned()));
            }
        }
        Ok(source)
    }

    fn parse_line_number(e: &BytesStart, word_prefixes: &[String]) -> Result<CT_LineNumber> {
        let mut numbering = CT_LineNumber::default();
        for attribute in e.attributes() {
            let attribute = attribute?;
            let key = attribute.key.as_ref();
            let value = std::str::from_utf8(&attribute.value)?;
            if is_word_attribute(key, b"countBy", word_prefixes) {
                store_unsigned_attribute(
                    value,
                    key,
                    &mut numbering.count_by,
                    &mut numbering.extra_attributes,
                )?;
            } else if is_word_attribute(key, b"start", word_prefixes) {
                store_unsigned_attribute(
                    value,
                    key,
                    &mut numbering.start,
                    &mut numbering.extra_attributes,
                )?;
            } else if is_word_attribute(key, b"distance", word_prefixes) {
                match value.parse() {
                    Ok(parsed) => numbering.distance = Some(Twips(parsed)),
                    Err(_) => numbering
                        .extra_attributes
                        .push((std::str::from_utf8(key)?.to_owned(), value.to_owned())),
                }
            } else if is_word_attribute(key, b"restart", word_prefixes) {
                match ST_LineNumberRestart::from_str(value) {
                    Some(parsed) => numbering.restart = Some(parsed),
                    None => numbering
                        .extra_attributes
                        .push((std::str::from_utf8(key)?.to_owned(), value.to_owned())),
                }
            } else {
                numbering
                    .extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), value.to_owned()));
            }
        }
        Ok(numbering)
    }

    fn parse_page_border_attrs(e: &BytesStart, word_prefixes: &[String]) -> Result<CT_PageBorders> {
        let mut borders = CT_PageBorders::default();
        for attribute in e.attributes() {
            let attribute = attribute?;
            let key = attribute.key.as_ref();
            let value = std::str::from_utf8(&attribute.value)?;
            let retained = if is_word_attribute(key, b"zOrder", word_prefixes) {
                borders.z_order = ST_PageBorderZOrder::from_str(value);
                borders.z_order.is_none()
            } else if is_word_attribute(key, b"display", word_prefixes) {
                borders.display = ST_PageBorderDisplay::from_str(value);
                borders.display.is_none()
            } else if is_word_attribute(key, b"offsetFrom", word_prefixes) {
                borders.offset_from = ST_PageBorderOffset::from_str(value);
                borders.offset_from.is_none()
            } else {
                true
            };
            if retained {
                borders
                    .extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), value.to_owned()));
            }
        }
        Ok(borders)
    }

    fn parse_page_borders(
        reader: &mut Reader<&[u8]>,
        e: &BytesStart,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
    ) -> Result<CT_PageBorders> {
        let mut borders = Self::parse_page_border_attrs(e, word_prefixes)?;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref child)) => {
                    let name = child.name();
                    let prefixes = word_prefixes_at(child, word_prefixes)?;
                    let edge = |child: &BytesStart| {
                        CT_BorderEdge::from_xml_attrs_with_prefixes(child, &prefixes)
                    };
                    if is_word_element(name.as_ref(), b"top", &prefixes) {
                        borders.top = Some(edge(child)?);
                    } else if is_word_element(name.as_ref(), b"left", &prefixes) {
                        borders.left = Some(edge(child)?);
                    } else if is_word_element(name.as_ref(), b"bottom", &prefixes) {
                        borders.bottom = Some(edge(child)?);
                    } else if is_word_element(name.as_ref(), b"right", &prefixes) {
                        borders.right = Some(edge(child)?);
                    } else {
                        borders
                            .extra_xml
                            .push(crate::text::raw_with_external_bindings(
                                &capture_empty_element(child)?,
                                owner_bindings,
                            )?);
                    }
                }
                Ok(Event::Start(ref child)) => {
                    borders
                        .extra_xml
                        .push(crate::text::raw_with_external_bindings(
                            &capture_element(reader, child)?,
                            owner_bindings,
                        )?);
                }
                Ok(Event::End(ref end))
                    if matches_local_name(end.name().as_ref(), b"pgBorders") =>
                {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(error) => return Err(error.into()),
                _ => {}
            }
            buf.clear();
        }
        Ok(borders)
    }

    pub(crate) fn parse_note_properties(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
        tag: &[u8],
    ) -> Result<CT_NoteProperties> {
        let mut properties = CT_NoteProperties::default();
        let mut slot = 0usize;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref child)) => {
                    let name = child.name();
                    let prefixes = word_prefixes_at(child, word_prefixes)?;
                    let modeled = if is_word_element(name.as_ref(), b"pos", &prefixes) {
                        slot = 1;
                        sole_val_attribute(child, &prefixes)?
                            .map(|value| properties.pos = Some(value))
                    } else if is_word_element(name.as_ref(), b"numFmt", &prefixes) {
                        slot = 2;
                        sole_val_attribute(child, &prefixes)?
                            .map(|value| properties.num_fmt = Some(value))
                    } else if is_word_element(name.as_ref(), b"numStart", &prefixes) {
                        slot = 3;
                        sole_val_attribute(child, &prefixes)?
                            .and_then(|value| value.parse().ok())
                            .map(|value| properties.num_start = Some(value))
                    } else if is_word_element(name.as_ref(), b"numRestart", &prefixes) {
                        slot = 4;
                        sole_val_attribute(child, &prefixes)?
                            .map(|value| properties.num_restart = Some(value))
                    } else if is_word_element(
                        name.as_ref(),
                        if matches_local_name(tag, b"footnotePr") {
                            b"footnote".as_slice()
                        } else {
                            b"endnote".as_slice()
                        },
                        &prefixes,
                    ) {
                        slot = 5;
                        child
                            .attributes()
                            .filter_map(|attribute| attribute.ok())
                            .find(|attribute| {
                                is_word_attribute(attribute.key.as_ref(), b"id", &prefixes)
                            })
                            .and_then(|attribute| {
                                std::str::from_utf8(&attribute.value)
                                    .ok()
                                    .and_then(|value| value.parse::<i32>().ok())
                            })
                            .map(|id| properties.special_references.push(id))
                    } else {
                        None
                    };
                    if modeled.is_none() {
                        properties.extra_xml.push((
                            slot,
                            crate::text::raw_with_external_bindings(
                                &capture_empty_element(child)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                }
                Ok(Event::Start(ref child)) => {
                    properties.extra_xml.push((
                        slot,
                        crate::text::raw_with_external_bindings(
                            &capture_element(reader, child)?,
                            owner_bindings,
                        )?,
                    ));
                }
                Ok(Event::End(ref end)) if matches_local_name(end.name().as_ref(), tag) => break,
                Ok(Event::Eof) => break,
                Err(error) => return Err(error.into()),
                _ => {}
            }
            buf.clear();
        }
        Ok(properties)
    }

    fn parse_cols_attrs(e: &BytesStart, word_prefixes: &[String]) -> Result<CT_Columns> {
        let mut cols = CT_Columns::default();
        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let val_str = std::str::from_utf8(&attr.value)?;
            if is_word_attribute(key, b"num", word_prefixes) {
                cols.num = Some(val_str.parse()?);
            } else if is_word_attribute(key, b"space", word_prefixes) {
                cols.space = Some(Twips(parse_integer_measurement(
                    e.name().as_ref(),
                    key,
                    val_str,
                )?));
            } else if is_word_attribute(key, b"equalWidth", word_prefixes) {
                cols.equal_width = Some(val_str == "1" || val_str == "true");
            } else if is_word_attribute(key, b"sep", word_prefixes) {
                cols.sep = Some(val_str == "1" || val_str == "true");
            }
        }
        Ok(cols)
    }

    fn parse_cols_empty(e: &BytesStart, word_prefixes: &[String]) -> Result<CT_Columns> {
        Self::parse_cols_attrs(e, word_prefixes)
    }

    fn parse_cols_start(
        reader: &mut Reader<&[u8]>,
        e: &BytesStart,
        word_prefixes: &[String],
    ) -> Result<CT_Columns> {
        let mut cols = Self::parse_cols_attrs(e, word_prefixes)?;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if !is_word_element(e.name().as_ref(), b"col", &prefixes) {
                        buf.clear();
                        continue;
                    }
                    let mut width = None;
                    let mut space = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = attr.key.as_ref();
                        let value = std::str::from_utf8(&attr.value)?;
                        let measurement =
                            || parse_integer_measurement(e.name().as_ref(), key, value);
                        if is_word_attribute(key, b"w", &prefixes) {
                            width = Some(Twips(measurement()?));
                        } else if is_word_attribute(key, b"space", &prefixes) {
                            space = Some(Twips(measurement()?));
                        }
                    }
                    cols.columns.push(CT_Column { width, space });
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"cols") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(cols)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let mut root = BytesStart::new("w:sectPr");
        for (raw, position) in self.extra_xml.iter().zip(&self.extra_xml_positions) {
            if Self::raw_position_is_root_attributes(position, raw) {
                push_root_attribute_record(&mut root, raw, None)?;
            }
        }
        writer.write_event(Event::Start(root))?;
        let ordered_raw = self.extra_xml_positions.len() == self.extra_xml.len();
        if ordered_raw {
            self.write_story_reference_boundary(writer, &self.header_refs, true, 0)?;
        }

        // headerReference elements
        for (index, hdr) in self.header_refs.iter().enumerate() {
            let mut e = BytesStart::new("w:headerReference");
            e.push_attribute(("w:type", hdr.hdr_ftr_type.to_str()));
            e.push_attribute(("r:id", hdr.rel_id.as_str()));
            writer.write_event(Event::Empty(e))?;
            if ordered_raw {
                self.write_story_reference_boundary(writer, &self.header_refs, true, index + 1)?;
            }
        }
        if ordered_raw {
            self.write_story_reference_boundary(writer, &self.footer_refs, false, 0)?;
        }

        // footerReference elements
        for (index, ftr) in self.footer_refs.iter().enumerate() {
            let mut e = BytesStart::new("w:footerReference");
            e.push_attribute(("w:type", ftr.hdr_ftr_type.to_str()));
            e.push_attribute(("r:id", ftr.rel_id.as_str()));
            writer.write_event(Event::Empty(e))?;
            if ordered_raw {
                self.write_story_reference_boundary(writer, &self.footer_refs, false, index + 1)?;
            }
        }
        // footnotePr and endnotePr. F-274 owns what the values do to notes.
        if let Some(footnote_pr) = &self.footnote_pr {
            footnote_pr.to_xml(writer, "w:footnotePr")?;
        }
        if ordered_raw {
            self.write_raw_position(writer, 2, 0)?;
        }
        if let Some(endnote_pr) = &self.endnote_pr {
            endnote_pr.to_xml(writer, "w:endnotePr")?;
        }
        if ordered_raw {
            self.write_raw_position(writer, 2, 1)?;
        }

        // type (section break type)
        if let Some(st) = self.section_type {
            let mut e = BytesStart::new("w:type");
            e.push_attribute(("w:val", st.to_str()));
            writer.write_event(Event::Empty(e))?;
        }
        if ordered_raw {
            self.write_raw_position(writer, 3, 0)?;
        }

        // pgSz
        if self.page_width.is_some() || self.page_height.is_some() || self.orientation.is_some() {
            let mut e = BytesStart::new("w:pgSz");
            if let Some(w) = self.page_width {
                e.push_attribute(("w:w", buf.format(w.0)));
            }
            if let Some(h) = self.page_height {
                e.push_attribute(("w:h", buf.format(h.0)));
            }
            if let Some(orient) = self.orientation
                && orient == ST_PageOrientation::Landscape
            {
                e.push_attribute(("w:orient", orient.to_str()));
            }
            writer.write_event(Event::Empty(e))?;
        }
        if ordered_raw {
            self.write_raw_position(writer, 4, 0)?;
        }

        // pgMar
        if self.margin_top.is_some()
            || self.margin_right.is_some()
            || self.margin_bottom.is_some()
            || self.margin_left.is_some()
            || self.gutter.is_some()
            || self.header_distance.is_some()
            || self.footer_distance.is_some()
        {
            let mut e = BytesStart::new("w:pgMar");
            if let Some(t) = self.margin_top {
                e.push_attribute(("w:top", buf.format(t.0)));
            }
            if let Some(r) = self.margin_right {
                e.push_attribute(("w:right", buf.format(r.0)));
            }
            if let Some(b) = self.margin_bottom {
                e.push_attribute(("w:bottom", buf.format(b.0)));
            }
            if let Some(l) = self.margin_left {
                e.push_attribute(("w:left", buf.format(l.0)));
            }
            if let Some(g) = self.gutter {
                e.push_attribute(("w:gutter", buf.format(g.0)));
            }
            if let Some(h) = self.header_distance {
                e.push_attribute(("w:header", buf.format(h.0)));
            }
            if let Some(f) = self.footer_distance {
                e.push_attribute(("w:footer", buf.format(f.0)));
            }
            writer.write_event(Event::Empty(e))?;
        }
        if ordered_raw {
            self.write_raw_position(writer, 5, 0)?;
        }

        // paperSrc. A printer tray, so pagination is identical with and without it.
        if let Some(paper_source) = &self.paper_source {
            let mut e = BytesStart::new("w:paperSrc");
            push_retained_attributes(&mut e, &paper_source.extra_attributes);
            if let Some(first) = paper_source.first {
                e.push_attribute(("w:first", buf.format(first)));
            }
            if let Some(other) = paper_source.other {
                e.push_attribute(("w:other", buf.format(other)));
            }
            writer.write_event(Event::Empty(e))?;
        }
        if ordered_raw {
            self.write_raw_position(writer, 5, 1)?;
        }

        // pgBorders
        if let Some(page_borders) = &self.page_borders {
            page_borders.to_xml(writer)?;
        }
        if ordered_raw {
            self.write_raw_position(writer, 5, 2)?;
        }

        // lnNumType
        if let Some(line_numbers) = &self.line_numbers {
            let mut e = BytesStart::new("w:lnNumType");
            push_retained_attributes(&mut e, &line_numbers.extra_attributes);
            if let Some(count_by) = line_numbers.count_by {
                e.push_attribute(("w:countBy", buf.format(count_by)));
            }
            if let Some(start) = line_numbers.start {
                e.push_attribute(("w:start", buf.format(start)));
            }
            if let Some(distance) = line_numbers.distance {
                e.push_attribute(("w:distance", buf.format(distance.0)));
            }
            if let Some(restart) = line_numbers.restart {
                e.push_attribute(("w:restart", restart.to_str()));
            }
            writer.write_event(Event::Empty(e))?;
        }
        if ordered_raw {
            self.write_raw_position(writer, 5, 3)?;
        }

        // pgNumType. M24 format and chapter attributes stay byte-exact unless start changes.
        if let Some(page_number) = &self.page_number {
            page_number.to_xml(writer)?;
        }
        if ordered_raw {
            self.write_raw_position(writer, 6, 1)?;
        }

        // cols
        if let Some(ref cols) = self.columns {
            if cols.columns.is_empty() {
                // Simple equal-width columns
                let mut e = BytesStart::new("w:cols");
                if let Some(num) = cols.num {
                    e.push_attribute(("w:num", buf.format(num)));
                }
                if let Some(space) = cols.space {
                    e.push_attribute(("w:space", buf.format(space.0)));
                }
                if let Some(eq) = cols.equal_width
                    && !eq
                {
                    e.push_attribute(("w:equalWidth", "0"));
                }
                if let Some(sep) = cols.sep
                    && sep
                {
                    e.push_attribute(("w:sep", "1"));
                }
                writer.write_event(Event::Empty(e))?;
            } else {
                // Individual column definitions
                let mut e = BytesStart::new("w:cols");
                if let Some(num) = cols.num {
                    e.push_attribute(("w:num", buf.format(num)));
                }
                if let Some(eq) = cols.equal_width {
                    e.push_attribute(("w:equalWidth", if eq { "1" } else { "0" }));
                }
                if let Some(sep) = cols.sep
                    && sep
                {
                    e.push_attribute(("w:sep", "1"));
                }
                writer.write_event(Event::Start(e))?;

                for col in &cols.columns {
                    let mut ce = BytesStart::new("w:col");
                    if let Some(w) = col.width {
                        ce.push_attribute(("w:w", buf.format(w.0)));
                    }
                    if let Some(s) = col.space {
                        ce.push_attribute(("w:space", buf.format(s.0)));
                    }
                    writer.write_event(Event::Empty(ce))?;
                }

                writer.write_event(Event::End(BytesEnd::new("w:cols")))?;
            }
        }
        if ordered_raw {
            self.write_raw_position(writer, 7, 0)?;
        }

        // vAlign
        if let Some(vertical_alignment) = self.vertical_alignment {
            let mut e = BytesStart::new("w:vAlign");
            e.push_attribute(("w:val", vertical_alignment.to_str()));
            writer.write_event(Event::Empty(e))?;
        }
        if ordered_raw {
            self.write_raw_position(writer, 7, 1)?;
        }

        // titlePg
        if let Some(title_pg) = self.title_pg {
            let mut e = BytesStart::new("w:titlePg");
            if !title_pg {
                e.push_attribute(("w:val", "0"));
            }
            writer.write_event(Event::Empty(e))?;
        }

        if ordered_raw {
            self.write_raw_position(writer, 8, 0)?;
        }

        // textDirection. Authored and preserved here, projected by F-266c.
        if let Some(text_direction) = &self.text_direction {
            let mut e = BytesStart::new("w:textDirection");
            e.push_attribute(("w:val", text_direction.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        // docGrid. The grid type, line pitch and character space this crate
        // models, with any producer attribute it does not written first. A
        // legacy caller that populates only `extra_xml` gets the modeled
        // element ahead of its raw dump, which is the order the rest of this
        // serialiser already writes a modeled child in.
        let mut write_doc_grid = |writer: &mut Writer<W>| -> Result<()> {
            if let Some(doc_grid) = &self.doc_grid {
                let mut e = BytesStart::new("w:docGrid");
                push_retained_attributes(&mut e, &doc_grid.extra_attributes);
                if let Some(grid_type) = doc_grid.grid_type {
                    e.push_attribute(("w:type", grid_type.to_str()));
                }
                if let Some(line_pitch) = doc_grid.line_pitch {
                    e.push_attribute(("w:linePitch", buf.format(line_pitch.0)));
                }
                if let Some(char_space) = doc_grid.char_space {
                    e.push_attribute(("w:charSpace", buf.format(char_space)));
                }
                writer.write_event(Event::Empty(e))?;
            }
            Ok(())
        };
        if ordered_raw {
            self.write_raw_position(writer, 8, 1)?;
            write_doc_grid(writer)?;
            self.write_raw_position(writer, 8, 2)?;
            self.write_raw_position(writer, 9, 0)?;
        } else {
            write_doc_grid(writer)?;
            // Legacy callers that populate only `extra_xml` retain the previous position.
            for raw in &self.extra_xml {
                if is_root_attribute_record(raw) {
                    continue;
                }
                writer.get_mut().write_all(raw)?;
            }
        }
        if let Some(change) = &self.change {
            change.write_xml(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:sectPr")))?;
        Ok(())
    }

    fn write_raw_position<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        slot: usize,
        occurrence: usize,
    ) -> Result<()> {
        for (position, raw) in self.extra_xml_positions.iter().zip(&self.extra_xml) {
            if matches!(
                position,
                CT_SectPrRawPosition::Schema {
                    slot: candidate_slot,
                    occurrence: candidate_occurrence,
                } if (*candidate_slot, *candidate_occurrence) == (slot, occurrence)
            ) {
                writer.get_mut().write_all(raw)?;
            }
        }
        Ok(())
    }

    fn write_story_reference_boundary<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        references: &[HdrFtrRef],
        header: bool,
        boundary: usize,
    ) -> Result<()> {
        for (position, raw) in self.extra_xml_positions.iter().zip(&self.extra_xml) {
            let anchored = match position {
                CT_SectPrRawPosition::Header {
                    preceding,
                    following,
                } if header => Some((preceding, following)),
                CT_SectPrRawPosition::Footer {
                    preceding,
                    following,
                } if !header => Some((preceding, following)),
                CT_SectPrRawPosition::Schema { slot, occurrence }
                    if (*slot == 0 && header) || (*slot == 1 && !header) =>
                {
                    if (*occurrence).min(references.len()) == boundary {
                        writer.get_mut().write_all(raw)?;
                    }
                    None
                }
                _ => None,
            };
            let Some((preceding, following)) = anchored else {
                continue;
            };
            let resolved = following
                .as_ref()
                .and_then(|reference| reference.resolve(references))
                .or_else(|| {
                    preceding
                        .as_ref()
                        .and_then(|reference| reference.resolve(references).map(|index| index + 1))
                })
                .unwrap_or(references.len());
            if resolved == boundary {
                writer.get_mut().write_all(raw)?;
            }
        }
        Ok(())
    }
}

impl CT_PageBorders {
    /// Write `w:pgBorders` with its edges in `xsd:sequence` order.
    ///
    /// Retained attributes are written ahead of the modeled ones and retained
    /// children after the modeled edges, which is what `CT_BorderEdge` does.
    fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut e = BytesStart::new("w:pgBorders");
        push_retained_attributes(&mut e, &self.extra_attributes);
        if let Some(z_order) = self.z_order {
            e.push_attribute(("w:zOrder", z_order.to_str()));
        }
        if let Some(display) = self.display {
            e.push_attribute(("w:display", display.to_str()));
        }
        if let Some(offset_from) = self.offset_from {
            e.push_attribute(("w:offsetFrom", offset_from.to_str()));
        }
        if self.top.is_none()
            && self.left.is_none()
            && self.bottom.is_none()
            && self.right.is_none()
            && self.extra_xml.is_empty()
        {
            writer.write_event(Event::Empty(e))?;
            return Ok(());
        }
        writer.write_event(Event::Start(e))?;
        for (edge, tag) in [
            (&self.top, "w:top"),
            (&self.left, "w:left"),
            (&self.bottom, "w:bottom"),
            (&self.right, "w:right"),
        ] {
            if let Some(edge) = edge {
                edge.to_xml(writer, tag)?;
            }
        }
        for raw in &self.extra_xml {
            writer.get_mut().write_all(raw)?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:pgBorders")))?;
        Ok(())
    }
}

impl CT_NoteProperties {
    /// Write section note properties under `tag`, in `xsd:sequence` order.
    pub(crate) fn to_xml<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        tag: &str,
    ) -> Result<()> {
        if self.pos.is_none()
            && self.num_fmt.is_none()
            && self.num_start.is_none()
            && self.num_restart.is_none()
            && self.special_references.is_empty()
            && self.extra_xml.is_empty()
        {
            writer.write_event(Event::Empty(BytesStart::new(tag)))?;
            return Ok(());
        }
        writer.write_event(Event::Start(BytesStart::new(tag)))?;
        self.write_retained_children(writer, 0)?;
        if let Some(pos) = &self.pos {
            let mut e = BytesStart::new("w:pos");
            e.push_attribute(("w:val", pos.as_str()));
            writer.write_event(Event::Empty(e))?;
        }
        self.write_retained_children(writer, 1)?;
        if let Some(num_fmt) = &self.num_fmt {
            let mut e = BytesStart::new("w:numFmt");
            e.push_attribute(("w:val", num_fmt.as_str()));
            writer.write_event(Event::Empty(e))?;
        }
        self.write_retained_children(writer, 2)?;
        if let Some(num_start) = self.num_start {
            let mut buf = itoa::Buffer::new();
            let mut e = BytesStart::new("w:numStart");
            e.push_attribute(("w:val", buf.format(num_start)));
            writer.write_event(Event::Empty(e))?;
        }
        self.write_retained_children(writer, 3)?;
        if let Some(num_restart) = &self.num_restart {
            let mut e = BytesStart::new("w:numRestart");
            e.push_attribute(("w:val", num_restart.as_str()));
            writer.write_event(Event::Empty(e))?;
        }
        self.write_retained_children(writer, 4)?;
        let child_tag = if tag.ends_with("footnotePr") {
            "w:footnote"
        } else {
            "w:endnote"
        };
        for id in &self.special_references {
            let mut buf = itoa::Buffer::new();
            let mut child = BytesStart::new(child_tag);
            child.push_attribute(("w:id", buf.format(*id)));
            writer.write_event(Event::Empty(child))?;
        }
        self.write_retained_children(writer, 5)?;
        writer.write_event(Event::End(BytesEnd::new(tag)))?;
        Ok(())
    }

    fn write_retained_children<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        slot: usize,
    ) -> Result<()> {
        for (candidate, raw) in &self.extra_xml {
            if *candidate == slot {
                writer.get_mut().write_all(raw)?;
            }
        }
        Ok(())
    }
}

impl CT_PageNumberType {
    fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if let Some(raw) = &self.raw_xml {
            if self.start == self.parsed_start {
                writer.get_mut().write_all(raw)?;
                return Ok(());
            }
            writer.get_mut().write_all(&replace_page_number_start(
                raw,
                self.start,
                self.start_attribute_name.as_deref(),
                self.start_insertion_prefix.as_deref(),
                &self.reserved_insertion_prefixes,
            )?)?;
            return Ok(());
        }

        let mut element = BytesStart::new("w:pgNumType");
        if let Some(start) = self.start {
            let mut buf = itoa::Buffer::new();
            element.push_attribute(("w:start", buf.format(start)));
        }
        writer.write_event(Event::Empty(element))?;
        Ok(())
    }
}

/// Store a parsed unsigned attribute, or retain it verbatim when the source
/// value falls outside what the field models, so nothing is dropped on write.
fn store_unsigned_attribute(
    value: &str,
    key: &[u8],
    field: &mut Option<u32>,
    retained: &mut Vec<(String, String)>,
) -> Result<()> {
    match value.parse() {
        Ok(parsed) => *field = Some(parsed),
        Err(_) => retained.push((std::str::from_utf8(key)?.to_owned(), value.to_owned())),
    }
    Ok(())
}

/// Return the `w:val` of an element whose only attribute is `w:val`.
///
/// `None` means the element carries something this crate does not model, and
/// the caller keeps the element raw rather than typing part of it.
fn sole_val_attribute(e: &BytesStart<'_>, word_prefixes: &[String]) -> Result<Option<String>> {
    let mut value = None;
    for attribute in e.attributes() {
        let attribute = attribute?;
        if !is_word_attribute(attribute.key.as_ref(), b"val", word_prefixes) {
            return Ok(None);
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, e.decoder())?
                .into_owned(),
        );
    }
    Ok(value)
}

/// Write the retained attributes of a section child, ahead of modeled ones.
fn push_retained_attributes(e: &mut BytesStart<'_>, retained: &[(String, String)]) {
    for (name, value) in retained {
        e.push_attribute(quick_xml::events::attributes::Attribute {
            key: quick_xml::name::QName(name.as_bytes()),
            value: std::borrow::Cow::Borrowed(value.as_bytes()),
        });
    }
}

fn replace_page_number_start(
    raw: &[u8],
    start: Option<u32>,
    attribute_name: Option<&[u8]>,
    insertion_prefix: Option<&[u8]>,
    reserved_prefixes: &[Vec<u8>],
) -> Result<Vec<u8>> {
    let replacement = start.map(|value| value.to_string());
    let mut output = raw.to_vec();
    let Some((tag_end, attributes)) = start_tag_attributes(raw) else {
        return Ok(output);
    };
    if let Some(name) = attribute_name
        && let Some((_, value, full)) = attributes
            .iter()
            .find(|(candidate, _, _)| &raw[candidate.clone()] == name)
    {
        if let Some(ref value_text) = replacement {
            output.splice(value.clone(), value_text.bytes());
        } else {
            output.drain(full.clone());
        }
        return Ok(output);
    }

    if let Some(value) = replacement {
        let insert_at = raw[..tag_end]
            .iter()
            .rposition(|byte| !byte.is_ascii_whitespace() && *byte != b'/')
            .map_or(tag_end, |index| index + 1);
        let inserted = if let Some(prefix) = insertion_prefix {
            format!(" {}:start=\"{value}\"", String::from_utf8_lossy(prefix))
        } else {
            let prefix = unused_page_number_prefix(raw, tag_end, &attributes, reserved_prefixes);
            format!(" xmlns:{prefix}=\"{W_NS}\" {prefix}:start=\"{value}\"")
        };
        output.splice(insert_at..insert_at, inserted.bytes());
    }
    Ok(output)
}

fn unused_page_number_prefix(
    raw: &[u8],
    tag_end: usize,
    attributes: &[AttributeSpans],
    reserved_prefixes: &[Vec<u8>],
) -> String {
    for suffix in 0usize.. {
        let prefix = if suffix == 0 {
            "rdocxWord".to_owned()
        } else {
            format!("rdocxWord{suffix}")
        };
        let declaration = format!("xmlns:{prefix}");
        let qualified = format!("{prefix}:");
        let used = reserved_prefixes
            .iter()
            .any(|candidate| candidate == prefix.as_bytes())
            || attributes.iter().any(|(name, _, _)| {
                &raw[name.clone()] == declaration.as_bytes()
                    || raw[name.clone()].starts_with(qualified.as_bytes())
            })
            || raw[1..tag_end].starts_with(qualified.as_bytes());
        if !used {
            return prefix;
        }
    }
    unreachable!("the finite start tag cannot use every generated prefix")
}

fn attribute_in_namespace(
    name: &[u8],
    local_name: &[u8],
    namespace: &str,
    scope: &[String],
) -> bool {
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return false;
    };
    if name.get(separator + 1..) != Some(local_name) {
        return false;
    }
    let prefix = &name[..separator];
    scope.iter().any(|binding| {
        binding
            .strip_prefix('\0')
            .and_then(|binding| binding.split_once('\0'))
            .is_some_and(|(candidate, value)| candidate.as_bytes() == prefix && value == namespace)
    })
}

fn page_size_attributes_are_modeled(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
) -> Result<bool> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        if namespace_declaration(key) {
            continue;
        }
        let value = std::str::from_utf8(&attribute.value)?;
        if is_word_attribute(key, b"w", word_prefixes)
            || is_word_attribute(key, b"h", word_prefixes)
        {
            if parse_integer_measurement::<i32>(element.name().as_ref(), key, value).is_err() {
                return Ok(false);
            }
        } else if is_word_attribute(key, b"orient", word_prefixes) {
            if ST_PageOrientation::from_str(value).is_err() {
                return Ok(false);
            }
        } else {
            return Ok(false);
        }
    }
    Ok(true)
}

fn page_margin_attributes_are_modeled(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
) -> Result<bool> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        if namespace_declaration(key) {
            continue;
        }
        let modeled = [
            b"top".as_slice(),
            b"right".as_slice(),
            b"end".as_slice(),
            b"bottom".as_slice(),
            b"left".as_slice(),
            b"start".as_slice(),
            b"gutter".as_slice(),
            b"header".as_slice(),
            b"footer".as_slice(),
        ]
        .iter()
        .any(|local| is_word_attribute(key, local, word_prefixes));
        let value = std::str::from_utf8(&attribute.value)?;
        if !modeled
            || parse_integer_measurement::<i32>(element.name().as_ref(), key, value).is_err()
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn story_reference_attributes_are_modeled(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
) -> Result<bool> {
    let mut type_count = 0usize;
    let mut relationship_count = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        if namespace_declaration(key) {
            continue;
        }
        let value = std::str::from_utf8(&attribute.value)?;
        if is_word_attribute(key, b"type", word_prefixes) {
            type_count += 1;
            if type_count > 1 || !matches!(value, "default" | "first" | "even") {
                return Ok(false);
            }
        } else if attribute_in_namespace(key, b"id", crate::namespace::R_NS, word_prefixes) {
            relationship_count += 1;
            if relationship_count > 1 || value.is_empty() {
                return Ok(false);
            }
        } else {
            return Ok(false);
        }
    }
    Ok(relationship_count == 1)
}

fn namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn raw_element_has_child_content(raw: &[u8]) -> Result<bool> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(_) => {
                if depth > 0 {
                    return Ok(true);
                }
                depth += 1;
            }
            Event::Empty(_) if depth > 0 => return Ok(true),
            Event::Text(text)
                if depth > 0 && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Ok(true);
            }
            Event::CData(text)
                if depth > 0 && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Ok(true);
            }
            Event::GeneralRef(_) if depth > 0 => return Ok(true),
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

type AttributeSpans = (
    std::ops::Range<usize>,
    std::ops::Range<usize>,
    std::ops::Range<usize>,
);

fn start_tag_attributes(raw: &[u8]) -> Option<(usize, Vec<AttributeSpans>)> {
    if raw.first() != Some(&b'<') {
        return None;
    }
    let mut quote = None;
    let tag_end = raw
        .iter()
        .enumerate()
        .find_map(|(index, byte)| match (*byte, quote) {
            (b'\'' | b'"', None) => {
                quote = Some(*byte);
                None
            }
            (value, Some(open)) if value == open => {
                quote = None;
                None
            }
            (b'>', None) => Some(index),
            _ => None,
        })?;

    let mut cursor = 1usize;
    while cursor < tag_end && !raw[cursor].is_ascii_whitespace() && raw[cursor] != b'/' {
        cursor += 1;
    }
    let mut attributes = Vec::new();
    while cursor < tag_end {
        let full_start = cursor;
        while cursor < tag_end && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor == tag_end || raw[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < tag_end
            && !raw[cursor].is_ascii_whitespace()
            && raw[cursor] != b'='
            && raw[cursor] != b'/'
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag_end && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if raw.get(cursor) != Some(&b'=') {
            return None;
        }
        cursor += 1;
        while cursor < tag_end && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let open_quote = *raw.get(cursor)?;
        if open_quote != b'\'' && open_quote != b'"' {
            return None;
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag_end && raw[cursor] != open_quote {
            cursor += 1;
        }
        if cursor == tag_end {
            return None;
        }
        let value_end = cursor;
        cursor += 1;
        attributes.push((
            name_start..name_end,
            value_start..value_end,
            full_start..cursor,
        ));
    }
    Some((tag_end, attributes))
}

/// `CT_Body` — The document body containing paragraphs, tables, and section properties.
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_Body {
    /// Mixed content: paragraphs and tables in document order.
    pub content: Vec<BodyContent>,
    pub sect_pr: Option<CT_SectPr>,
}

#[allow(non_snake_case)]
impl CT_Body {
    pub fn new() -> Self {
        CT_Body {
            content: Vec::new(),
            sect_pr: Some(CT_SectPr::default_letter()),
        }
    }

    /// Get an iterator over only the paragraphs.
    pub fn paragraphs(&self) -> impl Iterator<Item = &CT_P> {
        let mut paragraphs = Vec::new();
        for content in &self.content {
            match content {
                BodyContent::Paragraph(paragraph) => paragraphs.push(paragraph),
                BodyContent::ContentControl(sdt) => {
                    sdt.collect_paragraphs(SdtOwner::Body, &mut paragraphs)
                }
                BodyContent::Table(_) | BodyContent::RawXml(_) => {}
            }
        }
        paragraphs.into_iter()
    }

    /// Whether accepting every tracked change joins the body paragraph at
    /// `index` to the next body paragraph, as Word does when a paragraph mark
    /// is deleted or moved away (`w:del` or `w:moveFrom` in the mark `w:rPr`).
    ///
    /// The joined paragraph keeps the properties of the next paragraph, so a
    /// joining paragraph with no accepted content leaves nothing. The next
    /// body item must be a paragraph, the adjacency `Document::accept_all`
    /// requires before it merges a paragraph. `accept_all` refuses a removed
    /// mark before a table, a content control, preserved XML or the end of
    /// the body, and those paragraphs keep their marks here. A paragraph that
    /// ends a section also keeps its mark here, although `accept_all` merges
    /// its section break away.
    #[doc(hidden)]
    pub fn accepted_paragraph_joins_next(&self, index: usize) -> bool {
        let Some(BodyContent::Paragraph(paragraph)) = self.content.get(index) else {
            return false;
        };
        let Some(properties) = paragraph.properties.as_ref() else {
            return false;
        };
        let Some(mark) = properties.rpr.as_ref() else {
            return false;
        };
        // The model keeps a moved-away mark as preserved XML, read with the
        // fixed `w` prefix and any binding the XML declares itself. A mark
        // whose Word prefix is bound only outside it is not recognized, so
        // that paragraph keeps its mark.
        let moved_away = || {
            mark.revision_xml.iter().any(|raw| {
                let mut reader = Reader::from_reader(raw.as_slice());
                matches!(
                    reader.read_event(),
                    Ok(Event::Start(start) | Event::Empty(start))
                        if word_prefixes_at(&start, &["w".to_owned()]).is_ok_and(|prefixes| {
                            is_word_element(start.name().as_ref(), b"moveFrom", &prefixes)
                        })
                )
            })
        };
        let removed = mark
            .revision_markers
            .iter()
            .any(|marker| marker.kind() == RevisionKind::Deletion)
            || moved_away();
        removed
            && properties.sect_pr.is_none()
            && matches!(self.content.get(index + 1), Some(BodyContent::Paragraph(_)))
    }

    /// Get a mutable iterator over only the paragraphs.
    pub fn paragraphs_mut(&mut self) -> impl Iterator<Item = &mut CT_P> {
        self.content.iter_mut().filter_map(|c| match c {
            BodyContent::Paragraph(p) => Some(p),
            _ => None,
        })
    }

    /// Get an iterator over only the tables.
    pub fn tables(&self) -> impl Iterator<Item = &CT_Tbl> {
        let mut tables = Vec::new();
        for content in &self.content {
            match content {
                BodyContent::Table(table) => tables.push(table),
                BodyContent::ContentControl(sdt) => sdt.collect_tables(SdtOwner::Body, &mut tables),
                BodyContent::Paragraph(_) | BodyContent::RawXml(_) => {}
            }
        }
        tables.into_iter()
    }

    /// Get a mutable iterator over only the tables.
    pub fn tables_mut(&mut self) -> impl Iterator<Item = &mut CT_Tbl> {
        self.content.iter_mut().filter_map(|c| match c {
            BodyContent::Table(t) => Some(t),
            _ => None,
        })
    }

    /// Add a paragraph to the body.
    pub fn add_paragraph(&mut self, p: CT_P) {
        self.content.push(BodyContent::Paragraph(p));
    }

    /// Add a table to the body.
    pub fn add_table(&mut self, tbl: CT_Tbl) {
        self.content.push(BodyContent::Table(tbl));
    }

    /// Get the number of body content elements (paragraphs + tables).
    pub fn content_count(&self) -> usize {
        self.content.len()
    }

    /// Insert a paragraph at the given index.
    ///
    /// Panics if `index > content_count()`.
    pub fn insert_paragraph(&mut self, index: usize, p: CT_P) {
        self.content.insert(index, BodyContent::Paragraph(p));
    }

    /// Insert a table at the given index.
    ///
    /// Panics if `index > content_count()`.
    pub fn insert_table(&mut self, index: usize, tbl: CT_Tbl) {
        self.content.insert(index, BodyContent::Table(tbl));
    }

    /// Find the index of the first paragraph whose text contains the given substring.
    pub fn find_paragraph_index(&self, text: &str) -> Option<usize> {
        self.find_paragraph_indices(text).into_iter().next()
    }

    /// Find every matching body-content index, preferring direct paragraphs.
    pub fn find_paragraph_indices(&self, text: &str) -> Vec<usize> {
        let mut direct = Vec::new();
        let mut enclosing = Vec::new();
        for (index, content) in self.content.iter().enumerate() {
            match content {
                BodyContent::Paragraph(paragraph) if paragraph.text().contains(text) => {
                    direct.push(index);
                }
                BodyContent::ContentControl(sdt) => {
                    let mut paragraphs = Vec::new();
                    sdt.collect_paragraphs(SdtOwner::Body, &mut paragraphs);
                    if paragraphs
                        .iter()
                        .any(|paragraph| paragraph.text().contains(text))
                    {
                        enclosing.push(index);
                    }
                }
                _ => {}
            }
        }
        direct.extend(enclosing);
        direct
    }

    /// Return every content control in document order, including nested controls.
    pub fn content_controls(&self) -> Vec<&CT_Sdt> {
        let mut controls = Vec::new();
        for content in &self.content {
            match content {
                BodyContent::Paragraph(paragraph) => paragraph.collect_controls(&mut controls),
                BodyContent::Table(table) => table.collect_controls(&mut controls),
                BodyContent::ContentControl(sdt) => {
                    controls.push(sdt);
                    sdt.collect_controls(SdtOwner::Body, &mut controls);
                }
                BodyContent::RawXml(_) => {}
            }
        }
        controls
    }

    /// Remove and return the content at the given index, or `None` if out of bounds.
    pub fn remove(&mut self, index: usize) -> Option<BodyContent> {
        if index < self.content.len() {
            Some(self.content.remove(index))
        } else {
            None
        }
    }

    /// Get a reference to the content at the given index.
    pub fn get(&self, index: usize) -> Option<&BodyContent> {
        self.content.get(index)
    }

    /// Get a mutable reference to the content at the given index.
    pub fn get_mut(&mut self, index: usize) -> Option<&mut BodyContent> {
        self.content.get_mut(index)
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes_and_owner_bindings(reader, &["w".to_string()], &[])
    }

    pub(crate) fn from_xml_with_prefixes_and_owner_bindings(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
    ) -> Result<Self> {
        Self::from_xml_with_prefixes_and_owner_bindings_until(
            reader,
            word_prefixes,
            owner_bindings,
            b"body",
        )
    }

    pub(crate) fn from_xml_with_prefixes_and_owner_bindings_until(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
        end_local_name: &[u8],
    ) -> Result<Self> {
        let mut content = Vec::new();
        let mut sect_pr = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"p", &prefixes) {
                        content.push(BodyContent::Paragraph(
                            CT_P::from_xml_with_prefixes_and_root(reader, &prefixes, Some(e))?,
                        ));
                    } else if is_word_element(name.as_ref(), b"tbl", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let table_bindings = merged_owner_bindings(owner_bindings, &local_bindings);
                        content.push(BodyContent::Table(
                            CT_Tbl::from_xml_with_prefixes_and_owner_bindings(
                                reader,
                                &prefixes,
                                &table_bindings,
                            )?,
                        ));
                    } else if is_word_element(name.as_ref(), b"sdt", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(sdt) = CT_Sdt::from_body_raw(&raw, &prefixes)? {
                            content.push(BodyContent::ContentControl(sdt));
                        } else {
                            content.push(BodyContent::RawXml(raw));
                        }
                    } else if is_word_element(name.as_ref(), b"sectPr", &prefixes) {
                        if sect_pr.is_none() {
                            let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                            let section_bindings =
                                merged_owner_bindings(owner_bindings, &local_bindings);
                            sect_pr =
                                Some(CT_SectPr::from_xml_with_prefixes_owner_bindings_and_root(
                                    reader,
                                    &prefixes,
                                    &section_bindings,
                                    Some(e),
                                )?);
                        } else {
                            content.push(BodyContent::RawXml(
                                crate::text::raw_with_external_bindings(
                                    &capture_element(reader, e)?,
                                    owner_bindings,
                                )?,
                            ));
                        }
                    } else {
                        // Capture unknown elements as raw XML
                        content.push(BodyContent::RawXml(
                            crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"p", &prefixes) {
                        content.push(BodyContent::Paragraph(CT_P::from_empty_root(e, &prefixes)?));
                    } else if is_word_element(name.as_ref(), b"tbl", &prefixes) {
                        content.push(BodyContent::Table(CT_Tbl::new()));
                    } else if is_word_element(name.as_ref(), b"sectPr", &prefixes) {
                        if sect_pr.is_none() {
                            sect_pr = Some(CT_SectPr::from_empty_root(e, &prefixes)?);
                        } else {
                            content.push(BodyContent::RawXml(
                                crate::text::raw_with_external_bindings(
                                    &capture_empty_element(e)?,
                                    owner_bindings,
                                )?,
                            ));
                        }
                    } else {
                        content.push(BodyContent::RawXml(
                            crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                }
                Ok(Event::End(ref e))
                    if is_word_element(e.name().as_ref(), end_local_name, word_prefixes) =>
                {
                    break;
                }
                Ok(Event::Eof) => {
                    let end = String::from_utf8_lossy(end_local_name);
                    return Err(OxmlError::MissingElement(format!("w:{end} end")));
                }
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Body { content, sect_pr })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:body")))?;

        for item in &self.content {
            match item {
                BodyContent::Paragraph(p) => p.to_xml(writer)?,
                BodyContent::Table(t) => t.to_xml(writer)?,
                BodyContent::ContentControl(sdt) => sdt.to_xml(writer)?,
                BodyContent::RawXml(raw) => {
                    writer.get_mut().write_all(raw)?;
                }
            }
        }

        if let Some(ref sect) = self.sect_pr {
            sect.to_xml(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:body")))?;
        Ok(())
    }
}

impl Default for CT_Body {
    fn default() -> Self {
        Self::new()
    }
}

/// `CT_Document` — The root document element.
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_Document {
    pub body: CT_Body,
    /// Extra namespace declarations captured from the original document element.
    /// Each entry is (prefix, uri), e.g. ("xmlns:wp14", "http://...").
    pub extra_namespaces: Vec<(String, String)>,
    /// Raw XML for `<w:background>` element if present.
    pub background_xml: Option<Vec<u8>>,
    /// Foreign same-local-name background children retained before the body.
    #[doc(hidden)]
    pub background_extra_xml: Vec<Vec<u8>>,
    /// Non-namespace attributes of the original document element, such as
    /// `mc:Ignorable`, in source order.
    #[doc(hidden)]
    pub root_attributes: Vec<(String, String)>,
}

#[allow(non_snake_case)]
impl CT_Document {
    pub fn new() -> Self {
        CT_Document {
            body: CT_Body::new(),
            extra_namespaces: Vec::new(),
            background_xml: None,
            background_extra_xml: Vec::new(),
            root_attributes: Vec::new(),
        }
    }

    /// Parse from XML bytes (the content of word/document.xml).
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        crate::content_control::validate_story_part_nesting(xml)?;
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);

        let mut body = None;
        let mut extra_namespaces = Vec::new();
        let mut background_xml = None;
        let mut background_extra_xml = Vec::new();
        let mut root_attributes = Vec::new();
        let mut buf = Vec::new();
        let mut word_prefixes = Vec::new();
        let mut document_open = false;
        let mut document_closed = false;

        // Known namespace prefixes that we always emit ourselves
        let known_ns: &[&[u8]] = &[b"xmlns:w", b"xmlns:r", b"xmlns:mc", b"xmlns"];

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, &word_prefixes)?;
                    if is_word_element(name.as_ref(), b"document", &prefixes) {
                        if document_open || document_closed {
                            return Err(OxmlError::UnexpectedElement("w:document".to_owned()));
                        }
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            let is_namespace = key.starts_with(b"xmlns:") || key == b"xmlns";
                            if is_namespace && known_ns.contains(&key) {
                                continue;
                            }
                            let key_str = std::str::from_utf8(key).unwrap_or("").to_string();
                            let val_str = attr
                                .decoded_and_normalized_value(XmlVersion::Implicit1_0, e.decoder())?
                                .into_owned();
                            if is_namespace {
                                extra_namespaces.push((key_str, val_str));
                            } else {
                                root_attributes.push((key_str, val_str));
                            }
                        }
                        document_open = true;
                        word_prefixes = prefixes;
                    } else if !document_open || document_closed {
                        return Err(OxmlError::UnexpectedElement(
                            String::from_utf8_lossy(name.as_ref()).into_owned(),
                        ));
                    } else if is_word_element(name.as_ref(), b"body", &prefixes) {
                        if !document_open || document_closed || body.is_some() {
                            return Err(OxmlError::UnexpectedElement("w:body".to_owned()));
                        }
                        let owner_bindings = local_namespace_overrides(e, &word_prefixes)?;
                        body = Some(CT_Body::from_xml_with_prefixes_and_owner_bindings(
                            &mut reader,
                            &prefixes,
                            &owner_bindings,
                        )?);
                    } else if is_word_element(name.as_ref(), b"background", &prefixes) {
                        background_xml = Some(capture_element(&mut reader, e)?);
                    } else if matches_local_name(name.as_ref(), b"background") {
                        background_extra_xml.push(capture_element(&mut reader, e)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let prefixes = word_prefixes_at(e, &word_prefixes)?;
                    if is_word_element(e.name().as_ref(), b"document", &prefixes)
                        || !document_open
                        || document_closed
                    {
                        return Err(OxmlError::UnexpectedElement(
                            String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                        ));
                    } else if is_word_element(e.name().as_ref(), b"body", &prefixes) {
                        if body.is_some() {
                            return Err(OxmlError::UnexpectedElement("w:body".to_owned()));
                        }
                        body = Some(CT_Body::new());
                    } else if is_word_element(e.name().as_ref(), b"background", &prefixes) {
                        background_xml = Some(capture_empty_element(e)?);
                    } else if matches_local_name(e.name().as_ref(), b"background") {
                        background_extra_xml.push(capture_empty_element(e)?);
                    }
                }
                Ok(Event::End(ref e))
                    if is_word_element(e.name().as_ref(), b"document", &word_prefixes) =>
                {
                    if !document_open {
                        return Err(OxmlError::UnexpectedElement("w:document end".to_owned()));
                    }
                    document_open = false;
                    document_closed = true;
                }
                Ok(Event::Text(ref text)) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                    return Err(OxmlError::UnexpectedElement(
                        "text outside w:body".to_owned(),
                    ));
                }
                Ok(Event::Eof) => {
                    if document_open || !document_closed {
                        return Err(OxmlError::MissingElement("w:document end".to_owned()));
                    }
                    break;
                }
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Document {
            body: body.ok_or_else(|| OxmlError::MissingElement("w:body".to_owned()))?,
            extra_namespaces,
            background_xml,
            background_extra_xml,
            root_attributes,
        })
    }

    /// Serialize to XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let wp_ns = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        let wp_root_is_canonical = self
            .extra_namespaces
            .iter()
            .find(|(name, _)| name == "xmlns:wp")
            .is_none_or(|(_, namespace)| namespace == wp_ns);
        let _binding_scope = root_binding_scope(
            ROOT_R_BINDING
                | ROOT_MC_BINDING
                | if wp_root_is_canonical {
                    ROOT_WP_BINDING
                } else {
                    0
                },
        );
        let mut writer = Writer::new(Vec::new());

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut doc_start = BytesStart::new("w:document");
        doc_start.push_attribute(("xmlns:w", W_NS));
        doc_start.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        doc_start.push_attribute((
            "xmlns:mc",
            "http://schemas.openxmlformats.org/markup-compatibility/2006",
        ));

        // Always emit xmlns:wp for drawing elements
        let mut has_wp = false;
        for (key, _) in &self.extra_namespaces {
            if key == "xmlns:wp" {
                has_wp = true;
                break;
            }
        }
        if !has_wp {
            doc_start.push_attribute(("xmlns:wp", wp_ns));
        }

        // Replay captured extra namespaces, then the root attributes that
        // may name their prefixes, such as `mc:Ignorable`.
        for (key, val) in self.extra_namespaces.iter().chain(&self.root_attributes) {
            doc_start.push_attribute((key.as_str(), val.as_str()));
        }

        writer.write_event(Event::Start(doc_start))?;

        // Write background element if present
        if let Some(ref bg) = self.background_xml {
            writer.get_mut().extend_from_slice(bg);
        }
        for raw in &self.background_extra_xml {
            writer.get_mut().extend_from_slice(raw);
        }

        self.body.to_xml(&mut writer)?;

        writer.write_event(Event::End(BytesEnd::new("w:document")))?;

        let mut xml = writer.into_inner();
        declare_w14_on_part_root(&mut xml)?;
        Ok(xml)
    }
}

impl Default for CT_Document {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {

    // F-X018, an unmodelled enumerated value must not fail a document open.

    /// Every enumeration reachable from document parsing, with an unmodelled
    /// but plausibly spec-valid value, plus a sibling property on the same
    /// element that must survive.
    #[test]
    fn a_document_with_an_unmodelled_enumerated_value_still_opens() {
        let cases: [(&str, &str); 8] = [
            (
                "ST_Jc",
                r#"<w:pPr><w:jc w:val="thaiDistribute"/><w:keepNext/></w:pPr>"#,
            ),
            (
                "ST_Underline",
                r#"<w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:u w:val="wavyHeavy"/><w:b/></w:rPr><w:t>x</w:t></w:r>"#,
            ),
            (
                "ST_HighlightColor",
                r#"<w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:highlight w:val="chartreuse"/><w:b/></w:rPr><w:t>x</w:t></w:r>"#,
            ),
            (
                "ST_Border",
                r#"<w:pPr><w:keepNext/><w:pBdr><w:top w:val="tribal4" w:sz="8"/></w:pBdr></w:pPr>"#,
            ),
            (
                "ST_TabJc",
                r#"<w:pPr><w:keepNext/><w:tabs><w:tab w:val="numbering" w:pos="720"/></w:tabs></w:pPr>"#,
            ),
            (
                "ST_TabLeader",
                r#"<w:pPr><w:keepNext/><w:tabs><w:tab w:val="left" w:pos="720" w:leader="middleDot"/></w:tabs></w:pPr>"#,
            ),
            (
                "ST_SectionType",
                r#"<w:pPr><w:keepNext/><w:sectPr><w:type w:val="oddPage2"/></w:sectPr></w:pPr>"#,
            ),
            (
                "ST_PageOrientation",
                r#"<w:pPr><w:keepNext/><w:sectPr><w:pgSz w:w="12240" w:h="15840" w:orient="sideways"/></w:sectPr></w:pPr>"#,
            ),
        ];

        for (name, body) in cases {
            let xml = format!(
                r#"<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p>{body}</w:p></w:body></w:document>"#
            );
            let document = CT_Document::from_xml(xml.as_bytes()).unwrap_or_else(|e| {
                panic!("{name}: an unmodelled value must not fail the open, got {e}")
            });

            let BodyContent::Paragraph(paragraph) = &document.body.content[0] else {
                panic!("{name}: expected a paragraph");
            };
            let properties = paragraph
                .properties
                .as_ref()
                .unwrap_or_else(|| panic!("{name}: properties survive"));
            assert_eq!(
                properties.keep_next,
                Some(true),
                "{name}: an unmodelled value must not cost the element its siblings"
            );
        }
    }

    #[test]
    fn an_unmodelled_value_leaves_the_property_unset() {
        // None means "not specified", so the style chain still supplies it.
        // A guessed variant would override a style that does specify one.
        let xml = br#"<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:pPr><w:jc w:val="thaiDistribute"/></w:pPr></w:p></w:body></w:document>"#;
        let document = CT_Document::from_xml(xml).expect("opens");
        let BodyContent::Paragraph(paragraph) = &document.body.content[0] else {
            panic!("expected a paragraph");
        };
        assert_eq!(paragraph.properties.as_ref().unwrap().jc, None);
    }
    use super::*;

    #[test]
    fn parses_empty_body_paragraph_and_first_section_properties_as_modeled_content() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body><w:p/><w:sectPr/><w:sectPr><w:titlePg/></w:sectPr></w:body></w:document>"#
        );

        let document = CT_Document::from_xml(xml.as_bytes()).unwrap();
        assert!(matches!(
            document.body.content.as_slice(),
            [BodyContent::Paragraph(_), BodyContent::RawXml(raw)]
                if raw.windows(b"<w:titlePg/>".len())
                    .any(|window| window == b"<w:titlePg/>")
        ));
        let section = document.body.sect_pr.expect("first section is retained");
        assert_eq!(section.title_pg, None);
    }

    #[test]
    fn start_end_section_leaves_keep_unsupported_attributes_opaque() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:r="{}" xmlns:x="urn:producer"><w:body><w:sectPr><w:headerReference r:id="rId1" x:flag="keep"></w:headerReference><w:pgSz w:w="12240" w:h="15840" x:flag="keep"></w:pgSz><w:pgMar w:top="720" x:flag="not-a-number"></w:pgMar></w:sectPr></w:body></w:document>"#,
            crate::namespace::R_NS,
        );

        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let section = document.body.sect_pr.expect("section is retained");
        assert!(section.header_refs.is_empty());
        assert_eq!(section.page_width, None);
        assert_eq!(section.margin_top, None);
        assert_eq!(section.extra_xml.len(), 3);
        assert!(section.extra_xml.iter().all(|raw| {
            raw.windows(b"x:flag=\"".len())
                .any(|window| window == b"x:flag=\"")
        }));
    }

    #[test]
    fn expanded_section_children_preserve_modeled_facts_and_relationship_namespaces() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:q="{}" xmlns:x="urn:foreign"><w:body><w:sectPr><w:headerReference w:type="default" q:id="rId1"></w:headerReference><w:footerReference w:type="even" x:id="decoy"></w:footerReference><w:pgSz w:w="12240" w:h="15840"></w:pgSz><w:pgMar w:top="720"></w:pgMar></w:sectPr></w:body></w:document>"#,
            crate::namespace::R_NS,
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let section = document.body.sect_pr.expect("section is retained");
        assert_eq!(section.page_width, Some(Twips(12240)));
        assert_eq!(section.page_height, Some(Twips(15840)));
        assert_eq!(section.margin_top, Some(Twips(720)));
        assert_eq!(section.header_refs[0].rel_id, "rId1");
        assert!(section.footer_refs.is_empty());
    }

    #[test]
    fn document_reader_rejects_truncated_multiple_and_foreign_roots() {
        let truncated = format!(r#"<w:document xmlns:w="{W_NS}"><w:body><w:p/>"#);
        assert!(CT_Document::from_xml(truncated.as_bytes()).is_err());

        let multiple_roots = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body/></w:document><w:document xmlns:w="{W_NS}"><w:body/></w:document>"#
        );
        assert!(CT_Document::from_xml(multiple_roots.as_bytes()).is_err());

        let foreign =
            format!(r#"<x:document xmlns:x="urn:foreign" xmlns:w="{W_NS}"><x:body/></x:document>"#);
        assert!(CT_Document::from_xml(foreign.as_bytes()).is_err());

        let text_outside = format!(r#"outside<w:document xmlns:w="{W_NS}"><w:body/></w:document>"#);
        assert!(CT_Document::from_xml(text_outside.as_bytes()).is_err());

        let empty_body = format!(r#"<w:document xmlns:w="{W_NS}"><w:body/></w:document>"#);
        assert_eq!(
            CT_Document::from_xml(empty_body.as_bytes())
                .expect("self-closing body parses")
                .body,
            CT_Body::new()
        );
    }

    #[test]
    fn round_trip_document() {
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        p.add_run("Hello World");
        doc.body.add_paragraph(p);

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();

        let paras: Vec<_> = parsed.body.paragraphs().collect();
        assert_eq!(paras.len(), 1);
        assert_eq!(paras[0].text(), "Hello World");
    }

    #[test]
    fn foreign_document_background_lookalikes_remain_untyped() {
        for raw in [
            r#"<ext:background ext:color="red"/>"#,
            r#"<ext:background ext:color="red"><ext:payload/></ext:background>"#,
        ] {
            let xml = format!(
                r#"<w:document xmlns:w="{W_NS}" xmlns:ext="urn:producer">{raw}<w:body><w:p/></w:body></w:document>"#
            );
            let document = CT_Document::from_xml(xml.as_bytes()).expect("document opens");
            assert!(document.background_xml.is_none());

            let written = document.to_xml().expect("document writes");
            assert!(
                written
                    .windows(raw.len())
                    .any(|window| window == raw.as_bytes()),
                "foreign background bytes survive exactly"
            );
            let reopened = CT_Document::from_xml(&written).expect("written document reopens");
            assert!(reopened.background_xml.is_none());
            assert_eq!(reopened.background_extra_xml, vec![raw.as_bytes().to_vec()]);
        }
    }

    /// #160: a typed rewrite dropped `mc:Ignorable` and every other
    /// non-namespace attribute of the document root.
    #[test]
    fn root_attributes_survive_a_rewrite_after_the_namespace_declarations() {
        let xml = format!(
            r#"<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="{W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14" xmlns:x="urn:producer" x:root="a &amp; b"><w:body><w:p/></w:body></w:document>"#
        );
        let parsed = CT_Document::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(
            parsed.root_attributes,
            [
                ("mc:Ignorable".to_owned(), "w14".to_owned()),
                ("x:root".to_owned(), "a & b".to_owned()),
            ]
        );

        let written = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        let root = &written[written.find("<w:document").unwrap()..];
        let root = &root[..=root.find('>').unwrap()];
        assert!(
            root.ends_with(
                r#" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:x="urn:producer" mc:Ignorable="w14" x:root="a &amp; b">"#
            ),
            "{root}"
        );
        let reparsed = CT_Document::from_xml(written.as_bytes()).unwrap();
        assert_eq!(reparsed.root_attributes, parsed.root_attributes);
        assert_eq!(reparsed.to_xml().unwrap(), written.as_bytes());
    }

    #[test]
    fn default_namespace_document_paragraph_properties_parse_in_scope() {
        let xml = format!(
            r#"<document xmlns="{W_NS}" xmlns:q="{W_NS}" xmlns:ext="urn:producer"><body><p><pPr><ext:jc ext:val="right"/><jc q:val="center"/></pPr><r><t>Scoped</t></r></p></body></document>"#
        );
        let parsed = CT_Document::from_xml(xml.as_bytes()).unwrap();
        let paragraph = parsed.body.paragraphs().next().unwrap();
        assert_eq!(paragraph.text(), "Scoped");
        assert_eq!(
            paragraph.properties.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn body_owner_bindings_reach_direct_table_descendants_after_redundant_redeclaration() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}">
                  <w:body xmlns:ext="urn:body-owner">
                    <w:tbl xmlns:ext="urn:body-owner">
                      <w:tblPr>
                        <w:tblBorders>
                          <ext:diagonal ext:style="kept"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
                    </w:tbl>
                  </w:body>
                </w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");

        let written = String::from_utf8(document.to_xml().expect("document writes"))
            .expect("document XML is UTF-8");
        assert!(
            written.contains(r#"<ext:diagonal ext:style="kept" xmlns:ext="urn:body-owner"/>"#),
            "body binding survives a redundant table declaration: {written}"
        );

        let reopened = CT_Document::from_xml(written.as_bytes()).expect("document reopens");
        let BodyContent::Table(table) = &reopened.body.content[0] else {
            panic!("direct table reopens as typed content");
        };
        assert_eq!(
            table
                .properties
                .as_ref()
                .and_then(|properties| properties.borders.as_ref())
                .expect("table borders reopen")
                .extra_xml,
            vec![br#"<ext:diagonal ext:style="kept" xmlns:ext="urn:body-owner"/>"#.to_vec()]
        );
    }

    #[test]
    fn body_owner_bindings_survive_body_content_control_table_projection() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}">
                  <w:body xmlns:ext="urn:body-control">
                    <w:sdt><w:sdtContent>
                      <w:tbl>
                        <w:tblPr>
                          <w:tblBorders><ext:diagonal ext:style="kept"/></w:tblBorders>
                        </w:tblPr>
                        <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
                      </w:tbl>
                    </w:sdtContent></w:sdt>
                  </w:body>
                </w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");

        let written = String::from_utf8(document.to_xml().expect("document writes"))
            .expect("document XML is UTF-8");
        assert!(
            written.contains(r#"xmlns:ext="urn:body-control""#),
            "{written}"
        );
        assert!(written.contains("<ext:diagonal"), "{written}");

        let reopened = CT_Document::from_xml(written.as_bytes()).expect("document reopens");
        assert!(
            matches!(reopened.body.content[0], BodyContent::ContentControl(_)),
            "body control remains typed"
        );
        let rewritten = String::from_utf8(reopened.to_xml().expect("document rewrites"))
            .expect("document XML is UTF-8");
        assert!(
            rewritten.contains(r#"xmlns:ext="urn:body-control""#),
            "{rewritten}"
        );
        assert!(rewritten.contains("<ext:diagonal"), "{rewritten}");
    }

    #[test]
    fn body_owner_bindings_survive_direct_raw_children() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}">
                  <w:body xmlns:ext="urn:body-raw">
                    <ext:block ext:value="kept"><ext:nested/></ext:block>
                    <ext:empty ext:value="kept"/>
                  </w:body>
                </w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");

        let written = String::from_utf8(document.to_xml().expect("document writes"))
            .expect("document XML is UTF-8");
        assert!(
            written.contains(
                r#"<ext:block ext:value="kept" xmlns:ext="urn:body-raw"><ext:nested/></ext:block>"#
            ),
            "body binding reaches a raw start child: {written}"
        );
        assert!(
            written.contains(r#"<ext:empty ext:value="kept" xmlns:ext="urn:body-raw"/>"#),
            "body binding reaches a raw empty child: {written}"
        );

        let reopened = CT_Document::from_xml(written.as_bytes()).expect("document reopens");
        let rewritten = String::from_utf8(reopened.to_xml().expect("document rewrites"))
            .expect("document XML is UTF-8");
        assert!(
            rewritten.contains(
                r#"<ext:block ext:value="kept" xmlns:ext="urn:body-raw"><ext:nested/></ext:block>"#
            ),
            "body binding survives a second serialization: {rewritten}"
        );
        assert!(
            rewritten.contains(r#"<ext:empty ext:value="kept" xmlns:ext="urn:body-raw"/>"#),
            "empty body child survives a second serialization: {rewritten}"
        );
    }

    #[test]
    fn self_closing_modeled_body_children_are_typed_by_namespace() {
        let xml = format!(
            r#"<document xmlns="{W_NS}" xmlns:q="{W_NS}" xmlns:ext="urn:producer"><body><p/><q:tbl/><ext:p ext:flag="keep"/><ext:body/><q:sectPr/></body></document>"#
        );
        let parsed = CT_Document::from_xml(xml.as_bytes()).unwrap();

        assert!(matches!(parsed.body.content[0], BodyContent::Paragraph(_)));
        assert!(matches!(parsed.body.content[1], BodyContent::Table(_)));
        let BodyContent::RawXml(raw) = &parsed.body.content[2] else {
            panic!("foreign paragraph must remain raw");
        };
        assert_eq!(raw, br#"<ext:p ext:flag="keep"/>"#);
        let BodyContent::RawXml(raw) = &parsed.body.content[3] else {
            panic!("foreign body must remain raw");
        };
        assert_eq!(raw, br#"<ext:body/>"#);
        assert!(parsed.body.sect_pr.is_some());
        assert_eq!(parsed.body.content.len(), 4);

        let written = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert!(written.contains("<w:p"), "{written}");
        assert!(written.contains("<w:tbl>"));
        assert!(written.contains(r#"<ext:p ext:flag="keep"/>"#));
        assert!(written.contains("<ext:body/>"));
        assert!(written.contains("<w:sectPr"));
        let reparsed = CT_Document::from_xml(written.as_bytes()).unwrap();
        assert_eq!(reparsed.body.content.len(), 4);
        assert!(reparsed.body.sect_pr.is_some());
    }

    #[test]
    fn round_trip_with_section() {
        let doc = CT_Document::new();
        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();
        assert!(parsed.body.sect_pr.is_some());
        let sect = parsed.body.sect_pr.unwrap();
        assert_eq!(sect.page_width, Some(Twips(12240)));
    }

    #[test]
    fn round_trip_landscape() {
        let mut doc = CT_Document::new();
        let sect = doc.body.sect_pr.as_mut().unwrap();
        sect.orientation = Some(ST_PageOrientation::Landscape);
        sect.page_width = Some(Twips(15840)); // 11"
        sect.page_height = Some(Twips(12240)); // 8.5"

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();
        let sect = parsed.body.sect_pr.unwrap();
        assert_eq!(sect.orientation, Some(ST_PageOrientation::Landscape));
        assert_eq!(sect.page_width, Some(Twips(15840)));
    }

    #[test]
    fn round_trip_columns() {
        let mut doc = CT_Document::new();
        let sect = doc.body.sect_pr.as_mut().unwrap();
        sect.columns = Some(CT_Columns {
            num: Some(2),
            space: Some(Twips(720)),
            equal_width: Some(true),
            sep: Some(true),
            columns: Vec::new(),
        });

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();
        let cols = parsed.body.sect_pr.unwrap().columns.unwrap();
        assert_eq!(cols.num, Some(2));
        assert_eq!(cols.space, Some(Twips(720)));
        assert_eq!(cols.sep, Some(true));
    }

    #[test]
    fn round_trip_section_type() {
        let mut doc = CT_Document::new();
        let sect = doc.body.sect_pr.as_mut().unwrap();
        sect.section_type = Some(ST_SectionType::Continuous);
        sect.title_pg = Some(true);

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();
        let sect = parsed.body.sect_pr.unwrap();
        assert_eq!(sect.section_type, Some(ST_SectionType::Continuous));
        assert_eq!(sect.title_pg, Some(true));
    }

    #[test]
    fn page_number_start_mutation_targets_only_the_parsed_attribute() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:x="urn:producer"><w:body><w:sectPr><q:pgNumType x:note="q:start='producer'>still producer" q:start = '3' x:tail="keep"/></w:sectPr></w:body></w:document>"#
        );
        let mut document = CT_Document::from_xml(xml.as_bytes()).unwrap();
        document
            .body
            .sect_pr
            .as_mut()
            .unwrap()
            .page_number
            .as_mut()
            .unwrap()
            .start = Some(27);

        let written = String::from_utf8(document.to_xml().unwrap()).unwrap();
        assert!(
            written.contains(r#"x:note="q:start='producer'>still producer""#),
            "{written}"
        );
        assert!(written.contains("q:start = '27'"), "{written}");
        assert!(written.contains(r#"x:tail="keep""#), "{written}");
    }

    #[test]
    fn missing_page_number_start_uses_a_prefix_bound_to_word() {
        for (page_number, expected, inherited) in [
            (
                format!(r#"<q:pgNumType xmlns:q="{W_NS}" xmlns:w="urn:producer"/>"#),
                "q:start=\"27\"",
                "",
            ),
            (
                format!(r#"<pgNumType xmlns="{W_NS}" xmlns:w="urn:producer"/>"#),
                "rdocxWord:start=\"27\"",
                "",
            ),
            (
                format!(
                    r#"<pgNumType xmlns="{W_NS}" xmlns:w="urn:producer" xmlns:x="urn:extension" x:type="rdocxWord:ProducerType"/>"#
                ),
                "rdocxWord1:start=\"27\"",
                r#" xmlns:rdocxWord="urn:producer""#,
            ),
        ] {
            let xml = format!(
                r#"<w:document xmlns:w="{W_NS}"{inherited}><w:body><w:sectPr>{page_number}</w:sectPr></w:body></w:document>"#
            );
            let mut document = CT_Document::from_xml(xml.as_bytes()).unwrap();
            document
                .body
                .sect_pr
                .as_mut()
                .unwrap()
                .page_number
                .as_mut()
                .unwrap()
                .start = Some(27);

            let written = document.to_xml().unwrap();
            let written_text = String::from_utf8(written.clone()).unwrap();
            assert!(written_text.contains(expected), "{written_text}");
            assert!(!written_text.contains(r#" w:start="27""#), "{written_text}");
            if !inherited.is_empty() {
                assert!(
                    written_text.contains(r#"xmlns:rdocxWord="urn:producer""#),
                    "{written_text}"
                );
                assert!(
                    written_text.contains(r#"x:type="rdocxWord:ProducerType""#),
                    "{written_text}"
                );
            }
            let reopened = CT_Document::from_xml(&written).unwrap();
            assert_eq!(
                reopened.body.sect_pr.unwrap().page_number.unwrap().start,
                Some(27)
            );
        }
    }

    #[test]
    fn duplicate_page_number_elements_preserve_source_order() {
        for (page_numbers, expected_start) in [
            (
                r#"<w:pgNumType x:id="raw-first"/><w:pgNumType x:id="raw-second"/>"#,
                None,
            ),
            (
                r#"<w:pgNumType w:start="12" x:id="typed-first"/><w:pgNumType w:start="27" x:id="typed-second"/>"#,
                Some(12),
            ),
        ] {
            let xml = format!(
                r#"<w:document xmlns:w="{W_NS}" xmlns:x="urn:producer"><w:body><w:sectPr>{page_numbers}</w:sectPr></w:body></w:document>"#
            );
            let document = CT_Document::from_xml(xml.as_bytes()).unwrap();
            assert_eq!(
                document
                    .body
                    .sect_pr
                    .as_ref()
                    .unwrap()
                    .page_number
                    .as_ref()
                    .unwrap()
                    .start,
                expected_start
            );
            let written = String::from_utf8(document.to_xml().unwrap()).unwrap();
            let first = written.find("first").unwrap();
            let second = written.find("second").unwrap();
            assert!(first < second, "{written}");
        }
    }

    #[test]
    fn retained_children_keep_boundaries_between_repeated_story_references() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:producer"><w:body><w:sectPr><w:headerReference w:type="default" r:id="h1"/><x:betweenHeaders/><w:headerReference w:type="first" r:id="h2"/><w:footerReference w:type="default" r:id="f1"/><x:betweenFooters/><w:footerReference w:type="even" r:id="f2"/></w:sectPr></w:body></w:document>"#
        );
        let mut document = CT_Document::from_xml(xml.as_bytes()).unwrap();
        let written = String::from_utf8(document.to_xml().unwrap()).unwrap();
        let positions = ["h1", "betweenHeaders", "h2", "f1", "betweenFooters", "f2"]
            .map(|marker| written.find(marker).unwrap());
        assert!(
            positions.windows(2).all(|pair| pair[0] < pair[1]),
            "{written}"
        );

        let mut front_removed = document.clone();
        front_removed
            .body
            .sect_pr
            .as_mut()
            .unwrap()
            .header_refs
            .remove(0);
        let front_removed = String::from_utf8(front_removed.to_xml().unwrap()).unwrap();
        assert!(
            front_removed.find("betweenHeaders").unwrap() < front_removed.find("h2").unwrap(),
            "{front_removed}"
        );

        let mut reordered = document.clone();
        reordered
            .body
            .sect_pr
            .as_mut()
            .unwrap()
            .header_refs
            .swap(0, 1);
        let reordered = String::from_utf8(reordered.to_xml().unwrap()).unwrap();
        let reordered_positions =
            ["betweenHeaders", "h2", "h1"].map(|marker| reordered.find(marker).unwrap());
        assert!(
            reordered_positions.windows(2).all(|pair| pair[0] < pair[1]),
            "{reordered}"
        );

        let section = document.body.sect_pr.as_mut().unwrap();
        section.header_refs.clear();
        section.footer_refs.clear();
        let shortened = String::from_utf8(document.to_xml().unwrap()).unwrap();
        assert!(shortened.contains("betweenHeaders"), "{shortened}");
        assert!(shortened.contains("betweenFooters"), "{shortened}");
    }

    #[test]
    fn retained_boundaries_use_occurrences_for_equal_story_references() {
        fn positions(xml: &str, reference: &str, retained: &str) -> (Vec<usize>, usize) {
            (
                xml.match_indices(reference)
                    .map(|(index, _)| index)
                    .collect(),
                xml.find(retained).unwrap(),
            )
        }

        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:producer"><w:body><w:sectPr><w:headerReference w:type="default" r:id="sameHeader"/><x:betweenEqualHeaders/><w:headerReference w:type="default" r:id="sameHeader"/><w:footerReference w:type="default" r:id="sameFooter"/><x:betweenEqualFooters/><w:footerReference w:type="default" r:id="sameFooter"/></w:sectPr></w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).unwrap();

        let unchanged = String::from_utf8(document.to_xml().unwrap()).unwrap();
        let (headers, raw_header) = positions(&unchanged, "sameHeader", "betweenEqualHeaders");
        let (footers, raw_footer) = positions(&unchanged, "sameFooter", "betweenEqualFooters");
        assert!(
            headers[0] < raw_header && raw_header < headers[1],
            "{unchanged}"
        );
        assert!(
            footers[0] < raw_footer && raw_footer < footers[1],
            "{unchanged}"
        );

        let mut front_removed = document.clone();
        front_removed
            .body
            .sect_pr
            .as_mut()
            .unwrap()
            .header_refs
            .remove(0);
        let front_removed = String::from_utf8(front_removed.to_xml().unwrap()).unwrap();
        let (headers, raw_header) = positions(&front_removed, "sameHeader", "betweenEqualHeaders");
        assert!(headers[0] < raw_header, "{front_removed}");

        let mut tail_removed = document.clone();
        tail_removed
            .body
            .sect_pr
            .as_mut()
            .unwrap()
            .footer_refs
            .remove(1);
        let tail_removed = String::from_utf8(tail_removed.to_xml().unwrap()).unwrap();
        let (footers, raw_footer) = positions(&tail_removed, "sameFooter", "betweenEqualFooters");
        assert!(footers[0] < raw_footer, "{tail_removed}");

        let mut reordered = document.clone();
        let section = reordered.body.sect_pr.as_mut().unwrap();
        section.header_refs.swap(0, 1);
        section.footer_refs.swap(0, 1);
        let reordered = String::from_utf8(reordered.to_xml().unwrap()).unwrap();
        let (headers, raw_header) = positions(&reordered, "sameHeader", "betweenEqualHeaders");
        let (footers, raw_footer) = positions(&reordered, "sameFooter", "betweenEqualFooters");
        assert!(
            headers[0] < raw_header && raw_header < headers[1],
            "{reordered}"
        );
        assert!(
            footers[0] < raw_footer && raw_footer < footers[1],
            "{reordered}"
        );
    }

    #[test]
    fn ambiguous_equal_reference_mutations_are_byte_stable() {
        fn positions(xml: &str) -> (Vec<usize>, usize) {
            (
                xml.match_indices("sameHeader")
                    .map(|(index, _)| index)
                    .collect(),
                xml.find("betweenEqualHeaders").unwrap(),
            )
        }

        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:producer"><w:body><w:sectPr><w:headerReference w:type="default" r:id="sameHeader"/><x:betweenEqualHeaders/><w:headerReference w:type="default" r:id="sameHeader"/></w:sectPr></w:body></w:document>"#
        );
        let source = CT_Document::from_xml(xml.as_bytes()).unwrap();

        let mut edited = source.clone();
        let section = edited.body.sect_pr.as_mut().unwrap();
        let _: &Vec<HdrFtrRef> = &section.header_refs;
        let old_allocations = [
            section.header_refs[0].rel_id.as_ptr(),
            section.header_refs[1].rel_id.as_ptr(),
        ];
        let replacements = section
            .header_refs
            .iter()
            .map(|reference| String::from_utf8(reference.rel_id.as_bytes().to_vec()).unwrap())
            .collect::<Vec<_>>();
        assert!(replacements.iter().all(|replacement| {
            old_allocations
                .iter()
                .all(|allocation| replacement.as_ptr() != *allocation)
        }));
        for (reference, replacement) in section.header_refs.iter_mut().zip(replacements) {
            reference.rel_id = replacement;
        }
        section.header_refs.swap(0, 1);

        let written = String::from_utf8(edited.to_xml().unwrap()).unwrap();
        let (references, retained) = positions(&written);
        assert!(
            references[0] < retained && retained < references[1],
            "{written}"
        );
        assert_eq!(
            written,
            String::from_utf8(edited.clone().to_xml().unwrap()).unwrap()
        );

        let reopened = CT_Document::from_xml(written.as_bytes()).unwrap();
        let reopened_written = String::from_utf8(reopened.to_xml().unwrap()).unwrap();
        let (references, retained) = positions(&reopened_written);
        assert!(
            references[0] < retained && retained < references[1],
            "{reopened_written}"
        );
        assert_eq!(
            reopened_written,
            String::from_utf8(reopened.clone().to_xml().unwrap()).unwrap()
        );

        let mut replaced_source = source.clone();
        let section = replaced_source.body.sect_pr.as_mut().unwrap();
        let retired = section.header_refs.remove(1);
        drop(retired);
        section.header_refs.insert(
            0,
            HdrFtrRef {
                hdr_ftr_type: HdrFtrType::Default,
                rel_id: "sameHeader".to_owned(),
            },
        );
        let replaced_source = String::from_utf8(replaced_source.to_xml().unwrap()).unwrap();
        let (references, retained) = positions(&replaced_source);
        assert!(
            references[0] < retained && retained < references[1],
            "{replaced_source}"
        );
        let reopened = CT_Document::from_xml(replaced_source.as_bytes()).unwrap();
        assert_eq!(
            replaced_source,
            String::from_utf8(reopened.to_xml().unwrap()).unwrap()
        );
    }

    #[test]
    fn insert_paragraph_at_beginning() {
        let mut body = CT_Body::new();
        let mut p1 = CT_P::new();
        p1.add_run("First");
        body.add_paragraph(p1);

        let mut p0 = CT_P::new();
        p0.add_run("Inserted");
        body.insert_paragraph(0, p0);

        assert_eq!(body.content_count(), 2);
        match &body.content[0] {
            BodyContent::Paragraph(p) => assert_eq!(p.text(), "Inserted"),
            _ => panic!("expected paragraph"),
        }
        match &body.content[1] {
            BodyContent::Paragraph(p) => assert_eq!(p.text(), "First"),
            _ => panic!("expected paragraph"),
        }
    }

    #[test]
    fn insert_paragraph_in_middle() {
        let mut body = CT_Body::new();
        let mut p1 = CT_P::new();
        p1.add_run("First");
        body.add_paragraph(p1);
        let mut p2 = CT_P::new();
        p2.add_run("Third");
        body.add_paragraph(p2);

        let mut mid = CT_P::new();
        mid.add_run("Middle");
        body.insert_paragraph(1, mid);

        assert_eq!(body.content_count(), 3);
        let texts: Vec<_> = body.paragraphs().map(|p| p.text()).collect();
        assert_eq!(texts, vec!["First", "Middle", "Third"]);
    }

    #[test]
    fn find_paragraph_index_match() {
        let mut body = CT_Body::new();
        let mut p1 = CT_P::new();
        p1.add_run("Hello World");
        body.add_paragraph(p1);
        let mut p2 = CT_P::new();
        p2.add_run("INSERT_HERE");
        body.add_paragraph(p2);

        assert_eq!(body.find_paragraph_index("INSERT_HERE"), Some(1));
        assert_eq!(body.find_paragraph_index("NONEXISTENT"), None);

        let document = CT_Document::from_xml(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sdt><w:sdtContent><w:p><w:r><w:t>Background</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>Background</w:t></w:r></w:p></w:body></w:document>"#,
        )
        .unwrap();
        assert_eq!(document.body.find_paragraph_index("Background"), Some(1));
        assert_eq!(
            document.body.find_paragraph_indices("Background"),
            vec![1, 0]
        );
    }

    #[test]
    fn remove_content() {
        let mut body = CT_Body::new();
        let mut p1 = CT_P::new();
        p1.add_run("First");
        body.add_paragraph(p1);
        let mut p2 = CT_P::new();
        p2.add_run("Second");
        body.add_paragraph(p2);

        let removed = body.remove(0);
        assert!(removed.is_some());
        assert_eq!(body.content_count(), 1);
        match &body.content[0] {
            BodyContent::Paragraph(p) => assert_eq!(p.text(), "Second"),
            _ => panic!("expected paragraph"),
        }

        // Out of bounds
        assert!(body.remove(5).is_none());
    }

    #[test]
    fn get_and_get_mut() {
        let mut body = CT_Body::new();
        let mut p = CT_P::new();
        p.add_run("Test");
        body.add_paragraph(p);

        assert!(body.get(0).is_some());
        assert!(body.get(1).is_none());

        if let Some(BodyContent::Paragraph(p)) = body.get_mut(0) {
            p.add_run(" Modified");
        }
        match body.get(0).unwrap() {
            BodyContent::Paragraph(p) => assert_eq!(p.text(), "Test Modified"),
            _ => panic!("expected paragraph"),
        }
    }

    #[test]
    fn sect_pr_section_type_and_orientation_round_trip() {
        let mut doc = CT_Document::new();
        let sect = doc.body.sect_pr.as_mut().unwrap();
        sect.section_type = Some(ST_SectionType::NextPage);
        sect.orientation = Some(ST_PageOrientation::Landscape);
        sect.page_width = Some(Twips(15840));
        sect.page_height = Some(Twips(12240));

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();
        let sect2 = parsed.body.sect_pr.unwrap();
        assert_eq!(sect2.section_type, Some(ST_SectionType::NextPage));
        assert_eq!(sect2.orientation, Some(ST_PageOrientation::Landscape));
        assert_eq!(sect2.page_width, Some(Twips(15840)));
        assert_eq!(sect2.page_height, Some(Twips(12240)));
    }

    #[test]
    fn sect_pr_all_section_types() {
        for section_type in [
            ST_SectionType::NextPage,
            ST_SectionType::Continuous,
            ST_SectionType::EvenPage,
            ST_SectionType::OddPage,
        ] {
            let mut doc = CT_Document::new();
            let sect = doc.body.sect_pr.as_mut().unwrap();
            sect.section_type = Some(section_type);

            let xml = doc.to_xml().unwrap();
            let parsed = CT_Document::from_xml(&xml).unwrap();
            let sect2 = parsed.body.sect_pr.unwrap();
            assert_eq!(
                sect2.section_type,
                Some(section_type),
                "section type round-trip failed for {section_type:?}"
            );
        }
    }

    #[test]
    fn sect_pr_in_paragraph_ppr_round_trip() {
        // Section breaks inside paragraph properties (pPr/sectPr)
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        p.add_run("Section break paragraph");
        let mut ppr = crate::properties::CT_PPr::default();
        let mut sect = CT_SectPr::default_letter();
        sect.section_type = Some(ST_SectionType::NextPage);
        sect.orientation = Some(ST_PageOrientation::Landscape);
        sect.page_width = Some(Twips(15840));
        sect.page_height = Some(Twips(12240));
        ppr.sect_pr = Some(sect);
        p.properties = Some(ppr);
        doc.body.add_paragraph(p);

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();

        let paras: Vec<_> = parsed.body.paragraphs().collect();
        assert_eq!(paras.len(), 1);
        let ppr2 = paras[0].properties.as_ref().unwrap();
        let sect2 = ppr2.sect_pr.as_ref().unwrap();
        assert_eq!(sect2.section_type, Some(ST_SectionType::NextPage));
        assert_eq!(sect2.orientation, Some(ST_PageOrientation::Landscape));
    }
}
