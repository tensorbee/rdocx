//! Block-level layout: paragraphs and tables as positioned blocks.

use crate::table::TableBlock;
use oxml_layout::{
    Align, Color, FontMetrics, GroupElement, InlineItem, LayoutError, LayoutLine, LineBreakParams,
    MediaId, PositionedElement, SourceNodeId, StructureId, TextDirection,
};
use rdocx_oxml::borders::CT_PBdr;
use rdocx_oxml::drawing::{
    AnchorAlignH, AnchorAlignV, ST_RelativeFromH, ST_RelativeFromV, WrapType,
};
use std::ops::Deref;
use std::sync::Arc;

/// A floating drawing anchored to a paragraph.
///
/// Offsets are kept in points alongside the frame they are measured from. A
/// `wp:anchor` offset is meaningless on its own: the same number means a
/// different place depending on whether it is relative to the page, the
/// margin, the text column or the paragraph.
#[derive(Debug, Clone)]
pub struct AnchoredDrawing {
    /// Rotation of a modeled shape around its own physical center.
    pub rotation_degrees: f64,
    /// Render underneath the text rather than on top of it.
    pub behind_doc: bool,
    /// Frame the horizontal offset is measured from.
    pub rel_h: ST_RelativeFromH,
    /// Horizontal offset in points.
    pub off_h: f64,
    /// Frame the vertical offset is measured from.
    pub rel_v: ST_RelativeFromV,
    /// Vertical offset in points.
    pub off_v: f64,
    /// Width in points.
    pub width: f64,
    /// Height in points.
    pub height: f64,
    /// How text flows around the drawing.
    pub wrap: WrapType,
    /// Space kept between the drawing and the text wrapping around it, in
    /// points.
    pub dist_top: f64,
    pub dist_bottom: f64,
    pub dist_left: f64,
    pub dist_right: f64,
    /// Horizontal alignment, used instead of the offset when present.
    pub align_h: Option<AnchorAlignH>,
    /// Vertical alignment, used instead of the offset when present.
    pub align_v: Option<AnchorAlignV>,
    /// What the drawing actually holds.
    pub content: AnchoredContent,
    /// Source description for an informative drawing.
    pub alternate_text: Option<String>,
    /// Logical figure node allocated before pagination.
    pub structure_id: Option<StructureId>,
}

/// The drawable content of an anchored drawing.
#[derive(Debug, Clone)]
pub enum AnchoredContent {
    /// A picture resolved to its content-addressed shared media identity.
    Image { media_id: MediaId },
    /// A backend-neutral group rendered in child-local chart coordinates.
    Group(GroupElement),
    /// A shape: preset geometry, an optional fill, and optional text.
    ///
    /// The text arrives already laid out, because breaking it into lines needs
    /// a font manager and that only exists in the engine.
    Shape {
        /// Preset geometry we recognise.
        preset: ShapePreset,
        /// Fill colour, or `None` for `a:noFill` and for fills we cannot
        /// resolve. An unfilled shape draws no body, only its text.
        fill: Option<Color>,
        /// Laid-out paragraphs of the shape's text box.
        text: Vec<ParagraphBlock>,
    },
}

/// The preset geometries we can draw.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapePreset {
    /// A rectangle, drawn as a filled box.
    Rect,
    /// A straight line, drawn along the top edge of the extent.
    Line,
    /// Anything else. The body is not drawn, but text still is.
    Unsupported,
}

impl ShapePreset {
    /// Map an `a:prstGeom@prst` value onto what we can draw.
    pub fn from_prst(prst: Option<&str>) -> Self {
        match prst {
            Some("rect") => ShapePreset::Rect,
            Some("line") | Some("straightConnector1") => ShapePreset::Line,
            _ => ShapePreset::Unsupported,
        }
    }
}

/// A laid-out block element (paragraph or table).
#[derive(Debug, Clone)]
pub enum LayoutBlock {
    Paragraph(ParagraphBlock),
    Table(TableBlock),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) struct WordFieldOccurrence {
    pub source: oxml_layout::FieldSource,
    pub owner_id: Option<u64>,
    pub paint_scope: u64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum WordAnnotationKind {
    EastAsian(rdocx_oxml::properties::CT_EastAsianLayout),
    Emphasis(rdocx_oxml::properties::ST_Em),
}

/// Resolved producer values use exact bits, with no style lookup after paint.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WordAnnotationStyle {
    pub font_id: oxml_layout::FontId,
    pub font_size: u64,
    pub color: [u64; 4],
    pub bold: bool,
    pub italic: bool,
    pub baseline_offset: u64,
    pub spacing: u64,
    pub underline: Option<oxml_layout::Underline>,
    pub strike: bool,
    pub dstrike: bool,
    pub highlight: Option<[u64; 4]>,
    pub font_family: Option<String>,
    pub kern: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WordAnnotationRecipe {
    pub occurrence: WordFieldOccurrence,
    pub field_kind: Option<oxml_layout::FieldKind>,
    pub source_span: Option<oxml_layout::SourceSpan>,
    pub kind: WordAnnotationKind,
    pub fallback_emphasis: Option<rdocx_oxml::properties::ST_Em>,
    pub style: WordAnnotationStyle,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum WordAnnotationRootKind {
    Group,
    EmptySpan,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WordInputAnnotation {
    pub slot: usize,
    pub recipe: Arc<WordAnnotationRecipe>,
    pub policies: PaintPolicies,
    pub kind: WordAnnotationRootKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WordLineAnnotation {
    pub line: usize,
    pub item: usize,
    pub input: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct PaintAnnotation {
    path: Vec<usize>,
    recipe: Arc<WordAnnotationRecipe>,
    kind: WordAnnotationRootKind,
    base_count: usize,
    placement: [u64; 4],
}

impl WordAnnotationRecipe {
    fn retained_bytes(&self) -> Option<usize> {
        let mut bytes = std::mem::size_of::<Self>()
            .checked_add(2 * std::mem::size_of::<usize>())?
            .checked_add(self.style.font_family.as_ref().map_or(0, String::capacity))?;
        match &self.kind {
            WordAnnotationKind::EastAsian(layout) => {
                bytes = bytes
                    .checked_add(layout.combine_brackets.as_ref().map_or(0, String::capacity))?
                    .checked_add(
                        layout
                            .extra_attributes
                            .capacity()
                            .checked_mul(std::mem::size_of::<(String, String)>())?,
                    )?;
                for (name, value) in &layout.extra_attributes {
                    bytes = bytes
                        .checked_add(name.capacity())?
                        .checked_add(value.capacity())?;
                }
            }
            WordAnnotationKind::Emphasis(rdocx_oxml::properties::ST_Em::Other(value)) => {
                bytes = bytes.checked_add(value.capacity())?
            }
            _ => {}
        }
        if let Some(rdocx_oxml::properties::ST_Em::Other(value)) = &self.fallback_emphasis {
            bytes = bytes.checked_add(value.capacity())?;
        }
        Some(bytes)
    }
}

/// Private source facts retained alongside the exact Word paragraph occurrence.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct WordParagraphPaint {
    pub rich_layout: bool,
    pub input_annotations: Vec<WordInputAnnotation>,
    pub line_annotations: Vec<WordLineAnnotation>,
    pub field_occurrences: Vec<(usize, WordFieldOccurrence)>,
    pub anchors: Vec<WordAnchorPaint>,
    pub rows: Vec<WordRowPaint>,
    pub plain_field_policy: Option<(bool, bool)>,
    pub structural_count: usize,
    pub legacy_cache: Option<WordLegacyCache>,
}

impl WordParagraphPaint {
    pub(crate) fn bind_annotation_origins(
        &mut self,
        lines: &[LayoutLine],
        origins: &[(usize, usize, usize)],
    ) -> oxml_layout::Result<()> {
        if self
            .input_annotations
            .windows(2)
            .any(|pair| pair[0].slot >= pair[1].slot)
        {
            return Err(paint_error());
        }
        let mut bound = Vec::new();
        bound
            .try_reserve_exact(self.input_annotations.len())
            .map_err(|_| paint_error())?;
        for &(line, item, slot) in origins {
            if let Ok(input) = self
                .input_annotations
                .binary_search_by_key(&slot, |annotation| annotation.slot)
            {
                if !matches!(
                    lines.get(line).and_then(|line| line.items.get(item)),
                    Some(oxml_layout::LineItem::Group { .. })
                ) {
                    return Err(paint_error());
                }
                bound.push(WordLineAnnotation { line, item, input });
            }
        }
        if bound.len() != self.input_annotations.len() {
            return Err(paint_error());
        }
        self.line_annotations = bound;
        Ok(())
    }

    pub(crate) fn annotation_at(
        &self,
        line: usize,
        item: usize,
    ) -> oxml_layout::Result<Option<&WordInputAnnotation>> {
        let Some(bound) = self
            .line_annotations
            .iter()
            .find(|bound| bound.line == line && bound.item == item)
        else {
            return Ok(None);
        };
        self.input_annotations
            .get(bound.input)
            .map(Some)
            .ok_or_else(paint_error)
    }

    pub(crate) fn for_line_range(
        &self,
        lines: &[LayoutLine],
        range: std::ops::Range<usize>,
    ) -> oxml_layout::Result<Self> {
        if range.start > range.end || range.end > lines.len() {
            return Err(paint_error());
        }
        let mut selected = self.clone();
        selected
            .line_annotations
            .retain(|bound| range.contains(&bound.line));
        for bound in &mut selected.line_annotations {
            bound.line -= range.start;
        }
        selected.rows.retain(|row| range.contains(&row.line));
        for row in &mut selected.rows {
            row.line -= range.start;
        }
        selected.structural_count = lines[range]
            .iter()
            .flat_map(|line| &line.items)
            .filter(|item| match item {
                oxml_layout::LineItem::Text(base) | oxml_layout::LineItem::Marker(base) => {
                    base.field_kind.is_some() || base.field_source.is_some()
                }
                _ => false,
            })
            .count();
        Ok(selected)
    }

    pub(crate) fn retained_bytes(&self) -> Option<usize> {
        let mut bytes = std::mem::size_of::<Self>()
            .checked_add(
                self.anchors
                    .capacity()
                    .checked_mul(std::mem::size_of::<WordAnchorPaint>())?,
            )?
            .checked_add(
                self.rows
                    .capacity()
                    .checked_mul(std::mem::size_of::<WordRowPaint>())?,
            )?;
        bytes = bytes
            .checked_add(
                self.input_annotations
                    .capacity()
                    .checked_mul(std::mem::size_of::<WordInputAnnotation>())?,
            )?
            .checked_add(
                self.line_annotations
                    .capacity()
                    .checked_mul(std::mem::size_of::<WordLineAnnotation>())?,
            )?
            .checked_add(
                self.field_occurrences
                    .capacity()
                    .checked_mul(std::mem::size_of::<(usize, WordFieldOccurrence)>())?,
            )?;
        for annotation in &self.input_annotations {
            bytes = bytes
                .checked_add(annotation.recipe.retained_bytes()?)?
                .checked_add(
                    annotation
                        .policies
                        .retained_bytes()?
                        .checked_sub(std::mem::size_of::<PaintPolicies>())?,
                )?;
        }
        if let Some(proof) = &self.legacy_cache {
            bytes = bytes.checked_add(
                proof
                    .producers
                    .capacity()
                    .checked_mul(std::mem::size_of::<WordLegacyProducer>())?,
            )?;
            for producer in &proof.producers {
                bytes =
                    bytes.checked_add(producer.language.as_ref().map_or(0, String::capacity))?;
            }
        }
        for row in &self.rows {
            bytes = bytes.checked_add(
                row.policies
                    .retained_bytes()?
                    .checked_sub(std::mem::size_of::<PaintPolicies>())?,
            )?;
        }
        for anchor in &self.anchors {
            match anchor {
                WordAnchorPaint::Image => {}
                WordAnchorPaint::Group { policies } => {
                    bytes = bytes.checked_add(
                        policies
                            .retained_bytes()?
                            .checked_sub(std::mem::size_of::<PaintPolicies>())?,
                    )?
                }
                WordAnchorPaint::Shape { paragraphs } => {
                    bytes = bytes.checked_add(
                        paragraphs
                            .capacity()
                            .checked_mul(std::mem::size_of::<WordParagraphPaint>())?,
                    )?;
                    for paragraph in paragraphs {
                        bytes = bytes.checked_add(
                            paragraph
                                .retained_bytes()?
                                .checked_sub(std::mem::size_of::<WordParagraphPaint>())?,
                        )?;
                    }
                }
            }
        }
        Some(bytes)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WordLegacyCache {
    pub language_bytes_per_span: usize,
    pub producers: Vec<WordLegacyProducer>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WordLegacyProducer {
    pub kern: bool,
    pub language: Option<String>,
    pub structural: bool,
    pub atomic_marker: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum WordAnchorPaint {
    Image,
    Group { policies: PaintPolicies },
    Shape { paragraphs: Vec<WordParagraphPaint> },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WordRowPaint {
    pub line: usize,
    pub item: usize,
    pub behind_end: usize,
    pub flow_end: usize,
    pub children_len: usize,
    pub policies: PaintPolicies,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct PaintEntry {
    path: Vec<usize>,
    policy: Option<(bool, bool)>,
    occurrence: Option<WordFieldOccurrence>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct PaintPolicies {
    entries: Vec<PaintEntry>,
    annotations: Vec<PaintAnnotation>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct PaintedElements {
    elements: Vec<PositionedElement>,
    policies: PaintPolicies,
}

#[derive(Clone, Copy)]
pub(crate) enum FieldPaintOwner {
    Word {
        kern: bool,
        legacy: bool,
        occurrence: Option<WordFieldOccurrence>,
    },
    Neutral,
}

fn copy_paint_path(path: &[usize]) -> oxml_layout::Result<Vec<usize>> {
    let mut copied = Vec::new();
    copied
        .try_reserve_exact(path.len())
        .map_err(|_| paint_error())?;
    copied.extend_from_slice(path);
    Ok(copied)
}

fn paint_at_path<'a>(
    mut elements: &'a [PositionedElement],
    path: &[usize],
) -> Option<&'a PositionedElement> {
    let (last, parents) = path.split_last()?;
    for index in parents {
        elements = paint_children(elements.get(*index)?)?;
    }
    elements.get(*last)
}

fn paint_error() -> LayoutError {
    LayoutError::Layout("inconsistent positioned field paint ownership".to_owned())
}

fn paint_children(element: &PositionedElement) -> Option<&[PositionedElement]> {
    match element {
        PositionedElement::Group(group) => Some(&group.children),
        PositionedElement::MarkedContent { children, .. } => Some(children),
        _ => None,
    }
}

fn paint_children_mut(element: &mut PositionedElement) -> Option<&mut Vec<PositionedElement>> {
    match element {
        PositionedElement::Group(group) => Some(&mut group.children),
        PositionedElement::MarkedContent { children, .. } => Some(children),
        _ => None,
    }
}

fn paint_field(element: &PositionedElement) -> bool {
    match element {
        PositionedElement::Text(run) => run.field_kind.is_some() || run.field_source.is_some(),
        PositionedElement::MultilingualText(run) => {
            run.field_kind.is_some() || run.field_source.is_some()
        }
        _ => false,
    }
}

impl PaintPolicies {
    pub(crate) fn validate(&self, elements: &[PositionedElement]) -> oxml_layout::Result<()> {
        let mut stack = Vec::new();
        stack.try_reserve(1).map_err(|_| paint_error())?;
        stack.push((elements, 0usize));
        let mut path = Vec::new();
        let mut cursor = 0;
        while let Some((forest, next)) = stack.last_mut() {
            if *next == forest.len() {
                stack.pop();
                if !stack.is_empty() {
                    path.pop();
                }
                continue;
            }
            let index = *next;
            *next += 1;
            let element = &forest[index];
            path.try_reserve(1).map_err(|_| paint_error())?;
            path.push(index);
            if paint_field(element) {
                if self
                    .entries
                    .get(cursor)
                    .is_none_or(|entry| entry.path != path)
                {
                    return Err(paint_error());
                }
                cursor += 1;
            }
            if let Some(children) = paint_children(element) {
                stack.try_reserve(1).map_err(|_| paint_error())?;
                stack.push((children, 0));
            } else {
                path.pop();
            }
        }
        if cursor != self.entries.len() {
            return Err(paint_error());
        }
        self.validate_annotations(elements)
    }

    pub(crate) fn field_occurrence(
        &self,
        path: &[usize],
    ) -> oxml_layout::Result<Option<WordFieldOccurrence>> {
        let index = self
            .entries
            .binary_search_by(|entry| entry.path.as_slice().cmp(path))
            .map_err(|_| paint_error())?;
        Ok(self.entries[index].occurrence)
    }

    pub(crate) fn field_count(&self) -> usize {
        self.entries.len()
    }

    pub(crate) fn annotation(
        &self,
        path: &[usize],
    ) -> Option<(Arc<WordAnnotationRecipe>, WordAnnotationRootKind, [u64; 4])> {
        self.annotations
            .binary_search_by(|annotation| annotation.path.as_slice().cmp(path))
            .ok()
            .map(|index| {
                let annotation = &self.annotations[index];
                (
                    Arc::clone(&annotation.recipe),
                    annotation.kind,
                    annotation.placement,
                )
            })
    }

    pub(crate) fn annotation_occurrences(&self) -> oxml_layout::Result<Vec<WordFieldOccurrence>> {
        let mut occurrences = Vec::new();
        occurrences
            .try_reserve_exact(self.annotations.len())
            .map_err(|_| paint_error())?;
        occurrences.extend(
            self.annotations
                .iter()
                .map(|annotation| annotation.recipe.occurrence),
        );
        Ok(occurrences)
    }

    pub(crate) fn attach_annotation(
        &mut self,
        elements: &[PositionedElement],
        path: &[usize],
        recipe: Arc<WordAnnotationRecipe>,
        kind: WordAnnotationRootKind,
        base_count: usize,
        placement: [u64; 4],
    ) -> oxml_layout::Result<()> {
        self.validate(elements)?;
        let annotation = PaintAnnotation {
            path: copy_paint_path(path)?,
            recipe,
            kind,
            base_count,
            placement,
        };
        self.annotations.try_reserve(1).map_err(|_| paint_error())?;
        let index = self
            .annotations
            .binary_search_by(|annotation| annotation.path.as_slice().cmp(path))
            .err()
            .ok_or_else(paint_error)?;
        self.annotations.insert(index, annotation);
        if let Err(error) = self.validate_annotations(elements) {
            self.annotations.remove(index);
            return Err(error);
        }
        Ok(())
    }

    pub(crate) fn retire_annotations(&mut self) {
        self.annotations = Vec::new();
    }

    fn validate_annotations(&self, elements: &[PositionedElement]) -> oxml_layout::Result<()> {
        if self
            .annotations
            .windows(2)
            .any(|pair| pair[1].path.starts_with(&pair[0].path) || pair[0].path >= pair[1].path)
        {
            return Err(paint_error());
        }
        for annotation in &self.annotations {
            let root = paint_at_path(elements, &annotation.path).ok_or_else(paint_error)?;
            if !matches!(root, PositionedElement::Group(_)) || annotation.base_count == 0 {
                return Err(paint_error());
            }
            let mut count = 0usize;
            for entry in self
                .entries
                .iter()
                .filter(|entry| entry.path.starts_with(&annotation.path))
            {
                let leaf = paint_at_path(elements, &entry.path).ok_or_else(paint_error)?;
                let (source, kind, empty) = match leaf {
                    PositionedElement::Text(run) => (
                        run.field_source,
                        run.field_kind,
                        run.text.is_empty() && run.glyph_ids.is_empty(),
                    ),
                    PositionedElement::MultilingualText(run) => (
                        run.field_source,
                        run.field_kind,
                        run.logical_text.is_empty() && run.glyph_ids.is_empty(),
                    ),
                    _ => return Err(paint_error()),
                };
                if source != Some(annotation.recipe.occurrence.source)
                    || kind != annotation.recipe.field_kind
                    || entry.policy != Some((annotation.recipe.style.kern, true))
                    || entry.occurrence != Some(annotation.recipe.occurrence)
                    || annotation.kind == WordAnnotationRootKind::EmptySpan && !empty
                {
                    return Err(paint_error());
                }
                count = count.checked_add(1).ok_or_else(paint_error)?;
            }
            if count != annotation.base_count
                || annotation
                    .placement
                    .iter()
                    .any(|bits| !f64::from_bits(*bits).is_finite())
            {
                return Err(paint_error());
            }
        }
        Ok(())
    }

    pub(crate) fn translate_annotation_roots(
        &mut self,
        elements: &[PositionedElement],
        dx: f64,
        dy: f64,
    ) -> oxml_layout::Result<()> {
        self.validate(elements)?;
        for annotation in &mut self.annotations {
            let nested_group = (1..annotation.path.len()).any(|length| {
                matches!(
                    paint_at_path(elements, &annotation.path[..length]),
                    Some(PositionedElement::Group(_))
                )
            });
            if !nested_group {
                annotation.placement[0] = (f64::from_bits(annotation.placement[0]) + dx).to_bits();
                annotation.placement[1] = (f64::from_bits(annotation.placement[1]) + dy).to_bits();
                annotation.placement[2] = (f64::from_bits(annotation.placement[2]) + dy).to_bits();
            }
        }
        self.validate_annotations(elements)
    }

    pub(crate) fn required(&self, path: &[usize]) -> oxml_layout::Result<Option<(bool, bool)>> {
        let index = self
            .entries
            .binary_search_by(|entry| entry.path.as_slice().cmp(path))
            .map_err(|_| paint_error())?;
        Ok(self.entries[index].policy)
    }

    pub(crate) fn retained_bytes(&self) -> Option<usize> {
        let mut bytes = std::mem::size_of::<Self>()
            .checked_add(
                self.entries
                    .capacity()
                    .checked_mul(std::mem::size_of::<PaintEntry>())?,
            )?
            .checked_add(
                self.annotations
                    .capacity()
                    .checked_mul(std::mem::size_of::<PaintAnnotation>())?,
            )?;
        for entry in &self.entries {
            bytes = bytes.checked_add(
                entry
                    .path
                    .capacity()
                    .checked_mul(std::mem::size_of::<usize>())?,
            )?;
        }
        for annotation in &self.annotations {
            bytes = bytes
                .checked_add(
                    annotation
                        .path
                        .capacity()
                        .checked_mul(std::mem::size_of::<usize>())?,
                )?
                .checked_add(annotation.recipe.retained_bytes()?)?;
        }
        Some(bytes)
    }
}

impl PaintedElements {
    pub(crate) fn from_parts(
        elements: Vec<PositionedElement>,
        policies: PaintPolicies,
    ) -> oxml_layout::Result<Self> {
        policies.validate(&elements)?;
        Ok(Self { elements, policies })
    }

    pub(crate) fn into_parts(self) -> oxml_layout::Result<(Vec<PositionedElement>, PaintPolicies)> {
        self.policies.validate(&self.elements)?;
        Ok((self.elements, self.policies))
    }

    pub(crate) fn non_field(elements: Vec<PositionedElement>) -> oxml_layout::Result<Self> {
        Self::from_parts(elements, PaintPolicies::default())
    }

    /// Only the public neutral compatibility boundary may identify neutral fields.
    pub(crate) fn neutral(elements: Vec<PositionedElement>) -> oxml_layout::Result<Self> {
        fn collect(
            elements: &[PositionedElement],
            path: &mut Vec<usize>,
            entries: &mut Vec<PaintEntry>,
        ) -> oxml_layout::Result<()> {
            for (index, element) in elements.iter().enumerate() {
                path.try_reserve(1).map_err(|_| paint_error())?;
                path.push(index);
                if paint_field(element) {
                    entries.try_reserve(1).map_err(|_| paint_error())?;
                    entries.push(PaintEntry {
                        path: copy_paint_path(path)?,
                        policy: None,
                        occurrence: None,
                    });
                }
                if let Some(children) = paint_children(element) {
                    collect(children, path, entries)?;
                }
                path.pop();
            }
            Ok(())
        }
        let mut policies = PaintPolicies::default();
        collect(&elements, &mut Vec::new(), &mut policies.entries)?;
        Self::from_parts(elements, policies)
    }

    pub(crate) fn push_non_field(&mut self, element: PositionedElement) -> oxml_layout::Result<()> {
        PaintPolicies::default().validate(std::slice::from_ref(&element))?;
        self.elements.try_reserve(1).map_err(|_| paint_error())?;
        self.elements.push(element);
        Ok(())
    }

    pub(crate) fn push_field(
        &mut self,
        element: PositionedElement,
        owner: FieldPaintOwner,
    ) -> oxml_layout::Result<()> {
        if !paint_field(&element) {
            return Err(paint_error());
        }
        let mut path = Vec::new();
        path.try_reserve_exact(1).map_err(|_| paint_error())?;
        path.push(self.elements.len());
        self.elements.try_reserve(1).map_err(|_| paint_error())?;
        self.policies
            .entries
            .try_reserve(1)
            .map_err(|_| paint_error())?;
        let (policy, occurrence) = match owner {
            FieldPaintOwner::Word {
                kern,
                legacy,
                occurrence,
            } => (Some((kern, legacy)), occurrence),
            FieldPaintOwner::Neutral => (None, None),
        };
        self.elements.push(element);
        self.policies.entries.push(PaintEntry {
            path,
            policy,
            occurrence,
        });
        Ok(())
    }

    pub(crate) fn len(&self) -> usize {
        self.elements.len()
    }
    pub(crate) fn elements(&self) -> &[PositionedElement] {
        &self.elements
    }
    pub(crate) fn coordinate_elements_mut(&mut self) -> &mut [PositionedElement] {
        &mut self.elements
    }
    pub(crate) fn policies(&self) -> &PaintPolicies {
        &self.policies
    }

    /// Checked child-forest surgery. All fallible allocation and path work precedes moving glyph storage.
    pub(crate) fn splice_at(
        &mut self,
        parent: &[usize],
        range: std::ops::Range<usize>,
        replacement: &mut Self,
    ) -> oxml_layout::Result<Self> {
        self.policies.validate(&self.elements)?;
        replacement.policies.validate(&replacement.elements)?;
        let mut children = self.elements.as_slice();
        for index in parent {
            children = children
                .get(*index)
                .and_then(paint_children)
                .ok_or_else(paint_error)?;
        }
        if range.start > range.end || range.end > children.len() {
            return Err(paint_error());
        }
        if self
            .policies
            .annotations
            .iter()
            .any(|annotation| parent.starts_with(&annotation.path))
        {
            return Err(paint_error());
        }
        let removed_len = range.end - range.start;
        let replacement_len = replacement.len();
        let final_len = children
            .len()
            .checked_sub(removed_len)
            .and_then(|n| n.checked_add(replacement_len))
            .ok_or_else(paint_error)?;
        let mut retained = Vec::new();
        let mut removed = Vec::new();
        retained
            .try_reserve(
                self.policies
                    .entries
                    .len()
                    .checked_add(replacement.policies.entries.len())
                    .ok_or_else(paint_error)?,
            )
            .map_err(|_| paint_error())?;
        removed
            .try_reserve(self.policies.entries.len())
            .map_err(|_| paint_error())?;
        for entry in &self.policies.entries {
            let mut next = PaintEntry {
                path: copy_paint_path(&entry.path)?,
                policy: entry.policy,
                occurrence: entry.occurrence,
            };
            if entry.path.starts_with(parent) && entry.path.len() > parent.len() {
                let slot = entry.path[parent.len()];
                if range.contains(&slot) {
                    next.path.drain(..parent.len());
                    next.path[0] -= range.start;
                    removed.push(next);
                    continue;
                }
                if slot >= range.end {
                    next.path[parent.len()] = slot
                        .checked_sub(removed_len)
                        .and_then(|n| n.checked_add(replacement_len))
                        .ok_or_else(paint_error)?;
                }
            }
            retained.push(next);
        }
        for entry in &replacement.policies.entries {
            let mut path = Vec::new();
            path.try_reserve_exact(
                parent
                    .len()
                    .checked_add(entry.path.len())
                    .ok_or_else(paint_error)?,
            )
            .map_err(|_| paint_error())?;
            path.extend_from_slice(parent);
            path.push(
                entry.path[0]
                    .checked_add(range.start)
                    .ok_or_else(paint_error)?,
            );
            path.extend_from_slice(&entry.path[1..]);
            retained.push(PaintEntry {
                path,
                policy: entry.policy,
                occurrence: entry.occurrence,
            });
        }
        retained.sort_unstable_by(|a, b| a.path.cmp(&b.path));
        if retained.windows(2).any(|pair| pair[0].path >= pair[1].path) {
            return Err(paint_error());
        }
        let mut retained_annotations = Vec::new();
        let mut removed_annotations = Vec::new();
        retained_annotations
            .try_reserve(
                self.policies
                    .annotations
                    .len()
                    .checked_add(replacement.policies.annotations.len())
                    .ok_or_else(paint_error)?,
            )
            .map_err(|_| paint_error())?;
        removed_annotations
            .try_reserve(self.policies.annotations.len())
            .map_err(|_| paint_error())?;
        for annotation in &self.policies.annotations {
            let mut next = PaintAnnotation {
                path: copy_paint_path(&annotation.path)?,
                recipe: Arc::clone(&annotation.recipe),
                kind: annotation.kind,
                base_count: annotation.base_count,
                placement: annotation.placement,
            };
            if annotation.path.starts_with(parent) && annotation.path.len() > parent.len() {
                let slot = annotation.path[parent.len()];
                if range.contains(&slot) {
                    next.path.drain(..parent.len());
                    next.path[0] -= range.start;
                    removed_annotations.push(next);
                    continue;
                }
                if slot >= range.end {
                    next.path[parent.len()] = slot
                        .checked_sub(removed_len)
                        .and_then(|n| n.checked_add(replacement_len))
                        .ok_or_else(paint_error)?;
                }
            }
            retained_annotations.push(next);
        }
        for annotation in &replacement.policies.annotations {
            let mut path = Vec::new();
            path.try_reserve_exact(
                parent
                    .len()
                    .checked_add(annotation.path.len())
                    .ok_or_else(paint_error)?,
            )
            .map_err(|_| paint_error())?;
            path.extend_from_slice(parent);
            path.push(
                annotation.path[0]
                    .checked_add(range.start)
                    .ok_or_else(paint_error)?,
            );
            path.extend_from_slice(&annotation.path[1..]);
            retained_annotations.push(PaintAnnotation {
                path,
                recipe: Arc::clone(&annotation.recipe),
                kind: annotation.kind,
                base_count: annotation.base_count,
                placement: annotation.placement,
            });
        }
        retained_annotations.sort_unstable_by(|a, b| a.path.cmp(&b.path));
        let mut final_elements = Vec::new();
        let mut removed_elements = Vec::new();
        final_elements
            .try_reserve_exact(final_len)
            .map_err(|_| paint_error())?;
        removed_elements
            .try_reserve_exact(removed_len)
            .map_err(|_| paint_error())?;
        let mut destination = &mut self.elements;
        for index in parent {
            destination =
                paint_children_mut(&mut destination[*index]).expect("validated child forest");
        }
        for (index, element) in std::mem::take(destination).into_iter().enumerate() {
            if index == range.start {
                final_elements.append(&mut replacement.elements);
            }
            if range.contains(&index) {
                removed_elements.push(element);
            } else {
                final_elements.push(element);
            }
        }
        if range.start == children_len_before(final_len, removed_len, replacement_len) {
            final_elements.append(&mut replacement.elements);
        }
        *destination = final_elements;
        self.policies.entries = retained;
        self.policies.annotations = retained_annotations;
        replacement.policies.entries.clear();
        replacement.policies.annotations.clear();
        Ok(Self {
            elements: removed_elements,
            policies: PaintPolicies {
                entries: removed,
                annotations: removed_annotations,
            },
        })
    }

    pub(crate) fn append(&mut self, other: &mut Self) -> oxml_layout::Result<()> {
        self.splice_at(&[], self.len()..self.len(), other)?;
        Ok(())
    }

    pub(crate) fn split_off(&mut self, index: usize) -> oxml_layout::Result<Self> {
        self.splice_at(&[], index..self.len(), &mut Self::default())
    }

    pub(crate) fn permute_at(
        &mut self,
        parent: &[usize],
        positions: &[usize],
        old_for_destination: &[usize],
    ) -> oxml_layout::Result<()> {
        self.policies.validate(&self.elements)?;
        if self
            .policies
            .annotations
            .iter()
            .any(|annotation| parent.starts_with(&annotation.path))
        {
            return Err(paint_error());
        }
        let mut children = self.elements.as_slice();
        for index in parent {
            children = children
                .get(*index)
                .and_then(paint_children)
                .ok_or_else(paint_error)?;
        }
        if positions.len() != old_for_destination.len()
            || positions.windows(2).any(|p| p[0] >= p[1])
            || positions.last().is_some_and(|i| *i >= children.len())
        {
            return Err(paint_error());
        }
        let mut inverse = Vec::new();
        inverse
            .try_reserve_exact(positions.len())
            .map_err(|_| paint_error())?;
        inverse.resize(positions.len(), usize::MAX);
        for (destination, old) in old_for_destination.iter().copied().enumerate() {
            if old >= inverse.len() || inverse[old] != usize::MAX {
                return Err(paint_error());
            }
            inverse[old] = destination;
        }
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(self.policies.entries.len())
            .map_err(|_| paint_error())?;
        for entry in &self.policies.entries {
            entries.push(PaintEntry {
                path: copy_paint_path(&entry.path)?,
                policy: entry.policy,
                occurrence: entry.occurrence,
            });
        }
        for entry in &mut entries {
            if entry.path.starts_with(parent)
                && entry.path.len() > parent.len()
                && let Ok(old) = positions.binary_search(&entry.path[parent.len()])
            {
                entry.path[parent.len()] = positions[inverse[old]];
            }
        }
        entries.sort_unstable_by(|a, b| a.path.cmp(&b.path));
        let mut annotations = Vec::new();
        annotations
            .try_reserve_exact(self.policies.annotations.len())
            .map_err(|_| paint_error())?;
        for annotation in &self.policies.annotations {
            let mut path = copy_paint_path(&annotation.path)?;
            if path.starts_with(parent)
                && path.len() > parent.len()
                && let Ok(old) = positions.binary_search(&path[parent.len()])
            {
                path[parent.len()] = positions[inverse[old]];
            }
            annotations.push(PaintAnnotation {
                path,
                recipe: Arc::clone(&annotation.recipe),
                kind: annotation.kind,
                base_count: annotation.base_count,
                placement: annotation.placement,
            });
        }
        annotations.sort_unstable_by(|a, b| a.path.cmp(&b.path));
        let mut slots = Vec::new();
        slots
            .try_reserve_exact(positions.len())
            .map_err(|_| paint_error())?;
        let mut destination = &mut self.elements;
        for index in parent {
            destination =
                paint_children_mut(&mut destination[*index]).expect("validated child forest");
        }
        for position in positions {
            slots.push(Some(std::mem::replace(
                &mut destination[*position],
                PositionedElement::MarkedContent {
                    structure: None,
                    children: Vec::new(),
                },
            )));
        }
        for (position, old) in positions.iter().zip(old_for_destination) {
            destination[*position] = slots[*old].take().expect("validated permutation");
        }
        self.policies.entries = entries;
        self.policies.annotations = annotations;
        Ok(())
    }

    pub(crate) fn wrap(mut self, mut shell: PositionedElement) -> oxml_layout::Result<Self> {
        self.policies.validate(&self.elements)?;
        if paint_children(&shell).is_none_or(|children| !children.is_empty()) {
            return Err(paint_error());
        }
        for entry in &mut self.policies.entries {
            entry.path.try_reserve(1).map_err(|_| paint_error())?;
            entry.path.insert(0, 0);
        }
        for annotation in &mut self.policies.annotations {
            annotation.path.try_reserve(1).map_err(|_| paint_error())?;
            annotation.path.insert(0, 0);
        }
        let mut roots = Vec::new();
        roots.try_reserve_exact(1).map_err(|_| paint_error())?;
        *paint_children_mut(&mut shell).expect("validated wrapper") = self.elements;
        roots.push(shell);
        self.elements = roots;
        Ok(self)
    }
}

impl PaintPolicies {
    /// Emit from a known producer before its private context is erased.
    pub(crate) fn emit(
        &mut self,
        elements: &mut Vec<PositionedElement>,
        element: PositionedElement,
        owner: FieldPaintOwner,
    ) -> oxml_layout::Result<()> {
        self.validate(elements)?;
        let mut output = PaintedElements {
            elements: std::mem::take(elements),
            policies: std::mem::take(self),
        };
        let result = if paint_field(&element) {
            output.push_field(element, owner)
        } else {
            output.push_non_field(element)
        };
        *elements = output.elements;
        *self = output.policies;
        result
    }

    pub(crate) fn append(
        &mut self,
        elements: &mut Vec<PositionedElement>,
        mut other: PaintedElements,
    ) -> oxml_layout::Result<()> {
        self.validate(elements)?;
        let mut output = PaintedElements {
            elements: std::mem::take(elements),
            policies: std::mem::take(self),
        };
        let result = output.append(&mut other);
        *elements = output.elements;
        *self = output.policies;
        result
    }

    pub(crate) fn permute(
        &mut self,
        elements: &mut Vec<PositionedElement>,
        positions: &[usize],
        old_for_destination: &[usize],
    ) -> oxml_layout::Result<()> {
        self.validate(elements)?;
        let mut output = PaintedElements {
            elements: std::mem::take(elements),
            policies: std::mem::take(self),
        };
        let result = output.permute_at(&[], positions, old_for_destination);
        *elements = output.elements;
        *self = output.policies;
        result
    }
}

#[derive(Debug, Clone, Default)]
pub(crate) struct PaintLayers {
    pub behind: PaintedElements,
    pub regular: PaintedElements,
    pub front: PaintedElements,
}

impl PaintLayers {
    pub(crate) fn accept_note(
        &mut self,
        note: oxml_layout::Result<(f64, PaintLayers)>,
    ) -> oxml_layout::Result<f64> {
        let (height, mut layers) = note?;
        self.append(&mut layers)?;
        Ok(height)
    }

    pub(crate) fn append(&mut self, other: &mut Self) -> oxml_layout::Result<()> {
        self.behind.append(&mut other.behind)?;
        self.regular.append(&mut other.regular)?;
        self.front.append(&mut other.front)
    }

    pub(crate) fn flatten(mut self) -> oxml_layout::Result<PaintedElements> {
        self.behind.append(&mut self.regular)?;
        self.behind.append(&mut self.front)?;
        Ok(self.behind)
    }
}

impl WordParagraphPaint {
    pub(crate) fn remap_sources(
        &mut self,
        nodes: &std::collections::HashMap<SourceNodeId, SourceNodeId>,
    ) -> oxml_layout::Result<()> {
        for (_, occurrence) in &mut self.field_occurrences {
            occurrence.source.node = *nodes.get(&occurrence.source.node).ok_or_else(paint_error)?;
        }
        for annotation in &mut self.input_annotations {
            let recipe = Arc::make_mut(&mut annotation.recipe);
            recipe.occurrence.source.node = *nodes
                .get(&recipe.occurrence.source.node)
                .ok_or_else(paint_error)?;
            if let Some(span) = &mut recipe.source_span {
                span.node = *nodes.get(&span.node).ok_or_else(paint_error)?;
            }
            annotation.policies.remap_sources(nodes)?;
        }
        for row in &mut self.rows {
            row.policies.remap_sources(nodes)?;
        }
        for anchor in &mut self.anchors {
            match anchor {
                WordAnchorPaint::Image => {}
                WordAnchorPaint::Group { policies } => policies.remap_sources(nodes)?,
                WordAnchorPaint::Shape { paragraphs } => {
                    for paragraph in paragraphs {
                        paragraph.remap_sources(nodes)?;
                    }
                }
            }
        }
        Ok(())
    }

    pub(crate) fn rebind_source(&mut self, node: Option<SourceNodeId>) -> oxml_layout::Result<()> {
        let Some(node) = node else {
            if !self.input_annotations.is_empty() {
                return Err(paint_error());
            }
            self.field_occurrences.clear();
            return Ok(());
        };
        for (_, occurrence) in &mut self.field_occurrences {
            occurrence.source.node = node;
        }
        for annotation in &mut self.input_annotations {
            let recipe = Arc::make_mut(&mut annotation.recipe);
            recipe.occurrence.source.node = node;
            if let Some(span) = &mut recipe.source_span {
                span.node = node;
            }
            annotation.policies.rebind_source(node);
        }
        for row in &mut self.rows {
            row.policies.rebind_source(node);
        }
        Ok(())
    }
}

impl PaintPolicies {
    fn remap_sources(
        &mut self,
        nodes: &std::collections::HashMap<SourceNodeId, SourceNodeId>,
    ) -> oxml_layout::Result<()> {
        for entry in &mut self.entries {
            if let Some(occurrence) = &mut entry.occurrence {
                occurrence.source.node =
                    *nodes.get(&occurrence.source.node).ok_or_else(paint_error)?;
            }
        }
        for annotation in &mut self.annotations {
            let recipe = Arc::make_mut(&mut annotation.recipe);
            recipe.occurrence.source.node = *nodes
                .get(&recipe.occurrence.source.node)
                .ok_or_else(paint_error)?;
            if let Some(span) = &mut recipe.source_span {
                span.node = *nodes.get(&span.node).ok_or_else(paint_error)?;
            }
        }
        Ok(())
    }

    fn rebind_source(&mut self, node: SourceNodeId) {
        for entry in &mut self.entries {
            if let Some(occurrence) = &mut entry.occurrence {
                occurrence.source.node = node;
            }
        }
        for annotation in &mut self.annotations {
            let recipe = Arc::make_mut(&mut annotation.recipe);
            recipe.occurrence.source.node = node;
            if let Some(span) = &mut recipe.source_span {
                span.node = node;
            }
        }
    }
}

impl PaintedElements {
    pub(crate) fn clear_annotation_kind(&mut self, path: &[usize]) -> oxml_layout::Result<()> {
        self.policies.validate(&self.elements)?;
        let index = self
            .policies
            .annotations
            .binary_search_by(|annotation| annotation.path.as_slice().cmp(path))
            .map_err(|_| paint_error())?;
        let mut recipe = self.policies.annotations[index].recipe.as_ref().clone();
        recipe.field_kind = None;
        let recipe = Arc::new(recipe);
        for entry in self
            .policies
            .entries
            .iter()
            .filter(|entry| entry.path.starts_with(path))
        {
            match paint_at_path_mut(&mut self.elements, &entry.path)
                .expect("validated annotation base leaf")
            {
                PositionedElement::Text(run) => run.field_kind = None,
                PositionedElement::MultilingualText(run) => run.field_kind = None,
                _ => unreachable!("validated annotation base kind"),
            }
        }
        self.policies.annotations[index].recipe = recipe;
        Ok(())
    }
}

fn paint_at_path_mut<'a>(
    mut elements: &'a mut Vec<PositionedElement>,
    path: &[usize],
) -> Option<&'a mut PositionedElement> {
    let (last, parents) = path.split_last()?;
    for index in parents {
        elements = paint_children_mut(elements.get_mut(*index)?)?;
    }
    elements.get_mut(*last)
}

impl RowSemantics {
    pub(crate) fn with_repetition_scope(&self, scope: u64) -> oxml_layout::Result<Option<Self>> {
        let mut owners = std::collections::HashSet::new();
        collect_row_annotation_owners(self, &mut owners)?;
        if owners.is_empty() {
            return Ok(None);
        }
        let mut repeated = self.clone();
        rescope_row_annotations(&mut repeated, &owners, scope);
        Ok(Some(repeated))
    }
}

fn collect_row_annotation_owners(
    row: &RowSemantics,
    owners: &mut std::collections::HashSet<WordFieldOccurrence>,
) -> oxml_layout::Result<()> {
    for cell in &row.cells {
        for block in &cell.blocks {
            match block {
                CellBlockSemantics::Paragraph(paragraph) => {
                    collect_paint_annotation_owners(&paragraph.paint, owners)?
                }
                CellBlockSemantics::Table(table) => {
                    for row in &table.rows {
                        collect_row_annotation_owners(row, owners)?;
                    }
                }
            }
        }
    }
    Ok(())
}

fn collect_paint_annotation_owners(
    paint: &WordParagraphPaint,
    owners: &mut std::collections::HashSet<WordFieldOccurrence>,
) -> oxml_layout::Result<()> {
    owners
        .try_reserve(paint.input_annotations.len())
        .map_err(|_| paint_error())?;
    owners.extend(
        paint
            .input_annotations
            .iter()
            .map(|annotation| annotation.recipe.occurrence),
    );
    for row in &paint.rows {
        let occurrences = row.policies.annotation_occurrences()?;
        owners
            .try_reserve(occurrences.len())
            .map_err(|_| paint_error())?;
        owners.extend(occurrences);
    }
    for anchor in &paint.anchors {
        match anchor {
            WordAnchorPaint::Image => {}
            WordAnchorPaint::Group { policies } => {
                let occurrences = policies.annotation_occurrences()?;
                owners
                    .try_reserve(occurrences.len())
                    .map_err(|_| paint_error())?;
                owners.extend(occurrences);
            }
            WordAnchorPaint::Shape { paragraphs } => {
                for paragraph in paragraphs {
                    collect_paint_annotation_owners(paragraph, owners)?;
                }
            }
        }
    }
    Ok(())
}

fn rescope_row_annotations(
    row: &mut RowSemantics,
    owners: &std::collections::HashSet<WordFieldOccurrence>,
    scope: u64,
) {
    for cell in &mut row.cells {
        for block in &mut cell.blocks {
            match block {
                CellBlockSemantics::Paragraph(paragraph) => {
                    rescope_paint_annotations(Arc::make_mut(&mut paragraph.paint), owners, scope)
                }
                CellBlockSemantics::Table(table) => {
                    for row in &mut table.rows {
                        rescope_row_annotations(row, owners, scope);
                    }
                }
            }
        }
    }
}

fn rescope_paint_annotations(
    paint: &mut WordParagraphPaint,
    owners: &std::collections::HashSet<WordFieldOccurrence>,
    scope: u64,
) {
    for (_, occurrence) in &mut paint.field_occurrences {
        if owners.contains(occurrence) {
            occurrence.paint_scope = scope;
        }
    }
    for annotation in &mut paint.input_annotations {
        if owners.contains(&annotation.recipe.occurrence) {
            Arc::make_mut(&mut annotation.recipe).occurrence.paint_scope = scope;
        }
        annotation.policies.rescope_annotations(owners, scope);
    }
    for row in &mut paint.rows {
        row.policies.rescope_annotations(owners, scope);
    }
    for anchor in &mut paint.anchors {
        match anchor {
            WordAnchorPaint::Image => {}
            WordAnchorPaint::Group { policies } => policies.rescope_annotations(owners, scope),
            WordAnchorPaint::Shape { paragraphs } => {
                for paragraph in paragraphs {
                    rescope_paint_annotations(paragraph, owners, scope);
                }
            }
        }
    }
}

impl PaintPolicies {
    fn rescope_annotations(
        &mut self,
        owners: &std::collections::HashSet<WordFieldOccurrence>,
        scope: u64,
    ) {
        for entry in &mut self.entries {
            if let Some(occurrence) = &mut entry.occurrence
                && owners.contains(occurrence)
            {
                occurrence.paint_scope = scope;
            }
        }
        for annotation in &mut self.annotations {
            if owners.contains(&annotation.recipe.occurrence) {
                Arc::make_mut(&mut annotation.recipe).occurrence.paint_scope = scope;
            }
        }
    }
}

fn children_len_before(final_len: usize, removed_len: usize, replacement_len: usize) -> usize {
    final_len - replacement_len + removed_len
}

#[derive(Debug, Clone, PartialEq)]
pub(crate) struct ParagraphSemantics {
    pub paint: Arc<WordParagraphPaint>,
    pub tb_rl_exact_baseline: Option<(f64, bool)>,
    // Facing group edges are removed only on transient layout clones.
    pub border_group_before: bool,
    pub source_node: Option<SourceNodeId>,
    pub structure_id: Option<StructureId>,
    pub reflow_direction: TextDirection,
}

#[derive(Debug, Clone, PartialEq)]
pub(crate) enum CellBlockSemantics {
    Paragraph(ParagraphSemantics),
    Table(TableSemantics),
}

#[derive(Debug, Clone, PartialEq)]
pub(crate) struct CellSemantics {
    pub blocks: Vec<CellBlockSemantics>,
}

#[derive(Debug, Clone, PartialEq)]
pub(crate) struct RowSemantics {
    pub cells: Vec<CellSemantics>,
}

#[derive(Debug, Clone, PartialEq)]
pub(crate) struct TableSemantics {
    pub rows: Vec<RowSemantics>,
}

#[derive(Debug, Clone)]
pub(crate) enum SharedLayoutBlock {
    Owned {
        semantics: Option<ParagraphSemantics>,
        border_group_before: bool,
        block: Box<LayoutBlock>,
        reflow_direction: TextDirection,
        body_index: Option<usize>,
    },
    Paragraph {
        block: Arc<ParagraphBlock>,
        semantics: ParagraphSemantics,
        body_index: Option<usize>,
    },
    Table {
        block: Arc<TableBlock>,
        semantics: TableSemantics,
        body_index: Option<usize>,
    },
}

impl SharedLayoutBlock {
    pub(crate) fn set_body_index(&mut self, body_index: usize) {
        match self {
            Self::Owned {
                body_index: owner, ..
            }
            | Self::Paragraph {
                body_index: owner, ..
            }
            | Self::Table {
                body_index: owner, ..
            } => *owner = Some(body_index),
        }
    }
}

#[derive(Clone, Copy)]
pub(crate) struct ParagraphView<'a> {
    pub line_start: usize,
    pub paint: Option<&'a WordParagraphPaint>,
    pub baseline_policy: Option<(f64, bool, &'a oxml_layout::FontManager)>,
    pub line_gaps: &'a [f64],
    pub border_group_before: bool,
    pub block: &'a ParagraphBlock,
    pub semantics: Option<&'a ParagraphSemantics>,
    pub reflow_direction: TextDirection,
    pub reflow_allowed: bool,
}

impl Deref for ParagraphView<'_> {
    type Target = ParagraphBlock;

    fn deref(&self) -> &Self::Target {
        self.block
    }
}

impl ParagraphView<'_> {
    pub fn content_height(self) -> f64 {
        self.block.content_height() + self.line_gaps.iter().sum::<f64>()
    }

    pub fn source_node(self) -> Option<Option<SourceNodeId>> {
        self.semantics.map(|semantics| semantics.source_node)
    }

    pub fn structure_id(self) -> Option<StructureId> {
        self.semantics
            .map_or(self.block.structure_id, |semantics| semantics.structure_id)
    }
}

#[derive(Clone, Copy)]
pub(crate) struct TableView<'a> {
    pub block: &'a TableBlock,
    pub semantics: Option<&'a TableSemantics>,
}

impl Deref for TableView<'_> {
    type Target = TableBlock;

    fn deref(&self) -> &Self::Target {
        self.block
    }
}

pub(crate) trait LayoutBlockLike {
    fn paragraph(&self) -> Option<ParagraphView<'_>>;
    fn table(&self) -> Option<TableView<'_>>;

    fn body_index(&self) -> Option<usize> {
        None
    }

    fn paint_owner_valid(&self) -> bool {
        true
    }

    fn content_height(&self) -> f64 {
        self.paragraph().map_or_else(
            || self.table().unwrap().content_height(),
            |p| p.content_height(),
        )
    }

    fn space_before(&self) -> f64 {
        self.paragraph()
            .map_or(0.0, |paragraph| paragraph.space_before)
    }

    fn space_after(&self) -> f64 {
        self.paragraph()
            .map_or(0.0, |paragraph| paragraph.space_after)
    }

    fn page_break_before(&self) -> bool {
        self.paragraph()
            .is_some_and(|paragraph| paragraph.page_break_before)
    }
}

impl LayoutBlockLike for LayoutBlock {
    fn paragraph(&self) -> Option<ParagraphView<'_>> {
        match self {
            Self::Paragraph(block) => Some(ParagraphView {
                line_start: 0,
                paint: None,
                baseline_policy: None,
                line_gaps: &[],
                border_group_before: false,
                block,
                semantics: None,
                reflow_direction: TextDirection::Auto,
                reflow_allowed: true,
            }),
            Self::Table(_) => None,
        }
    }

    fn table(&self) -> Option<TableView<'_>> {
        match self {
            Self::Paragraph(_) => None,
            Self::Table(block) => Some(TableView {
                block,
                semantics: None,
            }),
        }
    }
}

impl LayoutBlockLike for SharedLayoutBlock {
    fn paint_owner_valid(&self) -> bool {
        match self {
            Self::Owned {
                block, semantics, ..
            } => block.paragraph().is_some() == semantics.is_some(),
            Self::Paragraph { .. } | Self::Table { .. } => true,
        }
    }

    fn paragraph(&self) -> Option<ParagraphView<'_>> {
        match self {
            Self::Owned {
                block,
                border_group_before,
                reflow_direction,
                semantics,
                ..
            } => block.paragraph().map(|mut paragraph| {
                paragraph.semantics = semantics.as_ref();
                paragraph.reflow_direction = *reflow_direction;
                paragraph.border_group_before = *border_group_before;
                paragraph
            }),
            Self::Paragraph {
                block, semantics, ..
            } => Some(ParagraphView {
                line_start: 0,
                paint: None,
                baseline_policy: None,
                line_gaps: &[],
                border_group_before: semantics.border_group_before,
                block,
                semantics: Some(semantics),
                reflow_direction: semantics.reflow_direction,
                reflow_allowed: true,
            }),
            Self::Table { .. } => None,
        }
    }

    fn table(&self) -> Option<TableView<'_>> {
        match self {
            Self::Owned { block, .. } => block.table(),
            Self::Paragraph { .. } => None,
            Self::Table {
                block, semantics, ..
            } => Some(TableView {
                block,
                semantics: Some(semantics),
            }),
        }
    }

    fn body_index(&self) -> Option<usize> {
        match self {
            Self::Owned { body_index, .. }
            | Self::Paragraph { body_index, .. }
            | Self::Table { body_index, .. } => *body_index,
        }
    }
}

impl LayoutBlock {
    /// Total height including spacing.
    pub fn total_height(&self) -> f64 {
        match self {
            LayoutBlock::Paragraph(p) => p.total_height(),
            LayoutBlock::Table(t) => t.total_height(),
        }
    }

    /// Content height without spacing.
    pub fn content_height(&self) -> f64 {
        match self {
            LayoutBlock::Paragraph(p) => p.content_height(),
            LayoutBlock::Table(t) => t.content_height(),
        }
    }

    pub fn space_before(&self) -> f64 {
        match self {
            LayoutBlock::Paragraph(p) => p.space_before,
            LayoutBlock::Table(_) => 0.0,
        }
    }

    pub fn space_after(&self) -> f64 {
        match self {
            LayoutBlock::Paragraph(p) => p.space_after,
            LayoutBlock::Table(_) => 0.0,
        }
    }

    pub fn keep_next(&self) -> bool {
        match self {
            LayoutBlock::Paragraph(p) => p.keep_next,
            LayoutBlock::Table(_) => false,
        }
    }

    pub fn keep_lines(&self) -> bool {
        match self {
            LayoutBlock::Paragraph(p) => p.keep_lines,
            LayoutBlock::Table(_) => false,
        }
    }

    pub fn page_break_before(&self) -> bool {
        match self {
            LayoutBlock::Paragraph(p) => p.page_break_before,
            LayoutBlock::Table(_) => false,
        }
    }

    pub fn widow_control(&self) -> bool {
        match self {
            LayoutBlock::Paragraph(p) => p.widow_control,
            LayoutBlock::Table(_) => false,
        }
    }
}

/// What a paragraph needs in order to be broken into lines again.
///
/// Whether a floating drawing overlaps a line is only known once the paragraph
/// has a position on a page, which is after layout and during pagination. The
/// inputs to line breaking therefore have to survive that far.
///
/// `InlineItem::Text` holds the same shaped glyphs `LayoutLine` already holds,
/// so this roughly doubles a paragraph's text memory. It is carried only when
/// the document contains a drawing that wraps, which nearly none do.
#[derive(Debug, Clone)]
pub struct ParagraphReflow {
    pub items: Vec<InlineItem>,
    pub params: LineBreakParams,
    /// The section grid pitch this paragraph's line advance snaps to.
    ///
    /// `LineBreakParams` has no grid, so a re-break around a float would
    /// leave the grid without this. It is `None` for every paragraph off the
    /// grid, including one on an exact `w:lineRule`, so an ungridded reflow
    /// takes the arithmetic it always took.
    pub grid_line_pitch_pt: Option<f64>,
    /// Word's measure of the paragraph mark, which a re-broken line holding
    /// only inline objects takes its proportional spacing from.
    pub paragraph_mark: Option<FontMetrics>,
}

/// A laid-out paragraph with its lines and spacing.
#[derive(Debug, Clone)]
pub struct ParagraphBlock {
    /// Laid-out lines.
    pub lines: Vec<LayoutLine>,
    /// The colour of the change bar the tracked projection draws beside the
    /// paragraph, `None` when it shows no revision.
    pub change_bar: Option<Color>,
    /// Floating drawings anchored to this paragraph.
    ///
    /// These travel with the paragraph so the paginator can resolve a
    /// paragraph-relative or line-relative offset once it knows where the
    /// paragraph actually landed.
    pub anchored: Vec<AnchoredDrawing>,
    /// Space before the paragraph in points.
    pub space_before: f64,
    /// Space after the paragraph in points.
    pub space_after: f64,
    /// Paragraph borders.
    pub borders: Option<CT_PBdr>,
    /// Background shading color.
    pub shading: Option<Color>,
    /// Left indent in points.
    pub indent_left: f64,
    /// Right indent in points.
    pub indent_right: f64,
    /// Paragraph justification.
    pub jc: Option<Align>,
    /// Keep with next paragraph.
    pub keep_next: bool,
    /// Keep all lines together on one page.
    pub keep_lines: bool,
    /// Force page break before this paragraph.
    pub page_break_before: bool,
    /// Widow/orphan control.
    pub widow_control: bool,
    /// Heading level (1-9) if this is a heading paragraph, for outline generation.
    pub heading_level: Option<u32>,
    /// Heading text for outline generation.
    pub heading_text: Option<String>,
    /// Numbering instance and zero-based nesting level for semantic lists.
    pub list: Option<(u32, u8)>,
    /// Logical paragraph node allocated before pagination.
    pub structure_id: Option<StructureId>,
    /// Inputs for re-breaking this paragraph around a floating drawing.
    ///
    /// `None` unless the document holds a drawing that wraps.
    pub reflow: Option<Box<ParagraphReflow>>,
    /// Vertical space kept clear above the first line, for a drawing this
    /// paragraph must clear rather than flow beside.
    pub content_offset_top: f64,
}

/// Complete painted width of a supported paragraph border edge.
pub(crate) fn paragraph_border_width(edge: Option<&rdocx_oxml::borders::CT_BorderEdge>) -> f64 {
    let Some(edge) = edge.filter(|edge| {
        !edge.val.is_none() && !matches!(edge.val, rdocx_oxml::shared::ST_Border::Art(_))
    }) else {
        return 0.0;
    };
    let width = edge.sz.unwrap_or(4) as f64 / 8.0;
    if edge.val == rdocx_oxml::shared::ST_Border::Double {
        width * 3.0
    } else {
        width
    }
}

pub(crate) fn paragraph_border_padding(edge: Option<&rdocx_oxml::borders::CT_BorderEdge>) -> f64 {
    let width = paragraph_border_width(edge);
    if width == 0.0 {
        0.0
    } else {
        width + edge.and_then(|edge| edge.space).unwrap_or(0) as f64
    }
}

/// Facing edges of identical adjacent frames belong to the outer group only.
pub(crate) fn paragraph_border_groups(paragraphs: &[Option<&ParagraphBlock>]) -> Vec<(bool, bool)> {
    let joins = |left: Option<&ParagraphBlock>, right: Option<&ParagraphBlock>| {
        matches!((left,right), (Some(left),Some(right)) if !right.page_break_before
            && left.borders.is_some() && left.borders == right.borders
            && left.borders.as_ref().is_some_and(|borders| borders.between.is_none())
            && left.indent_left == right.indent_left && left.indent_right == right.indent_right)
    };
    (0..paragraphs.len())
        .map(|index| {
            (
                index > 0 && joins(paragraphs[index - 1], paragraphs[index]),
                index + 1 < paragraphs.len() && joins(paragraphs[index], paragraphs[index + 1]),
            )
        })
        .collect()
}

pub(crate) fn collapse_shared_border_groups(blocks: &mut [SharedLayoutBlock]) {
    if !blocks.iter().any(|block| {
        block
            .paragraph()
            .is_some_and(|paragraph| paragraph.borders.is_some())
    }) {
        return;
    }
    let paragraphs = blocks
        .iter()
        .map(|block| block.paragraph().map(|paragraph| paragraph.block))
        .collect::<Vec<_>>();
    let groups = paragraph_border_groups(&paragraphs);
    for (block, (top, bottom)) in blocks.iter_mut().zip(groups) {
        if !top && !bottom {
            continue;
        }
        match block {
            SharedLayoutBlock::Paragraph {
                block, semantics, ..
            } => {
                semantics.border_group_before = top;
                Arc::make_mut(block).remove_group_border_edges(top, bottom)
            }
            SharedLayoutBlock::Owned {
                block,
                border_group_before,
                ..
            } => {
                *border_group_before = top;
                if let LayoutBlock::Paragraph(paragraph) = block.as_mut() {
                    paragraph.remove_group_border_edges(top, bottom);
                }
            }
            SharedLayoutBlock::Table { .. } => {}
        }
    }
}

impl ParagraphBlock {
    /// Price frame bands into the existing line stream, so every height
    /// consumer, including note flattening and cell splitting, sees them.
    pub(crate) fn add_border_padding(&mut self) {
        let Some(borders) = &self.borders else {
            return;
        };
        let top = paragraph_border_padding(borders.top.as_ref());
        let bottom = paragraph_border_padding(borders.bottom.as_ref());
        if let Some(first) = self.lines.first_mut() {
            first.height += top;
            first.ascent += top;
        }
        if let Some(last) = self.lines.last_mut() {
            last.height += bottom;
        }
    }

    pub(crate) fn remove_group_border_edges(&mut self, top: bool, bottom: bool) {
        if !top && !bottom {
            return;
        }
        let Some(borders) = &mut self.borders else {
            return;
        };
        if top {
            let padding = paragraph_border_padding(borders.top.as_ref());
            if let Some(first) = self.lines.first_mut() {
                first.height -= padding;
                first.ascent -= padding;
            }
            borders.top = None;
        }
        if bottom {
            let padding = paragraph_border_padding(borders.bottom.as_ref());
            if let Some(last) = self.lines.last_mut() {
                last.height -= padding;
            }
            borders.bottom = None;
        }
    }

    /// Total height of the paragraph lines (not including before/after spacing).
    pub fn content_height(&self) -> f64 {
        self.content_offset_top + self.lines.iter().map(|l| l.height).sum::<f64>()
    }

    /// Total height including spacing.
    pub fn total_height(&self) -> f64 {
        self.space_before + self.content_height() + self.space_after
    }

    /// Number of lines.
    pub fn line_count(&self) -> usize {
        self.lines.len()
    }
}

/// Build a ParagraphBlock from resolved properties and layout lines.
pub fn build_paragraph_block(
    lines: Vec<LayoutLine>,
    space_before: f64,
    space_after: f64,
    borders: Option<CT_PBdr>,
    shading: Option<Color>,
    indent_left: f64,
    indent_right: f64,
    jc: Option<Align>,
    keep_next: bool,
    keep_lines: bool,
    page_break_before: bool,
    widow_control: bool,
) -> ParagraphBlock {
    let mut block = ParagraphBlock {
        lines,
        change_bar: None,
        anchored: Vec::new(),
        space_before,
        space_after,
        borders,
        shading,
        indent_left,
        indent_right,
        jc,
        keep_next,
        keep_lines,
        page_break_before,
        widow_control,
        heading_level: None,
        heading_text: None,
        list: None,
        structure_id: None,
        reflow: None,
        content_offset_top: 0.0,
    };
    block.add_border_padding();
    block
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn paragraph_block_height() {
        let block = ParagraphBlock {
            anchored: Vec::new(),
            change_bar: None,
            lines: vec![
                LayoutLine {
                    items: vec![],
                    width: 0.0,
                    ascent: 10.0,
                    descent: 3.0,
                    line_gap: 0.0,
                    height: 13.0,
                    indent_left: 0.0,
                    available_width: 468.0,
                    is_last: false,
                    forced_break_after: None,
                },
                LayoutLine {
                    items: vec![],
                    width: 0.0,
                    ascent: 10.0,
                    descent: 3.0,
                    line_gap: 0.0,
                    height: 13.0,
                    indent_left: 0.0,
                    available_width: 468.0,
                    is_last: true,
                    forced_break_after: None,
                },
            ],
            space_before: 6.0,
            space_after: 8.0,
            borders: None,
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
            list: None,
            structure_id: None,
            reflow: None,
            content_offset_top: 0.0,
        };
        assert!((block.content_height() - 26.0).abs() < 0.01);
        assert!((block.total_height() - 40.0).abs() < 0.01);
    }

    /// Only the presets we can actually draw map to a drawable body. Anything
    /// else still renders its text, so it must not be treated as a rectangle.
    #[test]
    fn shape_presets_map_to_what_we_can_draw() {
        assert_eq!(ShapePreset::from_prst(Some("rect")), ShapePreset::Rect);
        assert_eq!(ShapePreset::from_prst(Some("line")), ShapePreset::Line);
        assert_eq!(
            ShapePreset::from_prst(Some("straightConnector1")),
            ShapePreset::Line
        );
        assert_eq!(
            ShapePreset::from_prst(Some("roundRect")),
            ShapePreset::Unsupported,
            "an unhandled preset must not silently draw as a plain rectangle"
        );
        assert_eq!(ShapePreset::from_prst(None), ShapePreset::Unsupported);
    }
}
