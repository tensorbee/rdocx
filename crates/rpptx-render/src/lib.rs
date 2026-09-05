//! Presentation rendering inputs and package assembly helpers.

use std::collections::HashMap;
use std::error::Error;
use std::fmt;
use std::sync::Arc;

use oxml_drawing::theme::CT_OfficeStyleSheet;
use oxml_layout::{
    Color, DocumentMetadata, FontFile, FontManager, GroupElement, LayoutResult, MediaId, PageFrame,
    Paint, Path, PathCommand, PathElement, Point, PositionedElement, Rect, Stroke, Transform,
};
use rpptx_layout::{
    CropRect, ResolvedBackground, ResolvedContent, ResolvedGeometry, ResolvedImage,
    ResolvedImagePlacement, ResolvedLineEnd, ResolvedLineEndKind, ResolvedLineEndSize,
    ResolvedRectAlignment, ResolvedShape, ResolvedSlide, ResolvedSlideTextDirections,
    ResolvedTable, ResolvedTableBorder, ResolvedTileFlip, ResolvedTilePlacement,
    ScopedHyperlinkTargets,
};
use rpptx_oxml::notes_parts::CT_NotesSlide;
use rpptx_oxml::slide_parts::{CT_Slide, CT_SlideLayout, CT_SlideMaster};

mod text;
pub mod timeline;

/// The source part whose relationship map owns an identifier.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum RelScope {
    Slide,
    Layout,
    Master,
}

impl fmt::Display for RelScope {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Slide => "slide",
            Self::Layout => "layout",
            Self::Master => "master",
        })
    }
}

/// A package relationship after its target has been resolved against its source part.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ResolvedRel {
    pub target: String,
    pub relationship_type: String,
    pub target_mode: Option<String>,
}

/// Relationship maps kept separate by their source-part scope.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct RelScopes {
    pub slide: HashMap<String, ResolvedRel>,
    pub layout: HashMap<String, ResolvedRel>,
    pub master: HashMap<String, ResolvedRel>,
}

impl RelScopes {
    /// Look up a relationship only in the explicitly selected source-part scope.
    pub fn get(
        &self,
        scope: RelScope,
        relationship_id: &str,
    ) -> Result<&ResolvedRel, RenderInputError> {
        let relationships = match scope {
            RelScope::Slide => &self.slide,
            RelScope::Layout => &self.layout,
            RelScope::Master => &self.master,
        };
        relationships
            .get(relationship_id)
            .ok_or_else(|| RenderInputError::MissingRelationship {
                scope,
                relationship_id: relationship_id.to_owned(),
            })
    }

    /// Project external hyperlink relationships into layout's source-scoped map.
    pub fn external_hyperlink_targets(&self) -> ScopedHyperlinkTargets {
        fn external_targets(
            relationships: &HashMap<String, ResolvedRel>,
        ) -> HashMap<String, String> {
            relationships
                .iter()
                .filter(|(_, relationship)| {
                    relationship.relationship_type == HYPERLINK_RELATIONSHIP
                        && relationship.target_mode.as_deref() == Some("External")
                })
                .map(|(id, relationship)| (id.clone(), relationship.target.clone()))
                .collect()
        }

        ScopedHyperlinkTargets {
            slide: external_targets(&self.slide),
            layout: external_targets(&self.layout),
            master: external_targets(&self.master),
        }
    }
}

const HYPERLINK_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/// Media bytes available to a renderer, with their package content type.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MediaData {
    pub bytes: Vec<u8>,
    pub content_type: String,
}

/// Package assembly failures that retain relationship source context.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum RenderInputError {
    MissingRelationship {
        scope: RelScope,
        relationship_id: String,
    },
    MissingMediaTarget {
        scope: RelScope,
        relationship_id: String,
        target: String,
    },
    SlideIndexOutOfBounds {
        index: usize,
        slide_count: usize,
    },
    MissingMedia {
        media: MediaId,
    },
    InvalidPicture {
        media: MediaId,
        detail: &'static str,
    },
    TileLimitExceeded {
        media: MediaId,
        requested: usize,
        limit: usize,
    },
    TextLayout {
        detail: String,
    },
}

impl fmt::Display for RenderInputError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::MissingRelationship {
                scope,
                relationship_id,
            } => write!(formatter, "missing {scope} relationship {relationship_id}"),
            Self::MissingMediaTarget {
                scope,
                relationship_id,
                target,
            } => write!(
                formatter,
                "missing media target {target} for {scope} relationship {relationship_id}"
            ),
            Self::SlideIndexOutOfBounds { index, slide_count } => write!(
                formatter,
                "slide index {index} is out of bounds for {slide_count} slides"
            ),
            Self::MissingMedia { media } => {
                write!(formatter, "missing picture media {}", media.0)
            }
            Self::InvalidPicture { media, detail } => {
                write!(formatter, "invalid picture media {}: {detail}", media.0)
            }
            Self::TileLimitExceeded {
                media,
                requested,
                limit,
            } => write!(
                formatter,
                "picture media {} requests {requested} tiles, above the {limit} tile limit",
                media.0
            ),
            Self::TextLayout { detail } => write!(formatter, "text layout failed: {detail}"),
        }
    }
}

impl Error for RenderInputError {}

/// Raw package parts assembled before inheritance resolution.
#[derive(Clone, Debug)]
pub struct SlideBundle {
    pub slide: CT_Slide,
    pub layout: Arc<CT_SlideLayout>,
    pub master: Arc<CT_SlideMaster>,
    pub theme: Arc<CT_OfficeStyleSheet>,
    pub notes: Option<CT_NotesSlide>,
    pub hidden: bool,
    pub relationships: RelScopes,
}

/// Frozen, format-neutral input consumed by the rendering stage.
#[derive(Clone, Debug)]
pub struct RenderInput {
    pub slides: Vec<ResolvedSlide>,
    pub media: HashMap<MediaId, MediaData>,
    pub fonts: Vec<FontFile>,
    pub metadata: Option<DocumentMetadata>,
}

/// Lower every resolved slide to one fixed-size page in presentation order.
pub fn layout_presentation(input: &RenderInput) -> Result<LayoutResult, RenderInputError> {
    let mut font_manager = FontManager::new();
    font_manager.load_additional_fonts(&input.fonts);
    layout_presentation_with_font_manager(input, font_manager)
}

/// Lower every resolved slide using only bundled and presentation-embedded fonts.
///
/// Unlike [`layout_presentation`], this entry point never discovers host system
/// fonts, so its raster output is suitable for deterministic comparison gates.
pub fn layout_presentation_deterministic(
    input: &RenderInput,
) -> Result<LayoutResult, RenderInputError> {
    let mut font_manager =
        FontManager::new_deterministic().map_err(|error| RenderInputError::TextLayout {
            detail: error.to_string(),
        })?;
    font_manager.load_additional_fonts(&input.fonts);
    layout_presentation_with_font_manager(input, font_manager)
}

/// Lowers every slide with the same manager that shaped any frozen group content.
pub fn layout_presentation_with_font_manager(
    input: &RenderInput,
    mut font_manager: FontManager,
) -> Result<LayoutResult, RenderInputError> {
    layout_presentation_with_font_manager_inner(input, &mut font_manager, None)
}

/// Lowers every slide with paragraph directions from the additive resolver path.
///
/// Direction entries follow slides, shapes, text bodies, and paragraphs in that
/// order. A table shape has one text-body entry per cell in row-major order.
pub fn layout_presentation_with_font_manager_and_text_directions(
    input: &RenderInput,
    mut font_manager: FontManager,
    text_directions: &[ResolvedSlideTextDirections],
) -> Result<LayoutResult, RenderInputError> {
    layout_presentation_with_font_manager_and_text_directions_mut(
        input,
        &mut font_manager,
        text_directions,
    )
}

/// Internal facade hook for `rpptx` prepared rendering.
///
/// `font_manager` must be the same instance that assigned every
/// [`oxml_layout::FontId`]
/// stored in `input`. Rasterization must use the font data retained by that
/// manager so each identifier resolves to the face that shaped it.
#[doc(hidden)]
pub fn layout_presentation_with_font_manager_and_text_directions_mut(
    input: &RenderInput,
    font_manager: &mut FontManager,
    text_directions: &[ResolvedSlideTextDirections],
) -> Result<LayoutResult, RenderInputError> {
    layout_presentation_with_font_manager_inner(input, font_manager, Some(text_directions))
}

fn layout_presentation_with_font_manager_inner(
    input: &RenderInput,
    font_manager: &mut FontManager,
    text_directions: Option<&[ResolvedSlideTextDirections]>,
) -> Result<LayoutResult, RenderInputError> {
    let pages = (0..input.slides.len())
        .map(|index| {
            layout_slide_with_fonts_and_text_directions(
                input,
                index,
                font_manager,
                text_directions
                    .and_then(|directions| directions.get(index))
                    .map(Vec::as_slice),
            )
            .map(Arc::new)
        })
        .collect::<Result<Vec<_>, _>>()?;
    let diagnostics = input
        .slides
        .iter()
        .flat_map(|slide| slide.diagnostics.iter().cloned())
        .collect();
    let mut layout = LayoutResult::new(
        pages,
        font_manager.all_font_data(),
        input.metadata.clone(),
        Vec::new(),
    );
    layout.diagnostics = diagnostics;
    Ok(layout)
}

/// Lower one zero-based resolved slide to a fixed-size page.
pub fn layout_slide(input: &RenderInput, index: usize) -> Result<PageFrame, RenderInputError> {
    let mut font_manager = FontManager::new();
    font_manager.load_additional_fonts(&input.fonts);
    layout_slide_with_fonts(input, index, &mut font_manager)
}

fn layout_slide_with_fonts(
    input: &RenderInput,
    index: usize,
    font_manager: &mut FontManager,
) -> Result<PageFrame, RenderInputError> {
    layout_slide_with_fonts_and_text_directions(input, index, font_manager, None)
}

fn layout_slide_with_fonts_and_text_directions(
    input: &RenderInput,
    index: usize,
    font_manager: &mut FontManager,
    text_directions: Option<&[Vec<Vec<oxml_layout::TextDirection>>]>,
) -> Result<PageFrame, RenderInputError> {
    layout_slide_with_fonts_text_directions_and_states(
        input,
        index,
        font_manager,
        text_directions,
        None,
    )
}

pub(crate) fn layout_slide_with_fonts_text_directions_and_states(
    input: &RenderInput,
    index: usize,
    font_manager: &mut FontManager,
    text_directions: Option<&[Vec<Vec<oxml_layout::TextDirection>>]>,
    shape_states: Option<&[rpptx_layout::timeline::EvaluatedShapeState]>,
) -> Result<PageFrame, RenderInputError> {
    let slide = input
        .slides
        .get(index)
        .ok_or(RenderInputError::SlideIndexOutOfBounds {
            index,
            slide_count: input.slides.len(),
        })?;
    layout_resolved_slide_with_fonts_text_directions_and_states(
        input,
        slide,
        index + 1,
        font_manager,
        text_directions,
        shape_states,
    )
}

pub(crate) fn layout_resolved_slide_with_fonts_text_directions_and_states(
    input: &RenderInput,
    slide: &ResolvedSlide,
    page_number: usize,
    font_manager: &mut FontManager,
    text_directions: Option<&[Vec<Vec<oxml_layout::TextDirection>>]>,
    shape_states: Option<&[rpptx_layout::timeline::EvaluatedShapeState]>,
) -> Result<PageFrame, RenderInputError> {
    let mut elements = Vec::new();
    let mut background_paint = None;
    match slide.background.as_ref() {
        Some(ResolvedBackground::Paint(paint)) => background_paint = Some(paint.clone()),
        Some(ResolvedBackground::Image(image)) => {
            let shape = ResolvedShape {
                group_transform: Transform::IDENTITY,
                bounds: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: slide.size.0,
                    height: slide.size.1,
                },
                rotation_deg: 0.0,
                flip_h: false,
                flip_v: false,
                geometry: ResolvedGeometry::Rectangle,
                fill: None,
                image_fill: None,
                line: None,
                head_end: None,
                tail_end: None,
                shadow: None,
                content: ResolvedContent::None,
                unsupported: None,
            };
            let paths = [Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: slide.size.0,
                height: slide.size.1,
            })];
            elements.extend(lower_picture(input, &shape, &paths, image)?);
        }
        None => {}
    }
    elements.extend(
        slide
            .shapes
            .iter()
            .enumerate()
            .map(|(shape_index, shape)| {
                let state = shape_states.and_then(|states| states.get(shape_index));
                lower_shape(
                    input,
                    shape,
                    font_manager,
                    page_number,
                    text_directions
                        .and_then(|directions| directions.get(shape_index))
                        .map(Vec::as_slice),
                )
                .map(|mut element| {
                    let Some(state) = state else {
                        return element;
                    };
                    let opacity = if state.visible {
                        f64::from(state.opacity.clamp(0.0, 1.0))
                    } else {
                        0.0
                    };
                    if let PositionedElement::Group(group) = &mut element {
                        group.opacity *= opacity;
                        for clip in
                            state.local_clip_paths((shape.bounds.width, shape.bounds.height))
                        {
                            if group.clip.is_none() {
                                group.clip = Some(clip);
                            } else {
                                group.children = vec![PositionedElement::Group(GroupElement {
                                    transform: Transform::IDENTITY,
                                    clip: Some(clip),
                                    opacity: 1.0,
                                    effects: Vec::new(),
                                    children: std::mem::take(&mut group.children),
                                })];
                            }
                        }
                    }
                    element
                })
            })
            .collect::<Result<Vec<_>, _>>()?,
    );
    let mut page = PageFrame::new(page_number, slide.size.0, slide.size.1, elements);
    page.background = background_paint;
    Ok(page)
}

fn lower_shape(
    input: &RenderInput,
    shape: &ResolvedShape,
    font_manager: &mut FontManager,
    page_number: usize,
    text_directions: Option<&[Vec<oxml_layout::TextDirection>]>,
) -> Result<PositionedElement, RenderInputError> {
    let paths = if matches!(shape.content, ResolvedContent::Table(_)) {
        Vec::new()
    } else {
        match &shape.geometry {
            ResolvedGeometry::Rectangle | ResolvedGeometry::BoundsFallback => {
                vec![Path::rect(Rect {
                    x: 0.0,
                    y: 0.0,
                    width: shape.bounds.width,
                    height: shape.bounds.height,
                })]
            }
            ResolvedGeometry::Custom { paths, .. } => paths.clone(),
        }
    };
    let stroke = match (&shape.geometry, &shape.fill, &shape.image_fill, &shape.line) {
        (ResolvedGeometry::BoundsFallback, None, None, None) => {
            Some(Stroke::new(Paint::Solid(Color::BLACK), 1.0))
        }
        _ => shape.line.clone(),
    };
    let mut text_children = Vec::new();
    let mut children = shape
        .image_fill
        .as_ref()
        .map(|image| lower_picture(input, shape, &paths, image))
        .transpose()?
        .unwrap_or_default();
    match &shape.content {
        ResolvedContent::Image(image) => {
            children.extend(lower_picture(input, shape, &paths, image)?)
        }
        ResolvedContent::Text(text_body) => {
            let content_box = text::content_box(shape, text_body);
            let (content_box, text_transform) =
                text::oriented_content_box(content_box, text_body.vertical);
            let paragraph_directions = text_directions
                .and_then(|directions| directions.first())
                .map(Vec::as_slice)
                .unwrap_or(&[]);
            let stacked = text::stack_text_for_page_with_directions(
                font_manager,
                content_box,
                text_body,
                page_number,
                paragraph_directions,
            )
            .map_err(|error| RenderInputError::TextLayout {
                detail: error.to_string(),
            })?;
            debug_assert!(stacked.width.is_finite() && stacked.height.is_finite());
            text_children = if let Some(transform) = text_transform {
                vec![PositionedElement::Group(GroupElement {
                    transform,
                    clip: None,
                    opacity: 1.0,
                    effects: Vec::new(),
                    children: stacked.elements,
                })]
            } else {
                stacked.elements
            };
        }
        ResolvedContent::Table(table) => {
            children.extend(lower_table(
                table,
                font_manager,
                page_number,
                text_directions.unwrap_or(&[]),
            )?);
        }
        ResolvedContent::Group(group) => {
            children.push(PositionedElement::Group(group.clone()));
        }
        _ => {}
    }
    children.extend(
        paths
            .iter()
            .cloned()
            .map(|path| {
                PositionedElement::Path(PathElement {
                    path,
                    fill: shape.fill.clone(),
                    stroke: stroke.clone(),
                })
            })
            .collect::<Vec<_>>(),
    );
    if let Some(line) = &shape.line {
        let (head_tangent, tail_tangent) = endpoint_tangents(&paths);
        if let Some(path) = shape
            .head_end
            .as_ref()
            .zip(head_tangent)
            .and_then(|(end, tangent)| line_end_path(end, tangent, line.width))
        {
            children.push(filled_line_end(path, &line.paint));
        }
        if let Some(path) = shape
            .tail_end
            .as_ref()
            .zip(tail_tangent)
            .and_then(|(end, tangent)| line_end_path(end, tangent, line.width))
        {
            children.push(filled_line_end(path, &line.paint));
        }
    }
    children.extend(text_children);
    Ok(PositionedElement::Group(GroupElement {
        transform: shape_transform(shape),
        clip: None,
        opacity: 1.0,
        effects: shape.shadow.iter().cloned().collect(),
        children,
    }))
}

fn lower_table(
    table: &ResolvedTable,
    font_manager: &mut FontManager,
    page_number: usize,
    text_directions: &[Vec<oxml_layout::TextDirection>],
) -> Result<Vec<PositionedElement>, RenderInputError> {
    let mut physical_column_widths = table.column_widths.clone();
    if table.right_to_left {
        physical_column_widths.reverse();
    }
    let column_offsets = cumulative_offsets(&physical_column_widths);
    let row_heights = table.rows.iter().map(|row| row.height).collect::<Vec<_>>();
    let row_offsets = cumulative_offsets(&row_heights);
    let mut fills = Vec::new();
    let mut texts = Vec::new();
    let mut borders: HashMap<(bool, usize, usize), TableBorderCandidate> = HashMap::new();
    let mut cell_index = 0usize;

    for (row_index, row) in table.rows.iter().enumerate() {
        for (column_index, cell) in row.cells.iter().enumerate() {
            let paragraph_directions = text_directions
                .get(cell_index)
                .map(Vec::as_slice)
                .unwrap_or(&[]);
            cell_index += 1;
            if cell.horizontal_merge
                || cell.vertical_merge
                || column_index >= table.column_widths.len()
            {
                continue;
            }
            let row_span = usize::try_from(cell.row_span)
                .unwrap_or(usize::MAX)
                .max(1)
                .min(table.rows.len().saturating_sub(row_index));
            let column_span = usize::try_from(cell.grid_span)
                .unwrap_or(usize::MAX)
                .max(1)
                .min(table.column_widths.len().saturating_sub(column_index));
            let visual_column = if table.right_to_left {
                table
                    .column_widths
                    .len()
                    .saturating_sub(column_index + column_span)
            } else {
                column_index
            };
            let rectangle = Rect {
                x: column_offsets[visual_column],
                y: row_offsets[row_index],
                width: column_offsets[visual_column + column_span] - column_offsets[visual_column],
                height: row_offsets[row_index + row_span] - row_offsets[row_index],
            };
            if let Some(fill) = &cell.fill {
                fills.push(PositionedElement::Path(PathElement {
                    path: Path::rect(rectangle),
                    fill: Some(fill.clone()),
                    stroke: None,
                }));
            }
            if let Some(text_body) = &cell.text {
                let mut text_body = text_body.clone();
                text_body.insets = cell.margins;
                let content = Rect {
                    x: rectangle.x + text_body.insets.left,
                    y: rectangle.y + text_body.insets.top,
                    width: (rectangle.width - text_body.insets.left - text_body.insets.right)
                        .max(0.0),
                    height: (rectangle.height - text_body.insets.top - text_body.insets.bottom)
                        .max(0.0),
                };
                let (content, transform) = text::oriented_content_box(content, text_body.vertical);
                let stacked = text::stack_text_for_page_with_directions(
                    font_manager,
                    content,
                    &text_body,
                    page_number,
                    paragraph_directions,
                )
                .map_err(|error| RenderInputError::TextLayout {
                    detail: error.to_string(),
                })?;
                if let Some(transform) = transform {
                    texts.push(PositionedElement::Group(GroupElement {
                        transform,
                        clip: None,
                        opacity: 1.0,
                        effects: Vec::new(),
                        children: stacked.elements,
                    }));
                } else {
                    texts.extend(stacked.elements);
                }
            }

            let last_row = row_index + row_span - 1;
            let last_column = column_index + column_span - 1;
            for logical_column in column_index..=last_column {
                let physical_column = if table.right_to_left {
                    table.column_widths.len() - logical_column - 1
                } else {
                    logical_column
                };
                let top = table.rows[row_index]
                    .cells
                    .get(logical_column)
                    .and_then(|covered| covered.top.as_ref())
                    .or(cell.top.as_ref());
                let bottom = table.rows[last_row]
                    .cells
                    .get(logical_column)
                    .and_then(|covered| covered.bottom.as_ref())
                    .or(cell.bottom.as_ref());
                insert_table_border(&mut borders, (true, row_index, physical_column), top, 4);
                insert_table_border(
                    &mut borders,
                    (true, row_index + row_span, physical_column),
                    bottom,
                    2,
                );
            }
            let left_column = if table.right_to_left {
                last_column
            } else {
                column_index
            };
            let right_column = if table.right_to_left {
                column_index
            } else {
                last_column
            };
            for covered_row in row_index..=last_row {
                let left = table.rows[covered_row]
                    .cells
                    .get(left_column)
                    .and_then(|covered| covered.left.as_ref())
                    .or(cell.left.as_ref());
                let right = table.rows[covered_row]
                    .cells
                    .get(right_column)
                    .and_then(|covered| covered.right.as_ref())
                    .or(cell.right.as_ref());
                insert_table_border(&mut borders, (false, visual_column, covered_row), left, 3);
                insert_table_border(
                    &mut borders,
                    (false, visual_column + column_span, covered_row),
                    right,
                    1,
                );
            }
        }
    }

    let mut ordered_borders = borders.into_iter().collect::<Vec<_>>();
    ordered_borders.sort_by_key(|(key, _)| *key);
    let mut border_elements = Vec::new();
    for ((horizontal, boundary, segment), candidate) in ordered_borders {
        let Some(stroke) = candidate.border.stroke else {
            continue;
        };
        let path = if horizontal {
            open_path(
                Point {
                    x: column_offsets[segment],
                    y: row_offsets[boundary],
                },
                Point {
                    x: column_offsets[segment + 1],
                    y: row_offsets[boundary],
                },
            )
        } else {
            open_path(
                Point {
                    x: column_offsets[boundary],
                    y: row_offsets[segment],
                },
                Point {
                    x: column_offsets[boundary],
                    y: row_offsets[segment + 1],
                },
            )
        };
        border_elements.push(PositionedElement::Path(PathElement {
            path,
            fill: None,
            stroke: Some(stroke),
        }));
    }
    fills.extend(texts);
    fills.extend(border_elements);
    Ok(fills)
}

struct TableBorderCandidate {
    border: ResolvedTableBorder,
    side_rank: u8,
}

fn insert_table_border(
    borders: &mut HashMap<(bool, usize, usize), TableBorderCandidate>,
    key: (bool, usize, usize),
    border: Option<&ResolvedTableBorder>,
    side_rank: u8,
) {
    let Some(border) = border else {
        return;
    };
    let candidate = TableBorderCandidate {
        border: border.clone(),
        side_rank,
    };
    let replace = borders.get(&key).is_none_or(|current| {
        let candidate_width = candidate
            .border
            .stroke
            .as_ref()
            .map_or(0.0, |stroke| stroke.width);
        let current_width = current
            .border
            .stroke
            .as_ref()
            .map_or(0.0, |stroke| stroke.width);
        (
            candidate.border.priority,
            ordered_width(candidate_width),
            candidate.side_rank,
        ) > (
            current.border.priority,
            ordered_width(current_width),
            current.side_rank,
        )
    });
    if replace {
        borders.insert(key, candidate);
    }
}

fn ordered_width(width: f64) -> u64 {
    if width.is_finite() && width >= 0.0 {
        width.to_bits()
    } else {
        0
    }
}

fn cumulative_offsets(lengths: &[f64]) -> Vec<f64> {
    let mut offsets = Vec::with_capacity(lengths.len() + 1);
    offsets.push(0.0);
    for length in lengths {
        offsets.push(offsets.last().copied().unwrap_or(0.0) + length.max(0.0));
    }
    offsets
}

fn open_path(start: Point, end: Point) -> Path {
    Path {
        commands: vec![PathCommand::MoveTo(start), PathCommand::LineTo(end)],
        fill_rule: oxml_layout::FillRule::NonZero,
    }
}

const MAX_TILE_ELEMENTS: usize = 4_096;
const MAX_TILED_IMAGE_BYTES: usize = 64 * 1024 * 1024;

fn lower_picture(
    input: &RenderInput,
    shape: &ResolvedShape,
    paths: &[Path],
    resolved_image: &ResolvedImage,
) -> Result<Vec<PositionedElement>, RenderInputError> {
    let media_id = resolved_image.media;
    let media = input
        .media
        .get(&media_id)
        .ok_or(RenderInputError::MissingMedia { media: media_id })?;
    let crop = normalized_insets(resolved_image.src_rect, media_id, "source crop")?;
    let elements = match &resolved_image.placement {
        ResolvedImagePlacement::Stretch { fill_rect } => {
            let fill_rect = normalized_insets(*fill_rect, media_id, "stretch fill rectangle")?;
            let destination = inset_rect(
                picture_coverage_rect(shape, resolved_image.rotate_with_shape),
                fill_rect,
            );
            let image = picture_image(media_id, media, expanded_crop_rect(destination, crop));
            let image = counter_rotate_image(image, shape, resolved_image.rotate_with_shape);
            let mut image = if crop.is_some() {
                clipped_group(Path::rect(destination), vec![image])
            } else {
                image
            };
            if !matches!(shape.geometry, ResolvedGeometry::Rectangle)
                || (!resolved_image.rotate_with_shape && shape.rotation_deg != 0.0)
            {
                image = clip_to_picture_shape(shape, paths, vec![image]);
            }
            vec![image]
        }
        ResolvedImagePlacement::Tile(tile) => lower_tiled_picture(
            shape,
            paths,
            media_id,
            media,
            crop,
            tile,
            resolved_image.dpi,
            resolved_image.rotate_with_shape,
        )?,
    };
    Ok(elements)
}

#[allow(clippy::too_many_arguments)]
fn lower_tiled_picture(
    shape: &ResolvedShape,
    paths: &[Path],
    media_id: MediaId,
    media: &MediaData,
    crop: Option<CropRect>,
    tile: &ResolvedTilePlacement,
    declared_dpi: Option<f64>,
    rotate_with_shape: bool,
) -> Result<Vec<PositionedElement>, RenderInputError> {
    let info = oxml_media::probe(&media.bytes).ok_or(RenderInputError::InvalidPicture {
        media: media_id,
        detail: "tile image metadata is unavailable",
    })?;
    let (native_width, native_height) = tile_native_size_points(info, declared_dpi, media_id)?;
    let tile_width = native_width * tile.scale_x;
    let tile_height = native_height * tile.scale_y;
    if !tile_width.is_finite()
        || !tile_height.is_finite()
        || tile_width <= 0.0
        || tile_height <= 0.0
        || !tile.translation.x.is_finite()
        || !tile.translation.y.is_finite()
    {
        return Err(RenderInputError::InvalidPicture {
            media: media_id,
            detail: "tile size or translation is not finite and positive",
        });
    }
    let shape_rect = local_shape_rect(shape);
    let coverage_rect = picture_coverage_rect(shape, rotate_with_shape);
    let anchor = tile_alignment_origin(shape_rect, tile_width, tile_height, tile.alignment);
    let translated_anchor = Point {
        x: anchor.x + tile.translation.x,
        y: anchor.y + tile.translation.y,
    };
    let origin_x = repeated_origin(translated_anchor.x, coverage_rect.x, tile_width);
    let origin_y = repeated_origin(translated_anchor.y, coverage_rect.y, tile_height);
    if !origin_x.is_finite() || !origin_y.is_finite() {
        return Err(RenderInputError::InvalidPicture {
            media: media_id,
            detail: "tile origin is not finite",
        });
    }
    let first_column = repeated_tile_index(origin_x, translated_anchor.x, tile_width, media_id)?;
    let first_row = repeated_tile_index(origin_y, translated_anchor.y, tile_height, media_id)?;
    let columns = repeat_count(
        origin_x,
        coverage_rect.x + coverage_rect.width,
        tile_width,
        media_id,
    )?;
    let rows = repeat_count(
        origin_y,
        coverage_rect.y + coverage_rect.height,
        tile_height,
        media_id,
    )?;
    let requested = columns
        .checked_mul(rows)
        .ok_or(RenderInputError::TileLimitExceeded {
            media: media_id,
            requested: usize::MAX,
            limit: MAX_TILE_ELEMENTS,
        })?;
    let byte_limit = MAX_TILED_IMAGE_BYTES / media.bytes.len().max(1);
    let limit = MAX_TILE_ELEMENTS.min(byte_limit.max(1));
    if requested > limit {
        return Err(RenderInputError::TileLimitExceeded {
            media: media_id,
            requested,
            limit,
        });
    }

    let mut tiles = Vec::with_capacity(requested);
    for row in 0..rows {
        for column in 0..columns {
            let rect = Rect {
                x: origin_x + column as f64 * tile_width,
                y: origin_y + row as f64 * tile_height,
                width: tile_width,
                height: tile_height,
            };
            let image = picture_image(media_id, media, expanded_crop_rect(rect, crop));
            let image = if crop.is_some() {
                clipped_group(Path::rect(rect), vec![image])
            } else {
                image
            };
            tiles.push(flip_tile(
                image,
                rect,
                tile.flip,
                (first_column.rem_euclid(2) == 1) != (column % 2 == 1),
                (first_row.rem_euclid(2) == 1) != (row % 2 == 1),
            ));
        }
    }
    let tiles = if rotate_with_shape || shape.rotation_deg == 0.0 {
        tiles
    } else {
        vec![PositionedElement::Group(GroupElement {
            transform: Transform::rotate_about(
                -shape.rotation_deg,
                shape.bounds.width / 2.0,
                shape.bounds.height / 2.0,
            ),
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: tiles,
        })]
    };
    Ok(vec![clip_to_picture_shape(shape, paths, tiles)])
}

fn normalized_insets(
    crop: Option<CropRect>,
    media: MediaId,
    detail: &'static str,
) -> Result<Option<CropRect>, RenderInputError> {
    let Some(crop) = crop else {
        return Ok(None);
    };
    let values = [crop.left, crop.top, crop.right, crop.bottom];
    if values.iter().any(|value| !value.is_finite()) {
        return Err(RenderInputError::InvalidPicture { media, detail });
    }
    let crop = CropRect {
        left: crop.left.clamp(0.0, 1.0),
        top: crop.top.clamp(0.0, 1.0),
        right: crop.right.clamp(0.0, 1.0),
        bottom: crop.bottom.clamp(0.0, 1.0),
    };
    if crop.left + crop.right >= 1.0 || crop.top + crop.bottom >= 1.0 {
        return Err(RenderInputError::InvalidPicture { media, detail });
    }
    Ok((crop != CropRect::default()).then_some(crop))
}

fn local_shape_rect(shape: &ResolvedShape) -> Rect {
    Rect {
        x: 0.0,
        y: 0.0,
        width: shape.bounds.width,
        height: shape.bounds.height,
    }
}

fn picture_coverage_rect(shape: &ResolvedShape, rotate_with_shape: bool) -> Rect {
    let rect = local_shape_rect(shape);
    if rotate_with_shape || shape.rotation_deg == 0.0 {
        return rect;
    }
    Transform::rotate_about(
        shape.rotation_deg,
        shape.bounds.width / 2.0,
        shape.bounds.height / 2.0,
    )
    .transform_rect_bbox(rect)
}

fn inset_rect(rect: Rect, insets: Option<CropRect>) -> Rect {
    let Some(insets) = insets else {
        return rect;
    };
    Rect {
        x: rect.x + rect.width * insets.left,
        y: rect.y + rect.height * insets.top,
        width: rect.width * (1.0 - insets.left - insets.right),
        height: rect.height * (1.0 - insets.top - insets.bottom),
    }
}

fn expanded_crop_rect(destination: Rect, crop: Option<CropRect>) -> Rect {
    let Some(crop) = crop else {
        return destination;
    };
    let retained_width = 1.0 - crop.left - crop.right;
    let retained_height = 1.0 - crop.top - crop.bottom;
    let width = destination.width / retained_width;
    let height = destination.height / retained_height;
    Rect {
        x: destination.x - width * crop.left,
        y: destination.y - height * crop.top,
        width,
        height,
    }
}

fn picture_image(media_id: MediaId, media: &MediaData, rect: Rect) -> PositionedElement {
    PositionedElement::Image {
        rect,
        data: media.bytes.clone(),
        content_type: media.content_type.clone(),
        media_id,
    }
}

fn counter_rotate_image(
    image: PositionedElement,
    shape: &ResolvedShape,
    rotate_with_shape: bool,
) -> PositionedElement {
    if rotate_with_shape || shape.rotation_deg == 0.0 {
        return image;
    }
    PositionedElement::Group(GroupElement {
        transform: Transform::rotate_about(
            -shape.rotation_deg,
            shape.bounds.width / 2.0,
            shape.bounds.height / 2.0,
        ),
        clip: None,
        opacity: 1.0,
        effects: Vec::new(),
        children: vec![image],
    })
}

fn clip_to_picture_shape(
    shape: &ResolvedShape,
    paths: &[Path],
    children: Vec<PositionedElement>,
) -> PositionedElement {
    if matches!(shape.geometry, ResolvedGeometry::Rectangle) {
        return clipped_group(Path::rect(local_shape_rect(shape)), children);
    }
    clipped_group(combined_clip_path(paths), children)
}

fn combined_clip_path(paths: &[Path]) -> Path {
    Path {
        commands: paths
            .iter()
            .flat_map(|path| path.commands.iter().cloned())
            .collect(),
        fill_rule: paths
            .first()
            .map_or(oxml_layout::FillRule::NonZero, |path| path.fill_rule),
    }
}

fn clipped_group(clip: Path, children: Vec<PositionedElement>) -> PositionedElement {
    PositionedElement::Group(GroupElement {
        transform: Transform::IDENTITY,
        clip: Some(clip),
        opacity: 1.0,
        effects: Vec::new(),
        children,
    })
}

fn tile_native_size_points(
    mut info: oxml_media::ImageInfo,
    declared_dpi: Option<f64>,
    media: MediaId,
) -> Result<(f64, f64), RenderInputError> {
    if let Some(dpi) = declared_dpi {
        if !dpi.is_finite() || dpi <= 0.0 {
            return Err(RenderInputError::InvalidPicture {
                media,
                detail: "declared picture DPI is not finite and positive",
            });
        }
        info.dpi_x = Some(dpi);
        info.dpi_y = Some(dpi);
    }
    let size = info
        .native_size(96.0)
        .ok_or(RenderInputError::InvalidPicture {
            media,
            detail: "picture DPI cannot produce a native size",
        })?;
    Ok((
        size.width_emu as f64 / 12_700.0,
        size.height_emu as f64 / 12_700.0,
    ))
}

fn tile_alignment_origin(
    rect: Rect,
    tile_width: f64,
    tile_height: f64,
    alignment: ResolvedRectAlignment,
) -> Point {
    let center_x = rect.x + (rect.width - tile_width) / 2.0;
    let right = rect.x + rect.width - tile_width;
    let center_y = rect.y + (rect.height - tile_height) / 2.0;
    let bottom = rect.y + rect.height - tile_height;
    match alignment {
        ResolvedRectAlignment::TopLeft => Point {
            x: rect.x,
            y: rect.y,
        },
        ResolvedRectAlignment::Top => Point {
            x: center_x,
            y: rect.y,
        },
        ResolvedRectAlignment::TopRight => Point {
            x: right,
            y: rect.y,
        },
        ResolvedRectAlignment::Left => Point {
            x: rect.x,
            y: center_y,
        },
        ResolvedRectAlignment::Center => Point {
            x: center_x,
            y: center_y,
        },
        ResolvedRectAlignment::Right => Point {
            x: right,
            y: center_y,
        },
        ResolvedRectAlignment::BottomLeft => Point {
            x: rect.x,
            y: bottom,
        },
        ResolvedRectAlignment::Bottom => Point {
            x: center_x,
            y: bottom,
        },
        ResolvedRectAlignment::BottomRight => Point {
            x: right,
            y: bottom,
        },
    }
}

fn repeated_origin(anchor: f64, coverage_start: f64, tile_size: f64) -> f64 {
    coverage_start - (coverage_start - anchor).rem_euclid(tile_size)
}

fn repeated_tile_index(
    origin: f64,
    anchor: f64,
    tile_size: f64,
    media: MediaId,
) -> Result<isize, RenderInputError> {
    let index = ((origin - anchor) / tile_size).round();
    if !index.is_finite() || index < isize::MIN as f64 || index > isize::MAX as f64 {
        return Err(RenderInputError::InvalidPicture {
            media,
            detail: "tile translation cannot preserve flip phase",
        });
    }
    Ok(index as isize)
}

fn repeat_count(
    origin: f64,
    coverage_end: f64,
    tile_size: f64,
    media: MediaId,
) -> Result<usize, RenderInputError> {
    let count = ((coverage_end - origin) / tile_size).ceil().max(0.0);
    if !count.is_finite() || count > usize::MAX as f64 {
        return Err(RenderInputError::TileLimitExceeded {
            media,
            requested: usize::MAX,
            limit: MAX_TILE_ELEMENTS,
        });
    }
    Ok(count as usize)
}

fn flip_tile(
    tile: PositionedElement,
    rect: Rect,
    flip: ResolvedTileFlip,
    odd_column: bool,
    odd_row: bool,
) -> PositionedElement {
    let flip_h =
        matches!(flip, ResolvedTileFlip::Horizontal | ResolvedTileFlip::Both) && odd_column;
    let flip_v = matches!(flip, ResolvedTileFlip::Vertical | ResolvedTileFlip::Both) && odd_row;
    if !flip_h && !flip_v {
        return tile;
    }
    PositionedElement::Group(GroupElement {
        transform: Transform {
            a: if flip_h { -1.0 } else { 1.0 },
            b: 0.0,
            c: 0.0,
            d: if flip_v { -1.0 } else { 1.0 },
            e: if flip_h {
                2.0 * rect.x + rect.width
            } else {
                0.0
            },
            f: if flip_v {
                2.0 * rect.y + rect.height
            } else {
                0.0
            },
        },
        clip: None,
        opacity: 1.0,
        effects: Vec::new(),
        children: vec![tile],
    })
}

#[derive(Clone, Copy)]
struct EndpointTangent {
    point: Point,
    outward: Point,
}

fn endpoint_tangents(paths: &[Path]) -> (Option<EndpointTangent>, Option<EndpointTangent>) {
    let mut head = None;
    let mut tail = None;
    for path in paths {
        let mut current = None;
        let mut subpath_start = None;
        for command in &path.commands {
            match *command {
                PathCommand::MoveTo(point) => {
                    current = Some(point);
                    subpath_start = Some(point);
                }
                PathCommand::LineTo(to) => {
                    if let Some(from) = current
                        && let Some(direction) = unit_direction(from, to)
                    {
                        head.get_or_insert(EndpointTangent {
                            point: from,
                            outward: Point {
                                x: -direction.x,
                                y: -direction.y,
                            },
                        });
                        tail = Some(EndpointTangent {
                            point: to,
                            outward: direction,
                        });
                    }
                    current = Some(to);
                }
                PathCommand::CurveTo { c1, c2, to } => {
                    if let Some(from) = current {
                        let start_direction = [c1, c2, to]
                            .into_iter()
                            .find_map(|candidate| unit_direction(from, candidate));
                        let end_direction = [c2, c1, from]
                            .into_iter()
                            .find_map(|candidate| unit_direction(candidate, to));
                        if let Some(direction) = start_direction {
                            head.get_or_insert(EndpointTangent {
                                point: from,
                                outward: Point {
                                    x: -direction.x,
                                    y: -direction.y,
                                },
                            });
                        }
                        if let Some(direction) = end_direction {
                            tail = Some(EndpointTangent {
                                point: to,
                                outward: direction,
                            });
                        }
                    }
                    current = Some(to);
                }
                PathCommand::Close => {
                    if let (Some(from), Some(to)) = (current, subpath_start)
                        && let Some(direction) = unit_direction(from, to)
                    {
                        head.get_or_insert(EndpointTangent {
                            point: from,
                            outward: Point {
                                x: -direction.x,
                                y: -direction.y,
                            },
                        });
                        tail = Some(EndpointTangent {
                            point: to,
                            outward: direction,
                        });
                    }
                    current = subpath_start;
                }
            }
        }
    }
    (head, tail)
}

fn unit_direction(from: Point, to: Point) -> Option<Point> {
    let dx = to.x - from.x;
    let dy = to.y - from.y;
    let length = dx.hypot(dy);
    (length.is_finite() && length > 1.0e-10).then_some(Point {
        x: dx / length,
        y: dy / length,
    })
}

fn line_end_path(
    end: &ResolvedLineEnd,
    tangent: EndpointTangent,
    stroke_width: f64,
) -> Option<Path> {
    if !stroke_width.is_finite() || stroke_width <= 0.0 {
        return None;
    }
    let width = line_end_factor(end.width) * stroke_width;
    let length = line_end_factor(end.length) * stroke_width;
    let point = |along: f64, across: f64| Point {
        x: tangent.point.x + tangent.outward.x * along - tangent.outward.y * across,
        y: tangent.point.y + tangent.outward.y * along + tangent.outward.x * across,
    };
    let path = match end.kind {
        ResolvedLineEndKind::Triangle => closed_polygon(vec![
            point(0.0, 0.0),
            point(-length, width / 2.0),
            point(-length, -width / 2.0),
        ]),
        ResolvedLineEndKind::Stealth => closed_polygon(vec![
            point(0.0, 0.0),
            point(-length, width / 2.0),
            point(-length / 2.0, 0.0),
            point(-length, -width / 2.0),
        ]),
        ResolvedLineEndKind::Diamond => closed_polygon(vec![
            point(0.0, 0.0),
            point(-length / 2.0, width / 2.0),
            point(-length, 0.0),
            point(-length / 2.0, -width / 2.0),
        ]),
        ResolvedLineEndKind::Oval => oval_line_end(&point, width, length),
        ResolvedLineEndKind::Arrow => {
            let arm = stroke_width.min(width / 2.0);
            closed_polygon(vec![
                point(0.0, 0.0),
                point(-length, width / 2.0),
                point(-length, width / 2.0 - arm),
                point(-arm, 0.0),
                point(-length, -width / 2.0 + arm),
                point(-length, -width / 2.0),
            ])
        }
    };
    path_is_finite(&path).then_some(path)
}

fn line_end_factor(size: ResolvedLineEndSize) -> f64 {
    match size {
        ResolvedLineEndSize::Small => 2.0,
        ResolvedLineEndSize::Medium => 3.0,
        ResolvedLineEndSize::Large => 5.0,
    }
}

fn closed_polygon(points: Vec<Point>) -> Path {
    let mut points = points.into_iter();
    let mut commands = points
        .next()
        .map(PathCommand::MoveTo)
        .into_iter()
        .collect::<Vec<_>>();
    commands.extend(points.map(PathCommand::LineTo));
    commands.push(PathCommand::Close);
    Path {
        commands,
        fill_rule: oxml_layout::FillRule::NonZero,
    }
}

fn oval_line_end(point: &impl Fn(f64, f64) -> Point, width: f64, length: f64) -> Path {
    const KAPPA: f64 = 0.552_284_749_830_793_6;
    let rx = length / 2.0;
    let ry = width / 2.0;
    let center = -rx;
    Path {
        commands: vec![
            PathCommand::MoveTo(point(0.0, 0.0)),
            PathCommand::CurveTo {
                c1: point(0.0, KAPPA * ry),
                c2: point(center + KAPPA * rx, ry),
                to: point(center, ry),
            },
            PathCommand::CurveTo {
                c1: point(center - KAPPA * rx, ry),
                c2: point(-length, KAPPA * ry),
                to: point(-length, 0.0),
            },
            PathCommand::CurveTo {
                c1: point(-length, -KAPPA * ry),
                c2: point(center - KAPPA * rx, -ry),
                to: point(center, -ry),
            },
            PathCommand::CurveTo {
                c1: point(center + KAPPA * rx, -ry),
                c2: point(0.0, -KAPPA * ry),
                to: point(0.0, 0.0),
            },
            PathCommand::Close,
        ],
        fill_rule: oxml_layout::FillRule::NonZero,
    }
}

fn path_is_finite(path: &Path) -> bool {
    path.commands.iter().all(|command| match command {
        PathCommand::MoveTo(point) | PathCommand::LineTo(point) => point_is_finite(*point),
        PathCommand::CurveTo { c1, c2, to } => {
            point_is_finite(*c1) && point_is_finite(*c2) && point_is_finite(*to)
        }
        PathCommand::Close => true,
    })
}

fn point_is_finite(point: Point) -> bool {
    point.x.is_finite() && point.y.is_finite()
}

fn filled_line_end(path: Path, paint: &Paint) -> PositionedElement {
    PositionedElement::Path(PathElement {
        path,
        fill: Some(paint.clone()),
        stroke: None,
    })
}

fn shape_transform(shape: &ResolvedShape) -> Transform {
    let center_x = shape.bounds.width / 2.0;
    let center_y = shape.bounds.height / 2.0;
    let rotation = Transform::rotate_about(shape.rotation_deg, center_x, center_y);
    let flip = Transform {
        a: if shape.flip_h { -1.0 } else { 1.0 },
        b: 0.0,
        c: 0.0,
        d: if shape.flip_v { -1.0 } else { 1.0 },
        e: if shape.flip_h {
            shape.bounds.width
        } else {
            0.0
        },
        f: if shape.flip_v {
            shape.bounds.height
        } else {
            0.0
        },
    };
    let translation = Transform {
        e: shape.bounds.x,
        f: shape.bounds.y,
        ..Transform::IDENTITY
    };
    rotation
        .then(flip)
        .then(translation)
        .then(shape.group_transform)
}

/// Resolve one scoped media relationship into the deck's content-addressed store.
pub fn resolve_media_relationship(
    relationships: &RelScopes,
    scope: RelScope,
    relationship_id: &str,
    package_media: &HashMap<String, MediaData>,
    deck_media: &mut HashMap<MediaId, MediaData>,
) -> Result<MediaId, RenderInputError> {
    let relationship = relationships.get(scope, relationship_id)?;
    let media = package_media.get(&relationship.target).ok_or_else(|| {
        RenderInputError::MissingMediaTarget {
            scope,
            relationship_id: relationship_id.to_owned(),
            target: relationship.target.clone(),
        }
    })?;
    let media_id = MediaId::from_bytes(&media.bytes);
    deck_media.entry(media_id).or_insert_with(|| media.clone());
    Ok(media_id)
}

#[cfg(test)]
mod tests {
    use super::*;
    use oxml_drawing::color::ColorMap;
    use oxml_drawing::text::CT_TextListStyle;
    use oxml_layout::{
        Color, Diagnostic, Effect, FieldKind, FillRule, GradientStop, GroupElement, Paint, Path,
        PathCommand, Point, PositionedElement, Rect, Stroke, Transform, walk,
    };
    use rpptx_layout::{
        ResolveCtx, ResolvedAutofit, ResolvedContent, ResolvedGeometry, ResolvedParagraph,
        ResolvedRunStyle, ResolvedShape, ResolvedTable, ResolvedTableBorder, ResolvedTableCell,
        ResolvedTableRow, ResolvedTextBody, ResolvedTextRun, TextAnchor, TextDirection, TextInsets,
    };

    const IMAGE_RELATIONSHIP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    fn media(bytes: &[u8]) -> MediaData {
        MediaData {
            bytes: bytes.to_vec(),
            content_type: "image/png".to_owned(),
        }
    }

    fn relationship(target: &str) -> ResolvedRel {
        ResolvedRel {
            target: target.to_owned(),
            relationship_type: IMAGE_RELATIONSHIP.to_owned(),
            target_mode: None,
        }
    }

    fn hyperlink_relationship(target: &str, target_mode: Option<&str>) -> ResolvedRel {
        ResolvedRel {
            target: target.to_owned(),
            relationship_type: HYPERLINK_RELATIONSHIP.to_owned(),
            target_mode: target_mode.map(str::to_owned),
        }
    }

    fn color(red: f64, green: f64, blue: f64) -> Color {
        Color {
            r: red,
            g: green,
            b: blue,
            a: 1.0,
        }
    }

    fn shape(
        bounds: Rect,
        geometry: ResolvedGeometry,
        fill: Option<Paint>,
        line: Option<Stroke>,
    ) -> ResolvedShape {
        ResolvedShape {
            group_transform: Transform::IDENTITY,
            bounds,
            rotation_deg: 0.0,
            flip_h: false,
            flip_v: false,
            geometry,
            fill,
            image_fill: None,
            line,
            head_end: None,
            tail_end: None,
            shadow: None,
            content: ResolvedContent::None,
            unsupported: None,
        }
    }

    fn table_shape(table: ResolvedTable) -> ResolvedShape {
        let mut table_shape = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 20.0,
                height: 20.0,
            },
            ResolvedGeometry::Rectangle,
            None,
            None,
        );
        table_shape.content = ResolvedContent::Table(table);
        table_shape
    }

    fn table_text(value: &str) -> ResolvedTextBody {
        ResolvedTextBody {
            insets: TextInsets::default(),
            anchor: TextAnchor::Top,
            wrap: true,
            vertical: TextDirection::Horizontal,
            space_first_last_paragraph: false,
            autofit: ResolvedAutofit::None,
            paragraphs: vec![ResolvedParagraph {
                runs: vec![ResolvedTextRun::Text {
                    text: value.to_owned(),
                    style: ResolvedRunStyle {
                        font_size: Some(6.0),
                        ..ResolvedRunStyle::default()
                    },
                }],
                ..ResolvedParagraph::default()
            }],
        }
    }

    fn table_border() -> ResolvedTableBorder {
        ResolvedTableBorder {
            stroke: Some(Stroke::new(Paint::Solid(Color::BLACK), 1.0)),
            priority: 1,
        }
    }

    fn linked_field_shape(group_transform: Transform) -> ResolvedShape {
        let mut linked = shape(
            Rect {
                x: 10.0,
                y: 20.0,
                width: 120.0,
                height: 40.0,
            },
            ResolvedGeometry::Rectangle,
            None,
            None,
        );
        linked.group_transform = group_transform;
        linked.content = ResolvedContent::Text(ResolvedTextBody {
            insets: TextInsets::default(),
            anchor: TextAnchor::Top,
            wrap: true,
            vertical: TextDirection::Horizontal,
            space_first_last_paragraph: false,
            autofit: ResolvedAutofit::None,
            paragraphs: vec![ResolvedParagraph {
                runs: vec![
                    ResolvedTextRun::Field {
                        text: "stored".to_owned(),
                        field_type: Some("slidenum".to_owned()),
                        style: ResolvedRunStyle {
                            font_size: Some(12.0),
                            ..ResolvedRunStyle::default()
                        },
                    },
                    ResolvedTextRun::Text {
                        text: " linked".to_owned(),
                        style: ResolvedRunStyle {
                            font_size: Some(12.0),
                            hyperlink_url: Some("https://example.com/deck".to_owned()),
                            ..ResolvedRunStyle::default()
                        },
                    },
                ],
                ..ResolvedParagraph::default()
            }],
        });
        linked
    }

    fn rendered_text_and_links(
        page: &PageFrame,
    ) -> (Vec<(String, Option<FieldKind>)>, Vec<String>) {
        let mut text = Vec::new();
        let mut links = Vec::new();
        walk(&page.elements, &mut |element, _| match element {
            PositionedElement::Text(run) => text.push((run.text.clone(), run.field_kind)),
            PositionedElement::LinkAnnotation { url, .. } => links.push(url.clone()),
            _ => {}
        });
        (text, links)
    }

    fn transformed_link_rect(page: &PageFrame) -> Rect {
        let mut result = None;
        walk(&page.elements, &mut |element, transform| {
            if let PositionedElement::LinkAnnotation { rect, .. } = element {
                let top_left = transform.apply(Point {
                    x: rect.x,
                    y: rect.y,
                });
                let bottom_right = transform.apply(Point {
                    x: rect.x + rect.width,
                    y: rect.y + rect.height,
                });
                result = Some(Rect {
                    x: top_left.x,
                    y: top_left.y,
                    width: bottom_right.x - top_left.x,
                    height: bottom_right.y - top_left.y,
                });
            }
        });
        result.expect("rendered hyperlink annotation")
    }

    #[test]
    fn slide_number_field_renders_current_page_and_hyperlink_emits_annotation() {
        let input = render_input(vec![
            slide((200.0, 100.0), Vec::new()),
            slide(
                (200.0, 100.0),
                vec![linked_field_shape(Transform::IDENTITY)],
            ),
        ]);
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");

        let page = layout_slide_with_fonts(&input, 1, &mut fonts).expect("render slide two");
        let (text, links) = rendered_text_and_links(&page);

        assert_eq!(text[0], ("2".to_owned(), Some(FieldKind::Page)));
        assert!(!links.is_empty());
        assert!(links.iter().all(|url| url == "https://example.com/deck"));
    }

    #[test]
    fn grouped_hyperlink_annotation_keeps_transformed_run_bounds() {
        let translation = Transform {
            e: 30.0,
            f: 40.0,
            ..Transform::IDENTITY
        };
        let plain_input = render_input(vec![slide(
            (240.0, 180.0),
            vec![linked_field_shape(Transform::IDENTITY)],
        )]);
        let grouped_input = render_input(vec![slide(
            (240.0, 180.0),
            vec![linked_field_shape(translation)],
        )]);
        let mut plain_fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let mut grouped_fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let plain = layout_slide_with_fonts(&plain_input, 0, &mut plain_fonts)
            .expect("render ungrouped hyperlink");
        let grouped = layout_slide_with_fonts(&grouped_input, 0, &mut grouped_fonts)
            .expect("render grouped hyperlink");
        let plain_rect = transformed_link_rect(&plain);
        let grouped_rect = transformed_link_rect(&grouped);

        assert!((grouped_rect.x - plain_rect.x - 30.0).abs() < 1.0e-10);
        assert!((grouped_rect.y - plain_rect.y - 40.0).abs() < 1.0e-10);
        assert!((grouped_rect.width - plain_rect.width).abs() < 1.0e-10);
        assert!((grouped_rect.height - plain_rect.height).abs() < 1.0e-10);
    }

    #[test]
    fn banded_merged_table_renders_correct_fills_without_duplicated_borders() {
        let table = ResolvedTable {
            right_to_left: false,
            column_widths: vec![10.0, 10.0],
            rows: vec![
                ResolvedTableRow {
                    height: 10.0,
                    cells: vec![
                        ResolvedTableCell {
                            fill: Some(Paint::Solid(Color::from_hex("FF0000"))),
                            text: Some(table_text("merged")),
                            left: Some(table_border()),
                            right: Some(table_border()),
                            top: Some(table_border()),
                            bottom: Some(table_border()),
                            grid_span: 2,
                            row_span: 1,
                            ..ResolvedTableCell::default()
                        },
                        ResolvedTableCell {
                            horizontal_merge: true,
                            ..ResolvedTableCell::default()
                        },
                    ],
                },
                ResolvedTableRow {
                    height: 10.0,
                    cells: vec![
                        ResolvedTableCell {
                            fill: Some(Paint::Solid(Color::from_hex("00FF00"))),
                            left: Some(table_border()),
                            right: Some(table_border()),
                            top: Some(table_border()),
                            bottom: Some(table_border()),
                            ..ResolvedTableCell::default()
                        },
                        ResolvedTableCell {
                            fill: Some(Paint::Solid(Color::from_hex("0000FF"))),
                            left: Some(table_border()),
                            right: Some(table_border()),
                            top: Some(table_border()),
                            bottom: Some(table_border()),
                            ..ResolvedTableCell::default()
                        },
                    ],
                },
            ],
        };
        let layout = layout_presentation(&render_input(vec![slide(
            (20.0, 20.0),
            vec![table_shape(table)],
        )]))
        .unwrap();
        let png = oxml_pdf::render_page_to_png(&layout, 0, 72.0).unwrap();
        let pixmap = tiny_skia::Pixmap::decode_png(&png).unwrap();

        let red = rgb_at(&pixmap, 5, 5);
        assert!(red.0 > 200 && red.1 < 30 && red.2 < 30, "{red:?}");
        assert_eq!(rgb_at(&pixmap, 5, 15), (0, 255, 0));
        assert_eq!(rgb_at(&pixmap, 15, 15), (0, 0, 255));
        let group = only_group(&layout.pages[0].elements[0]);
        assert_eq!(
            group
                .children
                .iter()
                .filter(|element| matches!(element, PositionedElement::Path(path) if path.stroke.is_some()))
                .count(),
            11,
            "the merged top row removes its internal vertical segment"
        );
        assert!(group.children.iter().any(
            |element| matches!(element, PositionedElement::Text(run) if run.text == "merged")
        ));
    }

    #[test]
    fn merged_continuation_cells_do_not_render_fill_border_or_text_twice() {
        let table = ResolvedTable {
            right_to_left: false,
            column_widths: vec![10.0, 10.0],
            rows: vec![ResolvedTableRow {
                height: 10.0,
                cells: vec![
                    ResolvedTableCell {
                        fill: Some(Paint::Solid(Color::BLACK)),
                        grid_span: 2,
                        ..ResolvedTableCell::default()
                    },
                    ResolvedTableCell {
                        fill: Some(Paint::Solid(Color::WHITE)),
                        horizontal_merge: true,
                        ..ResolvedTableCell::default()
                    },
                ],
            }],
        };
        let page = layout_slide(
            &render_input(vec![slide((20.0, 10.0), vec![table_shape(table)])]),
            0,
        )
        .unwrap();
        let group = only_group(&page.elements[0]);

        assert_eq!(group.children.iter().filter(|element| matches!(element, PositionedElement::Path(path) if path.fill.is_some())).count(), 1);
    }

    #[test]
    fn right_to_left_table_keeps_unequal_logical_column_widths() {
        let table = ResolvedTable {
            right_to_left: true,
            column_widths: vec![10.0, 20.0],
            rows: vec![ResolvedTableRow {
                height: 10.0,
                cells: vec![
                    ResolvedTableCell {
                        fill: Some(Paint::Solid(Color::from_hex("FF0000"))),
                        ..ResolvedTableCell::default()
                    },
                    ResolvedTableCell {
                        fill: Some(Paint::Solid(Color::from_hex("0000FF"))),
                        ..ResolvedTableCell::default()
                    },
                ],
            }],
        };
        let layout = layout_presentation(&render_input(vec![slide(
            (30.0, 10.0),
            vec![table_shape(table)],
        )]))
        .unwrap();
        let png = oxml_pdf::render_page_to_png(&layout, 0, 72.0).unwrap();
        let pixmap = tiny_skia::Pixmap::decode_png(&png).unwrap();

        assert_eq!(rgb_at(&pixmap, 5, 5), (0, 0, 255));
        assert_eq!(rgb_at(&pixmap, 25, 5), (255, 0, 0));
    }

    #[test]
    fn merged_table_uses_far_continuation_border() {
        let outer = ResolvedTableBorder {
            stroke: Some(Stroke::new(Paint::Solid(Color::from_hex("FF0000")), 3.0)),
            priority: 2,
        };
        let table = ResolvedTable {
            right_to_left: false,
            column_widths: vec![10.0, 10.0],
            rows: vec![ResolvedTableRow {
                height: 10.0,
                cells: vec![
                    ResolvedTableCell {
                        grid_span: 2,
                        right: Some(table_border()),
                        ..ResolvedTableCell::default()
                    },
                    ResolvedTableCell {
                        horizontal_merge: true,
                        right: Some(outer),
                        ..ResolvedTableCell::default()
                    },
                ],
            }],
        };
        let page = layout_slide(
            &render_input(vec![slide((20.0, 10.0), vec![table_shape(table)])]),
            0,
        )
        .unwrap();
        let group = only_group(&page.elements[0]);
        let far_border = group.children.iter().find_map(|element| {
            let PositionedElement::Path(path) = element else {
                return None;
            };
            match path.path.commands.as_slice() {
                [PathCommand::MoveTo(start), PathCommand::LineTo(end)]
                    if start.x == 20.0 && end.x == 20.0 =>
                {
                    path.stroke.as_ref()
                }
                _ => None,
            }
        });

        let far_border = far_border.expect("merged far edge should be emitted");
        assert_eq!(far_border.width, 3.0);
        assert_eq!(far_border.paint, Paint::Solid(Color::from_hex("FF0000")));
    }

    #[test]
    fn table_cell_margins_place_text_in_the_fixed_content_box() {
        let table = ResolvedTable {
            right_to_left: false,
            column_widths: vec![20.0],
            rows: vec![ResolvedTableRow {
                height: 20.0,
                cells: vec![ResolvedTableCell {
                    text: Some(table_text("cell")),
                    margins: TextInsets {
                        left: 2.0,
                        top: 3.0,
                        right: 4.0,
                        bottom: 5.0,
                    },
                    ..ResolvedTableCell::default()
                }],
            }],
        };
        let page = layout_slide(
            &render_input(vec![slide((20.0, 20.0), vec![table_shape(table)])]),
            0,
        )
        .unwrap();
        let group = only_group(&page.elements[0]);
        let PositionedElement::Text(run) = group
            .children
            .iter()
            .find(|element| matches!(element, PositionedElement::Text(_)))
            .expect("cell text should remain visible")
        else {
            unreachable!()
        };

        assert!(run.origin.x >= 2.0);
        assert!(run.origin.y >= 3.0);
    }

    fn slide(size: (f64, f64), shapes: Vec<ResolvedShape>) -> ResolvedSlide {
        ResolvedSlide {
            size,
            background: None,
            shapes,
            diagnostics: Vec::new(),
        }
    }

    fn render_input(slides: Vec<ResolvedSlide>) -> RenderInput {
        RenderInput {
            slides,
            media: HashMap::new(),
            fonts: Vec::new(),
            metadata: None,
        }
    }

    fn only_group(element: &PositionedElement) -> &GroupElement {
        let PositionedElement::Group(group) = element else {
            panic!("shape should lower to one group");
        };
        group
    }

    #[test]
    fn timeline_lowering_retains_hidden_identity_slots_and_uses_shape_local_clips() {
        let first = shape(
            Rect {
                x: 10.0,
                y: 20.0,
                width: 40.0,
                height: 20.0,
            },
            ResolvedGeometry::Rectangle,
            Some(Paint::Solid(Color::WHITE)),
            None,
        );
        let second = shape(
            Rect {
                x: 60.0,
                y: 20.0,
                width: 20.0,
                height: 20.0,
            },
            ResolvedGeometry::Rectangle,
            Some(Paint::Solid(Color::BLACK)),
            None,
        );
        let input = render_input(vec![slide((100.0, 60.0), vec![first, second])]);
        let mut hidden_state = rpptx_layout::timeline::EvaluatedShapeState::default();
        hidden_state.visible = false;
        hidden_state.clip = Some(Rect {
            x: 0.0,
            y: 0.0,
            width: 0.25,
            height: 1.0,
        });
        let states = [
            hidden_state,
            rpptx_layout::timeline::EvaluatedShapeState::default(),
        ];
        let mut fonts = FontManager::new_deterministic().unwrap();
        let page = layout_slide_with_fonts_text_directions_and_states(
            &input,
            0,
            &mut fonts,
            None,
            Some(&states),
        )
        .unwrap();

        assert_eq!(page.elements.len(), 2);
        let hidden = only_group(&page.elements[0]);
        assert_eq!(hidden.opacity, 0.0);
        assert_eq!(hidden.transform.e, 10.0);
        assert_eq!(
            hidden.clip,
            Some(Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 20.0,
            }))
        );
        assert_eq!(only_group(&page.elements[1]).opacity, 1.0);
    }

    #[test]
    fn resolved_outer_shadow_is_lowered_to_the_shape_group() {
        let effect = Effect::OuterShadow {
            dx: 3.0,
            dy: 4.0,
            blur: 2.0,
            color: Color {
                r: 0.5,
                g: 0.25,
                b: 0.0,
                a: 0.75,
            },
        };
        let mut shadowed = shape(
            Rect {
                x: 2.0,
                y: 3.0,
                width: 8.0,
                height: 6.0,
            },
            ResolvedGeometry::Rectangle,
            Some(Paint::Solid(Color::WHITE)),
            None,
        );
        shadowed.shadow = Some(effect.clone());

        let page =
            layout_slide(&render_input(vec![slide((20.0, 20.0), vec![shadowed])]), 0).unwrap();

        assert_eq!(only_group(&page.elements[0]).effects, vec![effect]);
    }

    fn assert_point_close(actual: Point, expected: Point) {
        const EPSILON: f64 = 1.0e-10;
        assert!(
            (actual.x - expected.x).abs() < EPSILON && (actual.y - expected.y).abs() < EPSILON,
            "expected ({}, {}), got ({}, {})",
            expected.x,
            expected.y,
            actual.x,
            actual.y
        );
    }

    #[test]
    fn rotated_shape_corners_match_hand_computed_coordinates() {
        let mut rotated = shape(
            Rect {
                x: 10.0,
                y: 20.0,
                width: 8.0,
                height: 4.0,
            },
            ResolvedGeometry::Rectangle,
            Some(Paint::Solid(Color::BLACK)),
            None,
        );
        rotated.rotation_deg = 30.0;
        let page = layout_slide(&render_input(vec![slide((40.0, 40.0), vec![rotated])]), 0)
            .expect("lower rotated shape");
        let transform = only_group(&page.elements[0]).transform;
        let radians = 30.0_f64.to_radians();
        let (sin, cos) = radians.sin_cos();

        for corner in [
            Point { x: 0.0, y: 0.0 },
            Point { x: 8.0, y: 0.0 },
            Point { x: 0.0, y: 4.0 },
            Point { x: 8.0, y: 4.0 },
        ] {
            let dx = corner.x - 4.0;
            let dy = corner.y - 2.0;
            let expected = Point {
                x: 10.0 + 4.0 + cos * dx - sin * dy,
                y: 20.0 + 2.0 + sin * dx + cos * dy,
            };
            assert_point_close(transform.apply(corner), expected);
        }
    }

    #[test]
    fn horizontal_and_vertical_flips_are_about_the_shape_centre() {
        let bounds = Rect {
            x: 10.0,
            y: 20.0,
            width: 8.0,
            height: 4.0,
        };
        let mut horizontal = shape(
            bounds,
            ResolvedGeometry::Rectangle,
            Some(Paint::Solid(Color::BLACK)),
            None,
        );
        horizontal.flip_h = true;
        let mut vertical = horizontal.clone();
        vertical.flip_h = false;
        vertical.flip_v = true;
        let page = layout_slide(
            &render_input(vec![slide((40.0, 40.0), vec![horizontal, vertical])]),
            0,
        )
        .expect("lower flipped shapes");
        let horizontal = only_group(&page.elements[0]).transform;
        let vertical = only_group(&page.elements[1]).transform;

        assert_point_close(
            horizontal.apply(Point { x: 4.0, y: 2.0 }),
            Point { x: 14.0, y: 22.0 },
        );
        assert_point_close(
            horizontal.apply(Point { x: 0.0, y: 0.0 }),
            Point { x: 18.0, y: 20.0 },
        );
        assert_point_close(
            horizontal.apply(Point { x: 8.0, y: 4.0 }),
            Point { x: 10.0, y: 24.0 },
        );
        assert_point_close(
            vertical.apply(Point { x: 4.0, y: 2.0 }),
            Point { x: 14.0, y: 22.0 },
        );
        assert_point_close(
            vertical.apply(Point { x: 0.0, y: 0.0 }),
            Point { x: 10.0, y: 24.0 },
        );
        assert_point_close(
            vertical.apply(Point { x: 8.0, y: 4.0 }),
            Point { x: 18.0, y: 20.0 },
        );
    }

    #[test]
    fn nested_group_transform_applies_child_before_parent() {
        let mut nested = shape(
            Rect {
                x: 10.0,
                y: 20.0,
                width: 8.0,
                height: 4.0,
            },
            ResolvedGeometry::Rectangle,
            Some(Paint::Solid(Color::BLACK)),
            None,
        );
        nested.rotation_deg = 90.0;
        nested.flip_h = true;
        nested.group_transform = Transform {
            a: 2.0,
            b: 0.0,
            c: 0.0,
            d: 3.0,
            e: 5.0,
            f: 7.0,
        };
        let page = layout_slide(&render_input(vec![slide((80.0, 100.0), vec![nested])]), 0)
            .expect("lower nested shape");
        let transform = only_group(&page.elements[0]).transform;

        assert_point_close(
            transform.apply(Point { x: 0.0, y: 0.0 }),
            Point { x: 29.0, y: 61.0 },
        );
        assert_point_close(
            transform.apply(Point { x: 8.0, y: 4.0 }),
            Point { x: 37.0, y: 85.0 },
        );
    }

    #[test]
    fn group_mapping_does_not_clip_a_child_outside_group_bounds() {
        let mut outside = shape(
            Rect {
                x: 40.0,
                y: 20.0,
                width: 8.0,
                height: 4.0,
            },
            ResolvedGeometry::Rectangle,
            Some(Paint::Solid(Color::BLACK)),
            None,
        );
        outside.group_transform = Transform {
            e: 30.0,
            f: 10.0,
            ..Transform::IDENTITY
        };
        let page = layout_slide(&render_input(vec![slide((60.0, 40.0), vec![outside])]), 0)
            .expect("lower grouped child outside nominal group bounds");
        let group = only_group(&page.elements[0]);

        assert_eq!(group.clip, None);
        assert_point_close(
            group.transform.apply(Point { x: 0.0, y: 0.0 }),
            Point { x: 70.0, y: 30.0 },
        );
    }

    #[test]
    fn rotated_gradient_and_outline_share_the_shape_transform() {
        let red = color(1.0, 0.0, 0.0);
        let blue = color(0.0, 0.0, 1.0);
        let mut rotated = shape(
            Rect {
                x: 8.0,
                y: 8.0,
                width: 12.0,
                height: 6.0,
            },
            ResolvedGeometry::Rectangle,
            Some(Paint::linear(
                Point { x: 0.0, y: 0.0 },
                Point { x: 12.0, y: 0.0 },
                vec![
                    GradientStop {
                        offset: 0.0,
                        color: red,
                    },
                    GradientStop {
                        offset: 0.49,
                        color: red,
                    },
                    GradientStop {
                        offset: 0.51,
                        color: blue,
                    },
                    GradientStop {
                        offset: 1.0,
                        color: blue,
                    },
                ],
                (true, true),
            )),
            Some(Stroke::new(Paint::Solid(Color::BLACK), 2.0)),
        );
        rotated.rotation_deg = 90.0;
        let layout = layout_presentation(&render_input(vec![slide((28.0, 24.0), vec![rotated])]))
            .expect("lower rotated gradient");
        let png =
            oxml_pdf::render_page_to_png(&layout, 0, 72.0).expect("rasterise rotated gradient");
        let pixmap = tiny_skia::Pixmap::decode_png(&png).expect("decode rotated gradient");
        let rgb = |x, y| {
            let pixel = pixmap.pixel(x, y).expect("sample lies inside page");
            (pixel.red(), pixel.green(), pixel.blue())
        };

        assert_eq!(rgb(14, 7), (255, 0, 0));
        assert_eq!(rgb(14, 15), (0, 0, 255));
        assert_eq!(rgb(11, 11), (0, 0, 0));
        assert_eq!(rgb(8, 8), (255, 255, 255));
    }

    #[test]
    fn solid_gradient_and_outlined_shapes_rasterise_at_sampled_pixels() {
        let red = color(1.0, 0.0, 0.0);
        let blue = color(0.0, 0.0, 1.0);
        let green = color(0.0, 1.0, 0.0);
        let gradient = Paint::linear(
            Point { x: 0.0, y: 0.0 },
            Point { x: 8.0, y: 0.0 },
            vec![
                GradientStop {
                    offset: 0.0,
                    color: red,
                },
                GradientStop {
                    offset: 0.49,
                    color: red,
                },
                GradientStop {
                    offset: 0.51,
                    color: blue,
                },
                GradientStop {
                    offset: 1.0,
                    color: blue,
                },
            ],
            (true, true),
        );
        let input = render_input(vec![slide(
            (40.0, 14.0),
            vec![
                shape(
                    Rect {
                        x: 2.0,
                        y: 2.0,
                        width: 8.0,
                        height: 8.0,
                    },
                    ResolvedGeometry::Rectangle,
                    Some(Paint::Solid(red)),
                    None,
                ),
                shape(
                    Rect {
                        x: 14.0,
                        y: 2.0,
                        width: 8.0,
                        height: 8.0,
                    },
                    ResolvedGeometry::Rectangle,
                    Some(gradient),
                    None,
                ),
                shape(
                    Rect {
                        x: 26.0,
                        y: 2.0,
                        width: 8.0,
                        height: 8.0,
                    },
                    ResolvedGeometry::Rectangle,
                    None,
                    Some(Stroke::new(Paint::Solid(green), 2.0)),
                ),
            ],
        )]);

        let layout = layout_presentation(&input).expect("lower shape slide");
        let png = oxml_pdf::render_page_to_png(&layout, 0, 72.0).expect("rasterise shape slide");
        let pixmap = tiny_skia::Pixmap::decode_png(&png).expect("decode shape slide");
        let rgb = |x, y| {
            let pixel = pixmap.pixel(x, y).expect("sample lies inside page");
            (pixel.red(), pixel.green(), pixel.blue())
        };

        assert_eq!(rgb(5, 5), (255, 0, 0));
        assert_eq!(rgb(15, 5), (255, 0, 0));
        assert_eq!(rgb(20, 5), (0, 0, 255));
        assert_eq!(rgb(26, 5), (0, 255, 0));
        assert_eq!(rgb(30, 5), (255, 255, 255));
        assert_eq!(rgb(38, 5), (255, 255, 255));
    }

    #[test]
    fn preset_and_custom_geometry_lower_to_ordered_paths() {
        let first = Path {
            commands: vec![
                PathCommand::MoveTo(Point { x: 0.0, y: 0.0 }),
                PathCommand::LineTo(Point { x: 4.0, y: 0.0 }),
            ],
            fill_rule: FillRule::NonZero,
        };
        let second = Path {
            commands: vec![
                PathCommand::MoveTo(Point { x: 0.0, y: 1.0 }),
                PathCommand::LineTo(Point { x: 4.0, y: 1.0 }),
            ],
            fill_rule: FillRule::EvenOdd,
        };
        let fill = Paint::Solid(Color::BLACK);
        let line = Stroke::new(Paint::Solid(Color::WHITE), 2.0);
        let input = render_input(vec![slide(
            (20.0, 20.0),
            vec![
                shape(
                    Rect {
                        x: 2.0,
                        y: 3.0,
                        width: 4.0,
                        height: 5.0,
                    },
                    ResolvedGeometry::Rectangle,
                    Some(fill.clone()),
                    Some(line.clone()),
                ),
                shape(
                    Rect {
                        x: 8.0,
                        y: 9.0,
                        width: 4.0,
                        height: 5.0,
                    },
                    ResolvedGeometry::Custom {
                        paths: vec![first.clone(), second.clone()],
                        text_rect: None,
                    },
                    Some(fill.clone()),
                    Some(line.clone()),
                ),
            ],
        )]);

        let page = layout_slide(&input, 0).expect("lower first slide");
        assert_eq!(page.elements.len(), 2);
        let rectangle = only_group(&page.elements[0]);
        assert_eq!(
            rectangle.transform,
            Transform {
                e: 2.0,
                f: 3.0,
                ..Transform::IDENTITY
            }
        );
        let PositionedElement::Path(rectangle) = &rectangle.children[0] else {
            panic!("rectangle should lower to a path");
        };
        assert_eq!(
            rectangle.path,
            Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 4.0,
                height: 5.0
            })
        );
        assert_eq!(rectangle.fill, Some(fill.clone()));
        assert_eq!(rectangle.stroke, Some(line.clone()));

        let custom = only_group(&page.elements[1]);
        assert_eq!(custom.children.len(), 2);
        for (element, expected) in custom.children.iter().zip([first, second]) {
            let PositionedElement::Path(element) = element else {
                panic!("custom geometry should lower to paths");
            };
            assert_eq!(element.path, expected);
            assert_eq!(element.fill, Some(fill.clone()));
            assert_eq!(element.stroke, Some(line.clone()));
        }
    }

    #[test]
    fn bounds_fallback_emits_a_visible_black_outline() {
        let input = render_input(vec![slide(
            (20.0, 20.0),
            vec![shape(
                Rect {
                    x: 2.0,
                    y: 3.0,
                    width: 4.0,
                    height: 5.0,
                },
                ResolvedGeometry::BoundsFallback,
                None,
                None,
            )],
        )]);

        let page = layout_slide(&input, 0).expect("lower fallback slide");
        let group = only_group(&page.elements[0]);
        let PositionedElement::Path(path) = &group.children[0] else {
            panic!("fallback should lower to a path");
        };
        assert_eq!(path.fill, None);
        assert_eq!(
            path.stroke,
            Some(Stroke::new(Paint::Solid(Color::BLACK), 1.0))
        );
    }

    #[test]
    fn triangular_tail_end_emits_an_extra_filled_path() {
        let paint = Paint::Solid(color(1.0, 0.0, 0.0));
        let mut arrow = shape(
            Rect {
                x: 2.0,
                y: 3.0,
                width: 10.0,
                height: 10.0,
            },
            ResolvedGeometry::Custom {
                paths: vec![open_line(
                    Point { x: 0.0, y: 5.0 },
                    Point { x: 10.0, y: 5.0 },
                )],
                text_rect: None,
            },
            None,
            Some(Stroke::new(paint.clone(), 2.0)),
        );
        arrow.tail_end = Some(line_end(ResolvedLineEndKind::Triangle));

        let layout = layout_presentation(&render_input(vec![slide((20.0, 20.0), vec![arrow])]))
            .expect("lower triangular tail");
        let group = only_group(&layout.pages[0].elements[0]);
        assert_eq!(group.children.len(), 2);
        let PositionedElement::Path(end) = &group.children[1] else {
            panic!("tail end should lower to a path");
        };
        assert_eq!(end.fill, Some(paint));
        assert_eq!(end.stroke, None);
        assert!(matches!(end.path.commands.last(), Some(PathCommand::Close)));
        assert_eq!(
            end.path.commands.first(),
            Some(&PathCommand::MoveTo(Point { x: 10.0, y: 5.0 }))
        );

        let png =
            oxml_pdf::render_page_to_png(&layout, 0, 72.0).expect("rasterise triangular tail");
        let pixmap = tiny_skia::Pixmap::decode_png(&png).expect("decode triangular tail");
        let endpoint_pixel = pixmap.pixel(7, 6).expect("sample lies inside page");
        assert_eq!(
            (
                endpoint_pixel.red(),
                endpoint_pixel.green(),
                endpoint_pixel.blue()
            ),
            (255, 0, 0)
        );
    }

    #[test]
    fn head_end_uses_the_reversed_start_tangent() {
        let mut arrow = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 20.0,
                height: 10.0,
            },
            ResolvedGeometry::Custom {
                paths: vec![open_line(
                    Point { x: 5.0, y: 5.0 },
                    Point { x: 15.0, y: 5.0 },
                )],
                text_rect: None,
            },
            None,
            Some(Stroke::new(Paint::Solid(Color::BLACK), 2.0)),
        );
        arrow.head_end = Some(line_end(ResolvedLineEndKind::Triangle));
        arrow.tail_end = Some(line_end(ResolvedLineEndKind::Triangle));

        let page = layout_slide(&render_input(vec![slide((20.0, 10.0), vec![arrow])]), 0)
            .expect("lower opposite ends");
        let group = only_group(&page.elements[0]);
        let PositionedElement::Path(head) = &group.children[1] else {
            panic!("head end should be a path");
        };
        let PositionedElement::Path(tail) = &group.children[2] else {
            panic!("tail end should be a path");
        };
        assert_eq!(
            head.path.commands[1],
            PathCommand::LineTo(Point { x: 11.0, y: 2.0 })
        );
        assert_eq!(
            tail.path.commands[1],
            PathCommand::LineTo(Point { x: 9.0, y: 8.0 })
        );
    }

    #[test]
    fn all_supported_line_end_kinds_produce_finite_geometry() {
        let tangent = EndpointTangent {
            point: Point { x: 10.0, y: 5.0 },
            outward: Point { x: 1.0, y: 0.0 },
        };
        for kind in [
            ResolvedLineEndKind::Triangle,
            ResolvedLineEndKind::Stealth,
            ResolvedLineEndKind::Diamond,
            ResolvedLineEndKind::Oval,
            ResolvedLineEndKind::Arrow,
        ] {
            let path = line_end_path(&line_end(kind), tangent, 2.0)
                .expect("supported endpoint should produce geometry");
            assert!(path_is_finite(&path));
            assert!(matches!(path.commands.last(), Some(PathCommand::Close)));
            let bounds = path.bounds().expect("endpoint should have bounds");
            assert!(bounds.x >= 4.0 && bounds.x + bounds.width <= 10.0);
            assert!(bounds.y >= 2.0 && bounds.y + bounds.height <= 8.0);
        }
        assert_eq!(line_end_factor(ResolvedLineEndSize::Small), 2.0);
        assert_eq!(line_end_factor(ResolvedLineEndSize::Medium), 3.0);
        assert_eq!(line_end_factor(ResolvedLineEndSize::Large), 5.0);
    }

    #[test]
    fn zero_length_segment_omits_arrowhead_without_panicking() {
        let point = Point { x: 5.0, y: 5.0 };
        let mut arrow = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            ResolvedGeometry::Custom {
                paths: vec![open_line(point, point)],
                text_rect: None,
            },
            None,
            Some(Stroke::new(Paint::Solid(Color::BLACK), 2.0)),
        );
        arrow.tail_end = Some(line_end(ResolvedLineEndKind::Triangle));

        let page = layout_slide(&render_input(vec![slide((10.0, 10.0), vec![arrow])]), 0)
            .expect("lower zero-length line");
        assert_eq!(only_group(&page.elements[0]).children.len(), 1);
    }

    #[test]
    fn cropped_picture_renders_only_its_crop_region() {
        let png = horizontal_png(&[
            [255, 0, 0, 255],
            [0, 255, 0, 255],
            [0, 255, 0, 255],
            [0, 255, 0, 255],
        ]);
        let media_id = MediaId::from_bytes(&png);
        let picture = picture_shape(
            Rect {
                x: 2.0,
                y: 2.0,
                width: 8.0,
                height: 4.0,
            },
            media_id,
            Some(CropRect {
                left: 0.25,
                ..CropRect::default()
            }),
            ResolvedImagePlacement::default(),
            None,
        );
        let input = render_input_with_media(
            vec![slide((12.0, 8.0), vec![picture])],
            HashMap::from([(media_id, media(&png))]),
        );

        let layout = layout_presentation(&input).expect("lower cropped picture");
        let rendered =
            oxml_pdf::render_page_to_png(&layout, 0, 72.0).expect("rasterise cropped picture");
        let pixmap = tiny_skia::Pixmap::decode_png(&rendered).expect("decode cropped picture");
        let rgb = |x, y| {
            let pixel = pixmap.pixel(x, y).expect("sample lies inside page");
            (pixel.red(), pixel.green(), pixel.blue())
        };

        assert_eq!(rgb(3, 4), (0, 255, 0));
        assert_eq!(rgb(8, 4), (0, 255, 0));
        assert_eq!(rgb(1, 4), (255, 255, 255));
    }

    #[test]
    fn crop_lowers_to_clipped_source_image_geometry() {
        let media_id = MediaId(21);
        let mut picture = picture_shape(
            Rect {
                x: 2.0,
                y: 3.0,
                width: 10.0,
                height: 8.0,
            },
            media_id,
            Some(CropRect {
                left: 0.25,
                right: 0.25,
                ..CropRect::default()
            }),
            ResolvedImagePlacement::default(),
            None,
        );
        picture.line = Some(Stroke::new(Paint::Solid(Color::BLACK), 1.0));
        let input = render_input_with_media(
            vec![slide((20.0, 20.0), vec![picture])],
            HashMap::from([(media_id, media(b"image"))]),
        );

        let page = layout_slide(&input, 0).expect("lower crop geometry");
        let shape_group = only_group(&page.elements[0]);
        assert_eq!(shape_group.children.len(), 2, "outline must follow picture");
        let crop_clip = only_group(&shape_group.children[0]);
        assert_eq!(
            crop_clip.clip,
            Some(Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 8.0,
            }))
        );
        let PositionedElement::Image {
            rect, media_id: id, ..
        } = &crop_clip.children[0]
        else {
            panic!("crop group should contain the expanded source image");
        };
        assert_eq!(*id, media_id);
        assert_eq!(
            *rect,
            Rect {
                x: -5.0,
                y: 0.0,
                width: 20.0,
                height: 8.0,
            }
        );
        let PositionedElement::Path(outline) = &shape_group.children[1] else {
            panic!("picture outline should remain above image content");
        };
        assert!(outline.stroke.is_some());
    }

    #[test]
    fn shape_picture_fill_is_clipped_below_stroke_and_text() {
        let media_id = MediaId(22);
        let clip_path = Path {
            commands: vec![
                PathCommand::MoveTo(Point { x: 0.0, y: 0.0 }),
                PathCommand::LineTo(Point { x: 10.0, y: 0.0 }),
                PathCommand::LineTo(Point { x: 5.0, y: 10.0 }),
                PathCommand::Close,
            ],
            fill_rule: FillRule::NonZero,
        };
        let mut filled = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            ResolvedGeometry::Custom {
                paths: vec![clip_path.clone()],
                text_rect: None,
            },
            None,
            Some(Stroke::new(Paint::Solid(Color::BLACK), 1.0)),
        );
        filled.image_fill = Some(ResolvedImage {
            media: media_id,
            src_rect: None,
            placement: ResolvedImagePlacement::default(),
            dpi: None,
            rotate_with_shape: true,
        });
        filled.content = ResolvedContent::Text(table_text("caption"));

        let mut fallback = shape(
            Rect {
                x: 10.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            ResolvedGeometry::BoundsFallback,
            None,
            None,
        );
        fallback.image_fill = filled.image_fill.clone();
        let input = render_input_with_media(
            vec![slide((20.0, 10.0), vec![filled, fallback])],
            HashMap::from([(media_id, media(b"image"))]),
        );

        let page = layout_slide(&input, 0).expect("lower shape picture fill");
        let filled_group = only_group(&page.elements[0]);
        assert!(
            filled_group.children.len() > 2,
            "text must follow image and stroke"
        );
        let image_clip = only_group(&filled_group.children[0]);
        assert_eq!(image_clip.clip, Some(clip_path));
        assert!(matches!(
            image_clip.children.as_slice(),
            [PositionedElement::Image { media_id: id, .. }] if *id == media_id
        ));
        let PositionedElement::Path(outline) = &filled_group.children[1] else {
            panic!("shape stroke must follow its image fill");
        };
        assert!(outline.fill.is_none());
        assert!(outline.stroke.is_some());

        let fallback_group = only_group(&page.elements[1]);
        let PositionedElement::Path(bounds) = &fallback_group.children[1] else {
            panic!("bounds path must follow its image fill");
        };
        assert!(
            bounds.stroke.is_none(),
            "image fill suppresses synthetic border"
        );
    }

    #[test]
    fn tile_picture_repeats_media_in_row_major_order_inside_shape_clip() {
        let png = horizontal_png(&[[255, 0, 0, 255]]);
        let media_id = MediaId::from_bytes(&png);
        let picture = picture_shape(
            Rect {
                x: 2.0,
                y: 2.0,
                width: 3.0,
                height: 2.0,
            },
            media_id,
            None,
            ResolvedImagePlacement::Tile(ResolvedTilePlacement {
                translation: Point { x: 0.0, y: 0.0 },
                scale_x: 1.0,
                scale_y: 1.0,
                flip: ResolvedTileFlip::None,
                alignment: ResolvedRectAlignment::TopLeft,
            }),
            Some(72.0),
        );
        let input = render_input_with_media(
            vec![slide((8.0, 6.0), vec![picture])],
            HashMap::from([(media_id, media(&png))]),
        );

        let layout = layout_presentation(&input).expect("lower tiled picture");
        let shape_group = only_group(&layout.pages[0].elements[0]);
        let tiles = only_group(&shape_group.children[0]);
        let rects = tiles
            .children
            .iter()
            .map(|tile| match tile {
                PositionedElement::Image { rect, .. } => *rect,
                _ => panic!("unflipped tile should be one image"),
            })
            .collect::<Vec<_>>();
        assert_eq!(
            rects,
            [
                Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0
                },
                Rect {
                    x: 1.0,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0
                },
                Rect {
                    x: 2.0,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0
                },
                Rect {
                    x: 0.0,
                    y: 1.0,
                    width: 1.0,
                    height: 1.0
                },
                Rect {
                    x: 1.0,
                    y: 1.0,
                    width: 1.0,
                    height: 1.0
                },
                Rect {
                    x: 2.0,
                    y: 1.0,
                    width: 1.0,
                    height: 1.0
                },
            ]
        );
        let rendered =
            oxml_pdf::render_page_to_png(&layout, 0, 72.0).expect("rasterise tiled picture");
        let pixmap = tiny_skia::Pixmap::decode_png(&rendered).expect("decode tiled picture");
        assert_eq!(rgb_at(&pixmap, 3, 3), (255, 0, 0));
        assert_eq!(rgb_at(&pixmap, 1, 3), (255, 255, 255));
        assert_eq!(rgb_at(&pixmap, 5, 3), (255, 255, 255));

        assert_eq!(
            tile_alignment_origin(
                Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 3.0,
                    height: 2.0,
                },
                1.0,
                1.0,
                ResolvedRectAlignment::BottomRight,
            ),
            Point { x: 2.0, y: 1.0 }
        );
        assert_eq!(repeated_origin(3.5, 0.0, 1.0), -0.5);
        assert_eq!(repeated_origin(0.0, -2.1, 1.0), -3.0);
        assert_eq!(repeated_tile_index(-0.5, 2.5, 1.0, media_id).unwrap(), -3);
        let flipped = flip_tile(
            picture_image(media_id, input.media.get(&media_id).unwrap(), rects[4]),
            rects[4],
            ResolvedTileFlip::Both,
            true,
            true,
        );
        let flipped = only_group(&flipped);
        assert_eq!(
            flipped.transform,
            Transform {
                a: -1.0,
                b: 0.0,
                c: 0.0,
                d: -1.0,
                e: 3.0,
                f: 3.0,
            }
        );
    }

    #[test]
    fn picture_rotation_policy_counter_rotates_only_image_content() {
        let png = horizontal_png(&[[255, 0, 0, 255]]);
        let media_id = MediaId::from_bytes(&png);
        let mut picture = picture_shape(
            Rect {
                x: 5.0,
                y: 5.0,
                width: 10.0,
                height: 10.0,
            },
            media_id,
            None,
            ResolvedImagePlacement::default(),
            None,
        );
        picture.rotation_deg = 45.0;
        let ResolvedContent::Image(image) = &mut picture.content else {
            unreachable!();
        };
        image.rotate_with_shape = false;
        let mut tiled_picture = picture_shape(
            Rect {
                x: 25.0,
                y: 5.0,
                width: 10.0,
                height: 10.0,
            },
            media_id,
            None,
            ResolvedImagePlacement::Tile(ResolvedTilePlacement {
                translation: Point { x: 0.0, y: 0.0 },
                scale_x: 1.0,
                scale_y: 1.0,
                flip: ResolvedTileFlip::None,
                alignment: ResolvedRectAlignment::TopLeft,
            }),
            Some(72.0),
        );
        tiled_picture.rotation_deg = 45.0;
        let ResolvedContent::Image(image) = &mut tiled_picture.content else {
            unreachable!();
        };
        image.rotate_with_shape = false;
        let input = render_input_with_media(
            vec![slide((40.0, 20.0), vec![picture, tiled_picture])],
            HashMap::from([(media_id, media(&png))]),
        );

        let layout = layout_presentation(&input).expect("lower picture rotation policy");
        let shape = only_group(&layout.pages[0].elements[0]);
        let clip = only_group(&shape.children[0]);
        assert_eq!(
            clip.clip,
            Some(Path::rect(local_shape_rect(&input.slides[0].shapes[0])))
        );
        let image = only_group(&clip.children[0]);
        assert_eq!(image.transform, Transform::rotate_about(-45.0, 5.0, 5.0));
        assert!(matches!(
            image.children.as_slice(),
            [PositionedElement::Image { .. }]
        ));

        let rendered =
            oxml_pdf::render_page_to_png(&layout, 0, 72.0).expect("rasterise picture rotation");
        let pixmap = tiny_skia::Pixmap::decode_png(&rendered).expect("decode picture rotation");
        assert_eq!(rgb_at(&pixmap, 10, 4), (255, 0, 0));
        assert_eq!(rgb_at(&pixmap, 4, 4), (255, 255, 255));
        assert_eq!(rgb_at(&pixmap, 30, 4), (255, 0, 0));
        assert_eq!(rgb_at(&pixmap, 24, 4), (255, 255, 255));
    }

    #[test]
    fn tile_dpi_prefers_declared_then_embedded_then_96() {
        let embedded = oxml_media::ImageInfo {
            format: oxml_media::ImageFormat::Png,
            width_px: 144,
            height_px: 72,
            dpi_x: Some(72.0),
            dpi_y: Some(72.0),
            bit_depth: 8,
            channels: 4,
            has_alpha: true,
        };
        assert_eq!(
            tile_native_size_points(embedded, Some(144.0), MediaId(1)).unwrap(),
            (72.0, 36.0)
        );
        assert_eq!(
            tile_native_size_points(embedded, None, MediaId(1)).unwrap(),
            (144.0, 72.0)
        );
        assert_eq!(
            tile_native_size_points(
                oxml_media::ImageInfo {
                    dpi_x: None,
                    dpi_y: None,
                    ..embedded
                },
                None,
                MediaId(1),
            )
            .unwrap(),
            (108.0, 54.0)
        );
    }

    #[test]
    fn equal_picture_bytes_reuse_one_media_id_across_elements() {
        let bytes = b"same picture";
        let media_id = MediaId::from_bytes(bytes);
        let pictures = [0.0, 5.0]
            .map(|x| {
                picture_shape(
                    Rect {
                        x,
                        y: 0.0,
                        width: 4.0,
                        height: 4.0,
                    },
                    media_id,
                    None,
                    ResolvedImagePlacement::default(),
                    None,
                )
            })
            .to_vec();
        let input = render_input_with_media(
            vec![slide((10.0, 5.0), pictures)],
            HashMap::from([(media_id, media(bytes))]),
        );

        let layout = layout_presentation(&input).expect("lower repeated picture media");
        let mut ids = Vec::new();
        collect_image_ids(&layout.pages[0].elements, &mut ids);
        assert_eq!(ids, [media_id, media_id]);
        assert_eq!(input.media.len(), 1);
    }

    #[test]
    fn missing_external_media_and_empty_crop_are_contextual() {
        let missing_id = MediaId(404);
        let missing = picture_shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 4.0,
                height: 4.0,
            },
            missing_id,
            None,
            ResolvedImagePlacement::default(),
            None,
        );
        assert_eq!(
            layout_slide(&render_input(vec![slide((5.0, 5.0), vec![missing])]), 0).unwrap_err(),
            RenderInputError::MissingMedia { media: missing_id }
        );

        let media_id = MediaId(405);
        let empty = picture_shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 4.0,
                height: 4.0,
            },
            media_id,
            Some(CropRect {
                left: 0.5,
                right: 0.5,
                ..CropRect::default()
            }),
            ResolvedImagePlacement::default(),
            None,
        );
        let input = render_input_with_media(
            vec![slide((5.0, 5.0), vec![empty])],
            HashMap::from([(media_id, media(b"image"))]),
        );
        assert_eq!(
            layout_slide(&input, 0).unwrap_err(),
            RenderInputError::InvalidPicture {
                media: media_id,
                detail: "source crop",
            }
        );

        let png = horizontal_png(&[[0, 0, 0, 255]]);
        let media_id = MediaId::from_bytes(&png);
        let excessive = picture_shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            media_id,
            None,
            ResolvedImagePlacement::Tile(ResolvedTilePlacement {
                translation: Point { x: 0.0, y: 0.0 },
                scale_x: 0.000_1,
                scale_y: 0.000_1,
                flip: ResolvedTileFlip::None,
                alignment: ResolvedRectAlignment::TopLeft,
            }),
            Some(72.0),
        );
        let input = render_input_with_media(
            vec![slide((10.0, 10.0), vec![excessive])],
            HashMap::from([(media_id, media(&png))]),
        );
        assert!(matches!(
            layout_slide(&input, 0),
            Err(RenderInputError::TileLimitExceeded { media, .. }) if media == media_id
        ));
    }

    fn picture_shape(
        bounds: Rect,
        media: MediaId,
        src_rect: Option<CropRect>,
        placement: ResolvedImagePlacement,
        dpi: Option<f64>,
    ) -> ResolvedShape {
        let mut picture = shape(bounds, ResolvedGeometry::Rectangle, None, None);
        picture.content = ResolvedContent::Image(ResolvedImage {
            media,
            src_rect,
            placement,
            dpi,
            rotate_with_shape: true,
        });
        picture
    }

    fn render_input_with_media(
        slides: Vec<ResolvedSlide>,
        media: HashMap<MediaId, MediaData>,
    ) -> RenderInput {
        RenderInput {
            slides,
            media,
            fonts: Vec::new(),
            metadata: None,
        }
    }

    fn horizontal_png(colors: &[[u8; 4]]) -> Vec<u8> {
        let mut pixmap = tiny_skia::Pixmap::new(colors.len() as u32, 1).expect("fixture pixmap");
        for (pixel, color) in pixmap.pixels_mut().iter_mut().zip(colors) {
            *pixel =
                tiny_skia::PremultipliedColorU8::from_rgba(color[0], color[1], color[2], color[3])
                    .expect("premultiplied fixture colour");
        }
        pixmap.encode_png().expect("encode fixture PNG")
    }

    fn rgb_at(pixmap: &tiny_skia::Pixmap, x: u32, y: u32) -> (u8, u8, u8) {
        let pixel = pixmap.pixel(x, y).expect("sample lies inside page");
        (pixel.red(), pixel.green(), pixel.blue())
    }

    fn collect_image_ids(elements: &[PositionedElement], ids: &mut Vec<MediaId>) {
        for element in elements {
            match element {
                PositionedElement::Image { media_id, .. } => ids.push(*media_id),
                PositionedElement::Group(group) => collect_image_ids(&group.children, ids),
                _ => {}
            }
        }
    }

    fn open_line(from: Point, to: Point) -> Path {
        Path {
            commands: vec![PathCommand::MoveTo(from), PathCommand::LineTo(to)],
            fill_rule: FillRule::NonZero,
        }
    }

    fn line_end(kind: ResolvedLineEndKind) -> ResolvedLineEnd {
        ResolvedLineEnd {
            kind,
            width: ResolvedLineEndSize::Medium,
            length: ResolvedLineEndSize::Medium,
        }
    }

    #[test]
    fn layout_slide_rejects_an_out_of_range_index() {
        let input = render_input(vec![slide((20.0, 10.0), Vec::new())]);

        assert_eq!(
            layout_slide(&input, 4).unwrap_err(),
            RenderInputError::SlideIndexOutOfBounds {
                index: 4,
                slide_count: 1,
            }
        );
    }

    #[test]
    fn resolved_rtl_numeric_forced_breaks_reach_rich_line_layout() {
        const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
        let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="1" name="rtl"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr rtl="1"/><a:r><a:rPr sz="1800"/><a:t>123</a:t></a:r><a:br/><a:r><a:rPr sz="1800"/><a:t>456</a:t></a:r></a:p></p:txBody></p:sp>"#;
        let slide = CT_Slide::from_xml(
            format!(
                "<p:sld xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{shape}</p:spTree></p:cSld></p:sld>"
            )
            .as_bytes(),
        )
        .unwrap();
        let empty_tree = "<p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree>";
        let layout = CT_SlideLayout::from_xml(
            format!(
                "<p:sldLayout xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"><p:cSld>{empty_tree}</p:cSld></p:sldLayout>"
            )
            .as_bytes(),
        )
        .unwrap();
        let master = CT_SlideMaster::from_xml(
            format!(
                "<p:sldMaster xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"><p:cSld>{empty_tree}</p:cSld><p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/></p:sldMaster>"
            )
            .as_bytes(),
        )
        .unwrap();
        let theme = CT_OfficeStyleSheet::office_default();
        let default_text_style = CT_TextListStyle::default();
        let media = rpptx_layout::ScopedMediaIds::default();
        let hyperlinks = rpptx_layout::ScopedHyperlinkTargets::default();
        let charts = rpptx_layout::ScopedChartResources::default();
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let (resolved, directions) = ResolveCtx::new(
            &theme,
            ColorMap::default(),
            &master,
            &layout,
            &slide,
            &default_text_style,
        )
        .resolve_slide_with_chart_resources_and_text_directions(
            (720.0, 144.0),
            &media,
            &hyperlinks,
            &charts,
            &mut fonts,
        )
        .unwrap();
        let rendered = layout_presentation_with_font_manager_and_text_directions(
            &render_input(vec![resolved]),
            fonts,
            &[directions],
        )
        .unwrap();
        let mut runs = Vec::new();
        walk(&rendered.pages[0].elements, &mut |element, _| {
            if let PositionedElement::MultilingualText(run) = element {
                runs.push((run.logical_text.clone(), run.bidi_level, run.origin.y));
            }
        });

        assert_eq!(
            runs.iter()
                .map(|(text, level, _)| (text.as_str(), *level))
                .collect::<Vec<_>>(),
            [("123", 2), ("456", 2)]
        );
        assert!(runs[0].2 < runs[1].2);
    }

    #[test]
    fn master_gradient_background_renders_when_slide_and_layout_omit_one() {
        const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
        let shape_tree = "<p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree>";
        let slide = CT_Slide::from_xml(
            format!(
                "<p:sld xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"><p:cSld>{shape_tree}</p:cSld></p:sld>"
            )
            .as_bytes(),
        )
        .unwrap();
        let layout = CT_SlideLayout::from_xml(
            format!(
                "<p:sldLayout xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"><p:cSld>{shape_tree}</p:cSld></p:sldLayout>"
            )
            .as_bytes(),
        )
        .unwrap();
        let background = r#"<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></p:bgPr></p:bg>"#;
        let master = CT_SlideMaster::from_xml(
            format!(
                "<p:sldMaster xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"><p:cSld>{background}{shape_tree}</p:cSld><p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/></p:sldMaster>"
            )
            .as_bytes(),
        )
        .unwrap();
        let theme = CT_OfficeStyleSheet::office_default();
        let default_text_style = CT_TextListStyle::default();
        let resolved = ResolveCtx::new(
            &theme,
            ColorMap::default(),
            &master,
            &layout,
            &slide,
            &default_text_style,
        )
        .resolve_slide((40.0, 20.0))
        .unwrap();
        let Some(ResolvedBackground::Paint(resolved_background)) = resolved.background.clone()
        else {
            panic!("expected resolved paint background");
        };
        let rendered = layout_presentation(&render_input(vec![resolved])).unwrap();

        assert_eq!(rendered.pages[0].background, Some(resolved_background));
        assert!(rendered.pages[0].elements.is_empty());
        let png = oxml_pdf::render_page_to_png(&rendered, 0, 72.0)
            .expect("rasterise inherited master gradient");
        let pixmap = tiny_skia::Pixmap::decode_png(&png).expect("decode background raster");
        let left = rgb_at(&pixmap, 2, 10);
        let right = rgb_at(&pixmap, 37, 10);
        assert!(
            left.0 > left.2,
            "left sample should be red-dominant: {left:?}"
        );
        assert!(
            right.2 > right.0,
            "right sample should be blue-dominant: {right:?}"
        );
    }

    #[test]
    fn absent_background_keeps_the_default_white_raster() {
        const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
        let shape_tree = "<p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree>";
        let slide = CT_Slide::from_xml(
            format!(
                "<p:sld xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"><p:cSld>{shape_tree}</p:cSld></p:sld>"
            )
            .as_bytes(),
        )
        .unwrap();
        let layout = CT_SlideLayout::from_xml(
            format!(
                "<p:sldLayout xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"><p:cSld>{shape_tree}</p:cSld></p:sldLayout>"
            )
            .as_bytes(),
        )
        .unwrap();
        let master = CT_SlideMaster::from_xml(
            format!(
                "<p:sldMaster xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"><p:cSld>{shape_tree}</p:cSld><p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/></p:sldMaster>"
            )
            .as_bytes(),
        )
        .unwrap();
        let mut theme = CT_OfficeStyleSheet::office_default();
        theme
            .theme_elements
            .format_scheme
            .background_fill_styles[0] = oxml_drawing::fill::Fill::from_xml(
            br#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="000000"/></a:solidFill>"#,
        )
        .unwrap();
        let default_text_style = CT_TextListStyle::default();
        let resolved = ResolveCtx::new(
            &theme,
            ColorMap::default(),
            &master,
            &layout,
            &slide,
            &default_text_style,
        )
        .resolve_slide((10.0, 10.0))
        .unwrap();

        assert!(resolved.background.is_none());
        let rendered = layout_presentation(&render_input(vec![resolved])).unwrap();
        assert!(rendered.pages[0].background.is_none());
        let png = oxml_pdf::render_page_to_png(&rendered, 0, 72.0)
            .expect("rasterise absent presentation background");
        let pixmap = tiny_skia::Pixmap::decode_png(&png).expect("decode white background raster");
        assert_eq!(rgb_at(&pixmap, 5, 5), (255, 255, 255));
    }

    #[test]
    fn background_is_not_duplicated_in_page_elements() {
        let mut resolved = slide((20.0, 10.0), Vec::new());
        resolved.background = Some(ResolvedBackground::Paint(Paint::Solid(Color::from_hex(
            "102030",
        ))));

        let page = layout_slide(&render_input(vec![resolved]), 0).unwrap();

        assert_eq!(
            page.background,
            Some(Paint::Solid(Color::from_hex("102030")))
        );
        assert!(page.elements.is_empty());
    }

    #[test]
    fn background_image_is_lowered_before_slide_shapes() {
        let png = horizontal_png(&[[0, 0, 255, 255]]);
        let media_id = MediaId::from_bytes(&png);
        let foreground = shape(
            Rect {
                x: 2.0,
                y: 2.0,
                width: 6.0,
                height: 6.0,
            },
            ResolvedGeometry::Rectangle,
            Some(Paint::Solid(Color::from_hex("FF0000"))),
            None,
        );
        let mut resolved = slide((10.0, 10.0), vec![foreground]);
        resolved.background = Some(ResolvedBackground::Image(ResolvedImage {
            media: media_id,
            src_rect: None,
            placement: ResolvedImagePlacement::default(),
            dpi: None,
            rotate_with_shape: true,
        }));
        let input =
            render_input_with_media(vec![resolved], HashMap::from([(media_id, media(&png))]));

        let page = layout_slide(&input, 0).unwrap();
        assert!(matches!(
            page.elements.first(),
            Some(PositionedElement::Image { media_id: id, .. }) if *id == media_id
        ));
        assert!(matches!(
            page.elements.get(1),
            Some(PositionedElement::Group(_))
        ));
        let rendered = oxml_pdf::render_page_to_png(
            &LayoutResult::new(vec![page.into()], Vec::new(), None, Vec::new()),
            0,
            72.0,
        )
        .expect("rasterise background image ordering");
        let pixmap = tiny_skia::Pixmap::decode_png(&rendered).unwrap();
        assert_eq!(rgb_at(&pixmap, 0, 0), (0, 0, 255));
        assert_eq!(rgb_at(&pixmap, 5, 5), (255, 0, 0));
    }

    #[test]
    fn layout_presentation_preserves_page_order_and_diagnostics() {
        let mut first = slide((20.0, 10.0), Vec::new());
        first.diagnostics.push(Diagnostic {
            message: "first diagnostic".to_owned(),
        });
        let mut second = slide((30.0, 15.0), Vec::new());
        second.diagnostics.push(Diagnostic {
            message: "second diagnostic".to_owned(),
        });
        let mut input = render_input(vec![first, second]);
        input.metadata = Some(DocumentMetadata {
            title: Some("shape deck".to_owned()),
            author: Some("rpptx-render".to_owned()),
            ..DocumentMetadata::default()
        });

        let layout = layout_presentation(&input).expect("lower presentation");
        assert_eq!(layout.pages.len(), 2);
        assert_eq!(
            (
                layout.pages[0].page_number,
                layout.pages[0].width,
                layout.pages[0].height
            ),
            (1, 20.0, 10.0)
        );
        assert_eq!(
            (
                layout.pages[1].page_number,
                layout.pages[1].width,
                layout.pages[1].height
            ),
            (2, 30.0, 15.0)
        );
        assert_eq!(
            layout
                .metadata
                .as_ref()
                .and_then(|metadata| metadata.title.as_deref()),
            Some("shape deck")
        );
        assert_eq!(
            layout
                .diagnostics
                .iter()
                .map(|diagnostic| diagnostic.message.as_str())
                .collect::<Vec<_>>(),
            vec!["first diagnostic", "second diagnostic"]
        );
        assert!(layout.fonts.is_empty());
        assert!(layout.outlines.is_empty());
        assert!(layout.structure.is_none());
    }

    #[test]
    fn same_relationship_id_resolves_independently_in_all_three_scopes() {
        let relationships = RelScopes {
            slide: HashMap::from([("rId2".to_owned(), relationship("slide.png"))]),
            layout: HashMap::from([("rId2".to_owned(), relationship("layout.png"))]),
            master: HashMap::from([("rId2".to_owned(), relationship("master.png"))]),
        };
        let package_media = HashMap::from([
            ("slide.png".to_owned(), media(b"slide image")),
            ("layout.png".to_owned(), media(b"layout image")),
            ("master.png".to_owned(), media(b"master image")),
        ]);
        let mut deck_media = HashMap::new();

        let slide = resolve_media_relationship(
            &relationships,
            RelScope::Slide,
            "rId2",
            &package_media,
            &mut deck_media,
        )
        .unwrap();
        let layout = resolve_media_relationship(
            &relationships,
            RelScope::Layout,
            "rId2",
            &package_media,
            &mut deck_media,
        )
        .unwrap();
        let master = resolve_media_relationship(
            &relationships,
            RelScope::Master,
            "rId2",
            &package_media,
            &mut deck_media,
        )
        .unwrap();

        assert_eq!(slide, MediaId::from_bytes(b"slide image"));
        assert_eq!(layout, MediaId::from_bytes(b"layout image"));
        assert_eq!(master, MediaId::from_bytes(b"master image"));
        assert_eq!(deck_media.len(), 3);
    }

    #[test]
    fn external_hyperlink_projection_keeps_scopes_and_excludes_internal_targets() {
        let relationships = RelScopes {
            slide: HashMap::from([
                (
                    "rId7".to_owned(),
                    hyperlink_relationship("https://slide.example", Some("External")),
                ),
                (
                    "rId8".to_owned(),
                    hyperlink_relationship("../slides/slide2.xml", None),
                ),
            ]),
            layout: HashMap::from([(
                "rId7".to_owned(),
                hyperlink_relationship("https://layout.example", Some("External")),
            )]),
            master: HashMap::from([(
                "rId7".to_owned(),
                hyperlink_relationship("https://master.example", Some("External")),
            )]),
        };

        let targets = relationships.external_hyperlink_targets();

        assert_eq!(
            targets.slide.get("rId7").map(String::as_str),
            Some("https://slide.example")
        );
        assert!(!targets.slide.contains_key("rId8"));
        assert_eq!(
            targets.layout.get("rId7").map(String::as_str),
            Some("https://layout.example")
        );
        assert_eq!(
            targets.master.get("rId7").map(String::as_str),
            Some("https://master.example")
        );
    }

    #[test]
    fn equal_media_bytes_deduplicate_to_one_media_entry() {
        let relationships = RelScopes {
            slide: HashMap::from([
                ("rId1".to_owned(), relationship("logo-a.png")),
                ("rId2".to_owned(), relationship("logo-b.png")),
            ]),
            ..RelScopes::default()
        };
        let package_media = HashMap::from([
            ("logo-a.png".to_owned(), media(b"shared logo")),
            ("logo-b.png".to_owned(), media(b"shared logo")),
        ]);
        let mut deck_media = HashMap::new();

        let first = resolve_media_relationship(
            &relationships,
            RelScope::Slide,
            "rId1",
            &package_media,
            &mut deck_media,
        )
        .unwrap();
        let second = resolve_media_relationship(
            &relationships,
            RelScope::Slide,
            "rId2",
            &package_media,
            &mut deck_media,
        )
        .unwrap();

        assert_eq!(first, second);
        assert_eq!(deck_media.len(), 1);
    }

    #[test]
    fn missing_relationship_reports_scope_and_id() {
        let error = resolve_media_relationship(
            &RelScopes::default(),
            RelScope::Layout,
            "rId9",
            &HashMap::new(),
            &mut HashMap::new(),
        )
        .unwrap_err();

        assert_eq!(
            error,
            RenderInputError::MissingRelationship {
                scope: RelScope::Layout,
                relationship_id: "rId9".to_owned(),
            }
        );
        assert!(error.to_string().contains("layout"));
        assert!(error.to_string().contains("rId9"));
    }

    #[test]
    fn render_input_contains_only_resolved_slides() {
        let input = RenderInput {
            slides: Vec::<ResolvedSlide>::new(),
            media: HashMap::new(),
            fonts: Vec::new(),
            metadata: None,
        };

        assert!(input.slides.is_empty());
        assert_eq!(
            std::any::type_name_of_val(&input.slides),
            "alloc::vec::Vec<rpptx_layout::ResolvedSlide>"
        );
    }

    #[test]
    fn rpptx_render_dependency_direction_is_one_way() {
        let manifest = include_str!("../Cargo.toml");
        let rpptx_manifest = include_str!("../../rpptx/Cargo.toml");
        let binding_manifest = include_str!("../../rpptx-py/Cargo.toml");
        assert!(manifest.contains(
            "[features]\ndefault = [\"system-fonts\"]\nsystem-fonts = [\"oxml-layout/system-fonts\"]"
        ));
        assert!(manifest.contains("oxml-layout = { workspace = true, default-features = false }"));
        assert!(
            rpptx_manifest
                .contains("default = [\"default-template\", \"render\", \"system-fonts\"]")
        );
        assert!(rpptx_manifest.contains(
            "system-fonts = [\"oxml-layout/system-fonts\", \"rpptx-render?/system-fonts\"]"
        ));
        assert!(rpptx_manifest.contains(
            "rpptx-render = { workspace = true, default-features = false, optional = true }"
        ));
        assert!(
            binding_manifest.contains(
                "rpptx = { workspace = true, features = [\"default-template\", \"render\", \"system-fonts\"] }"
            )
        );
        for dependency in [
            "oxml-drawing.workspace = true",
            "oxml-media.workspace = true",
            "rpptx-layout.workspace = true",
            "rpptx-oxml.workspace = true",
        ] {
            assert!(manifest.contains(dependency), "missing {dependency}");
        }
        for oxml_manifest in [
            include_str!("../../oxml-core/Cargo.toml"),
            include_str!("../../oxml-drawing/Cargo.toml"),
            include_str!("../../oxml-layout/Cargo.toml"),
            include_str!("../../oxml-media/Cargo.toml"),
            include_str!("../../oxml-opc/Cargo.toml"),
            include_str!("../../oxml-pdf/Cargo.toml"),
        ] {
            assert!(!oxml_manifest.contains("rpptx-render"));
        }
        assert!(manifest.contains("version = \"0.10.0\""));
        assert!(manifest.contains("publish = true"));
    }
}
