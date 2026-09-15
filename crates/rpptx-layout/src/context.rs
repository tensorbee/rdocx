use std::cell::RefCell;
use std::collections::BTreeMap;
use std::collections::HashMap;

use oxml_chart::{render_chart, render_chart_placeholder};
use oxml_drawing::color::{ColorChoice, ColorMap, ColorMapSlot, RgbColor, resolve_color};
use oxml_drawing::effect::CT_OuterShadowEffect;
use oxml_drawing::fill::{
    BlipFill, BlipMode, Fill, GradientGeometry, PathGradientKind, RelativeRect, SolidFill,
};
use oxml_drawing::geometry::EvaluatedPathCommand;
use oxml_drawing::line::{
    CT_LineProperties, LineCap as DrawingLineCap, LineDash, LineEnd, LineEndSize, LineEndType,
    LineJoin as DrawingLineJoin,
};
use oxml_drawing::style_ref::{FontCollectionIndex, StyleReference};
use oxml_drawing::table::{
    CT_TableBorders, CT_TableCellStyle, CT_TablePartStyle, CT_TableStyleList, CT_TableTextStyle,
};
use oxml_drawing::text::{
    CT_TextBody, CT_TextBodyProperties, CT_TextCharacterProperties, CT_TextListStyle,
    Coordinate32Value, TextAlignment, TextAnchor as DrawingTextAnchor, TextAutofit,
    TextBulletChoice, TextBulletSizeValue, TextFont, TextPointValue, TextRun, TextSpacing,
    TextVertical as DrawingTextVertical, TextWrap,
};
use oxml_drawing::theme::CT_OfficeStyleSheet;
use oxml_drawing::xfrm::CT_Transform2D;
use oxml_layout::{
    Color, Diagnostic, Effect, FillRule, FontManager, GradientStop, LineCap, LineJoin, Paint, Path,
    PathCommand, Point, Rect, Stroke, Transform,
};
use rpptx_oxml::graphic_frame::GraphicDataPayload;
use rpptx_oxml::picture::CT_Picture;
use rpptx_oxml::placeholder::{CT_Placeholder, PhType, PlaceholderKey};
use rpptx_oxml::shape_tree::{CT_Shape, ShapeTreeChild};
use rpptx_oxml::slide_parts::{BackgroundRendering, CT_Slide, CT_SlideLayout, CT_SlideMaster};

use crate::ResolveError;
use crate::style::{referenced_fill, substitute_fill};
use crate::text::EffectiveListStyle;
use crate::timeline::{
    EvaluatedFrameState, EvaluatedShapeState, ResolvedShapeIdentity, ResolvedTimelineSlide,
    TimelinePosition, evaluate_timeline,
};
use crate::{
    ChartResource, ParagraphAlignment, ResolvedAutofit, ResolvedBackground, ResolvedBullet,
    ResolvedBulletSize, ResolvedContent, ResolvedGeometry, ResolvedImage, ResolvedImagePlacement,
    ResolvedLineEnd, ResolvedLineEndKind, ResolvedLineEndSize, ResolvedParagraph,
    ResolvedRectAlignment, ResolvedRunStyle, ResolvedShape, ResolvedSlide,
    ResolvedSlideTextDirections, ResolvedTable, ResolvedTableBorder, ResolvedTableCell,
    ResolvedTableRow, ResolvedTextBody, ResolvedTextRun, ResolvedTextSpacing, ResolvedTileFlip,
    ResolvedTilePlacement, ScopedChartResources, ScopedHyperlinkTargets, ScopedMediaIds,
    TextAnchor, TextDirection, TextInsets,
};

/// The producer part that supplied the effective background.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum BackgroundSource {
    Slide,
    Layout,
    Master,
}

/// The borrowed background payload selected for one slide.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum BackgroundContent<'a> {
    Model(&'a BackgroundRendering),
}

/// The effective background and the per-master colour map used to resolve it.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct EffectiveBackground<'a> {
    pub source: BackgroundSource,
    pub content: BackgroundContent<'a>,
    pub color_map: &'a ColorMap,
}

/// The four ordered sources in a flattened slide view.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum FlattenedSource {
    Background,
    Master,
    Layout,
    Slide,
}

/// One borrowed entry in final draw order.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum FlattenedItem<'a> {
    Background(EffectiveBackground<'a>),
    Shape {
        source: FlattenedSource,
        child: &'a ShapeTreeChild,
        /// Coordinate scale accumulated from every non-shearing parent group.
        group_scale: (f64, f64),
        /// Rigid mapping accumulated from every non-shearing parent group.
        group_transform: Transform,
        /// Stable diagnostic flags accumulated from parent groups.
        group_issues: u8,
    },
}

const GROUP_ZERO_X: u8 = 1;
const GROUP_ZERO_Y: u8 = 2;
const GROUP_SHEAR: u8 = 4;

#[derive(Clone, Copy)]
struct ShapePlacement {
    source: FlattenedSource,
    group_scale: (f64, f64),
    group_transform: Transform,
}

impl FlattenedItem<'_> {
    pub const fn source(&self) -> FlattenedSource {
        match self {
            Self::Background(_) => FlattenedSource::Background,
            Self::Shape { source, .. } => *source,
        }
    }
}

/// The fixed presentation hierarchy and theme inputs for resolving one slide.
pub struct ResolveCtx<'a> {
    pub theme: &'a CT_OfficeStyleSheet,
    pub color_map: ColorMap,
    pub master: &'a CT_SlideMaster,
    pub layout: &'a CT_SlideLayout,
    pub slide: &'a CT_Slide,
    pub default_text_style: &'a CT_TextListStyle,
    pub table_styles: Option<&'a CT_TableStyleList>,
    pub(crate) list_style_cache: RefCell<HashMap<Option<PlaceholderKey>, EffectiveListStyle>>,
    media_poster_diagnostic_identity: bool,
}

impl<'a> ResolveCtx<'a> {
    pub fn new(
        theme: &'a CT_OfficeStyleSheet,
        color_map: ColorMap,
        master: &'a CT_SlideMaster,
        layout: &'a CT_SlideLayout,
        slide: &'a CT_Slide,
        default_text_style: &'a CT_TextListStyle,
    ) -> Self {
        Self {
            theme,
            color_map,
            master,
            layout,
            slide,
            default_text_style,
            table_styles: None,
            list_style_cache: RefCell::new(HashMap::new()),
            media_poster_diagnostic_identity: false,
        }
    }

    /// Adds the optional `ppt/tableStyles.xml` projection used by table frames.
    pub fn with_table_styles(mut self, table_styles: &'a CT_TableStyleList) -> Self {
        self.table_styles = Some(table_styles);
        self
    }

    /// Enables private source identity used while applying media fallback policy.
    #[doc(hidden)]
    pub fn with_media_poster_diagnostic_identity(mut self) -> Self {
        self.media_poster_diagnostic_identity = true;
        self
    }

    pub(crate) fn placeholder_chain<'ctx>(
        &'ctx self,
        shape: &CT_Shape,
    ) -> (Option<&'ctx CT_Shape>, Option<&'ctx CT_Shape>) {
        self.placeholder_chain_for(shape.placeholder.as_ref())
    }

    fn placeholder_chain_for<'ctx>(
        &'ctx self,
        placeholder: Option<&CT_Placeholder>,
    ) -> (Option<&'ctx CT_Shape>, Option<&'ctx CT_Shape>) {
        let Some(slide_key) = placeholder.map(CT_Placeholder::key) else {
            return (None, None);
        };
        let Some(layout_shape) = find_placeholder(
            &self.layout.common_slide_data.shape_tree.children,
            &slide_key,
        ) else {
            return (None, None);
        };
        let Some(layout_key) = layout_shape
            .placeholder
            .as_ref()
            .map(|placeholder| placeholder.key())
        else {
            return (None, None);
        };
        let master_shape = find_placeholder(
            &self.master.common_slide_data.shape_tree.children,
            &layout_key,
        );
        (Some(layout_shape), master_shape)
    }

    fn is_slide_number_placeholder(&self, shape: &CT_Shape) -> bool {
        if shape
            .placeholder
            .as_ref()
            .is_some_and(|placeholder| placeholder.ph_type == Some(PhType::SlideNumber))
        {
            return true;
        }
        let (layout, master) = self.placeholder_chain(shape);
        [layout, master].into_iter().flatten().any(|shape| {
            shape
                .placeholder
                .as_ref()
                .is_some_and(|placeholder| placeholder.ph_type == Some(PhType::SlideNumber))
        })
    }

    /// Resolves an owned transform from the slide, layout, then master shape.
    pub fn effective_xfrm(&self, shape: &CT_Shape) -> Option<CT_Transform2D> {
        let (layout, master) = self.placeholder_chain(shape);
        shape
            .shape_properties
            .transform
            .clone()
            .or_else(|| layout.and_then(|shape| shape.shape_properties.transform.as_ref().cloned()))
            .or_else(|| master.and_then(|shape| shape.shape_properties.transform.as_ref().cloned()))
    }

    /// Resolves picture bounds from the slide picture, layout placeholder, then master placeholder.
    pub fn effective_picture_xfrm(
        &self,
        picture: &rpptx_oxml::picture::CT_Picture,
    ) -> Option<CT_Transform2D> {
        let (layout, master) = self.placeholder_chain_for(picture.placeholder.as_ref());
        picture
            .shape_properties
            .transform
            .clone()
            .or_else(|| layout.and_then(|shape| shape.shape_properties.transform.as_ref().cloned()))
            .or_else(|| master.and_then(|shape| shape.shape_properties.transform.as_ref().cloned()))
    }

    /// Resolves body properties per field over defaults, master, layout, and slide.
    pub fn effective_body_pr(&self, shape: &CT_Shape) -> CT_TextBodyProperties {
        let (layout, master) = self.placeholder_chain(shape);
        let mut effective = default_body_properties();
        for source in [master, layout, Some(shape)].into_iter().flatten() {
            if let Some(text_body) = &source.text_body {
                merge_body_properties(&mut effective, &text_body.body_properties);
            }
        }
        effective
    }

    /// Selects the first background in slide, layout, and master order.
    pub fn effective_background(&self) -> Option<EffectiveBackground<'_>> {
        for (source, background) in [
            (
                BackgroundSource::Slide,
                self.slide.common_slide_data.background.as_ref(),
            ),
            (
                BackgroundSource::Layout,
                self.layout.common_slide_data.background.as_ref(),
            ),
            (
                BackgroundSource::Master,
                self.master.common_slide_data.background.as_ref(),
            ),
        ] {
            if let Some(background) = background {
                return Some(EffectiveBackground {
                    source,
                    content: BackgroundContent::Model(background.rendering()),
                    color_map: &self.color_map,
                });
            }
        }
        None
    }

    /// Returns background and shape-tree leaves in final draw order.
    pub fn flatten(&self) -> Vec<FlattenedItem<'_>> {
        let slide_children = &self.slide.common_slide_data.shape_tree.children;
        let layout_children = &self.layout.common_slide_data.shape_tree.children;
        let master_children = &self.master.common_slide_data.shape_tree.children;
        let mut slide_latent = Vec::new();
        let mut layout_latent = Vec::new();
        collect_occupied_latent(slide_children, &mut slide_latent);
        collect_occupied_latent(layout_children, &mut layout_latent);

        let mut flattened = Vec::new();
        if let Some(background) = self.effective_background() {
            flattened.push(FlattenedItem::Background(background));
        }
        let allow_latent = LatentPolicy::from_context(self);
        let inherited_latent_enabled =
            self.layout.header_footer.is_some() || self.master.header_footer.is_some();
        let mut master_deeper = layout_latent.clone();
        master_deeper.extend(slide_latent.iter().cloned());
        emit_tree(
            master_children,
            PassRules {
                source: FlattenedSource::Master,
                emit_non_placeholders: self.layout.show_master_shapes.unwrap_or(true),
                emit_inherited_latent: inherited_latent_enabled,
                deeper_latent: &master_deeper,
                latent_policy: allow_latent,
            },
            &mut flattened,
        );
        emit_tree(
            layout_children,
            PassRules {
                source: FlattenedSource::Layout,
                emit_non_placeholders: self.slide.show_master_shapes.unwrap_or(true),
                emit_inherited_latent: inherited_latent_enabled,
                deeper_latent: &slide_latent,
                latent_policy: allow_latent,
            },
            &mut flattened,
        );
        emit_tree(
            slide_children,
            PassRules {
                source: FlattenedSource::Slide,
                emit_non_placeholders: true,
                emit_inherited_latent: true,
                deeper_latent: &[],
                latent_policy: allow_latent,
            },
            &mut flattened,
        );
        flattened
    }

    /// Resolves one owned renderer-facing slide at the supplied point size.
    pub fn resolve_slide(&self, size: (f64, f64)) -> Result<ResolvedSlide, ResolveError> {
        self.resolve_slide_inner(size, None, None, None, None)
            .map(|(slide, _)| slide)
    }

    /// Resolves one slide through the additive slide-local timeline path.
    pub fn resolve_slide_at(
        &self,
        size: (f64, f64),
        position: TimelinePosition,
    ) -> Result<ResolvedTimelineSlide, ResolveError> {
        let (slide, _, identities) =
            self.resolve_slide_inner_with_identities(size, None, None, None, None)?;
        let state = evaluate_timeline(
            self.slide.timing.as_ref(),
            self.slide.transition.as_ref(),
            position,
        )?;
        Ok(apply_timeline_state(
            slide,
            identities,
            state,
            timeline_group_geometries(&self.slide.common_slide_data.shape_tree.children),
        ))
    }

    /// Resolves one slide using embedded media identifiers scoped to source parts.
    pub fn resolve_slide_with_media(
        &self,
        size: (f64, f64),
        media: &ScopedMediaIds,
    ) -> Result<ResolvedSlide, ResolveError> {
        self.resolve_slide_inner(size, Some(media), None, None, None)
            .map(|(slide, _)| slide)
    }

    /// Resolves one slide with source-scoped media and external hyperlink targets.
    pub fn resolve_slide_with_resources(
        &self,
        size: (f64, f64),
        media: &ScopedMediaIds,
        hyperlinks: &ScopedHyperlinkTargets,
    ) -> Result<ResolvedSlide, ResolveError> {
        self.resolve_slide_inner(size, Some(media), Some(hyperlinks), None, None)
            .map(|(slide, _)| slide)
    }

    /// Resolves one slide with source-scoped charts and the caller's font manager.
    pub fn resolve_slide_with_chart_resources(
        &self,
        size: (f64, f64),
        media: &ScopedMediaIds,
        hyperlinks: &ScopedHyperlinkTargets,
        charts: &ScopedChartResources,
        fonts: &mut FontManager,
    ) -> Result<ResolvedSlide, ResolveError> {
        self.resolve_slide_inner(
            size,
            Some(media),
            Some(hyperlinks),
            Some(charts),
            Some(fonts),
        )
        .map(|(slide, _)| slide)
    }

    /// Resolves one slide and returns paragraph directions in shape and text-body order.
    ///
    /// The outer vector follows `ResolvedSlide::shapes`. Each shape entry contains
    /// one paragraph-direction vector for a text shape, or one entry per table cell
    /// in row-major order. Non-text shapes have an empty entry.
    pub fn resolve_slide_with_chart_resources_and_text_directions(
        &self,
        size: (f64, f64),
        media: &ScopedMediaIds,
        hyperlinks: &ScopedHyperlinkTargets,
        charts: &ScopedChartResources,
        fonts: &mut FontManager,
    ) -> Result<(ResolvedSlide, ResolvedSlideTextDirections), ResolveError> {
        self.resolve_slide_inner(
            size,
            Some(media),
            Some(hyperlinks),
            Some(charts),
            Some(fonts),
        )
    }

    /// Resolves one timestamped slide while retaining text directions and target identity.
    pub fn resolve_slide_with_chart_resources_and_text_directions_at(
        &self,
        size: (f64, f64),
        media: &ScopedMediaIds,
        hyperlinks: &ScopedHyperlinkTargets,
        charts: &ScopedChartResources,
        fonts: &mut FontManager,
        position: TimelinePosition,
    ) -> Result<(ResolvedTimelineSlide, ResolvedSlideTextDirections), ResolveError> {
        let (slide, directions, identities) = self.resolve_slide_inner_with_identities(
            size,
            Some(media),
            Some(hyperlinks),
            Some(charts),
            Some(fonts),
        )?;
        let state = evaluate_timeline(
            self.slide.timing.as_ref(),
            self.slide.transition.as_ref(),
            position,
        )?;
        Ok((
            apply_timeline_state(
                slide,
                identities,
                state,
                timeline_group_geometries(&self.slide.common_slide_data.shape_tree.children),
            ),
            directions,
        ))
    }

    fn resolve_slide_inner(
        &self,
        size: (f64, f64),
        media: Option<&ScopedMediaIds>,
        hyperlinks: Option<&ScopedHyperlinkTargets>,
        charts: Option<&ScopedChartResources>,
        fonts: Option<&mut FontManager>,
    ) -> Result<(ResolvedSlide, ResolvedSlideTextDirections), ResolveError> {
        self.resolve_slide_inner_common(size, media, hyperlinks, charts, fonts, None)
    }

    fn resolve_slide_inner_with_identities(
        &self,
        size: (f64, f64),
        media: Option<&ScopedMediaIds>,
        hyperlinks: Option<&ScopedHyperlinkTargets>,
        charts: Option<&ScopedChartResources>,
        fonts: Option<&mut FontManager>,
    ) -> Result<
        (
            ResolvedSlide,
            ResolvedSlideTextDirections,
            Vec<ResolvedShapeIdentity>,
        ),
        ResolveError,
    > {
        let mut identities = Vec::new();
        let (slide, directions) = self.resolve_slide_inner_common(
            size,
            media,
            hyperlinks,
            charts,
            fonts,
            Some(&mut identities),
        )?;
        Ok((slide, directions, identities))
    }

    fn resolve_slide_inner_common(
        &self,
        size: (f64, f64),
        media: Option<&ScopedMediaIds>,
        hyperlinks: Option<&ScopedHyperlinkTargets>,
        charts: Option<&ScopedChartResources>,
        mut fonts: Option<&mut FontManager>,
        mut identities: Option<&mut Vec<ResolvedShapeIdentity>>,
    ) -> Result<(ResolvedSlide, ResolvedSlideTextDirections), ResolveError> {
        let mut slide = ResolvedSlide {
            size,
            background: None,
            shapes: Vec::new(),
            diagnostics: Vec::new(),
        };
        let mut text_directions = Vec::new();
        for item in self.flatten() {
            match item {
                FlattenedItem::Background(background) => {
                    let source = match background.source {
                        BackgroundSource::Slide => FlattenedSource::Slide,
                        BackgroundSource::Layout => FlattenedSource::Layout,
                        BackgroundSource::Master => FlattenedSource::Master,
                    };
                    let (resolved_background, unsupported) = match background.content {
                        BackgroundContent::Model(BackgroundRendering::Properties(fill)) => {
                            match fill {
                                Some(fill) => self.concrete_background_fill(
                                    fill,
                                    size,
                                    source,
                                    media,
                                    &mut slide.diagnostics,
                                )?,
                                None => (None, None),
                            }
                        }
                        BackgroundContent::Model(BackgroundRendering::Reference {
                            index,
                            color,
                        }) => self.concrete_background_reference(
                            *index,
                            color.as_ref(),
                            size,
                            &mut slide.diagnostics,
                        )?,
                        BackgroundContent::Model(BackgroundRendering::Unsupported(detail)) => {
                            (None, Some(*detail))
                        }
                    };
                    slide.background = resolved_background;
                    if let Some(unsupported) = unsupported {
                        slide.diagnostics.push(Diagnostic {
                            message: format!("unsupported background {unsupported}"),
                        });
                    }
                }
                FlattenedItem::Shape {
                    source,
                    child,
                    group_scale,
                    group_transform,
                    group_issues,
                } => {
                    push_group_diagnostics(group_issues, &mut slide.diagnostics);
                    let mut shape_text_directions = Vec::new();
                    if let Some(shape) = self.resolve_flattened_shape(
                        ShapePlacement {
                            source,
                            group_scale,
                            group_transform,
                        },
                        child,
                        media,
                        (hyperlinks, charts),
                        fonts.as_deref_mut(),
                        (&mut slide.diagnostics, &mut shape_text_directions),
                    )? {
                        if let Some(identities) = identities.as_deref_mut() {
                            identities.push(self.shape_identity(source, child));
                        }
                        slide.shapes.push(shape);
                        text_directions.push(shape_text_directions);
                    }
                }
            }
        }
        Ok((slide, text_directions))
    }

    fn shape_identity(
        &self,
        source: FlattenedSource,
        child: &ShapeTreeChild,
    ) -> ResolvedShapeIdentity {
        let children: &[ShapeTreeChild] = match source {
            FlattenedSource::Slide => &self.slide.common_slide_data.shape_tree.children,
            FlattenedSource::Layout => &self.layout.common_slide_data.shape_tree.children,
            FlattenedSource::Master => &self.master.common_slide_data.shape_tree.children,
            FlattenedSource::Background => &[],
        };
        let mut containing_group_ids = Vec::new();
        find_group_lineage(children, child, &mut Vec::new(), &mut containing_group_ids);
        ResolvedShapeIdentity {
            source,
            shape_id: child.non_visual_id(),
            containing_group_ids,
            name: child.non_visual_name(),
        }
    }

    fn resolve_flattened_shape(
        &self,
        placement: ShapePlacement,
        child: &ShapeTreeChild,
        media: Option<&ScopedMediaIds>,
        scoped_resources: (
            Option<&ScopedHyperlinkTargets>,
            Option<&ScopedChartResources>,
        ),
        fonts: Option<&mut FontManager>,
        output: (
            &mut Vec<Diagnostic>,
            &mut Vec<Vec<oxml_layout::TextDirection>>,
        ),
    ) -> Result<Option<ResolvedShape>, ResolveError> {
        let (diagnostics, text_directions) = output;
        let (hyperlinks, charts) = scoped_resources;
        let ShapePlacement {
            source,
            group_scale,
            group_transform,
        } = placement;
        match child {
            ShapeTreeChild::Shape(shape) => self.resolve_ordinary_shape(
                shape,
                placement,
                media,
                hyperlinks,
                diagnostics,
                text_directions,
            ),
            ShapeTreeChild::Picture(picture) => {
                let Some((bounds, rotation_deg, flip_h, flip_v)) =
                    transform_values(self.effective_picture_xfrm(picture).as_ref())
                else {
                    return Ok(None);
                };
                let bounds = scaled_group_bounds(bounds, group_scale);
                let line_properties = picture.shape_properties.line.as_ref();
                let line = line_properties
                    .map(|line| self.concrete_line(line, (bounds.width, bounds.height)))
                    .transpose()?
                    .flatten();
                let (head_end, tail_end) = line_properties
                    .map(resolved_line_ends)
                    .unwrap_or((None, None));
                let shadow = picture
                    .shape_properties
                    .effects
                    .as_ref()
                    .and_then(|effects| effects.outer_shadow.as_ref())
                    .map(|shadow| self.concrete_shadow(shadow))
                    .transpose()?;
                let (geometry, geometry_unsupported) =
                    self.concrete_picture_geometry(picture, (bounds.width, bounds.height));
                if let Some(category) = geometry_unsupported {
                    diagnostics.push(Diagnostic {
                        message: format!("unsupported {category} retained as picture bounds"),
                    });
                }
                let diagnostic_start = diagnostics.len();
                let (content, media_unsupported) =
                    resolve_picture_content(picture, source, media, diagnostics);
                if self.media_poster_diagnostic_identity
                    && media_unsupported.is_some()
                    && picture.media.is_some()
                    && let Some(shape_id) = child.non_visual_id()
                    && let Some(diagnostic) = diagnostics.get_mut(diagnostic_start)
                {
                    diagnostic.message = format!(
                        "unresolved {} media poster for shape {shape_id}: {}",
                        flattened_source_name(source),
                        diagnostic.message
                    );
                }
                Ok(Some(ResolvedShape {
                    group_transform,
                    bounds,
                    rotation_deg,
                    flip_h,
                    flip_v,
                    geometry,
                    fill: None,
                    image_fill: None,
                    line,
                    head_end,
                    tail_end,
                    shadow,
                    content,
                    unsupported: media_unsupported.or(geometry_unsupported),
                }))
            }
            ShapeTreeChild::GraphicFrame(frame) => {
                let Some((bounds, rotation_deg, flip_h, flip_v)) =
                    transform_values(Some(&frame.transform))
                else {
                    return Ok(None);
                };
                let bounds = scaled_group_bounds(bounds, group_scale);
                let (content, unsupported, bounds_fallback) = match frame.graphic_data.payload() {
                    GraphicDataPayload::Table(table) => (
                        ResolvedContent::Table(self.resolve_table(
                            table,
                            source,
                            hyperlinks,
                            diagnostics,
                            text_directions,
                        )?),
                        None,
                        false,
                    ),
                    GraphicDataPayload::Chart(_) => self.resolve_chart_content(
                        frame,
                        None,
                        source,
                        bounds,
                        media,
                        charts,
                        fonts,
                        diagnostics,
                    )?,
                    GraphicDataPayload::SmartArt(_) => {
                        (ResolvedContent::None, Some("SmartArt"), true)
                    }
                    GraphicDataPayload::Ole { preview, .. } => {
                        if let Some(image) = resolve_ole_preview(preview.as_deref(), source, media)
                        {
                            diagnostics.push(Diagnostic {
                                message: "OLE object rendered as a static PNG preview. Embedded OLE interactivity is not rendered".to_owned(),
                            });
                            (
                                ResolvedContent::Image(image),
                                Some("embedded OLE interactivity"),
                                false,
                            )
                        } else {
                            (ResolvedContent::None, Some("OLE"), true)
                        }
                    }
                    GraphicDataPayload::Other(_) => {
                        (ResolvedContent::None, Some("unknown graphic frame"), true)
                    }
                };
                if bounds_fallback {
                    let category = unsupported.expect("bounds fallback has a category");
                    diagnostics.push(Diagnostic {
                        message: format!("unsupported {category} content retained as bounds"),
                    });
                }
                Ok(Some(ResolvedShape {
                    group_transform,
                    bounds,
                    rotation_deg,
                    flip_h,
                    flip_v,
                    geometry: if bounds_fallback {
                        ResolvedGeometry::BoundsFallback
                    } else {
                        ResolvedGeometry::Rectangle
                    },
                    fill: None,
                    image_fill: None,
                    line: None,
                    head_end: None,
                    tail_end: None,
                    shadow: None,
                    content,
                    unsupported,
                }))
            }
            ShapeTreeChild::Connector(connector) => {
                let Some((bounds, rotation_deg, flip_h, flip_v)) =
                    connector_transform_values(connector.shape_properties.transform.as_ref())
                else {
                    return Ok(None);
                };
                let bounds = scaled_group_bounds(bounds, group_scale);
                let (fill, fill_unsupported) = connector
                    .shape_properties
                    .fill
                    .as_ref()
                    .map(|fill| self.concrete_fill(fill, (bounds.width, bounds.height)))
                    .transpose()?
                    .unwrap_or((None, None));
                let has_direct_line = connector.shape_properties.line.is_some();
                let mut line = connector
                    .shape_properties
                    .line
                    .as_ref()
                    .map(|line| self.concrete_line(line, (bounds.width, bounds.height)))
                    .transpose()?
                    .flatten();
                let (head_end, tail_end) = connector
                    .shape_properties
                    .line
                    .as_ref()
                    .map(resolved_line_ends)
                    .unwrap_or((None, None));
                let (geometry, geometry_unsupported, geometry_diagnostic) = if let Some(custom) =
                    &connector.shape_properties.custom_geometry
                {
                    match self.concrete_custom_geometry(custom, (bounds.width, bounds.height)) {
                        Ok(geometry) => (geometry, None, None),
                        Err(error) => (
                            ResolvedGeometry::BoundsFallback,
                            Some("connector custom geometry evaluation"),
                            Some(format!(
                                "connector custom geometry evaluation failed: {error}; retained as shape bounds"
                            )),
                        ),
                    }
                } else if let Some(preset) = &connector.shape_properties.preset_geometry {
                    {
                        match self.concrete_preset_geometry(preset, (bounds.width, bounds.height)) {
                            Ok(Some(geometry)) => (geometry, None, None),
                            Ok(None) => (
                                ResolvedGeometry::BoundsFallback,
                                Some("unknown connector preset geometry"),
                                Some(format!(
                                    "unknown connector preset geometry `{}` retained as shape bounds",
                                    preset.preset
                                )),
                            ),
                            Err(error) => (
                                ResolvedGeometry::BoundsFallback,
                                Some("connector preset geometry evaluation"),
                                Some(format!(
                                    "connector preset geometry `{}` evaluation failed: {error}; retained as shape bounds",
                                    preset.preset
                                )),
                            ),
                        }
                    }
                } else {
                    (
                        ResolvedGeometry::BoundsFallback,
                        Some("connector geometry"),
                        Some("unsupported connector geometry retained as shape bounds".to_owned()),
                    )
                };
                if let Some(category) = fill_unsupported {
                    diagnostics.push(Diagnostic {
                        message: format!(
                            "unsupported connector {category} retained as shape bounds"
                        ),
                    });
                }
                if let Some(message) = geometry_diagnostic {
                    diagnostics.push(Diagnostic { message });
                }
                let line_unsupported = (!has_direct_line).then_some("connector line style");
                if line_unsupported.is_some() {
                    line = Some(Stroke::new(Paint::Solid(Color::BLACK), 1.0));
                    diagnostics.push(Diagnostic {
                        message: "unsupported connector line style retained as visible default"
                            .to_owned(),
                    });
                }
                let unsupported = fill_unsupported
                    .or(geometry_unsupported)
                    .or(line_unsupported);
                Ok(Some(ResolvedShape {
                    group_transform,
                    bounds,
                    rotation_deg,
                    flip_h,
                    flip_v,
                    geometry,
                    fill,
                    image_fill: None,
                    line,
                    head_end,
                    tail_end,
                    shadow: None,
                    content: ResolvedContent::None,
                    unsupported,
                }))
            }
            ShapeTreeChild::AlternateContent(alternate) => {
                let Some(frame) = alternate.chart_choice() else {
                    return Ok(None);
                };
                let Some((bounds, rotation_deg, flip_h, flip_v)) =
                    transform_values(Some(&frame.transform))
                else {
                    return Ok(None);
                };
                let bounds = scaled_group_bounds(bounds, group_scale);
                let (content, unsupported, bounds_fallback) = if charts.is_some() && fonts.is_some()
                {
                    self.resolve_chart_content(
                        frame,
                        alternate.picture_fallback(),
                        source,
                        bounds,
                        media,
                        charts,
                        fonts,
                        diagnostics,
                    )?
                } else if let Some(picture) = alternate.picture_fallback() {
                    let (content, _) = resolve_picture_content(picture, source, media, diagnostics);
                    let has_image = matches!(content, ResolvedContent::Image(_));
                    (content, Some("chart"), !has_image)
                } else {
                    (ResolvedContent::None, Some("chart"), true)
                };
                if bounds_fallback {
                    diagnostics.push(Diagnostic {
                        message: "unsupported chart content retained as bounds".to_owned(),
                    });
                }
                Ok(Some(ResolvedShape {
                    group_transform,
                    bounds,
                    rotation_deg,
                    flip_h,
                    flip_v,
                    geometry: if bounds_fallback {
                        ResolvedGeometry::BoundsFallback
                    } else {
                        ResolvedGeometry::Rectangle
                    },
                    fill: None,
                    image_fill: None,
                    line: None,
                    head_end: None,
                    tail_end: None,
                    shadow: None,
                    content,
                    unsupported,
                }))
            }
            ShapeTreeChild::GroupShape(_) => Ok(None),
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn resolve_chart_content(
        &self,
        frame: &rpptx_oxml::graphic_frame::CT_GraphicFrame,
        fallback: Option<&CT_Picture>,
        source: FlattenedSource,
        bounds: Rect,
        media: Option<&ScopedMediaIds>,
        charts: Option<&ScopedChartResources>,
        fonts: Option<&mut FontManager>,
        diagnostics: &mut Vec<Diagnostic>,
    ) -> Result<(ResolvedContent, Option<&'static str>, bool), ResolveError> {
        let Some(charts) = charts else {
            return Ok((ResolvedContent::None, Some("chart"), true));
        };
        let Some(fonts) = fonts else {
            return Ok((ResolvedContent::None, Some("chart"), true));
        };
        let relationship_id = frame.chart_relationship_id();
        let resource = relationship_id.and_then(|id| charts.get(source, id));
        let failure = match (relationship_id, resource) {
            (None, _) => "chart payload has no relationship identifier".to_owned(),
            (Some(id), None) => format!(
                "missing {} chart relationship `{id}`",
                flattened_source_name(source)
            ),
            (Some(_), Some(ChartResource::External(target))) => format!(
                "external {} chart relationship `{}` targets `{target}`",
                flattened_source_name(source),
                relationship_id.expect("matched a relationship id")
            ),
            (Some(_), Some(ChartResource::MissingTarget(target))) => format!(
                "missing {} chart target `{target}` for relationship `{}`",
                flattened_source_name(source),
                relationship_id.expect("matched a relationship id")
            ),
            (Some(_), Some(ChartResource::Invalid(detail))) => format!(
                "invalid {} chart relationship `{}`: {detail}",
                flattened_source_name(source),
                relationship_id.expect("matched a relationship id")
            ),
            (Some(_), Some(ChartResource::Parsed(chart_space))) => {
                let local_bounds = Rect {
                    x: 0.0,
                    y: 0.0,
                    width: bounds.width,
                    height: bounds.height,
                };
                match render_chart(
                    &chart_space.chart,
                    local_bounds,
                    self.theme,
                    &self.color_map,
                    fonts,
                ) {
                    Ok(group) => {
                        return Ok((ResolvedContent::Group(group), None, false));
                    }
                    Err(error) => error.to_string(),
                }
            }
        };

        if let Some(picture) = fallback {
            let (content, _) = resolve_picture_content(picture, source, media, diagnostics);
            let relationship_id = picture
                .blip_fill
                .as_ref()
                .and_then(|fill| fill.blip.as_ref())
                .and_then(|blip| blip.embed.as_deref());
            let renderer_compatible = relationship_id
                .and_then(|id| media.and_then(|media| media.content_type(source, id)))
                .is_some_and(renderer_compatible_image_content_type);
            if matches!(content, ResolvedContent::Image(_)) && renderer_compatible {
                diagnostics.push(Diagnostic {
                    message: format!("unsupported chart rendered as cached image: {failure}"),
                });
                return Ok((content, Some("chart"), false));
            }
            if matches!(content, ResolvedContent::Image(_)) {
                diagnostics.push(Diagnostic {
                    message: format!(
                        "unsupported chart cached image is not renderer-compatible: {failure}"
                    ),
                });
            }
        }

        let placeholder = render_chart_placeholder(
            Rect {
                x: 0.0,
                y: 0.0,
                width: bounds.width,
                height: bounds.height,
            },
            fonts,
        )
        .map_err(|error| ResolveError::ConcreteValue {
            kind: "chart placeholder",
            detail: error.to_string(),
        })?;
        diagnostics.push(Diagnostic {
            message: format!("unsupported chart rendered as labelled placeholder: {failure}"),
        });
        Ok((ResolvedContent::Group(placeholder), Some("chart"), true))
    }

    fn resolve_ordinary_shape(
        &self,
        shape: &CT_Shape,
        placement: ShapePlacement,
        media: Option<&ScopedMediaIds>,
        hyperlinks: Option<&ScopedHyperlinkTargets>,
        diagnostics: &mut Vec<Diagnostic>,
        text_directions: &mut Vec<Vec<oxml_layout::TextDirection>>,
    ) -> Result<Option<ResolvedShape>, ResolveError> {
        let ShapePlacement {
            source,
            group_scale,
            group_transform,
        } = placement;
        let transform = self.effective_xfrm(shape);
        let Some((unscaled_bounds, rotation_deg, flip_h, flip_v)) =
            transform_values(transform.as_ref())
        else {
            return Ok(None);
        };
        let mut bounds = scaled_group_bounds(unscaled_bounds, group_scale);
        let mut group_transform = group_transform;
        let effective = self.effective_shape_style(shape)?;
        let (fill, image_fill, fill_unsupported, fill_diagnostic) = match effective.fill.as_ref() {
            Some(Fill::Blip(fill))
                if matches!(shape.shape_properties.fill, Some(Fill::Blip(_))) =>
            {
                match resolve_image_fill(fill, source, media, "shape") {
                    Ok(image) => (None, Some(image), None, None),
                    Err((category, message)) => (None, None, Some(category), Some(message)),
                }
            }
            Some(Fill::Blip(_)) => (
                None,
                None,
                Some("theme-referenced shape picture fill"),
                Some(
                    "theme-referenced shape picture fill requires theme relationship scope"
                        .to_owned(),
                ),
            ),
            Some(fill) => {
                let normalized = match fill {
                    Fill::Gradient(gradient)
                        if gradient.rotate_with_shape == Some(false)
                            && rotation_deg.rem_euclid(360.0) == 0.0
                            && !flip_h
                            && !flip_v
                            && group_transform.a > 0.0
                            && group_transform.b.abs() < 1.0e-10
                            && group_transform.c.abs() < 1.0e-10
                            && group_transform.d > 0.0 =>
                    {
                        let mut normalized = fill.clone();
                        let Fill::Gradient(gradient) = &mut normalized else {
                            unreachable!("matched a gradient fill")
                        };
                        gradient.rotate_with_shape = None;
                        Some(normalized)
                    }
                    _ => None,
                };
                let (fill, unsupported) = self.concrete_fill(
                    normalized.as_ref().unwrap_or(fill),
                    (bounds.width, bounds.height),
                )?;
                (fill, None, unsupported, None)
            }
            None => (None, None, None, None),
        };
        let line = effective
            .line
            .as_ref()
            .map(|line| self.concrete_line(line, (bounds.width, bounds.height)))
            .transpose()?
            .flatten();
        let (head_end, tail_end) = effective
            .line
            .as_ref()
            .map(resolved_line_ends)
            .unwrap_or((None, None));
        let shadow = effective
            .effects
            .as_ref()
            .and_then(|effects| effects.outer_shadow.as_ref())
            .map(|shadow| self.concrete_shadow(shadow))
            .transpose()?;
        let transparent_text_only = shape.text_body.is_some()
            && fill.is_none()
            && image_fill.is_none()
            && fill_unsupported.is_none()
            && line.is_none()
            && shadow.is_none();
        if transparent_text_only {
            bounds = unscaled_bounds;
            group_transform = Transform {
                a: group_scale.0,
                d: group_scale.1,
                ..Transform::IDENTITY
            }
            .then(group_transform);
        }
        let (geometry, geometry_unsupported, geometry_diagnostic) = if let Some(custom) =
            &shape.shape_properties.custom_geometry
        {
            match self.concrete_custom_geometry(custom, (bounds.width, bounds.height)) {
                Ok(geometry) => (geometry, None, None),
                Err(error) => (
                    ResolvedGeometry::BoundsFallback,
                    Some("custom geometry evaluation"),
                    Some(format!(
                        "custom geometry evaluation failed: {error}; retained as shape bounds"
                    )),
                ),
            }
        } else if let Some(preset) = &shape.shape_properties.preset_geometry {
            match self.concrete_preset_geometry(preset, (bounds.width, bounds.height)) {
                Ok(Some(geometry)) => (geometry, None, None),
                Ok(None) => (
                    ResolvedGeometry::BoundsFallback,
                    Some("unknown preset geometry"),
                    Some(format!(
                        "unknown preset geometry `{}` retained as shape bounds",
                        preset.preset
                    )),
                ),
                Err(error) => (
                    ResolvedGeometry::BoundsFallback,
                    Some("preset geometry evaluation"),
                    Some(format!(
                        "preset geometry `{}` evaluation failed: {error}; retained as shape bounds",
                        preset.preset
                    )),
                ),
            }
        } else if shape.text_body.is_some() {
            (ResolvedGeometry::Rectangle, None, None)
        } else {
            (
                ResolvedGeometry::BoundsFallback,
                Some("preset geometry pending evaluation"),
                None,
            )
        };
        let content = if let Some(body) = shape.text_body.as_ref() {
            let (body, directions) =
                self.resolve_text_body(shape, body, source, hyperlinks, diagnostics)?;
            text_directions.push(directions);
            ResolvedContent::Text(body)
        } else {
            ResolvedContent::None
        };
        if let ResolvedContent::Text(text) = &content
            && let Some(message) = vertical_text_diagnostic(text.vertical)
        {
            diagnostics.push(Diagnostic {
                message: message.to_owned(),
            });
        }
        if let Some(category) = fill_unsupported {
            diagnostics.push(Diagnostic {
                message: fill_diagnostic
                    .unwrap_or_else(|| format!("unsupported {category} retained as shape bounds")),
            });
        }
        if let Some(category) = geometry_unsupported {
            diagnostics.push(Diagnostic {
                message: geometry_diagnostic
                    .unwrap_or_else(|| format!("unsupported {category} retained as shape bounds")),
            });
        }
        Ok(Some(ResolvedShape {
            group_transform,
            bounds,
            rotation_deg,
            flip_h,
            flip_v,
            geometry,
            fill,
            image_fill,
            line,
            head_end,
            tail_end,
            shadow,
            content,
            unsupported: fill_unsupported.or(geometry_unsupported),
        }))
    }

    fn concrete_picture_geometry(
        &self,
        picture: &rpptx_oxml::picture::CT_Picture,
        size: (f64, f64),
    ) -> (ResolvedGeometry, Option<&'static str>) {
        if let Some(custom) = &picture.shape_properties.custom_geometry {
            return self
                .concrete_custom_geometry(custom, size)
                .map(|geometry| (geometry, None))
                .unwrap_or((ResolvedGeometry::Rectangle, Some("picture custom geometry")));
        }
        if let Some(preset) = &picture.shape_properties.preset_geometry {
            return match self.concrete_preset_geometry(preset, size) {
                Ok(Some(geometry)) => (geometry, None),
                Ok(None) => (
                    ResolvedGeometry::Rectangle,
                    Some("unknown picture preset geometry"),
                ),
                Err(_) => (
                    ResolvedGeometry::Rectangle,
                    Some("picture preset geometry evaluation"),
                ),
            };
        }
        (ResolvedGeometry::Rectangle, None)
    }
}

fn renderer_compatible_image_content_type(content_type: &str) -> bool {
    content_type.eq_ignore_ascii_case("image/png")
        || content_type.eq_ignore_ascii_case("image/jpeg")
        || content_type.eq_ignore_ascii_case("image/jpg")
}

fn resolve_picture_content(
    picture: &rpptx_oxml::picture::CT_Picture,
    source: FlattenedSource,
    media: Option<&ScopedMediaIds>,
    diagnostics: &mut Vec<Diagnostic>,
) -> (ResolvedContent, Option<&'static str>) {
    let Some(fill) = picture.blip_fill.as_ref() else {
        return unsupported_picture(
            "alternate picture media",
            "alternate picture media is not resolved",
            diagnostics,
        );
    };
    match resolve_image_fill(fill, source, media, "picture") {
        Ok(image) => (ResolvedContent::Image(image), None),
        Err((category, message)) => unsupported_picture(category, &message, diagnostics),
    }
}

fn resolve_ole_preview(
    preview: Option<&rpptx_oxml::picture::CT_Picture>,
    source: FlattenedSource,
    media: Option<&ScopedMediaIds>,
) -> Option<ResolvedImage> {
    let fill = preview?.blip_fill.as_ref()?;
    let relationship_id = fill.blip.as_ref()?.embed.as_deref()?;
    let media = media?;
    if !media
        .content_type(source, relationship_id)
        .is_some_and(|content_type| content_type.eq_ignore_ascii_case("image/png"))
    {
        return None;
    }
    resolve_image_fill(fill, source, Some(media), "picture").ok()
}

fn resolve_image_fill(
    fill: &BlipFill,
    source: FlattenedSource,
    media: Option<&ScopedMediaIds>,
    role: &str,
) -> Result<ResolvedImage, (&'static str, String)> {
    let missing_blip_category = match role {
        "picture" => "missing picture blip",
        "shape" => "missing shape picture blip",
        _ => "missing background blip",
    };
    let external_media_category = match role {
        "picture" => "external picture media",
        "shape" => "external shape picture media",
        _ => "external background media",
    };
    let Some(blip) = fill.blip.as_ref() else {
        return Err((
            missing_blip_category,
            format!("{role} fill has no modelled embedded blip"),
        ));
    };
    let Some(relationship_id) = blip.embed.as_deref() else {
        let message = if let Some(link) = blip.link.as_deref() {
            format!("external {role} relationship `{link}` is unsupported")
        } else {
            format!("{role} blip has no embedded relationship")
        };
        return Err((external_media_category, message));
    };
    let Some(media) = media else {
        return Err((
            "image media pending relationship resolution",
            format!("{role} image media pending relationship resolution"),
        ));
    };
    let Some(media_id) = media.get(source, relationship_id) else {
        return Err((
            "missing image relationship",
            format!(
                "missing {} {role} image relationship `{relationship_id}`",
                flattened_source_name(source)
            ),
        ));
    };
    Ok(ResolvedImage {
        media: media_id,
        src_rect: fill.source_rect.as_ref().map(resolved_crop_rect),
        placement: match fill.mode.as_ref() {
            Some(BlipMode::Tile(tile)) => ResolvedImagePlacement::Tile(ResolvedTilePlacement {
                translation: Point {
                    x: emu_to_points(tile.translation_x.unwrap_or(0)),
                    y: emu_to_points(tile.translation_y.unwrap_or(0)),
                },
                scale_x: tile
                    .scale_x
                    .map_or(1.0, |value| f64::from(value.0) / 100_000.0),
                scale_y: tile
                    .scale_y
                    .map_or(1.0, |value| f64::from(value.0) / 100_000.0),
                flip: resolved_tile_flip(tile.flip.as_deref()),
                alignment: resolved_rect_alignment(tile.alignment.as_deref()),
            }),
            Some(BlipMode::Stretch { fill_rect, .. }) => ResolvedImagePlacement::Stretch {
                fill_rect: fill_rect.as_ref().map(resolved_crop_rect),
            },
            None => ResolvedImagePlacement::default(),
        },
        dpi: fill.dpi.filter(|dpi| *dpi > 0).map(f64::from),
        rotate_with_shape: fill.rotate_with_shape.unwrap_or(true),
        opacity: blip.alpha_modulation_fix().map_or(1.0, |amount| {
            (f64::from(amount.0) / 100_000.0).clamp(0.0, 1.0)
        }),
    })
}

fn unsupported_picture(
    category: &'static str,
    message: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> (ResolvedContent, Option<&'static str>) {
    diagnostics.push(Diagnostic {
        message: message.to_owned(),
    });
    (ResolvedContent::None, Some(category))
}

fn flattened_source_name(source: FlattenedSource) -> &'static str {
    match source {
        FlattenedSource::Background => "background",
        FlattenedSource::Master => "master",
        FlattenedSource::Layout => "layout",
        FlattenedSource::Slide => "slide",
    }
}

fn resolve_direct_hyperlink(
    properties: Option<&CT_TextCharacterProperties>,
    source: FlattenedSource,
    hyperlinks: Option<&ScopedHyperlinkTargets>,
    diagnostics: &mut Vec<Diagnostic>,
) -> Option<String> {
    let hyperlink = properties?.hyperlink_click.as_ref()?;
    if let Some(action) = hyperlink.action.as_deref() {
        diagnostics.push(Diagnostic {
            message: format!(
                "unsupported {} hyperlink action `{action}`",
                flattened_source_name(source)
            ),
        });
        return None;
    }
    let Some(relationship_id) = hyperlink.relationship_id.as_deref() else {
        diagnostics.push(Diagnostic {
            message: format!(
                "unsupported {} hyperlink without an external relationship",
                flattened_source_name(source)
            ),
        });
        return None;
    };
    let Some(target) = hyperlinks.and_then(|targets| targets.get(source, relationship_id)) else {
        diagnostics.push(Diagnostic {
            message: format!(
                "missing {} hyperlink relationship `{relationship_id}`",
                flattened_source_name(source)
            ),
        });
        return None;
    };
    Some(target.to_owned())
}

fn resolved_crop_rect(rect: &RelativeRect) -> crate::CropRect {
    crate::CropRect {
        left: rect
            .left
            .map_or(0.0, |value| f64::from(value.0) / 100_000.0),
        top: rect.top.map_or(0.0, |value| f64::from(value.0) / 100_000.0),
        right: rect
            .right
            .map_or(0.0, |value| f64::from(value.0) / 100_000.0),
        bottom: rect
            .bottom
            .map_or(0.0, |value| f64::from(value.0) / 100_000.0),
    }
}

fn resolved_tile_flip(value: Option<&str>) -> ResolvedTileFlip {
    match value {
        Some("x") => ResolvedTileFlip::Horizontal,
        Some("y") => ResolvedTileFlip::Vertical,
        Some("xy") => ResolvedTileFlip::Both,
        Some("none") | None => ResolvedTileFlip::None,
        Some(_) => unreachable!("DrawingML tile flip is validated while parsing"),
    }
}

fn resolved_rect_alignment(value: Option<&str>) -> ResolvedRectAlignment {
    match value {
        Some("t") => ResolvedRectAlignment::Top,
        Some("tr") => ResolvedRectAlignment::TopRight,
        Some("l") => ResolvedRectAlignment::Left,
        Some("ctr") => ResolvedRectAlignment::Center,
        Some("r") => ResolvedRectAlignment::Right,
        Some("bl") => ResolvedRectAlignment::BottomLeft,
        Some("b") => ResolvedRectAlignment::Bottom,
        Some("br") => ResolvedRectAlignment::BottomRight,
        Some("tl") | None => ResolvedRectAlignment::TopLeft,
        Some(_) => unreachable!("DrawingML rectangle alignment is validated while parsing"),
    }
}

fn resolved_line_ends(
    line: &CT_LineProperties,
) -> (Option<ResolvedLineEnd>, Option<ResolvedLineEnd>) {
    (
        line.head_end.as_ref().and_then(resolved_line_end),
        line.tail_end.as_ref().and_then(resolved_line_end),
    )
}

fn resolved_line_end(end: &LineEnd) -> Option<ResolvedLineEnd> {
    let kind = match end.kind? {
        LineEndType::None => return None,
        LineEndType::Triangle => ResolvedLineEndKind::Triangle,
        LineEndType::Stealth => ResolvedLineEndKind::Stealth,
        LineEndType::Diamond => ResolvedLineEndKind::Diamond,
        LineEndType::Oval => ResolvedLineEndKind::Oval,
        LineEndType::Arrow => ResolvedLineEndKind::Arrow,
    };
    Some(ResolvedLineEnd {
        kind,
        width: resolved_line_end_size(end.width),
        length: resolved_line_end_size(end.length),
    })
}

fn resolved_line_end_size(size: Option<LineEndSize>) -> ResolvedLineEndSize {
    match size.unwrap_or(LineEndSize::Medium) {
        LineEndSize::Small => ResolvedLineEndSize::Small,
        LineEndSize::Medium => ResolvedLineEndSize::Medium,
        LineEndSize::Large => ResolvedLineEndSize::Large,
    }
}

impl ResolveCtx<'_> {
    fn concrete_background_fill(
        &self,
        fill: &Fill,
        size: (f64, f64),
        source: FlattenedSource,
        media: Option<&ScopedMediaIds>,
        diagnostics: &mut Vec<Diagnostic>,
    ) -> Result<(Option<ResolvedBackground>, Option<&'static str>), ResolveError> {
        let mut fill = fill.clone();
        let slot = self.color_map.theme_slot(ColorMapSlot::Background1);
        let reference = self.theme.theme_elements.color_scheme.color(slot);
        substitute_fill(&mut fill, Some(reference), "background")?;
        self.concrete_background_value(&fill, size, source, media, diagnostics)
    }

    fn concrete_background_reference(
        &self,
        index: u32,
        color: Option<&ColorChoice>,
        size: (f64, f64),
        diagnostics: &mut Vec<Diagnostic>,
    ) -> Result<(Option<ResolvedBackground>, Option<&'static str>), ResolveError> {
        let matrix = &self.theme.theme_elements.format_scheme;
        let Some(mut fill) =
            referenced_fill(index, &matrix.fill_styles, &matrix.background_fill_styles)?
        else {
            return Ok((None, None));
        };
        substitute_fill(&mut fill, color, "background")?;
        if matches!(fill, Fill::Blip(_)) {
            diagnostics.push(Diagnostic {
                message: "theme-referenced background picture requires theme relationship scope"
                    .to_owned(),
            });
            return Ok((None, None));
        }
        let (paint, unsupported) = self.concrete_background_paint(&fill, size)?;
        Ok((paint.map(ResolvedBackground::Paint), unsupported))
    }

    fn concrete_background_value(
        &self,
        fill: &Fill,
        size: (f64, f64),
        source: FlattenedSource,
        media: Option<&ScopedMediaIds>,
        diagnostics: &mut Vec<Diagnostic>,
    ) -> Result<(Option<ResolvedBackground>, Option<&'static str>), ResolveError> {
        if let Fill::Blip(fill) = fill {
            return Ok(
                match resolve_image_fill(fill, source, media, "background") {
                    Ok(image) => (Some(ResolvedBackground::Image(image)), None),
                    Err((_category, message)) => {
                        diagnostics.push(Diagnostic { message });
                        (None, None)
                    }
                },
            );
        }
        let (paint, unsupported) = self.concrete_background_paint(fill, size)?;
        Ok((paint.map(ResolvedBackground::Paint), unsupported))
    }

    fn concrete_background_paint(
        &self,
        fill: &Fill,
        size: (f64, f64),
    ) -> Result<(Option<Paint>, Option<&'static str>), ResolveError> {
        let mut fill = fill.clone();
        if let Fill::Gradient(gradient) = &mut fill {
            gradient.rotate_with_shape = None;
        }
        self.concrete_fill(&fill, size)
    }

    fn concrete_fill(
        &self,
        fill: &Fill,
        size: (f64, f64),
    ) -> Result<(Option<Paint>, Option<&'static str>), ResolveError> {
        match fill {
            Fill::NoFill(_) => Ok((None, None)),
            Fill::Solid(solid) => {
                let Some(choice) = solid.color.as_ref() else {
                    return Ok((None, Some("solid fill without colour")));
                };
                Ok((Some(Paint::Solid(self.concrete_color(choice)?)), None))
            }
            Fill::Gradient(gradient) => {
                if gradient.rotate_with_shape == Some(false) {
                    return Ok((None, Some("gradient independent of shape rotation")));
                }
                match gradient.flip.as_deref() {
                    Some("x") => return Ok((None, Some("horizontal gradient flip"))),
                    Some("y") => return Ok((None, Some("vertical gradient flip"))),
                    Some("xy") => return Ok((None, Some("horizontal and vertical gradient flip"))),
                    Some("none") | None => {}
                    Some(_) => unreachable!("DrawingML parser validates gradient flip tokens"),
                }
                if gradient.tile_rect.is_some() {
                    return Ok((None, Some("gradient tile rectangle")));
                }
                if let Some(GradientGeometry::Path(path)) = &gradient.geometry {
                    let unsupported = match path.kind {
                        PathGradientKind::Circle => "circle path gradient",
                        PathGradientKind::Rectangle => "rectangle path gradient",
                        PathGradientKind::Shape => "shape path gradient",
                    };
                    return Ok((None, Some(unsupported)));
                }
                let mut stops = Vec::with_capacity(gradient.stops.len());
                for stop in &gradient.stops {
                    let Some(choice) = stop.color.as_ref() else {
                        return Ok((None, Some("gradient stop without colour")));
                    };
                    stops.push(GradientStop {
                        offset: f64::from(stop.position.0) / 100_000.0,
                        color: self.concrete_color(choice)?,
                    });
                }
                if stops.is_empty() {
                    return Ok((None, Some("gradient without concrete stops")));
                }
                let paint = match &gradient.geometry {
                    Some(GradientGeometry::Linear(linear)) => {
                        let radians = f64::from(linear.angle.0) / 60_000.0_f64;
                        let radians = radians.to_radians();
                        let center = Point {
                            x: size.0 / 2.0,
                            y: size.1 / 2.0,
                        };
                        let mut direction = Point {
                            x: radians.cos(),
                            y: radians.sin(),
                        };
                        if linear.scaled == Some(true) {
                            direction.x *= size.0;
                            direction.y *= size.1;
                        }
                        let length = direction.x.hypot(direction.y);
                        if length > 0.0 {
                            direction.x /= length;
                            direction.y /= length;
                        }
                        let extent =
                            (direction.x.abs() * size.0 + direction.y.abs() * size.1) / 2.0;
                        let delta = Point {
                            x: direction.x * extent,
                            y: direction.y * extent,
                        };
                        Paint::linear(
                            Point {
                                x: center.x - delta.x,
                                y: center.y - delta.y,
                            },
                            Point {
                                x: center.x + delta.x,
                                y: center.y + delta.y,
                            },
                            stops,
                            (true, true),
                        )
                    }
                    None => Paint::linear(
                        Point { x: 0.0, y: 0.0 },
                        Point { x: size.0, y: 0.0 },
                        stops,
                        (true, true),
                    ),
                    Some(GradientGeometry::Path(_)) => unreachable!("rejected above"),
                };
                Ok((Some(paint), None))
            }
            Fill::Pattern(_) => Ok((None, Some("pattern fill"))),
            Fill::Blip(_) => Ok((None, Some("picture fill media"))),
        }
    }

    fn concrete_line(
        &self,
        line: &CT_LineProperties,
        size: (f64, f64),
    ) -> Result<Option<Stroke>, ResolveError> {
        let Some(fill) = line.fill.as_ref() else {
            return Ok(None);
        };
        let (Some(paint), _) = self.concrete_fill(fill, size)? else {
            return Ok(None);
        };
        let width = emu_to_points(i64::from(line.width.unwrap_or(12_700)));
        let cap = match line.cap.unwrap_or(DrawingLineCap::Flat) {
            DrawingLineCap::Round => LineCap::Round,
            DrawingLineCap::Square => LineCap::Square,
            DrawingLineCap::Flat => LineCap::Butt,
        };
        let join = match line.join.as_ref() {
            Some(DrawingLineJoin::Round { .. }) => LineJoin::Round,
            Some(DrawingLineJoin::Bevel { .. }) => LineJoin::Bevel,
            Some(DrawingLineJoin::Miter { .. }) | None => LineJoin::Miter,
        };
        let dash = line.dash.as_ref().and_then(|dash| match dash {
            LineDash::Preset(preset) => {
                let values = preset
                    .value
                    .dash_array()
                    .iter()
                    .map(|value| f64::from(*value) * width)
                    .collect::<Vec<_>>();
                (!values.is_empty()).then_some(values)
            }
            LineDash::Custom(custom) => {
                let values = custom
                    .stops
                    .iter()
                    .flat_map(|stop| [stop.dash, stop.space])
                    .map(|value| f64::from(value) / 100_000.0 * width)
                    .collect::<Vec<_>>();
                (!values.is_empty()).then_some(values)
            }
        });
        Ok(Some(Stroke {
            paint,
            width,
            cap,
            join,
            dash,
        }))
    }

    fn concrete_shadow(&self, shadow: &CT_OuterShadowEffect) -> Result<Effect, ResolveError> {
        let color = shadow
            .color
            .as_ref()
            .map(|choice| self.concrete_color(choice))
            .transpose()?
            .unwrap_or(Color::BLACK);
        let distance = emu_to_points(shadow.distance.unwrap_or(0));
        let angle = f64::from(shadow.direction.unwrap_or_default().0) / 60_000.0;
        let radians = angle.to_radians();
        Ok(Effect::OuterShadow {
            dx: radians.cos() * distance,
            dy: radians.sin() * distance,
            blur: emu_to_points(shadow.blur_radius.unwrap_or(0)),
            color,
        })
    }

    fn concrete_color(&self, choice: &ColorChoice) -> Result<Color, ResolveError> {
        let owned_lookup = self.theme_lookup();
        let lookup = owned_lookup
            .iter()
            .map(|(name, color)| (name.as_str(), *color))
            .collect::<Vec<_>>();
        let color = resolve_color(choice, &self.color_map, &lookup).map_err(|error| {
            ResolveError::ConcreteValue {
                kind: "colour",
                detail: error.to_string(),
            }
        })?;
        Ok(Color {
            r: f64::from(color.red) / 255.0,
            g: f64::from(color.green) / 255.0,
            b: f64::from(color.blue) / 255.0,
            a: f64::from(color.alpha) / 255.0,
        })
    }

    fn theme_lookup(&self) -> Vec<(String, RgbColor)> {
        let mut lookup = self
            .theme
            .theme_elements
            .color_scheme
            .iter()
            .filter_map(|(slot, choice)| {
                base_rgb(choice).map(|color| (slot.as_str().to_owned(), color))
            })
            .collect::<Vec<_>>();
        lookup.push(("black".to_owned(), RgbColor::new(0, 0, 0)));
        lookup.push(("white".to_owned(), RgbColor::new(255, 255, 255)));
        lookup
    }

    fn concrete_custom_geometry(
        &self,
        geometry: &oxml_drawing::geometry::CT_CustomGeometry2D,
        size: (f64, f64),
    ) -> Result<ResolvedGeometry, ResolveError> {
        let evaluated = geometry
            .evaluate_with_size(&BTreeMap::new(), size)
            .map_err(|error| ResolveError::ConcreteValue {
                kind: "custom geometry",
                detail: error.to_string(),
            })?;
        let paths = evaluated
            .paths
            .into_iter()
            .zip(geometry.paths())
            .map(|(commands, source)| {
                let scale_x = source.width.map_or(1.0, |width| size.0 / width);
                let scale_y = source.height.map_or(1.0, |height| size.1 / height);
                Path {
                    commands: commands
                        .into_iter()
                        .map(|command| concrete_path_command(command, scale_x, scale_y))
                        .collect(),
                    fill_rule: FillRule::NonZero,
                }
            })
            .collect();
        let first_path = geometry.paths().first();
        let text_scale_x = first_path
            .and_then(|path| path.width)
            .map_or(1.0, |width| size.0 / width);
        let text_scale_y = first_path
            .and_then(|path| path.height)
            .map_or(1.0, |height| size.1 / height);
        let text_rect = evaluated.text_rectangle.map(|rectangle| Rect {
            x: rectangle.left * text_scale_x,
            y: rectangle.top * text_scale_y,
            width: (rectangle.right - rectangle.left) * text_scale_x,
            height: (rectangle.bottom - rectangle.top) * text_scale_y,
        });
        Ok(ResolvedGeometry::Custom { paths, text_rect })
    }

    fn concrete_preset_geometry(
        &self,
        geometry: &oxml_drawing::geometry::CT_PresetGeometry2D,
        size: (f64, f64),
    ) -> Result<Option<ResolvedGeometry>, ResolveError> {
        let evaluated = geometry
            .evaluate(size)
            .map_err(|error| ResolveError::ConcreteValue {
                kind: "preset geometry",
                detail: error.to_string(),
            })?;
        Ok(evaluated.map(|evaluated| {
            let paths = evaluated
                .paths
                .into_iter()
                .map(|commands| Path {
                    commands: commands
                        .into_iter()
                        .map(|command| concrete_path_command(command, 1.0, 1.0))
                        .collect(),
                    fill_rule: FillRule::NonZero,
                })
                .collect();
            let text_rect = evaluated.text_rectangle.map(|rectangle| Rect {
                x: rectangle.left,
                y: rectangle.top,
                width: rectangle.right - rectangle.left,
                height: rectangle.bottom - rectangle.top,
            });
            ResolvedGeometry::Custom { paths, text_rect }
        }))
    }

    fn resolve_text_body(
        &self,
        shape: &CT_Shape,
        body: &CT_TextBody,
        source: FlattenedSource,
        hyperlinks: Option<&ScopedHyperlinkTargets>,
        diagnostics: &mut Vec<Diagnostic>,
    ) -> Result<(ResolvedTextBody, Vec<oxml_layout::TextDirection>), ResolveError> {
        let properties = self.effective_body_pr(shape);
        let slide_number_placeholder = self.is_slide_number_placeholder(shape);
        let resolved = body
            .paragraphs()
            .iter()
            .map(|paragraph| {
                self.resolve_paragraph(
                    shape,
                    paragraph,
                    source,
                    hyperlinks,
                    slide_number_placeholder,
                    diagnostics,
                )
            })
            .collect::<Result<Vec<_>, _>>()?;
        let (paragraphs, directions) = resolved.into_iter().unzip();
        Ok((resolved_text_body(&properties, paragraphs)?, directions))
    }

    fn resolve_paragraph(
        &self,
        shape: &CT_Shape,
        paragraph: &oxml_drawing::text::CT_TextParagraph,
        source: FlattenedSource,
        hyperlinks: Option<&ScopedHyperlinkTargets>,
        slide_number_placeholder: bool,
        diagnostics: &mut Vec<Diagnostic>,
    ) -> Result<(ResolvedParagraph, oxml_layout::TextDirection), ResolveError> {
        let effective = self.effective_text_properties(shape, paragraph.properties.as_ref(), None);
        let paragraph_properties = &effective.paragraph;
        let mut runs = Vec::new();
        for run in &paragraph.runs {
            match run {
                TextRun::Run(run) => {
                    let mut style = self.resolve_run_style(
                        shape,
                        paragraph.properties.as_ref(),
                        run.properties.as_ref(),
                    )?;
                    style.hyperlink_url = resolve_direct_hyperlink(
                        run.properties.as_ref(),
                        source,
                        hyperlinks,
                        diagnostics,
                    );
                    runs.push(ResolvedTextRun::Text {
                        text: run.text.value.clone(),
                        style,
                    });
                }
                TextRun::Break(_) => runs.push(ResolvedTextRun::Break),
                TextRun::Field(field) => {
                    let mut style = self.resolve_run_style(
                        shape,
                        paragraph.properties.as_ref(),
                        field.run_properties.as_ref(),
                    )?;
                    style.hyperlink_url = resolve_direct_hyperlink(
                        field.run_properties.as_ref(),
                        source,
                        hyperlinks,
                        diagnostics,
                    );
                    runs.push(ResolvedTextRun::Field {
                        text: field
                            .text
                            .as_ref()
                            .map(|text| text.value.clone())
                            .unwrap_or_default(),
                        field_type: field
                            .field_type
                            .clone()
                            .or_else(|| slide_number_placeholder.then(|| "slidenum".to_owned())),
                        style,
                    });
                }
            }
        }
        let direction = paragraph_base_direction(paragraph_properties.right_to_left);
        Ok((
            ResolvedParagraph {
                level: paragraph_properties.level.unwrap_or(0),
                left_margin: emu_to_points(i64::from(
                    paragraph_properties.left_margin.unwrap_or(0),
                )),
                right_margin: emu_to_points(i64::from(
                    paragraph_properties.right_margin.unwrap_or(0),
                )),
                indent: emu_to_points(i64::from(paragraph_properties.indent.unwrap_or(0))),
                alignment: paragraph_alignment(paragraph_properties.alignment),
                line_spacing: resolved_text_spacing(paragraph_properties.line_spacing.as_ref()),
                space_before: resolved_text_spacing(paragraph_properties.space_before.as_ref()),
                space_after: resolved_text_spacing(paragraph_properties.space_after.as_ref()),
                bullet: paragraph_properties
                    .bullet
                    .as_ref()
                    .and_then(|bullet| self.resolve_bullet(bullet).transpose())
                    .transpose()?,
                end_style: self.resolve_run_style(
                    shape,
                    paragraph.properties.as_ref(),
                    paragraph.end_properties.as_ref(),
                )?,
                runs,
            },
            direction,
        ))
    }

    fn resolve_run_style(
        &self,
        shape: &CT_Shape,
        paragraph: Option<&oxml_drawing::text::CT_TextParagraphProperties>,
        run: Option<&CT_TextCharacterProperties>,
    ) -> Result<ResolvedRunStyle, ResolveError> {
        let effective = self.effective_text_properties(shape, paragraph, run).run;
        self.resolved_run_style(effective)
    }

    fn resolve_table_run_style(
        &self,
        table_style: &CT_TextCharacterProperties,
        paragraph: Option<&oxml_drawing::text::CT_TextParagraphProperties>,
        run: Option<&CT_TextCharacterProperties>,
    ) -> Result<ResolvedRunStyle, ResolveError> {
        let effective = self
            .effective_table_text_properties(Some(table_style), paragraph, run)
            .run;
        self.resolved_run_style(effective)
    }

    fn resolved_run_style(
        &self,
        effective: CT_TextCharacterProperties,
    ) -> Result<ResolvedRunStyle, ResolveError> {
        let fill = effective
            .fill
            .as_ref()
            .map(|fill| self.concrete_fill(fill, (1.0, 1.0)))
            .transpose()?
            .and_then(|(paint, _)| paint);
        Ok(ResolvedRunStyle {
            font_size: effective.font_size.map(|size| f64::from(size) / 100.0),
            bold: effective.bold.unwrap_or(false),
            italic: effective.italic.unwrap_or(false),
            all_caps: effective.all_caps.unwrap_or(false),
            underline: effective
                .underline
                .is_some_and(|value| value != oxml_drawing::text::TextUnderline::None),
            strike: effective
                .strike
                .is_some_and(|value| value != oxml_drawing::text::TextStrike::None),
            spacing: effective.spacing.as_ref().and_then(text_point_value),
            baseline: effective.baseline.as_deref().and_then(parse_percent),
            fill,
            latin_typeface: effective
                .latin
                .as_ref()
                .map(|font| self.resolve_typeface(&font.typeface, None)),
            east_asian_typeface: effective
                .east_asian
                .as_ref()
                .map(|font| self.resolve_typeface(&font.typeface, None)),
            complex_script_typeface: effective
                .complex_script
                .as_ref()
                .map(|font| self.resolve_typeface(&font.typeface, None)),
            symbol_typeface: effective
                .symbol
                .as_ref()
                .map(|font| self.resolve_typeface(&font.typeface, None)),
            hyperlink_url: None,
        })
    }

    fn resolve_bullet(
        &self,
        bullet: &oxml_drawing::text::TextBullet,
    ) -> Result<Option<ResolvedBullet>, ResolveError> {
        let color = bullet
            .color
            .as_ref()
            .map(|color| self.concrete_color(&color.color))
            .transpose()?;
        let font = bullet
            .font
            .as_ref()
            .map(|font| self.resolve_typeface(&font.typeface, None));
        let size = bullet.size.as_ref().and_then(|size| match &size.value {
            TextBulletSizeValue::Percent(value) => {
                parse_percent(value).map(ResolvedBulletSize::Percent)
            }
            TextBulletSizeValue::Points(value) => {
                Some(ResolvedBulletSize::Points(f64::from(*value) / 100.0))
            }
        });
        Ok(match bullet.choice.as_ref() {
            Some(TextBulletChoice::Character(character)) => Some(ResolvedBullet::Character {
                character: character.character.clone(),
                font,
                color,
                size,
            }),
            Some(TextBulletChoice::AutoNumber(number)) => Some(ResolvedBullet::AutoNumber {
                scheme: number.scheme.as_str().to_owned(),
                start_at: u32::from(number.start_at.unwrap_or(1)),
                font,
                color,
                size,
            }),
            Some(TextBulletChoice::None(_)) | None => None,
        })
    }

    fn resolve_table(
        &self,
        table: &oxml_drawing::table::CT_Table,
        source: FlattenedSource,
        hyperlinks: Option<&ScopedHyperlinkTargets>,
        diagnostics: &mut Vec<Diagnostic>,
        text_directions: &mut Vec<Vec<oxml_layout::TextDirection>>,
    ) -> Result<ResolvedTable, ResolveError> {
        let column_widths = table
            .grid
            .columns
            .iter()
            .map(|width| emu_to_points(width.0))
            .collect();
        let style = self.table_styles.and_then(|styles| {
            styles.style(
                table
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.style_id.as_deref()),
            )
        });
        let row_count = table.rows.len();
        let column_count = table.grid.columns.len();
        let table_properties = table.properties.as_ref();
        let rows = table
            .rows
            .iter()
            .enumerate()
            .map(|(row_index, row)| {
                let cells = row
                    .cells
                    .iter()
                    .enumerate()
                    .map(|(column_index, cell)| {
                        let mut cascade = TableCellCascade::default();
                        let position = TableCellPosition {
                            row: row_index,
                            column: column_index,
                            row_count,
                            column_count,
                        };
                        if let Some(style) = style {
                            let mut priority = 1u8;
                            cascade.apply_region(
                                style.whole_table.as_ref(),
                                true,
                                position,
                                priority,
                            );
                            priority += 1;
                            if table_properties.is_some_and(|properties| properties.band_rows) {
                                let offset = usize::from(
                                    table_properties.is_some_and(|properties| properties.first_row),
                                );
                                let region = if row_index.saturating_sub(offset) % 2 == 0 {
                                    style.band1_horizontal.as_ref()
                                } else {
                                    style.band2_horizontal.as_ref()
                                };
                                cascade.apply_region(region, false, position, priority);
                                priority += 1;
                            }
                            if table_properties.is_some_and(|properties| properties.band_columns) {
                                let offset = usize::from(
                                    table_properties
                                        .is_some_and(|properties| properties.first_column),
                                );
                                let region = if column_index.saturating_sub(offset) % 2 == 0 {
                                    style.band1_vertical.as_ref()
                                } else {
                                    style.band2_vertical.as_ref()
                                };
                                cascade.apply_region(region, false, position, priority);
                                priority += 1;
                            }
                            if column_index == 0
                                && table_properties
                                    .is_some_and(|properties| properties.first_column)
                            {
                                cascade.apply_region(
                                    style.first_column.as_ref(),
                                    false,
                                    position,
                                    priority,
                                );
                                priority += 1;
                            }
                            if column_index + 1 == column_count
                                && table_properties.is_some_and(|properties| properties.last_column)
                            {
                                cascade.apply_region(
                                    style.last_column.as_ref(),
                                    false,
                                    position,
                                    priority,
                                );
                                priority += 1;
                            }
                            if row_index == 0
                                && table_properties.is_some_and(|properties| properties.first_row)
                            {
                                cascade.apply_region(
                                    style.first_row.as_ref(),
                                    false,
                                    position,
                                    priority,
                                );
                                priority += 1;
                            }
                            if row_index + 1 == row_count
                                && table_properties.is_some_and(|properties| properties.last_row)
                            {
                                cascade.apply_region(
                                    style.last_row.as_ref(),
                                    false,
                                    position,
                                    priority,
                                );
                                priority += 1;
                            }
                            let first_row =
                                table_properties.is_some_and(|properties| properties.first_row);
                            let last_row =
                                table_properties.is_some_and(|properties| properties.last_row);
                            let first_column =
                                table_properties.is_some_and(|properties| properties.first_column);
                            let last_column =
                                table_properties.is_some_and(|properties| properties.last_column);
                            let corner = match (row_index, column_index) {
                                (0, 0) if first_row && first_column => {
                                    style.north_west_cell.as_ref()
                                }
                                (0, column)
                                    if column + 1 == column_count && first_row && last_column =>
                                {
                                    style.north_east_cell.as_ref()
                                }
                                (row, 0) if row + 1 == row_count && last_row && first_column => {
                                    style.south_west_cell.as_ref()
                                }
                                (row, column)
                                    if row + 1 == row_count
                                        && column + 1 == column_count
                                        && last_row
                                        && last_column =>
                                {
                                    style.south_east_cell.as_ref()
                                }
                                _ => None,
                            };
                            cascade.apply_region(corner, false, position, priority);
                        }
                        if let Some(properties) = &cell.properties {
                            cascade.apply_direct(properties);
                            for unsupported in &properties.unsupported {
                                push_table_diagnostic(diagnostics, unsupported);
                            }
                        }
                        for unsupported in &cascade.unsupported {
                            push_table_diagnostic(diagnostics, unsupported);
                        }
                        let referenced_fill = cascade
                            .fill_reference
                            .as_ref()
                            .map(|reference| self.table_referenced_fill(reference))
                            .transpose()?
                            .flatten();
                        let (fill, fill_unsupported) = referenced_fill
                            .as_ref()
                            .or(cascade.fill.as_ref())
                            .map(|fill| self.concrete_fill(fill, (1.0, 1.0)))
                            .transpose()?
                            .unwrap_or((None, None));
                        if let Some(unsupported) = fill_unsupported {
                            push_table_diagnostic(diagnostics, unsupported);
                        }
                        let table_text_style = table_character_properties(&cascade.text_style)?;
                        let mut text = if let Some(body) = cell.text_body.as_ref() {
                            let (body, directions) = resolve_standalone_text_body(
                                self,
                                body,
                                &table_text_style,
                                source,
                                hyperlinks,
                                diagnostics,
                            )?;
                            text_directions.push(directions);
                            Some(body)
                        } else {
                            text_directions.push(Vec::new());
                            None
                        };
                        if let Some(body) = text.as_mut()
                            && body.autofit != ResolvedAutofit::None
                        {
                            body.autofit = ResolvedAutofit::None;
                            let message = "table cell autofit is unsupported and was ignored";
                            if !diagnostics
                                .iter()
                                .any(|diagnostic| diagnostic.message == message)
                            {
                                diagnostics.push(Diagnostic {
                                    message: message.to_owned(),
                                });
                            }
                        }
                        Ok(ResolvedTableCell {
                            text,
                            fill,
                            margins: TextInsets {
                                left: emu_to_points(cascade.margin_left.unwrap_or(91_440)),
                                top: emu_to_points(cascade.margin_top.unwrap_or(45_720)),
                                right: emu_to_points(cascade.margin_right.unwrap_or(91_440)),
                                bottom: emu_to_points(cascade.margin_bottom.unwrap_or(45_720)),
                            },
                            left: self.resolve_table_border(cascade.left, diagnostics)?,
                            right: self.resolve_table_border(cascade.right, diagnostics)?,
                            top: self.resolve_table_border(cascade.top, diagnostics)?,
                            bottom: self.resolve_table_border(cascade.bottom, diagnostics)?,
                            row_span: cell.row_span,
                            grid_span: cell.grid_span,
                            horizontal_merge: cell.horizontal_merge,
                            vertical_merge: cell.vertical_merge,
                        })
                    })
                    .collect::<Result<Vec<_>, ResolveError>>()?;
                Ok(ResolvedTableRow {
                    height: emu_to_points(row.height.0),
                    cells,
                })
            })
            .collect::<Result<Vec<_>, ResolveError>>()?;
        Ok(ResolvedTable {
            right_to_left: table_properties.is_some_and(|properties| properties.right_to_left),
            column_widths,
            rows,
        })
    }

    fn resolve_table_border(
        &self,
        border: Option<(CT_LineProperties, u8)>,
        diagnostics: &mut Vec<Diagnostic>,
    ) -> Result<Option<ResolvedTableBorder>, ResolveError> {
        border
            .map(|(line, priority)| {
                let stroke = self.concrete_line(&line, (1.0, 1.0))?;
                if stroke.is_none()
                    && let Some(fill) = &line.fill
                {
                    let (_, unsupported) = self.concrete_fill(fill, (1.0, 1.0))?;
                    if let Some(unsupported) = unsupported {
                        push_table_diagnostic(diagnostics, unsupported);
                    }
                }
                Ok(ResolvedTableBorder { stroke, priority })
            })
            .transpose()
    }

    fn table_referenced_fill(
        &self,
        reference: &StyleReference,
    ) -> Result<Option<Fill>, ResolveError> {
        let StyleReference::Fill(reference) = reference else {
            return Ok(None);
        };
        let matrix = &self.theme.theme_elements.format_scheme;
        let Some(mut fill) = referenced_fill(
            reference.index,
            &matrix.fill_styles,
            &matrix.background_fill_styles,
        )?
        else {
            return Ok(None);
        };
        substitute_fill(&mut fill, reference.color.as_ref(), "table cell")?;
        Ok(Some(fill))
    }
}

#[derive(Default)]
struct TableCellCascade {
    fill: Option<Fill>,
    fill_reference: Option<StyleReference>,
    left: Option<(CT_LineProperties, u8)>,
    right: Option<(CT_LineProperties, u8)>,
    top: Option<(CT_LineProperties, u8)>,
    bottom: Option<(CT_LineProperties, u8)>,
    margin_left: Option<i64>,
    margin_right: Option<i64>,
    margin_top: Option<i64>,
    margin_bottom: Option<i64>,
    text_style: TableTextCascade,
    unsupported: Vec<String>,
}

#[derive(Default)]
struct TableTextCascade {
    bold: Option<bool>,
    italic: Option<bool>,
    font_collection: Option<FontCollectionIndex>,
    color: Option<ColorChoice>,
}

#[derive(Clone, Copy)]
struct TableCellPosition {
    row: usize,
    column: usize,
    row_count: usize,
    column_count: usize,
}

fn table_character_properties(
    style: &TableTextCascade,
) -> Result<CT_TextCharacterProperties, ResolveError> {
    let mut properties = CT_TextCharacterProperties::default();
    properties.bold = style.bold;
    properties.italic = style.italic;
    if let Some(color) = &style.color {
        let mut solid = SolidFill::default();
        solid.color = Some(color.clone());
        properties.fill = Some(Fill::Solid(solid));
    }
    let font_tokens = match style.font_collection {
        Some(FontCollectionIndex::Major) => Some(("+mj-lt", "+mj-ea", "+mj-cs")),
        Some(FontCollectionIndex::Minor) => Some(("+mn-lt", "+mn-ea", "+mn-cs")),
        Some(FontCollectionIndex::None) | None => None,
    };
    if let Some((latin, east_asian, complex_script)) = font_tokens {
        let font = |typeface| {
            TextFont::new(typeface).map_err(|error| ResolveError::ConcreteValue {
                kind: "table font",
                detail: error.to_string(),
            })
        };
        properties.latin = Some(font(latin)?);
        properties.east_asian = Some(font(east_asian)?);
        properties.complex_script = Some(font(complex_script)?);
    }
    Ok(properties)
}

impl TableCellCascade {
    fn apply_region(
        &mut self,
        region: Option<&CT_TablePartStyle>,
        whole_table: bool,
        position: TableCellPosition,
        priority: u8,
    ) {
        let Some(region) = region else {
            return;
        };
        if let Some(cell_style) = &region.cell_style {
            self.apply_cell_style(cell_style, whole_table, position, priority);
        }
        if let Some(text_style) = &region.text_style {
            self.apply_text_style(text_style);
        }
    }

    fn apply_cell_style(
        &mut self,
        style: &CT_TableCellStyle,
        whole_table: bool,
        position: TableCellPosition,
        priority: u8,
    ) {
        if let Some(fill) = &style.fill {
            self.fill = Some(fill.clone());
            self.fill_reference = None;
        }
        if let Some(reference) = &style.fill_reference {
            self.fill_reference = Some(reference.clone());
            self.fill = None;
        }
        if let Some(borders) = &style.borders {
            self.apply_borders(borders, whole_table, position, priority);
            self.unsupported.extend(borders.unsupported.iter().cloned());
        }
        self.unsupported.extend(style.unsupported.iter().cloned());
    }

    fn apply_borders(
        &mut self,
        borders: &CT_TableBorders,
        whole_table: bool,
        position: TableCellPosition,
        priority: u8,
    ) {
        let left = if whole_table {
            if position.column > 0 {
                borders.inside_vertical.as_ref()
            } else {
                borders.left.as_ref()
            }
        } else {
            borders.left.as_ref().or(borders.inside_vertical.as_ref())
        };
        let right = if whole_table {
            if position.column + 1 < position.column_count {
                borders.inside_vertical.as_ref()
            } else {
                borders.right.as_ref()
            }
        } else {
            borders.right.as_ref().or(borders.inside_vertical.as_ref())
        };
        let top = if whole_table {
            if position.row > 0 {
                borders.inside_horizontal.as_ref()
            } else {
                borders.top.as_ref()
            }
        } else {
            borders.top.as_ref().or(borders.inside_horizontal.as_ref())
        };
        let bottom = if whole_table {
            if position.row + 1 < position.row_count {
                borders.inside_horizontal.as_ref()
            } else {
                borders.bottom.as_ref()
            }
        } else {
            borders
                .bottom
                .as_ref()
                .or(borders.inside_horizontal.as_ref())
        };
        apply_table_edge(&mut self.left, left, priority);
        apply_table_edge(&mut self.right, right, priority);
        apply_table_edge(&mut self.top, top, priority);
        apply_table_edge(&mut self.bottom, bottom, priority);
    }

    fn apply_text_style(&mut self, style: &CT_TableTextStyle) {
        if style.bold.is_some() {
            self.text_style.bold = style.bold;
        }
        if style.italic.is_some() {
            self.text_style.italic = style.italic;
        }
        if let Some(StyleReference::Font(reference)) = &style.font_reference {
            self.text_style.font_collection = Some(reference.index);
            if reference.color.is_some() {
                self.text_style.color = reference.color.clone();
            }
        }
        if style.color.is_some() {
            self.text_style.color = style.color.clone();
        }
    }

    fn apply_direct(&mut self, properties: &oxml_drawing::table::CT_TableCellProperties) {
        if let Some(fill) = &properties.fill {
            self.fill = Some(fill.clone());
            self.fill_reference = None;
        }
        if let Some(value) = properties.margin_left {
            self.margin_left = Some(value.0);
        }
        if let Some(value) = properties.margin_right {
            self.margin_right = Some(value.0);
        }
        if let Some(value) = properties.margin_top {
            self.margin_top = Some(value.0);
        }
        if let Some(value) = properties.margin_bottom {
            self.margin_bottom = Some(value.0);
        }
        apply_table_edge(&mut self.left, properties.left.as_ref(), u8::MAX);
        apply_table_edge(&mut self.right, properties.right.as_ref(), u8::MAX);
        apply_table_edge(&mut self.top, properties.top.as_ref(), u8::MAX);
        apply_table_edge(&mut self.bottom, properties.bottom.as_ref(), u8::MAX);
    }
}

fn apply_table_edge(
    target: &mut Option<(CT_LineProperties, u8)>,
    source: Option<&CT_LineProperties>,
    priority: u8,
) {
    if let Some(source) = source {
        *target = Some((source.clone(), priority));
    }
}

fn push_table_diagnostic(diagnostics: &mut Vec<Diagnostic>, unsupported: &str) {
    let message = format!("unsupported table cell {unsupported} was ignored");
    if !diagnostics
        .iter()
        .any(|diagnostic| diagnostic.message == message)
    {
        diagnostics.push(Diagnostic { message });
    }
}

fn push_group_diagnostics(issues: u8, diagnostics: &mut Vec<Diagnostic>) {
    let messages = [
        (
            GROUP_ZERO_X,
            "zero horizontal group child extent used finite unit scale",
        ),
        (
            GROUP_ZERO_Y,
            "zero vertical group child extent used finite unit scale",
        ),
        (
            GROUP_SHEAR,
            "unsupported sheared or singular group transform retained as affine fallback",
        ),
    ];
    for (flag, message) in messages {
        if issues & flag != 0
            && !diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message == message)
        {
            diagnostics.push(Diagnostic {
                message: message.to_owned(),
            });
        }
    }
}

fn scaled_group_bounds(bounds: Rect, scale: (f64, f64)) -> Rect {
    Rect {
        x: bounds.x * scale.0,
        y: bounds.y * scale.1,
        width: bounds.width * scale.0,
        height: bounds.height * scale.1,
    }
}

fn transform_values(transform: Option<&CT_Transform2D>) -> Option<(Rect, f64, bool, bool)> {
    transform_values_with_degenerate_line(transform, false)
}

fn connector_transform_values(
    transform: Option<&CT_Transform2D>,
) -> Option<(Rect, f64, bool, bool)> {
    transform_values_with_degenerate_line(transform, true)
}

fn transform_values_with_degenerate_line(
    transform: Option<&CT_Transform2D>,
    allow_degenerate_line: bool,
) -> Option<(Rect, f64, bool, bool)> {
    let transform = transform?;
    let extent = transform.extent?;
    let invalid = if allow_degenerate_line {
        extent.cx.0 < 0 || extent.cy.0 < 0 || (extent.cx.0 == 0 && extent.cy.0 == 0)
    } else {
        extent.cx.0 <= 0 || extent.cy.0 <= 0
    };
    if invalid {
        return None;
    }
    let offset = transform.offset.unwrap_or_default();
    Some((
        Rect {
            x: emu_to_points(offset.x.0),
            y: emu_to_points(offset.y.0),
            width: emu_to_points(extent.cx.0),
            height: emu_to_points(extent.cy.0),
        },
        f64::from(transform.rotation.0) / 60_000.0,
        transform.flip_horizontal,
        transform.flip_vertical,
    ))
}

fn emu_to_points(value: i64) -> f64 {
    value as f64 / 12_700.0
}

fn base_rgb(choice: &ColorChoice) -> Option<RgbColor> {
    let resolved = resolve_color(choice, &ColorMap::default(), &[]).ok()?;
    Some(RgbColor::new(resolved.red, resolved.green, resolved.blue))
}

fn concrete_path_command(command: EvaluatedPathCommand, scale_x: f64, scale_y: f64) -> PathCommand {
    match command {
        EvaluatedPathCommand::MoveTo { x, y } => PathCommand::MoveTo(Point {
            x: x * scale_x,
            y: y * scale_y,
        }),
        EvaluatedPathCommand::LineTo { x, y } => PathCommand::LineTo(Point {
            x: x * scale_x,
            y: y * scale_y,
        }),
        EvaluatedPathCommand::CubicTo {
            x1,
            y1,
            x2,
            y2,
            x,
            y,
        } => PathCommand::CurveTo {
            c1: Point {
                x: x1 * scale_x,
                y: y1 * scale_y,
            },
            c2: Point {
                x: x2 * scale_x,
                y: y2 * scale_y,
            },
            to: Point {
                x: x * scale_x,
                y: y * scale_y,
            },
        },
        EvaluatedPathCommand::Close => PathCommand::Close,
    }
}

fn coordinate_points(value: Option<&Coordinate32Value>) -> Result<f64, ResolveError> {
    match value {
        Some(Coordinate32Value::Emu(value)) => Ok(emu_to_points(i64::from(*value))),
        Some(Coordinate32Value::UniversalMeasure(value)) => universal_measure_points(value),
        None => Ok(0.0),
    }
}

fn universal_measure_points(value: &str) -> Result<f64, ResolveError> {
    let split = value
        .find(|character: char| character.is_ascii_alphabetic())
        .unwrap_or(value.len());
    let (number, unit) = value.split_at(split);
    let number = number
        .parse::<f64>()
        .map_err(|error| ResolveError::ConcreteValue {
            kind: "universal measure",
            detail: error.to_string(),
        })?;
    match unit {
        "pt" => Ok(number),
        "in" => Ok(number * 72.0),
        "cm" => Ok(number * 72.0 / 2.54),
        "mm" => Ok(number * 72.0 / 25.4),
        _ => Err(ResolveError::ConcreteValue {
            kind: "universal measure",
            detail: format!("unsupported unit {unit}"),
        }),
    }
}

fn parse_percent(value: &str) -> Option<f64> {
    value.parse::<f64>().ok().map(|value| value / 100_000.0)
}

fn text_point_value(value: &TextPointValue) -> Option<f64> {
    match value {
        TextPointValue::Centipoints(value) => Some(f64::from(*value) / 100.0),
        TextPointValue::UniversalMeasure(value) => universal_measure_points(value).ok(),
    }
}

fn resolved_text_spacing(value: Option<&TextSpacing>) -> Option<ResolvedTextSpacing> {
    match value? {
        TextSpacing::Percent(value) => parse_percent(value).map(ResolvedTextSpacing::Percent),
        TextSpacing::Points(value) => Some(ResolvedTextSpacing::Points(f64::from(*value) / 100.0)),
    }
}

fn paragraph_alignment(alignment: Option<TextAlignment>) -> ParagraphAlignment {
    match alignment.unwrap_or(TextAlignment::Left) {
        TextAlignment::Left => ParagraphAlignment::Left,
        TextAlignment::Center => ParagraphAlignment::Center,
        TextAlignment::Right => ParagraphAlignment::Right,
        TextAlignment::Justified | TextAlignment::JustifiedLow => ParagraphAlignment::Justified,
        TextAlignment::Distributed | TextAlignment::ThaiDistributed => {
            ParagraphAlignment::Distributed
        }
    }
}

fn paragraph_base_direction(right_to_left: Option<bool>) -> oxml_layout::TextDirection {
    match right_to_left {
        Some(true) => oxml_layout::TextDirection::RightToLeft,
        Some(false) => oxml_layout::TextDirection::LeftToRight,
        None => oxml_layout::TextDirection::Auto,
    }
}

fn vertical_text_diagnostic(direction: TextDirection) -> Option<&'static str> {
    match direction {
        TextDirection::EastAsianVertical => {
            Some("east Asian vertical text rendered as rotated vertical text")
        }
        TextDirection::MongolianVertical => {
            Some("Mongolian vertical text rendered as rotated vertical-270 text")
        }
        TextDirection::WordArtVertical => {
            Some("WordArt vertical text rendered as rotated vertical text")
        }
        TextDirection::WordArtVerticalRtl => {
            Some("right-to-left WordArt vertical text rendered as rotated vertical-270 text")
        }
        TextDirection::Horizontal | TextDirection::Vertical | TextDirection::Vertical270 => None,
    }
}

fn resolve_standalone_text_body(
    context: &ResolveCtx<'_>,
    body: &CT_TextBody,
    table_style: &CT_TextCharacterProperties,
    source: FlattenedSource,
    hyperlinks: Option<&ScopedHyperlinkTargets>,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<(ResolvedTextBody, Vec<oxml_layout::TextDirection>), ResolveError> {
    let mut properties = default_body_properties();
    merge_body_properties(&mut properties, &body.body_properties);
    let resolved = body
        .paragraphs()
        .iter()
        .map(|paragraph| {
            let effective = context.effective_table_text_properties(
                Some(table_style),
                paragraph.properties.as_ref(),
                None,
            );
            let paragraph_properties = &effective.paragraph;
            let runs = paragraph
                .runs
                .iter()
                .map(|run| match run {
                    TextRun::Run(run) => {
                        let mut style = context.resolve_table_run_style(
                            table_style,
                            paragraph.properties.as_ref(),
                            run.properties.as_ref(),
                        )?;
                        style.hyperlink_url = resolve_direct_hyperlink(
                            run.properties.as_ref(),
                            source,
                            hyperlinks,
                            diagnostics,
                        );
                        Ok(ResolvedTextRun::Text {
                            text: run.text.value.clone(),
                            style,
                        })
                    }
                    TextRun::Break(_) => Ok(ResolvedTextRun::Break),
                    TextRun::Field(field) => {
                        let mut style = context.resolve_table_run_style(
                            table_style,
                            paragraph.properties.as_ref(),
                            field.run_properties.as_ref(),
                        )?;
                        style.hyperlink_url = resolve_direct_hyperlink(
                            field.run_properties.as_ref(),
                            source,
                            hyperlinks,
                            diagnostics,
                        );
                        Ok(ResolvedTextRun::Field {
                            text: field
                                .text
                                .as_ref()
                                .map(|text| text.value.clone())
                                .unwrap_or_default(),
                            field_type: field.field_type.clone(),
                            style,
                        })
                    }
                })
                .collect::<Result<Vec<_>, ResolveError>>()?;
            let direction = paragraph_base_direction(paragraph_properties.right_to_left);
            Ok((
                ResolvedParagraph {
                    level: paragraph_properties.level.unwrap_or(0),
                    left_margin: emu_to_points(i64::from(
                        paragraph_properties.left_margin.unwrap_or(0),
                    )),
                    right_margin: emu_to_points(i64::from(
                        paragraph_properties.right_margin.unwrap_or(0),
                    )),
                    indent: emu_to_points(i64::from(paragraph_properties.indent.unwrap_or(0))),
                    alignment: paragraph_alignment(paragraph_properties.alignment),
                    line_spacing: resolved_text_spacing(paragraph_properties.line_spacing.as_ref()),
                    space_before: resolved_text_spacing(paragraph_properties.space_before.as_ref()),
                    space_after: resolved_text_spacing(paragraph_properties.space_after.as_ref()),
                    bullet: paragraph_properties
                        .bullet
                        .as_ref()
                        .and_then(|bullet| context.resolve_bullet(bullet).transpose())
                        .transpose()?,
                    end_style: context.resolve_table_run_style(
                        table_style,
                        paragraph.properties.as_ref(),
                        paragraph.end_properties.as_ref(),
                    )?,
                    runs,
                },
                direction,
            ))
        })
        .collect::<Result<Vec<_>, ResolveError>>()?;
    let (paragraphs, directions) = resolved.into_iter().unzip();
    Ok((resolved_text_body(&properties, paragraphs)?, directions))
}

fn resolved_text_body(
    properties: &CT_TextBodyProperties,
    paragraphs: Vec<ResolvedParagraph>,
) -> Result<ResolvedTextBody, ResolveError> {
    let insets = TextInsets {
        left: coordinate_points(properties.left_inset.as_ref())?,
        top: coordinate_points(properties.top_inset.as_ref())?,
        right: coordinate_points(properties.right_inset.as_ref())?,
        bottom: coordinate_points(properties.bottom_inset.as_ref())?,
    };
    let anchor = match properties.anchor.unwrap_or(DrawingTextAnchor::Top) {
        DrawingTextAnchor::Top => TextAnchor::Top,
        DrawingTextAnchor::Center => TextAnchor::Center,
        DrawingTextAnchor::Bottom => TextAnchor::Bottom,
        DrawingTextAnchor::Justified => TextAnchor::Justified,
        DrawingTextAnchor::Distributed => TextAnchor::Distributed,
    };
    let vertical = match properties
        .vertical
        .unwrap_or(DrawingTextVertical::Horizontal)
    {
        DrawingTextVertical::Horizontal => TextDirection::Horizontal,
        DrawingTextVertical::Vertical => TextDirection::Vertical,
        DrawingTextVertical::Vertical270 => TextDirection::Vertical270,
        DrawingTextVertical::EastAsianVertical => TextDirection::EastAsianVertical,
        DrawingTextVertical::MongolianVertical => TextDirection::MongolianVertical,
        DrawingTextVertical::WordArtVertical => TextDirection::WordArtVertical,
        DrawingTextVertical::WordArtVerticalRtl => TextDirection::WordArtVerticalRtl,
    };
    let autofit = match properties.autofit.clone().unwrap_or(TextAutofit::NoAutofit) {
        TextAutofit::NoAutofit => ResolvedAutofit::None,
        TextAutofit::ShapeAutofit => ResolvedAutofit::Shape,
        TextAutofit::Normal(normal) => ResolvedAutofit::Normal {
            font_scale: normal.font_scale.as_deref().and_then(parse_percent),
            line_spacing_reduction: normal
                .line_spacing_reduction
                .as_deref()
                .and_then(parse_percent),
        },
    };
    Ok(ResolvedTextBody {
        insets,
        anchor,
        wrap: properties.wrap.unwrap_or(TextWrap::Square) != TextWrap::None,
        vertical,
        space_first_last_paragraph: properties.space_first_last_paragraph.unwrap_or(false),
        autofit,
        paragraphs,
    })
}

#[derive(Clone, Copy)]
struct LatentPolicy {
    date_time: bool,
    footer: bool,
    slide_number: bool,
}

impl LatentPolicy {
    fn from_context(context: &ResolveCtx<'_>) -> Self {
        let layout = context.layout.header_footer.as_ref();
        let master = context.master.header_footer.as_ref();
        Self {
            date_time: layout.is_none_or(|hf| hf.date_time_enabled())
                && master.is_none_or(|hf| hf.date_time_enabled()),
            footer: layout.is_none_or(|hf| hf.footer_enabled())
                && master.is_none_or(|hf| hf.footer_enabled()),
            slide_number: layout.is_none_or(|hf| hf.slide_number_enabled())
                && master.is_none_or(|hf| hf.slide_number_enabled()),
        }
    }

    fn permits(self, ph_type: &PhType) -> bool {
        match ph_type {
            PhType::DateTime => self.date_time,
            PhType::Footer => self.footer,
            PhType::SlideNumber => self.slide_number,
            _ => true,
        }
    }
}

#[derive(Clone, Copy)]
struct PassRules<'a> {
    source: FlattenedSource,
    emit_non_placeholders: bool,
    emit_inherited_latent: bool,
    deeper_latent: &'a [PlaceholderKey],
    latent_policy: LatentPolicy,
}

fn collect_occupied_latent(children: &[ShapeTreeChild], keys: &mut Vec<PlaceholderKey>) {
    for child in children {
        match child {
            ShapeTreeChild::GroupShape(group) => collect_occupied_latent(&group.children, keys),
            ShapeTreeChild::AlternateContent(alternate) => {
                if let Some(fallback) = alternate.selected_fallback() {
                    collect_occupied_latent(fallback, keys);
                }
            }
            ShapeTreeChild::Shape(shape) => {
                if let Some(placeholder) = shape.placeholder.as_ref()
                    && is_latent(&placeholder.effective_type())
                    && shape
                        .text_body
                        .as_ref()
                        .is_some_and(|body| !body.plain_text().trim().is_empty())
                {
                    push_unmatched(keys, placeholder.key());
                }
            }
            _ => {}
        }
    }
}

fn push_unmatched(keys: &mut Vec<PlaceholderKey>, key: PlaceholderKey) {
    if !keys
        .iter()
        .any(|existing| placeholder_keys_match(&key, existing))
    {
        keys.push(key);
    }
}

fn find_group_lineage(
    children: &[ShapeTreeChild],
    target: &ShapeTreeChild,
    lineage: &mut Vec<u32>,
    found: &mut Vec<u32>,
) -> bool {
    for child in children {
        if std::ptr::eq(child, target) {
            *found = lineage.clone();
            return true;
        }
        match child {
            ShapeTreeChild::GroupShape(group) => {
                if let Some(id) = child.non_visual_id() {
                    lineage.push(id);
                }
                if find_group_lineage(&group.children, target, lineage, found) {
                    return true;
                }
                if child.non_visual_id().is_some() {
                    lineage.pop();
                }
            }
            ShapeTreeChild::AlternateContent(alternate) => {
                if let Some(fallback) = alternate.selected_fallback()
                    && find_group_lineage(fallback, target, lineage, found)
                {
                    return true;
                }
            }
            _ => {}
        }
    }
    false
}

fn apply_timeline_state(
    mut slide: ResolvedSlide,
    identities: Vec<ResolvedShapeIdentity>,
    mut state: EvaluatedFrameState,
    mut group_geometries: BTreeMap<u32, TimelineGroupGeometry>,
) -> ResolvedTimelineSlide {
    let invalid_targets = state
        .shapes
        .iter()
        .filter_map(|(target, evaluated)| (!evaluated.is_finite()).then_some(*target))
        .collect::<Vec<_>>();
    for target in invalid_targets {
        state.shapes.insert(target, EvaluatedShapeState::default());
        state.diagnostics.push(Diagnostic {
            message: format!("non-finite timeline state ignored for target {target}"),
        });
    }
    let slide_size = slide.size;
    for (group_id, bounds) in timeline_group_bounds(&slide, &identities) {
        group_geometries
            .entry(group_id)
            .or_insert_with(|| TimelineGroupGeometry::from_page_bounds(bounds));
    }
    let shape_states = identities
        .iter()
        .zip(&mut slide.shapes)
        .map(|(identity, shape)| {
            if identity.source != FlattenedSource::Slide {
                return EvaluatedShapeState::default();
            }
            let mut evaluated = EvaluatedShapeState::default();
            let original_bounds = shape.bounds;
            let original_shape = shape.clone();
            let mut group_transform = Transform::IDENTITY;
            let mut group_page_clips = Vec::new();
            for group_id in &identity.containing_group_ids {
                let Some(target_state) = state.shapes.get(group_id) else {
                    continue;
                };
                evaluated.visible &= target_state.visible;
                evaluated.opacity *= target_state.opacity;
                let Some(geometry) = group_geometries.get(group_id).copied() else {
                    continue;
                };
                let target_transform =
                    target_state.resolved_page_transform(geometry.center(), slide_size);
                group_transform = target_transform.then(group_transform);
                if let Some(clip) = target_state.clip {
                    group_page_clips.push(group_clip_page_points(clip, geometry, group_transform));
                }
            }
            let direct_state = identity.shape_id.and_then(|id| state.shapes.get(&id));
            if let Some(direct_state) = direct_state {
                evaluated.visible &= direct_state.visible;
                evaluated.opacity *= direct_state.opacity;
                evaluated.clip = intersect_optional_rect(evaluated.clip, direct_state.clip);
            }
            let (scale_x, scale_y, rotation_deg) = direct_state.map_or(
                (1.0, 1.0, 0.0),
                EvaluatedShapeState::resolved_shape_geometry,
            );
            let original_page_bounds = resolved_shape_page_bounds(&original_shape);
            let original_page_center = (
                original_page_bounds.x + original_page_bounds.width / 2.0,
                original_page_bounds.y + original_page_bounds.height / 2.0,
            );
            let page_motion = direct_state.map_or(Transform::IDENTITY, |direct_state| {
                direct_state.resolved_motion_translation(original_page_center, slide_size)
            });
            evaluated.transform = page_motion.then(group_transform);
            let animated_bounds = Rect {
                x: original_bounds.x + original_bounds.width * (1.0 - scale_x) / 2.0,
                y: original_bounds.y + original_bounds.height * (1.0 - scale_y) / 2.0,
                width: original_bounds.width * scale_x,
                height: original_bounds.height * scale_y,
            };
            let animated_rotation = shape.rotation_deg + rotation_deg;
            let mut animated_shape = shape.clone();
            animated_shape.bounds = animated_bounds;
            animated_shape.rotation_deg = animated_rotation;
            animated_shape.group_transform = animated_shape
                .group_transform
                .then(page_motion)
                .then(group_transform);
            for page_clip in group_page_clips {
                match page_clip_for_shape(page_clip, &animated_shape) {
                    Some(clip) => evaluated.push_oriented_clip(clip),
                    None => state.diagnostics.push(Diagnostic {
                        message: format!(
                            "group timeline clip ignored for shape {}",
                            identity.shape_id.unwrap_or(0)
                        ),
                    }),
                }
            }
            if evaluated.is_finite()
                && [
                    animated_bounds.x,
                    animated_bounds.y,
                    animated_bounds.width,
                    animated_bounds.height,
                    animated_rotation,
                ]
                .into_iter()
                .all(f64::is_finite)
            {
                *shape = animated_shape;
            } else {
                state.diagnostics.push(Diagnostic {
                    message: format!(
                        "non-finite timeline state ignored for shape {}",
                        identity.shape_id.unwrap_or(0)
                    ),
                });
                evaluated = EvaluatedShapeState::default();
            }
            evaluated
        })
        .collect();
    ResolvedTimelineSlide {
        slide,
        identities,
        shape_states,
        state,
    }
}

#[derive(Clone, Copy)]
struct TimelineGroupGeometry {
    normalized_to_page: Transform,
}

impl TimelineGroupGeometry {
    fn from_page_bounds(bounds: Rect) -> Self {
        Self {
            normalized_to_page: Transform {
                a: bounds.width,
                d: bounds.height,
                e: bounds.x,
                f: bounds.y,
                ..Transform::IDENTITY
            },
        }
    }

    fn center(self) -> (f64, f64) {
        let center = self.normalized_to_page.apply(Point { x: 0.5, y: 0.5 });
        (center.x, center.y)
    }
}

fn timeline_group_geometries(children: &[ShapeTreeChild]) -> BTreeMap<u32, TimelineGroupGeometry> {
    let mut geometries = BTreeMap::new();
    collect_timeline_group_geometries(children, Transform::IDENTITY, &mut geometries);
    geometries
}

fn collect_timeline_group_geometries(
    children: &[ShapeTreeChild],
    parent_to_page: Transform,
    geometries: &mut BTreeMap<u32, TimelineGroupGeometry>,
) {
    for child in children {
        match child {
            ShapeTreeChild::GroupShape(group) => {
                let child_to_parent = group
                    .group_transform()
                    .and_then(group_affine)
                    .map_or(Transform::IDENTITY, |(transform, _)| transform);
                if let Some((group_id, geometry)) = child
                    .non_visual_id()
                    .zip(group.group_transform().and_then(group_extent_geometry))
                {
                    geometries.insert(
                        group_id,
                        TimelineGroupGeometry {
                            normalized_to_page: geometry.then(parent_to_page),
                        },
                    );
                }
                collect_timeline_group_geometries(
                    &group.children,
                    child_to_parent.then(parent_to_page),
                    geometries,
                );
            }
            ShapeTreeChild::AlternateContent(alternate) => {
                if let Some(fallback) = alternate.selected_fallback() {
                    collect_timeline_group_geometries(fallback, parent_to_page, geometries);
                }
            }
            _ => {}
        }
    }
}

fn group_extent_geometry(transform: &CT_Transform2D) -> Option<Transform> {
    let offset = transform.offset?;
    let extent = transform.extent?;
    let x = emu_to_points(offset.x.0);
    let y = emu_to_points(offset.y.0);
    let width = emu_to_points(extent.cx.0);
    let height = emu_to_points(extent.cy.0);
    let center_x = x + width / 2.0;
    let center_y = y + height / 2.0;
    let mut geometry = Transform {
        a: width,
        d: height,
        e: x,
        f: y,
        ..Transform::IDENTITY
    }
    .then(Transform::rotate_about(
        f64::from(transform.rotation.0) / 60_000.0,
        center_x,
        center_y,
    ));
    if transform.flip_horizontal || transform.flip_vertical {
        geometry = geometry.then(Transform {
            a: if transform.flip_horizontal { -1.0 } else { 1.0 },
            d: if transform.flip_vertical { -1.0 } else { 1.0 },
            e: if transform.flip_horizontal {
                2.0 * center_x
            } else {
                0.0
            },
            f: if transform.flip_vertical {
                2.0 * center_y
            } else {
                0.0
            },
            ..Transform::IDENTITY
        });
    }
    [
        geometry.a, geometry.b, geometry.c, geometry.d, geometry.e, geometry.f,
    ]
    .into_iter()
    .all(f64::is_finite)
    .then_some(geometry)
}

fn timeline_group_bounds(
    slide: &ResolvedSlide,
    identities: &[ResolvedShapeIdentity],
) -> BTreeMap<u32, Rect> {
    let mut bounds = BTreeMap::new();
    for (identity, shape) in identities.iter().zip(&slide.shapes) {
        if identity.source != FlattenedSource::Slide {
            continue;
        }
        let shape_bounds = resolved_shape_page_bounds(shape);
        for group_id in &identity.containing_group_ids {
            bounds
                .entry(*group_id)
                .and_modify(|group_bounds| *group_bounds = union_rect(*group_bounds, shape_bounds))
                .or_insert(shape_bounds);
        }
    }
    bounds
}

fn resolved_shape_page_bounds(shape: &ResolvedShape) -> Rect {
    resolved_shape_page_transform(shape).transform_rect_bbox(Rect {
        x: 0.0,
        y: 0.0,
        width: shape.bounds.width,
        height: shape.bounds.height,
    })
}

fn resolved_shape_page_transform(shape: &ResolvedShape) -> Transform {
    let center_x = shape.bounds.width / 2.0;
    let center_y = shape.bounds.height / 2.0;
    let flip = Transform {
        a: if shape.flip_h { -1.0 } else { 1.0 },
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
        ..Transform::IDENTITY
    };
    Transform::rotate_about(shape.rotation_deg, center_x, center_y)
        .then(flip)
        .then(timeline_translation(shape.bounds.x, shape.bounds.y))
        .then(shape.group_transform)
}

fn group_clip_page_points(
    clip: Rect,
    group_geometry: TimelineGroupGeometry,
    group_transform: Transform,
) -> [Point; 4] {
    let normalized_to_page = group_geometry.normalized_to_page.then(group_transform);
    [
        Point {
            x: clip.x,
            y: clip.y,
        },
        Point {
            x: clip.x + clip.width,
            y: clip.y,
        },
        Point {
            x: clip.x + clip.width,
            y: clip.y + clip.height,
        },
        Point {
            x: clip.x,
            y: clip.y + clip.height,
        },
    ]
    .map(|point| normalized_to_page.apply(point))
}

fn page_clip_for_shape(page_clip: [Point; 4], shape: &ResolvedShape) -> Option<[Point; 4]> {
    if shape.bounds.width.abs() <= f64::EPSILON || shape.bounds.height.abs() <= f64::EPSILON {
        return None;
    }
    let page_to_local = inverse_transform(resolved_shape_page_transform(shape))?;
    Some(page_clip.map(|point| {
        let local = page_to_local.apply(point);
        Point {
            x: local.x / shape.bounds.width,
            y: local.y / shape.bounds.height,
        }
    }))
}

fn inverse_transform(transform: Transform) -> Option<Transform> {
    let determinant = transform.a * transform.d - transform.b * transform.c;
    if !determinant.is_finite() || determinant.abs() <= f64::EPSILON {
        return None;
    }
    let inverse = Transform {
        a: transform.d / determinant,
        b: -transform.b / determinant,
        c: -transform.c / determinant,
        d: transform.a / determinant,
        e: (transform.c * transform.f - transform.d * transform.e) / determinant,
        f: (transform.b * transform.e - transform.a * transform.f) / determinant,
    };
    [
        inverse.a, inverse.b, inverse.c, inverse.d, inverse.e, inverse.f,
    ]
    .into_iter()
    .all(f64::is_finite)
    .then_some(inverse)
}

fn union_rect(left: Rect, right: Rect) -> Rect {
    let x = left.x.min(right.x);
    let y = left.y.min(right.y);
    let far_x = (left.x + left.width).max(right.x + right.width);
    let far_y = (left.y + left.height).max(right.y + right.height);
    Rect {
        x,
        y,
        width: far_x - x,
        height: far_y - y,
    }
}

fn intersect_rect(left: Rect, right: Rect) -> Rect {
    let x = left.x.max(right.x);
    let y = left.y.max(right.y);
    let far_x = (left.x + left.width).min(right.x + right.width);
    let far_y = (left.y + left.height).min(right.y + right.height);
    Rect {
        x,
        y,
        width: (far_x - x).max(0.0),
        height: (far_y - y).max(0.0),
    }
}

fn timeline_translation(x: f64, y: f64) -> Transform {
    Transform {
        e: x,
        f: y,
        ..Transform::IDENTITY
    }
}

fn intersect_optional_rect(left: Option<Rect>, right: Option<Rect>) -> Option<Rect> {
    match (left, right) {
        (None, None) => None,
        (Some(rect), None) | (None, Some(rect)) => Some(rect),
        (Some(left), Some(right)) => Some(intersect_rect(left, right)),
    }
}

fn emit_tree<'a>(
    children: &'a [ShapeTreeChild],
    rules: PassRules<'_>,
    output: &mut Vec<FlattenedItem<'a>>,
) {
    let mut emitted_latent = Vec::new();
    emit_tree_inner(
        children,
        rules,
        Transform::IDENTITY,
        0,
        &mut emitted_latent,
        output,
    );
}

fn emit_tree_inner<'a>(
    children: &'a [ShapeTreeChild],
    rules: PassRules<'_>,
    parent_transform: Transform,
    parent_issues: u8,
    emitted_latent: &mut Vec<PlaceholderKey>,
    output: &mut Vec<FlattenedItem<'a>>,
) {
    for child in children {
        match child {
            ShapeTreeChild::GroupShape(group) => {
                let (transform, issues) = group
                    .group_transform()
                    .and_then(group_affine)
                    .unwrap_or((Transform::IDENTITY, 0));
                emit_tree_inner(
                    &group.children,
                    rules,
                    transform.then(parent_transform),
                    parent_issues | issues,
                    emitted_latent,
                    output,
                );
            }
            ShapeTreeChild::AlternateContent(alternate) => {
                if alternate.chart_choice().is_some() {
                    emit_leaf(
                        child,
                        rules,
                        parent_transform,
                        parent_issues,
                        emitted_latent,
                        output,
                    );
                } else if let Some(fallback) = alternate.selected_fallback() {
                    emit_tree_inner(
                        fallback,
                        rules,
                        parent_transform,
                        parent_issues,
                        emitted_latent,
                        output,
                    );
                }
            }
            _ => emit_leaf(
                child,
                rules,
                parent_transform,
                parent_issues,
                emitted_latent,
                output,
            ),
        }
    }
}

fn emit_leaf<'a>(
    child: &'a ShapeTreeChild,
    rules: PassRules<'_>,
    group_transform: Transform,
    group_issues: u8,
    emitted_latent: &mut Vec<PlaceholderKey>,
    output: &mut Vec<FlattenedItem<'a>>,
) {
    let placeholder = child_placeholder(child);
    if rules.source == FlattenedSource::Slide {
        if let Some(placeholder) = placeholder {
            let ph_type = placeholder.effective_type();
            if is_latent(&ph_type) {
                let occupied = child_is_occupied(child);
                let already_emitted = emitted_latent
                    .iter()
                    .any(|key| placeholder_keys_match(&placeholder.key(), key));
                if !rules.latent_policy.permits(&ph_type) || !occupied || already_emitted {
                    return;
                }
                push_unmatched(emitted_latent, placeholder.key());
            }
        }
        push_flattened_shape(output, rules.source, child, group_transform, group_issues);
        return;
    }

    let Some(placeholder) = placeholder else {
        if rules.emit_non_placeholders {
            push_flattened_shape(output, rules.source, child, group_transform, group_issues);
        }
        return;
    };
    let ph_type = placeholder.effective_type();
    let key = placeholder.key();
    if !is_latent(&ph_type)
        || !rules.emit_inherited_latent
        || !rules.latent_policy.permits(&ph_type)
        || !child_is_occupied(child)
        || rules
            .deeper_latent
            .iter()
            .any(|deeper| placeholder_keys_match(&key, deeper))
        || emitted_latent
            .iter()
            .any(|emitted| placeholder_keys_match(&key, emitted))
    {
        return;
    }
    push_unmatched(emitted_latent, key);
    push_flattened_shape(output, rules.source, child, group_transform, group_issues);
}

fn push_flattened_shape<'a>(
    output: &mut Vec<FlattenedItem<'a>>,
    source: FlattenedSource,
    child: &'a ShapeTreeChild,
    group_transform: Transform,
    group_issues: u8,
) {
    let (group_scale, group_transform, group_issues) =
        split_group_transform(group_transform, group_issues);
    output.push(FlattenedItem::Shape {
        source,
        child,
        group_scale,
        group_transform,
        group_issues,
    });
}

fn split_group_transform(transform: Transform, mut issues: u8) -> ((f64, f64), Transform, u8) {
    let scale_x = transform.a.hypot(transform.b);
    let scale_y = transform.c.hypot(transform.d);
    let dot = transform.a * transform.c + transform.b * transform.d;
    let denominator = scale_x * scale_y;
    let is_rigid_after_scale = [
        transform.a,
        transform.b,
        transform.c,
        transform.d,
        transform.e,
        transform.f,
        scale_x,
        scale_y,
    ]
    .into_iter()
    .all(f64::is_finite)
        && scale_x > f64::EPSILON
        && scale_y > f64::EPSILON
        && (dot / denominator).abs() <= 1.0e-9;

    if !is_rigid_after_scale {
        issues |= GROUP_SHEAR;
        return ((1.0, 1.0), transform, issues);
    }

    (
        (scale_x, scale_y),
        Transform {
            a: transform.a / scale_x,
            b: transform.b / scale_x,
            c: transform.c / scale_y,
            d: transform.d / scale_y,
            e: transform.e,
            f: transform.f,
        },
        issues,
    )
}

fn group_affine(transform: &CT_Transform2D) -> Option<(Transform, u8)> {
    let offset = transform.offset?;
    let extent = transform.extent?;
    let child_offset = transform.child_offset?;
    let child_extent = transform.child_extent?;
    let mut issues = 0;
    let x = emu_to_points(offset.x.0);
    let y = emu_to_points(offset.y.0);
    let width = emu_to_points(extent.cx.0);
    let height = emu_to_points(extent.cy.0);
    let child_x = emu_to_points(child_offset.x.0);
    let child_y = emu_to_points(child_offset.y.0);
    let scale_x = if child_extent.cx.0 == 0 {
        issues |= GROUP_ZERO_X;
        1.0
    } else {
        extent.cx.0 as f64 / child_extent.cx.0 as f64
    };
    let scale_y = if child_extent.cy.0 == 0 {
        issues |= GROUP_ZERO_Y;
        1.0
    } else {
        extent.cy.0 as f64 / child_extent.cy.0 as f64
    };
    let mut affine = Transform {
        a: scale_x,
        b: 0.0,
        c: 0.0,
        d: scale_y,
        e: x - child_x * scale_x,
        f: y - child_y * scale_y,
    };
    let center_x = x + width / 2.0;
    let center_y = y + height / 2.0;
    affine = affine.then(Transform::rotate_about(
        f64::from(transform.rotation.0) / 60_000.0,
        center_x,
        center_y,
    ));
    if transform.flip_horizontal || transform.flip_vertical {
        affine = affine.then(Transform {
            a: if transform.flip_horizontal { -1.0 } else { 1.0 },
            b: 0.0,
            c: 0.0,
            d: if transform.flip_vertical { -1.0 } else { 1.0 },
            e: if transform.flip_horizontal {
                2.0 * center_x
            } else {
                0.0
            },
            f: if transform.flip_vertical {
                2.0 * center_y
            } else {
                0.0
            },
        });
    }
    Some((affine, issues))
}

fn child_placeholder(child: &ShapeTreeChild) -> Option<&CT_Placeholder> {
    match child {
        ShapeTreeChild::Shape(shape) => shape.placeholder.as_ref(),
        ShapeTreeChild::Picture(picture) => picture.placeholder.as_ref(),
        _ => None,
    }
}

fn child_is_occupied(child: &ShapeTreeChild) -> bool {
    match child {
        ShapeTreeChild::Shape(shape) => shape
            .text_body
            .as_ref()
            .is_some_and(|body| !body.plain_text().trim().is_empty()),
        ShapeTreeChild::Picture(picture) => picture.blip_fill.is_some(),
        _ => false,
    }
}

fn is_latent(ph_type: &PhType) -> bool {
    matches!(
        ph_type,
        PhType::DateTime | PhType::Footer | PhType::SlideNumber
    )
}

fn placeholder_keys_match(left: &PlaceholderKey, right: &PlaceholderKey) -> bool {
    if is_latent(&left.ph_type) && is_latent(&right.ph_type) {
        left.ph_type == right.ph_type
    } else {
        left.matches(right)
    }
}

fn default_body_properties() -> CT_TextBodyProperties {
    let mut properties = CT_TextBodyProperties::default();
    properties.left_inset = Some(Coordinate32Value::Emu(91_440));
    properties.top_inset = Some(Coordinate32Value::Emu(45_720));
    properties.right_inset = Some(Coordinate32Value::Emu(91_440));
    properties.bottom_inset = Some(Coordinate32Value::Emu(45_720));
    properties.anchor = Some(DrawingTextAnchor::Top);
    properties.wrap = Some(TextWrap::Square);
    properties.vertical = Some(DrawingTextVertical::Horizontal);
    properties.space_first_last_paragraph = Some(false);
    properties.autofit = Some(TextAutofit::NoAutofit);
    properties
}

fn merge_body_properties(target: &mut CT_TextBodyProperties, source: &CT_TextBodyProperties) {
    if let Some(value) = &source.left_inset {
        target.left_inset = Some(value.clone());
    }
    if let Some(value) = &source.top_inset {
        target.top_inset = Some(value.clone());
    }
    if let Some(value) = &source.right_inset {
        target.right_inset = Some(value.clone());
    }
    if let Some(value) = &source.bottom_inset {
        target.bottom_inset = Some(value.clone());
    }
    if let Some(value) = source.anchor {
        target.anchor = Some(value);
    }
    if let Some(value) = source.wrap {
        target.wrap = Some(value);
    }
    if let Some(value) = source.vertical {
        target.vertical = Some(value);
    }
    if let Some(value) = source.space_first_last_paragraph {
        target.space_first_last_paragraph = Some(value);
    }
    if let Some(value) = &source.autofit {
        target.autofit = Some(value.clone());
    }
}

fn find_placeholder<'a>(
    children: &'a [ShapeTreeChild],
    key: &PlaceholderKey,
) -> Option<&'a CT_Shape> {
    for child in children {
        let found = match child {
            ShapeTreeChild::Shape(shape) => shape
                .placeholder
                .as_ref()
                .is_some_and(|placeholder| placeholder_keys_match(key, &placeholder.key()))
                .then_some(shape),
            ShapeTreeChild::GroupShape(group) => find_placeholder(&group.children, key),
            ShapeTreeChild::AlternateContent(alternate) => alternate
                .selected_fallback()
                .and_then(|fallback| find_placeholder(fallback, key)),
            _ => None,
        };
        if found.is_some() {
            return found;
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;
    use std::path::{Path, PathBuf};

    use oxml_chart::CT_ChartSpace;
    use oxml_drawing::color::ColorMap;
    use oxml_drawing::fill::Fill;
    use oxml_drawing::text::{
        CT_TextListStyle, Coordinate32Value, TextAnchor, TextAutofit, TextVertical, TextWrap,
    };
    use oxml_drawing::theme::CT_OfficeStyleSheet;
    use oxml_opc::OpcPackage;
    use oxml_opc::relationship::rel_types;
    use rpptx_oxml::placeholder::PhType;
    use rpptx_oxml::presentation::CT_Presentation;
    use rpptx_oxml::shape_tree::{CT_Shape, ShapeTreeChild};
    use rpptx_oxml::slide_parts::{CT_Slide, CT_SlideLayout, CT_SlideMaster, ColorMapOverrideKind};

    use super::{
        BackgroundSource, FlattenedItem, FlattenedSource, ResolveCtx, resolved_shape_page_bounds,
        resolved_shape_page_transform, transform_values,
    };
    use crate::{
        ChartResource, Diagnostic, ParagraphAlignment, ResolvedAutofit, ResolvedBackground,
        ResolvedBullet, ResolvedBulletSize, ResolvedContent, ResolvedGeometry,
        ResolvedImagePlacement, ResolvedLineEnd, ResolvedLineEndKind, ResolvedLineEndSize,
        ResolvedParagraph, ResolvedRectAlignment, ResolvedRunStyle, ResolvedSlide, ResolvedTextRun,
        ResolvedTextSpacing, ResolvedTileFlip, ScopedChartResources, ScopedHyperlinkTargets,
        ScopedMediaIds, TextAnchor as ResolvedTextAnchor, TextDirection,
    };
    use oxml_layout::{
        Color, Effect, FontManager, MediaId, Paint, PathCommand, Point, PositionedElement, Rect,
        Transform, walk,
    };

    const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const EXPECTED_CORPUS_DECKS: usize = 50;

    fn assert_close(actual: f64, expected: f64) {
        assert!(
            (actual - expected).abs() < 1.0e-10,
            "expected {expected}, got {actual}"
        );
    }

    struct Fixture {
        theme: CT_OfficeStyleSheet,
        master: CT_SlideMaster,
        layout: CT_SlideLayout,
        slide: CT_Slide,
        default_text_style: CT_TextListStyle,
    }

    impl Fixture {
        fn new(slide_children: &str, layout_children: &str, master_children: &str) -> Self {
            Self::from_xml(
                &slide_xml(slide_children),
                &layout_xml(layout_children),
                &master_xml(master_children),
            )
        }

        fn from_xml(slide_xml: &str, layout_xml: &str, master_xml: &str) -> Self {
            Self {
                theme: CT_OfficeStyleSheet::office_default(),
                master: CT_SlideMaster::from_xml(master_xml.as_bytes()).unwrap(),
                layout: CT_SlideLayout::from_xml(layout_xml.as_bytes()).unwrap(),
                slide: CT_Slide::from_xml(slide_xml.as_bytes()).unwrap(),
                default_text_style: CT_TextListStyle::default(),
            }
        }

        fn context(&self) -> ResolveCtx<'_> {
            ResolveCtx::new(
                &self.theme,
                ColorMap::default(),
                &self.master,
                &self.layout,
                &self.slide,
                &self.default_text_style,
            )
        }

        fn slide_shape(&self, index: usize) -> &CT_Shape {
            let ShapeTreeChild::Shape(shape) =
                &self.slide.common_slide_data.shape_tree.children[index]
            else {
                panic!("expected an ordinary slide shape");
            };
            shape
        }
    }

    #[test]
    fn slide_placeholder_resolves_to_layout_and_master_counterparts() {
        let fixture = Fixture::new(
            &shape(Some("title"), Some(7)),
            &shape(Some("body"), Some(7)),
            &shape(Some("pic"), Some(7)),
        );

        let context = fixture.context();
        let (layout, master) = context.placeholder_chain(fixture.slide_shape(0));

        assert_eq!(
            layout.unwrap().placeholder.as_ref().unwrap().ph_type,
            Some(PhType::Body)
        );
        assert_eq!(
            master.unwrap().placeholder.as_ref().unwrap().ph_type,
            Some(PhType::Picture)
        );
    }

    #[test]
    fn master_match_uses_the_layout_placeholder_key() {
        let fixture = Fixture::new(
            &shape(Some("title"), Some(8)),
            &shape(Some("ctrTitle"), None),
            &shape(Some("title"), Some(5)),
        );

        let context = fixture.context();
        let (layout, master) = context.placeholder_chain(fixture.slide_shape(0));

        assert!(layout.is_some());
        assert_eq!(master.unwrap().placeholder.as_ref().unwrap().idx, Some(5));
    }

    #[test]
    fn placeholder_lookup_walks_groups_and_selected_fallbacks() {
        let nested_layout = format!(
            "<p:grpSp><p:nvGrpSpPr/><p:grpSpPr/>{}</p:grpSp>",
            shape(Some("body"), Some(11))
        );
        let fallback_master = format!(
            "<mc:AlternateContent><mc:Choice Requires=\"p14\"><p:sp/></mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent>",
            shape(Some("body"), Some(11))
        );
        let fixture = Fixture::new(
            &shape(Some("body"), Some(11)),
            &nested_layout,
            &fallback_master,
        );

        let context = fixture.context();
        let (layout, master) = context.placeholder_chain(fixture.slide_shape(0));

        assert!(layout.is_some());
        assert!(master.is_some());
    }

    #[test]
    fn group_targets_apply_to_every_resolved_descendant() {
        let first = shape_with_details(None, None, &transform(0), None).replacen(
            "<p:cNvPr/>",
            "<p:cNvPr id=\"7\" name=\"First\"/>",
            1,
        );
        let second = shape_with_details(None, None, &transform(127_000), None).replacen(
            "<p:cNvPr/>",
            "<p:cNvPr id=\"8\" name=\"Second\"/>",
            1,
        );
        let unrelated = shape_with_details(None, None, &transform(254_000), None).replacen(
            "<p:cNvPr/>",
            "<p:cNvPr id=\"9\" name=\"Other\"/>",
            1,
        );
        let group = format!(
            "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"10\" name=\"Group\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{first}{second}</p:grpSp>"
        );
        let slide = format!(
            "<p:sld xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"><p:cSld>{}</p:cSld><p:timing><p:tnLst><p:set><p:cBhvr><p:cTn id=\"1\" dur=\"1\" fill=\"hold\"/><p:tgtEl><p:spTgt spid=\"10\"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val=\"50000\"/></p:to></p:set></p:tnLst></p:timing></p:sld>",
            shape_tree(&format!("{group}{unrelated}"))
        );
        let fixture = Fixture::from_xml(&slide, &layout_xml(""), &master_xml(""));
        let resolved = fixture
            .context()
            .resolve_slide_at(
                (720.0, 540.0),
                crate::timeline::TimelinePosition {
                    elapsed_ms: 1,
                    click_count: 0,
                },
            )
            .unwrap();

        assert_eq!(resolved.shape_states.len(), 3);
        assert_eq!(resolved.shape_states[0].opacity, 0.5);
        assert_eq!(resolved.shape_states[1].opacity, 0.5);
        assert_eq!(resolved.shape_states[2].opacity, 1.0);
        assert_eq!(resolved.identities[0].containing_group_ids, vec![10]);
        assert_eq!(resolved.identities[1].containing_group_ids, vec![10]);
        assert!(resolved.identities[2].containing_group_ids.is_empty());
    }

    #[test]
    fn alternate_content_chart_choice_keeps_timeline_identity() {
        let alternate = chart_alternate("chart", None).replacen(
            "<p:nvGraphicFramePr/>",
            "<p:nvGraphicFramePr><p:cNvPr id=\"42\" name=\"!!Chart &amp; Data\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>",
            1,
        );
        let slide = format!(
            r#"<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld>{}</p:cSld><p:timing><p:tnLst><p:anim from="100000" to="50000"><p:cBhvr><p:cTn id="1" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="42"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr></p:anim></p:tnLst></p:timing></p:sld>"#,
            shape_tree(&alternate)
        );
        let fixture = Fixture::from_xml(&slide, &layout_xml(""), &master_xml(""));
        let resolved = fixture
            .context()
            .resolve_slide_at(
                (720.0, 540.0),
                crate::timeline::TimelinePosition {
                    elapsed_ms: 500,
                    click_count: 0,
                },
            )
            .unwrap();

        assert_eq!(resolved.identities[0].shape_id, Some(42));
        assert_eq!(
            resolved.identities[0].name.as_deref(),
            Some("!!Chart & Data")
        );
        assert_eq!(resolved.shape_states[0].opacity, 0.75);
    }

    #[test]
    fn timeline_geometry_keeps_shape_centres_and_resolves_motion_origins() {
        let first = shape_with_details(None, None, &transform(127_000), None).replacen(
            "<p:cNvPr/>",
            "<p:cNvPr id=\"7\" name=\"!!Scale &amp; Spin\"/>",
            1,
        );
        let second = shape_with_details(None, None, &transform(254_000), None).replacen(
            "<p:cNvPr/>",
            "<p:cNvPr id=\"8\" name=\"Target motion\"/>",
            1,
        );
        let slide = format!(
            r#"<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld>{}</p:cSld><p:timing><p:tnLst>
            <p:anim from="100000" to="200000"><p:cBhvr><p:cTn id="1" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl><p:attrNameLst><p:attrName>scale</p:attrName></p:attrNameLst></p:cBhvr></p:anim>
            <p:anim from="0" to="5400000"><p:cBhvr><p:cTn id="2" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl><p:attrNameLst><p:attrName>spin</p:attrName></p:attrNameLst></p:cBhvr></p:anim>
            <p:animMotion origin="parent" path="M 0 0 L .5 .25 E"><p:cBhvr><p:cTn id="3" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cBhvr></p:animMotion>
            <p:animMotion origin="layout" path="M 0 0 L .5 .25 E"><p:cBhvr><p:cTn id="4" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="8"/></p:tgtEl></p:cBhvr></p:animMotion>
            </p:tnLst></p:timing></p:sld>"#,
            shape_tree(&format!("{first}{second}"))
        );
        let fixture = Fixture::from_xml(&slide, &layout_xml(""), &master_xml(""));
        let static_slide = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let resolved = fixture
            .context()
            .resolve_slide_at(
                (720.0, 540.0),
                crate::timeline::TimelinePosition {
                    elapsed_ms: 1_000,
                    click_count: 0,
                },
            )
            .unwrap();

        let original = static_slide.shapes[0].bounds;
        let animated = resolved.slide.shapes[0].bounds;
        assert_close(animated.width, original.width * 2.0);
        assert_close(animated.height, original.height * 2.0);
        let animated_page = resolved_shape_page_bounds(&resolved.slide.shapes[0]);
        assert_close(animated_page.x + animated_page.width / 2.0, 360.0);
        assert_close(animated_page.y + animated_page.height / 2.0, 135.0);
        assert_close(resolved.slide.shapes[0].rotation_deg, 90.0);
        let target_original = resolved_shape_page_bounds(&static_slide.shapes[1]);
        let target_animated = resolved_shape_page_bounds(&resolved.slide.shapes[1]);
        assert_close(target_animated.x - target_original.x, 360.0);
        assert_close(target_animated.y - target_original.y, 135.0);
        assert_eq!(
            resolved.identities[0].name.as_deref(),
            Some("!!Scale & Spin")
        );
        assert!(resolved.shape_states.iter().all(|state| state.is_finite()));
    }

    #[test]
    fn group_timeline_geometry_uses_one_shared_group_coordinate_space() {
        let rotated_transform = transform(50).replacen("<a:xfrm>", "<a:xfrm rot=\"2700000\">", 1);
        let first = shape_with_details(None, None, &rotated_transform, None).replacen(
            "<p:cNvPr/>",
            "<p:cNvPr id=\"7\" name=\"First\"/>",
            1,
        );
        let second = shape_with_details(None, None, &transform(100), None).replacen(
            "<p:cNvPr/>",
            "<p:cNvPr id=\"8\" name=\"Second\"/>",
            1,
        );
        let group = format!(
            "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"10\" name=\"Group\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot=\"5400000\"><a:off x=\"1270000\" y=\"635000\"/><a:ext cx=\"2540000\" cy=\"1270000\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"200\" cy=\"100\"/></a:xfrm></p:grpSpPr>{first}{second}</p:grpSp>"
        );
        let slide = format!(
            r#"<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld>{}</p:cSld><p:timing><p:tnLst>
            <p:anim from="100000" to="200000"><p:cBhvr><p:cTn id="1" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="10"/></p:tgtEl><p:attrNameLst><p:attrName>scale</p:attrName></p:attrNameLst></p:cBhvr></p:anim>
            <p:animMotion origin="parent" path="M 0 0 L .1 .1 E"><p:cBhvr><p:cTn id="2" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="10"/></p:tgtEl></p:cBhvr></p:animMotion>
            <p:animEffect transition="in" filter="wipe(right)"><p:cBhvr><p:cTn id="3" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="10"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:tnLst></p:timing></p:sld>"#,
            shape_tree(&group)
        );
        let fixture = Fixture::from_xml(&slide, &layout_xml(""), &master_xml(""));
        let static_slide = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let resolved = fixture
            .context()
            .resolve_slide_at(
                (720.0, 540.0),
                crate::timeline::TimelinePosition {
                    elapsed_ms: 500,
                    click_count: 0,
                },
            )
            .unwrap();
        let first_before = resolved_shape_page_bounds(&static_slide.shapes[0]);
        let second_before = resolved_shape_page_bounds(&static_slide.shapes[1]);
        let group_center_x = 200.0;
        let group_center_y = 100.0;

        for (before, shape) in [first_before, second_before]
            .into_iter()
            .zip(&resolved.slide.shapes)
        {
            let after = resolved_shape_page_bounds(shape);
            assert_close(after.width, before.width * 1.5);
            assert_close(
                after.x + after.width / 2.0,
                36.0 + (before.x + before.width / 2.0 - group_center_x) * 1.5,
            );
            assert_close(
                after.y + after.height / 2.0,
                27.0 + (before.y + before.height / 2.0 - group_center_y) * 1.5,
            );
        }
        let first_clip = resolved.shape_states[0].local_clip_paths((100.0, 100.0));
        let [PathCommand::MoveTo(start), PathCommand::LineTo(next), ..] =
            first_clip[0].commands.as_slice()
        else {
            panic!("group wipe should remain a polygon")
        };
        assert!(start.x != next.x && start.y != next.y);
        assert_eq!(
            resolved.shape_states[1]
                .local_clip_paths((100.0, 100.0))
                .len(),
            1
        );
    }

    #[test]
    fn parent_group_wipe_does_not_follow_descendant_or_nested_group_animation() {
        let child = shape_with_details(None, None, &transform(50), None).replacen(
            "<p:cNvPr/>",
            "<p:cNvPr id=\"7\" name=\"Animated child\"/>",
            1,
        );
        let nested = format!(
            "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"11\" name=\"Nested\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"200\" cy=\"100\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"200\" cy=\"100\"/></a:xfrm></p:grpSpPr>{child}</p:grpSp>"
        );
        let group = format!(
            "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"10\" name=\"Parent\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot=\"2700000\"><a:off x=\"1270000\" y=\"635000\"/><a:ext cx=\"2540000\" cy=\"1270000\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"200\" cy=\"100\"/></a:xfrm></p:grpSpPr>{nested}</p:grpSp>"
        );
        let resolve = |descendant_timing: &str| {
            let slide = format!(
                r#"<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld>{}</p:cSld><p:timing><p:tnLst>
                <p:animEffect transition="in" filter="wipe(right)"><p:cBhvr><p:cTn id="1" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="10"/></p:tgtEl></p:cBhvr></p:animEffect>
                {descendant_timing}</p:tnLst></p:timing></p:sld>"#,
                shape_tree(&group)
            );
            Fixture::from_xml(&slide, &layout_xml(""), &master_xml(""))
                .context()
                .resolve_slide_at(
                    (720.0, 540.0),
                    crate::timeline::TimelinePosition {
                        elapsed_ms: 500,
                        click_count: 0,
                    },
                )
                .unwrap()
        };
        let baseline = resolve("");
        let animated = resolve(
            r#"<p:animMotion origin="layout" path="M 0 0 L .2 .1 E"><p:cBhvr><p:cTn id="2" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="11"/></p:tgtEl></p:cBhvr></p:animMotion>
            <p:anim from="0" to="5400000"><p:cBhvr><p:cTn id="3" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl><p:attrNameLst><p:attrName>spin</p:attrName></p:attrNameLst></p:cBhvr></p:anim>"#,
        );
        let page_clip = |resolved: &crate::timeline::ResolvedTimelineSlide| {
            let shape = &resolved.slide.shapes[0];
            let transform = resolved_shape_page_transform(shape);
            resolved.shape_states[0].local_clip_paths((shape.bounds.width, shape.bounds.height))[0]
                .commands
                .iter()
                .filter_map(|command| match command {
                    PathCommand::MoveTo(point) | PathCommand::LineTo(point) => {
                        Some(transform.apply(*point))
                    }
                    _ => None,
                })
                .collect::<Vec<_>>()
        };
        let baseline_clip = page_clip(&baseline);
        let animated_clip = page_clip(&animated);

        assert_ne!(
            resolved_shape_page_transform(&baseline.slide.shapes[0]),
            resolved_shape_page_transform(&animated.slide.shapes[0])
        );
        assert_eq!(baseline_clip.len(), animated_clip.len());
        for (baseline, animated) in baseline_clip.into_iter().zip(animated_clip) {
            assert_close(animated.x, baseline.x);
            assert_close(animated.y, baseline.y);
        }
    }

    #[test]
    fn missing_or_non_placeholder_shape_has_no_chain() {
        let slide_children = format!("{}{}", shape(None, None), shape(Some("body"), Some(42)));
        let fixture = Fixture::new(&slide_children, "", &shape(Some("body"), Some(42)));
        let context = fixture.context();

        assert_eq!(
            context.placeholder_chain(fixture.slide_shape(0)),
            (None, None)
        );
        assert_eq!(
            context.placeholder_chain(fixture.slide_shape(1)),
            (None, None)
        );
    }

    #[test]
    fn slide_placeholder_without_transform_inherits_layout_position() {
        let fixture = Fixture::new(
            &shape(Some("body"), Some(1)),
            &shape_with_details(Some("body"), Some(1), &transform(10), None),
            &shape_with_details(Some("body"), Some(1), &transform(20), None),
        );
        let transform = fixture
            .context()
            .effective_xfrm(fixture.slide_shape(0))
            .unwrap();

        assert_eq!(transform.offset.unwrap().x.0, 10);
    }

    #[test]
    fn effective_transform_uses_slide_layout_master_precedence() {
        let slide_children = [
            shape_with_details(Some("body"), Some(1), &transform(11), None),
            shape(Some("body"), Some(2)),
            shape(Some("body"), Some(3)),
            shape(Some("body"), Some(4)),
        ]
        .join("");
        let layout_children = [
            shape_with_details(Some("body"), Some(1), &transform(21), None),
            shape_with_details(Some("body"), Some(2), &transform(22), None),
            shape(Some("body"), Some(3)),
            shape(Some("body"), Some(4)),
        ]
        .join("");
        let master_children = [
            shape_with_details(Some("body"), Some(1), &transform(31), None),
            shape_with_details(Some("body"), Some(2), &transform(32), None),
            shape_with_details(Some("body"), Some(3), &transform(33), None),
            shape(Some("body"), Some(4)),
        ]
        .join("");
        let fixture = Fixture::new(&slide_children, &layout_children, &master_children);
        let context = fixture.context();

        let offsets: Vec<_> = (0..3)
            .map(|index| {
                context
                    .effective_xfrm(fixture.slide_shape(index))
                    .unwrap()
                    .offset
                    .unwrap()
                    .x
                    .0
            })
            .collect();

        assert_eq!(offsets, [11, 22, 33]);
        assert!(context.effective_xfrm(fixture.slide_shape(3)).is_none());
    }

    #[test]
    fn body_properties_merge_per_field_across_the_chain() {
        let fixture = Fixture::new(
            &shape_with_details(
                Some("body"),
                Some(4),
                "",
                Some(r#"<a:bodyPr bIns="50" anchor="ctr"/>"#),
            ),
            &shape_with_details(
                Some("body"),
                Some(4),
                "",
                Some(r#"<a:bodyPr lIns="30" rIns="40" wrap="none"/>"#),
            ),
            &shape_with_details(
                Some("body"),
                Some(4),
                "",
                Some(
                    r#"<a:bodyPr lIns="10" tIns="20" anchor="b" vert="vert" spcFirstLastPara="1"><a:spAutoFit/></a:bodyPr>"#,
                ),
            ),
        );
        let properties = fixture.context().effective_body_pr(fixture.slide_shape(0));

        assert_eq!(properties.left_inset, Some(Coordinate32Value::Emu(30)));
        assert_eq!(properties.top_inset, Some(Coordinate32Value::Emu(20)));
        assert_eq!(properties.right_inset, Some(Coordinate32Value::Emu(40)));
        assert_eq!(properties.bottom_inset, Some(Coordinate32Value::Emu(50)));
        assert_eq!(properties.anchor, Some(TextAnchor::Center));
        assert_eq!(properties.wrap, Some(TextWrap::None));
        assert_eq!(properties.vertical, Some(TextVertical::Vertical));
        assert_eq!(properties.space_first_last_paragraph, Some(true));
        assert_eq!(properties.autofit, Some(TextAutofit::ShapeAutofit));
    }

    #[test]
    fn body_property_defaults_use_exact_emu_values() {
        let fixture = Fixture::new(&shape(None, None), "", "");
        let properties = fixture.context().effective_body_pr(fixture.slide_shape(0));

        assert_eq!(properties.left_inset, Some(Coordinate32Value::Emu(91_440)));
        assert_eq!(properties.right_inset, Some(Coordinate32Value::Emu(91_440)));
        assert_eq!(properties.top_inset, Some(Coordinate32Value::Emu(45_720)));
        assert_eq!(
            properties.bottom_inset,
            Some(Coordinate32Value::Emu(45_720))
        );
        assert_eq!(properties.anchor, Some(TextAnchor::Top));
        assert_eq!(properties.wrap, Some(TextWrap::Square));
        assert_eq!(properties.vertical, Some(TextVertical::Horizontal));
        assert_eq!(properties.space_first_last_paragraph, Some(false));
        assert_eq!(properties.autofit, Some(TextAutofit::NoAutofit));
    }

    #[test]
    fn east_asian_vertical_text_degrades_to_rotated_with_a_diagnostic() {
        let shape = shape_with_details(
            None,
            None,
            &transform(0),
            Some(r#"<a:bodyPr vert="eaVert"/>"#),
        )
        .replace("<a:p/>", "<a:p><a:r><a:t>visible</a:t></a:r></a:p>");
        let resolved = Fixture::new(&shape, "", "")
            .context()
            .resolve_slide((720.0, 540.0))
            .expect("resolve East Asian vertical text");

        let ResolvedContent::Text(text) = &resolved.shapes[0].content else {
            panic!("fallback text should remain visible");
        };
        assert_eq!(text.vertical, TextDirection::EastAsianVertical);
        assert!(matches!(
            &text.paragraphs[0].runs[0],
            ResolvedTextRun::Text { text, .. } if text == "visible"
        ));
        assert!(resolved.diagnostics.iter().any(|diagnostic| {
            diagnostic.message == "east Asian vertical text rendered as rotated vertical text"
        }));
    }

    #[test]
    fn other_vertical_variants_remain_visible_with_diagnostics() {
        let cases = [
            (
                "mongolianVert",
                TextDirection::MongolianVertical,
                "Mongolian vertical text rendered as rotated vertical-270 text",
            ),
            (
                "wordArtVert",
                TextDirection::WordArtVertical,
                "WordArt vertical text rendered as rotated vertical text",
            ),
            (
                "wordArtVertRtl",
                TextDirection::WordArtVerticalRtl,
                "right-to-left WordArt vertical text rendered as rotated vertical-270 text",
            ),
        ];

        for (value, expected_direction, expected_diagnostic) in cases {
            let shape = shape_with_details(
                None,
                None,
                &transform(0),
                Some(&format!(r#"<a:bodyPr vert="{value}"/>"#)),
            )
            .replace("<a:p/>", "<a:p><a:r><a:t>visible</a:t></a:r></a:p>");
            let resolved = Fixture::new(&shape, "", "")
                .context()
                .resolve_slide((720.0, 540.0))
                .expect("resolve visible vertical fallback");

            let ResolvedContent::Text(text) = &resolved.shapes[0].content else {
                panic!("fallback text should remain visible");
            };
            assert_eq!(text.vertical, expected_direction);
            assert!(matches!(
                &text.paragraphs[0].runs[0],
                ResolvedTextRun::Text { text, .. } if text == "visible"
            ));
            assert!(
                resolved
                    .diagnostics
                    .iter()
                    .any(|diagnostic| { diagnostic.message == expected_diagnostic })
            );
        }
    }

    #[test]
    fn flattener_omits_template_prompt_and_emits_master_logo_once() {
        let fixture = Fixture::new(
            &shape_with_text(Some("title"), Some(1), "Slide title"),
            "",
            &[
                shape_with_text(Some("title"), Some(1), "Click to edit Master title style"),
                shape_with_text(None, None, "Master logo"),
            ]
            .join(""),
        );

        let texts: Vec<_> = fixture
            .context()
            .flatten()
            .iter()
            .filter_map(item_text)
            .collect();

        assert!(!texts.contains(&"Click to edit Master title style".to_owned()));
        assert_eq!(
            texts.iter().filter(|text| *text == "Master logo").count(),
            1
        );
    }

    #[test]
    fn flattener_emits_the_four_sources_in_draw_order() {
        let master_group = format!(
            "<p:grpSp><p:nvGrpSpPr/><p:grpSpPr/>{}</p:grpSp>",
            shape_with_text(None, None, "master")
        );
        let layout_fallback = format!(
            "<mc:AlternateContent><mc:Fallback>{}</mc:Fallback></mc:AlternateContent>",
            shape_with_text(None, None, "layout")
        );
        let fixture = Fixture::from_xml(
            &slide_xml_with(
                "",
                "<p:bg><p:bgPr/></p:bg>",
                &shape_with_text(None, None, "slide"),
            ),
            &layout_xml_with("", "", &layout_fallback, ""),
            &master_xml_with("", &master_group, ""),
        );
        let context = fixture.context();
        let flattened = context.flatten();

        assert_eq!(
            flattened
                .iter()
                .map(FlattenedItem::source)
                .collect::<Vec<_>>(),
            [
                FlattenedSource::Background,
                FlattenedSource::Master,
                FlattenedSource::Layout,
                FlattenedSource::Slide,
            ]
        );
        let FlattenedItem::Background(background) = flattened[0] else {
            panic!("expected background first");
        };
        assert_eq!(background.source, BackgroundSource::Slide);
    }

    #[test]
    fn background_precedence_stops_at_master_when_all_are_absent() {
        let slide_first = Fixture::from_xml(
            &slide_xml_with("", "<p:bg><p:bgPr/></p:bg>", ""),
            &layout_xml_with("", "<p:bg><p:bgPr/></p:bg>", "", ""),
            &master_xml_with("<p:bg><p:bgPr/></p:bg>", "", ""),
        );
        let layout_first = Fixture::from_xml(
            &slide_xml_with("", "", ""),
            &layout_xml_with("", "<p:bg><p:bgPr/></p:bg>", "", ""),
            &master_xml_with("<p:bg><p:bgPr/></p:bg>", "", ""),
        );
        let master_first = Fixture::from_xml(
            &slide_xml_with("", "", ""),
            &layout_xml_with("", "", "", ""),
            &master_xml_with("<p:bg><p:bgPr/></p:bg>", "", ""),
        );
        let absent = Fixture::new("", "", "");

        assert_eq!(background_source(&slide_first), BackgroundSource::Slide);
        assert_eq!(background_source(&layout_first), BackgroundSource::Layout);
        assert_eq!(background_source(&master_first), BackgroundSource::Master);
        assert!(absent.context().effective_background().is_none());
    }

    #[test]
    fn master_gradient_background_resolves_when_slide_and_layout_omit_one() {
        let background = r#"<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></p:bgPr></p:bg>"#;
        let fixture = Fixture::from_xml(
            &slide_xml_with("", "", ""),
            &layout_xml_with("", "", "", ""),
            &master_xml_with(background, "", ""),
        );

        let resolved = fixture.context().resolve_slide((40.0, 20.0)).unwrap();
        let Some(ResolvedBackground::Paint(Paint::Linear { stops, .. })) = resolved.background
        else {
            panic!("expected inherited master gradient paint");
        };
        assert_eq!(stops.len(), 2);
        assert_eq!(stops[0].color, Color::from_hex("FF0000"));
        assert_eq!(stops[1].color, Color::from_hex("0000FF"));
        assert!(resolved.diagnostics.is_empty());
    }

    #[test]
    fn background_gradient_ignores_shape_rotation_policy() {
        let background = r#"<p:bg><p:bgPr><a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></p:bgPr></p:bg>"#;
        let fixture = Fixture::from_xml(
            &slide_xml_with("", background, ""),
            &layout_xml_with("", "", "", ""),
            &master_xml_with("", "", ""),
        );

        let resolved = fixture.context().resolve_slide((40.0, 20.0)).unwrap();
        assert!(matches!(
            resolved.background,
            Some(ResolvedBackground::Paint(Paint::Linear { .. }))
        ));
        assert!(resolved.diagnostics.is_empty());
    }

    #[test]
    fn background_reference_resolves_phclr_through_the_master_colour_map() {
        let master = master_xml_with(
            r#"<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>"#,
            "",
            "",
        )
        .replace("bg1=\"lt1\"", "bg1=\"accent6\"");
        let fixture = Fixture::from_xml(
            &slide_xml_with("", "", ""),
            &layout_xml_with("", "", "", ""),
            &master,
        );
        let context = ResolveCtx::new(
            &fixture.theme,
            fixture.master.color_map.clone(),
            &fixture.master,
            &fixture.layout,
            &fixture.slide,
            &fixture.default_text_style,
        );

        let resolved = context.resolve_slide((40.0, 20.0)).unwrap();
        assert_eq!(
            resolved.background,
            Some(ResolvedBackground::Paint(Paint::Solid(Color::from_hex(
                "4EA72E",
            ))))
        );
        assert!(resolved.diagnostics.is_empty());

        let mut transformed = Fixture::from_xml(
            &slide_xml_with("", "", ""),
            &layout_xml_with("", "", "", ""),
            &master_xml_with(
                r#"<p:bg><p:bgRef idx="1001"><a:srgbClr val="C0504D"><a:tint val="60000"/></a:srgbClr></p:bgRef></p:bg>"#,
                "",
                "",
            ),
        );
        transformed
            .theme
            .theme_elements
            .format_scheme
            .background_fill_styles[0] = oxml_drawing::fill::Fill::from_xml(
            br#"<a:solidFill><a:schemeClr val="phClr"><a:shade val="70000"/></a:schemeClr></a:solidFill>"#,
        )
        .unwrap();

        let transformed = transformed.context().resolve_slide((40.0, 20.0)).unwrap();
        assert_eq!(
            transformed.background,
            Some(ResolvedBackground::Paint(Paint::Solid(Color::from_hex(
                "BC9897",
            ))))
        );
    }

    #[test]
    fn background_picture_relationship_uses_its_producer_scope() {
        let background = r#"<p:bg><p:bgPr><a:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>"#;
        let fixtures = [
            Fixture::from_xml(
                &slide_xml_with("", background, ""),
                &layout_xml_with("", "", "", ""),
                &master_xml_with("", "", ""),
            ),
            Fixture::from_xml(
                &slide_xml_with("", "", ""),
                &layout_xml_with("", background, "", ""),
                &master_xml_with("", "", ""),
            ),
            Fixture::from_xml(
                &slide_xml_with("", "", ""),
                &layout_xml_with("", "", "", ""),
                &master_xml_with(background, "", ""),
            ),
        ];
        let media = ScopedMediaIds {
            slide: HashMap::from([("rId9".to_owned(), MediaId(1))]),
            layout: HashMap::from([("rId9".to_owned(), MediaId(2))]),
            master: HashMap::from([("rId9".to_owned(), MediaId(3))]),
            ..ScopedMediaIds::default()
        };

        for (fixture, expected) in fixtures.iter().zip([MediaId(1), MediaId(2), MediaId(3)]) {
            let resolved = fixture
                .context()
                .resolve_slide_with_media((40.0, 20.0), &media)
                .unwrap();
            let Some(ResolvedBackground::Image(image)) = resolved.background else {
                panic!("expected source-scoped background image");
            };
            assert_eq!(image.media, expected);
            assert!(resolved.diagnostics.is_empty());
        }
    }

    #[test]
    fn master_relationship_cannot_satisfy_a_theme_background_blip() {
        let mut fixture = Fixture::from_xml(
            &slide_xml_with("", "", ""),
            &layout_xml_with("", "", "", ""),
            &master_xml_with(
                r#"<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>"#,
                "",
                "",
            ),
        );
        fixture
            .theme
            .theme_elements
            .format_scheme
            .background_fill_styles[0] = Fill::from_xml(
            br#"<a:blipFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>"#,
        )
        .unwrap();
        let media = ScopedMediaIds {
            master: HashMap::from([("rId9".to_owned(), MediaId(9))]),
            ..ScopedMediaIds::default()
        };

        let resolved = fixture
            .context()
            .resolve_slide_with_media((40.0, 20.0), &media)
            .unwrap();
        assert!(resolved.background.is_none());
        assert_eq!(
            resolved.diagnostics,
            [Diagnostic {
                message: "theme-referenced background picture requires theme relationship scope"
                    .to_owned(),
            }]
        );
    }

    #[test]
    fn background_picture_preserves_crop_stretch_and_tile_placement() {
        let tile = r#"<p:bg><p:bgPr><a:blipFill dpi="144" rotWithShape="0"><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:srcRect l="10000"/><a:tile tx="12700" ty="-25400" sx="50000" sy="200000" flip="xy" algn="br"/></a:blipFill></p:bgPr></p:bg>"#;
        let stretch = r#"<p:bg><p:bgPr><a:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId2"/><a:stretch><a:fillRect l="10000" t="20000" r="30000" b="40000"/></a:stretch></a:blipFill></p:bgPr></p:bg>"#;
        let media = ScopedMediaIds {
            slide: HashMap::from([
                ("rId1".to_owned(), MediaId(1)),
                ("rId2".to_owned(), MediaId(2)),
            ]),
            ..ScopedMediaIds::default()
        };

        let tile_fixture = Fixture::from_xml(
            &slide_xml_with("", tile, ""),
            &layout_xml_with("", "", "", ""),
            &master_xml_with("", "", ""),
        );
        let resolved = tile_fixture
            .context()
            .resolve_slide_with_media((40.0, 20.0), &media)
            .unwrap();
        let Some(ResolvedBackground::Image(image)) = resolved.background else {
            panic!("expected tiled background image");
        };
        assert_eq!(image.src_rect.unwrap().left, 0.1);
        assert_eq!(image.dpi, Some(144.0));
        assert!(!image.rotate_with_shape);
        let ResolvedImagePlacement::Tile(tile) = image.placement else {
            panic!("expected tile placement");
        };
        assert_eq!(tile.translation, Point { x: 1.0, y: -2.0 });
        assert_eq!((tile.scale_x, tile.scale_y), (0.5, 2.0));
        assert_eq!(tile.flip, ResolvedTileFlip::Both);
        assert_eq!(tile.alignment, ResolvedRectAlignment::BottomRight);

        let stretch_fixture = Fixture::from_xml(
            &slide_xml_with("", stretch, ""),
            &layout_xml_with("", "", "", ""),
            &master_xml_with("", "", ""),
        );
        let resolved = stretch_fixture
            .context()
            .resolve_slide_with_media((40.0, 20.0), &media)
            .unwrap();
        let Some(ResolvedBackground::Image(image)) = resolved.background else {
            panic!("expected stretched background image");
        };
        assert_eq!(
            image.placement,
            ResolvedImagePlacement::Stretch {
                fill_rect: Some(crate::CropRect {
                    left: 0.1,
                    top: 0.2,
                    right: 0.3,
                    bottom: 0.4,
                }),
            }
        );
    }

    #[test]
    fn missing_background_picture_relationship_is_diagnosed() {
        let background = r#"<p:bg><p:bgPr><a:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId404"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>"#;
        let fixture = Fixture::from_xml(
            &slide_xml_with("", background, ""),
            &layout_xml_with("", "", "", ""),
            &master_xml_with("", "", ""),
        );

        let resolved = fixture
            .context()
            .resolve_slide_with_media((40.0, 20.0), &ScopedMediaIds::default())
            .unwrap();
        assert!(resolved.background.is_none());
        assert_eq!(
            resolved.diagnostics,
            [Diagnostic {
                message: "missing slide background image relationship `rId404`".to_owned(),
            }]
        );
    }

    #[test]
    fn unsupported_background_fill_records_a_specific_diagnostic() {
        let fixture = Fixture::from_xml(
            &slide_xml_with(
                "",
                r#"<p:bg><p:bgPr><a:pattFill prst="pct5"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></p:bgPr></p:bg>"#,
                "",
            ),
            &layout_xml_with("", "", "", ""),
            &master_xml_with("", "", ""),
        );

        let resolved = fixture.context().resolve_slide((40.0, 20.0)).unwrap();
        assert!(resolved.background.is_none());
        assert_eq!(
            resolved.diagnostics,
            vec![oxml_layout::Diagnostic {
                message: "unsupported background pattern fill".to_owned(),
            }]
        );

        let group_fill = Fixture::from_xml(
            &slide_xml_with("", r#"<p:bg><p:bgPr><a:grpFill/></p:bgPr></p:bg>"#, ""),
            &layout_xml_with("", "", "", ""),
            &master_xml_with("", "", ""),
        )
        .context()
        .resolve_slide((40.0, 20.0))
        .unwrap();
        assert!(group_fill.background.is_none());
        assert_eq!(
            group_fill.diagnostics,
            vec![oxml_layout::Diagnostic {
                message: "unsupported background group fill".to_owned(),
            }]
        );
    }

    #[test]
    fn show_master_shapes_suppresses_the_owned_pass() {
        let master_suppressed = Fixture::from_xml(
            &slide_xml_with("", "", &shape_with_text(None, None, "slide")),
            &layout_xml_with(
                "showMasterSp=\"0\"",
                "",
                &shape_with_text(None, None, "layout"),
                "",
            ),
            &master_xml_with("", &shape_with_text(None, None, "master"), ""),
        );
        let layout_suppressed = Fixture::from_xml(
            &slide_xml_with(
                "showMasterSp=\"0\"",
                "",
                &shape_with_text(None, None, "slide"),
            ),
            &layout_xml_with("", "", &shape_with_text(None, None, "layout"), ""),
            &master_xml_with("", &shape_with_text(None, None, "master"), ""),
        );

        let master_suppressed_texts: Vec<_> = master_suppressed
            .context()
            .flatten()
            .iter()
            .filter_map(item_text)
            .collect();
        let layout_suppressed_texts: Vec<_> = layout_suppressed
            .context()
            .flatten()
            .iter()
            .filter_map(item_text)
            .collect();

        assert_eq!(master_suppressed_texts, ["layout", "slide"]);
        assert_eq!(layout_suppressed_texts, ["master", "slide"]);
    }

    #[test]
    fn slide_placeholder_suppresses_layout_and_master_matches() {
        let fixture = Fixture::from_xml(
            &slide_xml_with(
                "",
                "",
                &shape_with_text(Some("ftr"), Some(9), "slide footer"),
            ),
            &layout_xml_with(
                "",
                "",
                &shape_with_text(Some("ftr"), Some(9), "layout footer"),
                "<p:hf/>",
            ),
            &master_xml_with(
                "",
                &shape_with_text(Some("ftr"), Some(9), "master footer"),
                "<p:hf/>",
            ),
        );

        let texts: Vec<_> = fixture
            .context()
            .flatten()
            .iter()
            .filter_map(item_text)
            .collect();

        assert_eq!(texts, ["slide footer"]);
    }

    #[test]
    fn latent_placeholders_obey_header_footer_flags() {
        let latent_shapes = [
            shape_with_text(Some("dt"), Some(1), "date"),
            shape_with_text(Some("ftr"), Some(2), "footer"),
            shape_with_text(Some("sldNum"), Some(3), "number"),
        ]
        .join("");
        let fixture = Fixture::from_xml(
            &slide_xml_with("", "", ""),
            &layout_xml_with("", "", &latent_shapes, "<p:hf dt=\"0\"/>"),
            &master_xml_with("", "", "<p:hf sldNum=\"0\"/>"),
        );

        let texts: Vec<_> = fixture
            .context()
            .flatten()
            .iter()
            .filter_map(item_text)
            .collect();

        assert_eq!(texts, ["footer"]);
    }

    #[test]
    fn absent_header_footer_container_hides_inherited_latent_placeholders() {
        let slide = shape_with_text(Some("ftr"), Some(11), "slide footer");
        let layout = [
            shape_with_text(Some("dt"), Some(10), "layout date"),
            shape_with_text(Some("sldNum"), Some(12), "layout number"),
        ]
        .join("");
        let fixture = Fixture::from_xml(&slide_xml(&slide), &layout_xml(&layout), &master_xml(""));

        let texts = fixture
            .context()
            .flatten()
            .into_iter()
            .filter_map(|item| item_text(&item))
            .collect::<Vec<_>>();

        assert_eq!(texts, ["slide footer"]);
    }

    #[test]
    fn corpus_slide_resolves_without_theme_references() {
        let stats = resolve_pinned_corpus();

        assert_eq!(stats.decks, EXPECTED_CORPUS_DECKS);
        assert!(stats.slides > EXPECTED_CORPUS_DECKS);
        assert_eq!(stats.contextual_errors, 0, "{}", stats.errors.join("\n"));
        assert_eq!(stats.resolved, stats.slides);
        assert_eq!(stats.theme_references, 0);
    }

    #[test]
    fn resolved_contract_contains_no_presentation_or_drawing_types() {
        fn assert_owned_and_static(_: ResolvedSlide) {}

        let fixture = Fixture::new(&shape_with_details(None, None, &transform(0), None), "", "");
        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        assert_owned_and_static(resolved);
        for type_name in [
            std::any::type_name::<ResolvedSlide>(),
            std::any::type_name::<crate::ResolvedShape>(),
            std::any::type_name::<crate::ResolvedContent>(),
        ] {
            assert!(!type_name.contains("rpptx_oxml"));
            assert!(!type_name.contains("oxml_drawing"));
        }
    }

    #[test]
    fn line_end_resolution_keeps_kind_width_and_length() {
        let properties = format!(
            r#"{}<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="25400"><a:solidFill><a:prstClr val="black"/></a:solidFill><a:headEnd type="diamond" w="sm"/><a:tailEnd type="triangle" len="lg"/></a:ln>"#,
            transform(0)
        );
        let fixture = Fixture::new(&shape_with_details(None, None, &properties, None), "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let shape = &resolved.shapes[0];
        assert_eq!(
            shape.head_end,
            Some(ResolvedLineEnd {
                kind: ResolvedLineEndKind::Diamond,
                width: ResolvedLineEndSize::Small,
                length: ResolvedLineEndSize::Medium,
            })
        );
        assert_eq!(
            shape.tail_end,
            Some(ResolvedLineEnd {
                kind: ResolvedLineEndKind::Triangle,
                width: ResolvedLineEndSize::Medium,
                length: ResolvedLineEndSize::Large,
            })
        );
        assert_eq!(shape.line.as_ref().map(|line| line.width), Some(2.0));
    }

    #[test]
    fn same_relationship_id_resolves_to_distinct_media_in_each_source_scope() {
        let fixture = Fixture::new(
            &picture("rId2", "", "", 25_400),
            &picture("rId2", "", "", 12_700),
            &picture("rId2", "", "", 0),
        );
        let media = ScopedMediaIds {
            slide: HashMap::from([("rId2".to_owned(), MediaId(1))]),
            layout: HashMap::from([("rId2".to_owned(), MediaId(2))]),
            master: HashMap::from([("rId2".to_owned(), MediaId(3))]),
            ..ScopedMediaIds::default()
        };

        let resolved = fixture
            .context()
            .resolve_slide_with_media((720.0, 540.0), &media)
            .unwrap();
        let ids = resolved
            .shapes
            .iter()
            .map(|shape| match shape.content {
                ResolvedContent::Image(ref image) => image.media,
                _ => panic!("scoped picture should resolve to image content"),
            })
            .collect::<Vec<_>>();

        assert_eq!(ids, [MediaId(3), MediaId(2), MediaId(1)]);
    }

    #[test]
    fn picture_placeholder_inherits_layout_bounds_and_keeps_slide_media_scope() {
        let slide_picture = r#"<p:pic><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr><p:ph type="pic" idx="1"/></p:nvPr></p:nvPicPr><p:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId2"/><a:srcRect t="10000" b="20000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>"#;
        let layout_placeholder = shape_with_details(
            Some("pic"),
            Some(1),
            r#"<a:xfrm><a:off x="12700" y="25400"/><a:ext cx="38100" cy="50800"/></a:xfrm>"#,
            None,
        );
        let master_placeholder = shape_with_details(
            Some("pic"),
            Some(1),
            r#"<a:xfrm><a:off x="127000" y="254000"/><a:ext cx="381000" cy="508000"/></a:xfrm>"#,
            None,
        );
        let fixture = Fixture::new(slide_picture, &layout_placeholder, &master_placeholder);
        let context = fixture.context();
        let flattened = context.flatten();
        let bounded_picture_sources = flattened
            .iter()
            .filter(|item| {
                matches!(
                    item,
                    FlattenedItem::Shape {
                        child: ShapeTreeChild::Picture(picture),
                        ..
                    } if context.effective_picture_xfrm(picture).is_some()
                )
            })
            .count();
        let media = ScopedMediaIds {
            slide: HashMap::from([("rId2".to_owned(), MediaId(7))]),
            layout: HashMap::from([("rId2".to_owned(), MediaId(8))]),
            master: HashMap::from([("rId2".to_owned(), MediaId(9))]),
            ..ScopedMediaIds::default()
        };

        let resolved = context
            .resolve_slide_with_media((720.0, 540.0), &media)
            .unwrap();
        assert_eq!(bounded_picture_sources, 1);
        assert_eq!(resolved.shapes.len(), bounded_picture_sources);
        let picture = &resolved.shapes[0];
        assert_eq!(
            picture.bounds,
            Rect {
                x: 1.0,
                y: 2.0,
                width: 3.0,
                height: 4.0,
            }
        );
        let ResolvedContent::Image(image) = &picture.content else {
            panic!("picture placeholder did not retain image content");
        };
        assert_eq!(image.media, MediaId(7));
        assert_eq!(image.src_rect.as_ref().map(|crop| crop.top), Some(0.1));
        assert_eq!(image.src_rect.as_ref().map(|crop| crop.bottom), Some(0.2));
        assert_eq!(
            image.placement,
            ResolvedImagePlacement::Stretch {
                fill_rect: Some(crate::CropRect::default()),
            }
        );
        assert!(resolved.diagnostics.is_empty());
    }

    #[test]
    fn direct_shape_picture_fill_uses_its_producer_scope() {
        let fixture = Fixture::new(
            &shape_picture_fill("rId2", 25_400, None),
            &shape_picture_fill("rId2", 12_700, None),
            &shape_picture_fill("rId2", 0, None),
        );
        let media = ScopedMediaIds {
            slide: HashMap::from([("rId2".to_owned(), MediaId(1))]),
            layout: HashMap::from([("rId2".to_owned(), MediaId(2))]),
            master: HashMap::from([("rId2".to_owned(), MediaId(3))]),
            ..ScopedMediaIds::default()
        };

        let resolved = fixture
            .context()
            .resolve_slide_with_media((720.0, 540.0), &media)
            .unwrap();
        let ids = resolved
            .shapes
            .iter()
            .map(|shape| shape.image_fill.as_ref().unwrap().media)
            .collect::<Vec<_>>();

        assert_eq!(ids, [MediaId(3), MediaId(2), MediaId(1)]);
        assert!(resolved.diagnostics.is_empty());
    }

    #[test]
    fn shape_picture_fill_and_text_are_resolved_independently() {
        let shape = shape_picture_fill(
            "rId7",
            0,
            Some(r#"<a:bodyPr/><a:p><a:r><a:t>caption</a:t></a:r></a:p>"#),
        );
        let fixture = Fixture::new(&shape, "", "");
        let media = ScopedMediaIds {
            slide: HashMap::from([("rId7".to_owned(), MediaId(7))]),
            ..ScopedMediaIds::default()
        };

        let resolved = fixture
            .context()
            .resolve_slide_with_media((720.0, 540.0), &media)
            .unwrap();
        let shape = &resolved.shapes[0];
        assert_eq!(shape.image_fill.as_ref().unwrap().media, MediaId(7));
        let ResolvedContent::Text(text) = &shape.content else {
            panic!("picture-filled shape lost its text");
        };
        assert!(matches!(
            &text.paragraphs[0].runs[0],
            ResolvedTextRun::Text { text, .. } if text == "caption"
        ));
        assert!(resolved.diagnostics.is_empty());
    }

    #[test]
    fn picture_fill_opacity_comes_from_alpha_modulation_fix() {
        let opaque = shape_picture_fill("rId7", 0, None);
        let faded = opaque.replace(
            r#"r:embed="rId7"/>"#,
            r#"r:embed="rId7"><a:alphaModFix amt="30000"/></a:blip>"#,
        );
        assert_ne!(faded, opaque);
        let media = ScopedMediaIds {
            slide: HashMap::from([("rId7".to_owned(), MediaId(7))]),
            ..ScopedMediaIds::default()
        };

        for (shape, opacity) in [(opaque, 1.0), (faded, 0.3)] {
            let fixture = Fixture::new(&shape, "", "");
            let resolved = fixture
                .context()
                .resolve_slide_with_media((720.0, 540.0), &media)
                .unwrap();
            let image = resolved.shapes[0].image_fill.as_ref().unwrap();
            assert_eq!(image.opacity, opacity);
            assert!(resolved.diagnostics.is_empty());
        }
    }

    #[test]
    fn theme_shape_picture_fill_cannot_use_a_part_relationship() {
        let mut fixture = Fixture::new(
            &shape_with_details(
                None,
                None,
                &format!(
                    r#"{}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#,
                    transform(0)
                ),
                None,
            )
            .replace(
                "</p:sp>",
                r#"<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style></p:sp>"#,
            ),
            "",
            "",
        );
        fixture.theme.theme_elements.format_scheme.fill_styles[0] = Fill::from_xml(
            br#"<a:blipFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>"#,
        )
        .unwrap();
        let media = ScopedMediaIds {
            slide: HashMap::from([("rId9".to_owned(), MediaId(9))]),
            master: HashMap::from([("rId9".to_owned(), MediaId(99))]),
            ..ScopedMediaIds::default()
        };

        let resolved = fixture
            .context()
            .resolve_slide_with_media((720.0, 540.0), &media)
            .unwrap();

        assert!(resolved.shapes[0].image_fill.is_none());
        assert_eq!(
            resolved.shapes[0].unsupported,
            Some("theme-referenced shape picture fill")
        );
        assert_eq!(
            resolved.diagnostics,
            [Diagnostic {
                message: "theme-referenced shape picture fill requires theme relationship scope"
                    .to_owned(),
            }]
        );
    }

    #[test]
    fn missing_direct_shape_picture_relationship_keeps_text() {
        let shape = shape_picture_fill(
            "rId404",
            0,
            Some(r#"<a:bodyPr/><a:p><a:r><a:t>keep me</a:t></a:r></a:p>"#),
        );
        let resolved = Fixture::new(&shape, "", "")
            .context()
            .resolve_slide_with_media((720.0, 540.0), &ScopedMediaIds::default())
            .unwrap();

        assert!(resolved.shapes[0].image_fill.is_none());
        assert!(matches!(
            resolved.shapes[0].content,
            ResolvedContent::Text(_)
        ));
        assert_eq!(
            resolved.shapes[0].unsupported,
            Some("missing image relationship")
        );
        assert_eq!(
            resolved.diagnostics,
            [Diagnostic {
                message: "missing slide shape image relationship `rId404`".to_owned(),
            }]
        );
    }

    #[test]
    fn picture_model_resolves_to_neutral_stretch_and_tile_placement() {
        let fixture = Fixture::new(
            &format!(
                "{}{}{}{}",
                picture("rId1", "", "", 0),
                picture(
                    "rId2",
                    "dpi=\"144\" rotWithShape=\"0\"",
                    r#"<a:srcRect l="10000"/><a:tile tx="12700" ty="-25400" sx="50000" sy="200000" flip="xy" algn="br"/>"#,
                    12_700,
                ),
                picture(
                    "rId4",
                    "",
                    r#"<a:stretch><a:fillRect l="10000" t="20000" r="30000" b="40000"/></a:stretch>"#,
                    25_400,
                ),
                linked_picture("https://example.invalid/image.png", 38_100),
            ),
            "",
            "",
        );
        let media = ScopedMediaIds {
            slide: HashMap::from([
                ("rId1".to_owned(), MediaId(11)),
                ("rId2".to_owned(), MediaId(12)),
                ("rId4".to_owned(), MediaId(14)),
            ]),
            ..ScopedMediaIds::default()
        };

        let resolved = fixture
            .context()
            .resolve_slide_with_media((720.0, 540.0), &media)
            .unwrap();
        let ResolvedContent::Image(image) = &resolved.shapes[0].content else {
            panic!("default picture should resolve");
        };
        assert_eq!(image.placement, ResolvedImagePlacement::default());
        assert_eq!(image.dpi, None);
        assert!(image.rotate_with_shape);

        let ResolvedContent::Image(image) = &resolved.shapes[1].content else {
            panic!("tile picture should resolve");
        };
        let ResolvedImagePlacement::Tile(tile) = &image.placement else {
            panic!("expected tile placement");
        };
        assert_eq!(image.src_rect.unwrap().left, 0.1);
        assert_eq!(tile.translation, Point { x: 1.0, y: -2.0 });
        assert_eq!((tile.scale_x, tile.scale_y), (0.5, 2.0));
        assert_eq!(tile.flip, ResolvedTileFlip::Both);
        assert_eq!(tile.alignment, ResolvedRectAlignment::BottomRight);
        assert_eq!(image.dpi, Some(144.0));
        assert!(!image.rotate_with_shape);

        let ResolvedContent::Image(image) = &resolved.shapes[2].content else {
            panic!("explicit stretch picture should resolve");
        };
        let ResolvedImagePlacement::Stretch { fill_rect } = &image.placement else {
            panic!("expected stretch placement");
        };
        assert_eq!(
            *fill_rect,
            Some(crate::CropRect {
                left: 0.1,
                top: 0.2,
                right: 0.3,
                bottom: 0.4,
            })
        );
        assert_eq!(resolved.shapes[3].content, ResolvedContent::None);
        assert_eq!(
            resolved.shapes[3].unsupported,
            Some("external picture media")
        );
        assert!(resolved.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .message
                .contains("external picture relationship `https://example.invalid/image.png`")
        }));
    }

    #[test]
    fn transform_resolves_to_points_rotation_and_flips() {
        let fixture = Fixture::new(
            &shape_with_details(
                None,
                None,
                r#"<a:xfrm rot="5400000" flipH="1" flipV="1"><a:off x="12700" y="25400"/><a:ext cx="38100" cy="50800"/></a:xfrm>"#,
                None,
            ),
            "",
            "",
        );

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let shape = &resolved.shapes[0];
        assert_eq!(shape.bounds.x, 1.0);
        assert_eq!(shape.bounds.y, 2.0);
        assert_eq!(shape.bounds.width, 3.0);
        assert_eq!(shape.bounds.height, 4.0);
        assert_eq!(shape.rotation_deg, 90.0);
        assert!(shape.flip_h);
        assert!(shape.flip_v);
    }

    #[test]
    fn custom_geometry_scales_each_path_coordinate_space_to_shape_points() {
        let geometry = r#"<a:custGeom><a:avLst/><a:rect l="5400" t="2700" r="16200" b="8100"/><a:pathLst><a:path w="21600" h="10800"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="21600" y="10800"/></a:lnTo></a:path></a:pathLst></a:custGeom>"#;
        let shape = shape_with_details(
            None,
            None,
            &format!(
                "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"127000\" cy=\"254000\"/></a:xfrm>{geometry}"
            ),
            None,
        );
        let fixture = Fixture::new(&shape, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let ResolvedGeometry::Custom { paths, text_rect } = &resolved.shapes[0].geometry else {
            panic!("expected custom geometry");
        };
        assert_eq!(
            paths[0].commands[1],
            PathCommand::LineTo(oxml_layout::Point { x: 10.0, y: 20.0 })
        );
        assert_eq!(
            *text_rect,
            Some(Rect {
                x: 2.5,
                y: 5.0,
                width: 5.0,
                height: 10.0,
            })
        );
    }

    #[test]
    fn custom_geometry_uses_shape_size_for_omitted_path_dimensions() {
        let omitted_both = r#"<a:custGeom><a:avLst/><a:pathLst><a:path><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="r" y="b"/></a:lnTo></a:path></a:pathLst></a:custGeom>"#;
        let omitted_width = r#"<a:custGeom><a:avLst/><a:pathLst><a:path h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="r" y="100"/></a:lnTo></a:path></a:pathLst></a:custGeom>"#;
        let shapes = [omitted_both, omitted_width]
            .map(|geometry| {
                shape_with_details(
                    None,
                    None,
                    &format!(
                        "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"127000\" cy=\"254000\"/></a:xfrm>{geometry}"
                    ),
                    None,
                )
            })
            .join("");
        let resolved = Fixture::new(&shapes, "", "")
            .context()
            .resolve_slide((720.0, 540.0))
            .unwrap();

        for shape in &resolved.shapes {
            let ResolvedGeometry::Custom { paths, .. } = &shape.geometry else {
                panic!("expected concrete custom geometry");
            };
            assert_eq!(
                paths[0].commands[1],
                PathCommand::LineTo(oxml_layout::Point { x: 10.0, y: 20.0 })
            );
        }
    }

    #[test]
    fn invalid_custom_geometry_retains_a_diagnosed_bounds_fallback() {
        let geometry = r#"<a:custGeom><a:avLst/><a:pathLst><a:path w="1" h="1"><a:moveTo><a:pt x="missing" y="0"/></a:moveTo></a:path></a:pathLst></a:custGeom>"#;
        let shape = shape_with_details(None, None, &format!("{}{}", transform(0), geometry), None);
        let fixture = Fixture::new(&shape, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        assert_eq!(
            resolved.shapes[0].geometry,
            ResolvedGeometry::BoundsFallback
        );
        assert_eq!(
            resolved.shapes[0].unsupported,
            Some("custom geometry evaluation")
        );
        assert!(
            resolved
                .diagnostics
                .iter()
                .any(|diagnostic| { diagnostic.message.contains("unknown guide: missing") })
        );
    }

    #[test]
    fn connector_presets_reuse_geometry_and_preserve_line_ends() {
        let line = r#"<a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:headEnd type="triangle"/></a:ln>"#;
        let fixture = Fixture::new(
            &[
                connector("line", 127_000, 254_000, "", line),
                connector("straightConnector1", 127_000, 254_000, "", line),
                connector(
                    "bentConnector3",
                    127_000,
                    254_000,
                    r#"<a:gd name="adj1" fmla="val 25000"/>"#,
                    line,
                ),
                connector(
                    "curvedConnector3",
                    127_000,
                    254_000,
                    r#"<a:gd name="adj1" fmla="val 50000"/>"#,
                    line,
                ),
            ]
            .join(""),
            "",
            "",
        );

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();

        assert_eq!(resolved.shapes.len(), 4);
        assert!(resolved.diagnostics.is_empty());
        for shape in &resolved.shapes {
            assert_eq!(shape.unsupported, None);
            assert!(shape.line.is_some());
            assert_eq!(
                shape.head_end.as_ref().map(|end| end.kind),
                Some(ResolvedLineEndKind::Triangle)
            );
        }
        for index in [0, 1] {
            let ResolvedGeometry::Custom { paths, .. } = &resolved.shapes[index].geometry else {
                panic!("expected straight connector geometry");
            };
            assert_eq!(
                paths[0].commands,
                [
                    PathCommand::MoveTo(Point { x: 0.0, y: 0.0 }),
                    PathCommand::LineTo(Point { x: 10.0, y: 20.0 }),
                ]
            );
        }
        let ResolvedGeometry::Custom { paths, .. } = &resolved.shapes[2].geometry else {
            panic!("expected bent connector geometry");
        };
        assert_eq!(
            paths[0].commands,
            [
                PathCommand::MoveTo(Point { x: 0.0, y: 0.0 }),
                PathCommand::LineTo(Point { x: 2.5, y: 0.0 }),
                PathCommand::LineTo(Point { x: 2.5, y: 20.0 }),
                PathCommand::LineTo(Point { x: 10.0, y: 20.0 }),
            ]
        );
        let ResolvedGeometry::Custom { paths, .. } = &resolved.shapes[3].geometry else {
            panic!("expected curved connector geometry");
        };
        assert!(matches!(
            paths[0].commands.as_slice(),
            [
                PathCommand::MoveTo(Point { x: 0.0, y: 0.0 }),
                PathCommand::CurveTo { to, .. },
                PathCommand::CurveTo { .. }
            ] if *to == Point { x: 5.0, y: 10.0 }
        ));
    }

    #[test]
    fn connector_custom_geometry_reuses_the_shared_path_evaluator() {
        let connector = r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="254000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="50"/></a:lnTo><a:lnTo><a:pt x="100" y="50"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></p:spPr></p:cxnSp>"#;
        let resolved = Fixture::new(connector, "", "")
            .context()
            .resolve_slide((720.0, 540.0))
            .unwrap();

        assert!(resolved.diagnostics.is_empty());
        assert_eq!(resolved.shapes[0].unsupported, None);
        let ResolvedGeometry::Custom { paths, .. } = &resolved.shapes[0].geometry else {
            panic!("expected concrete connector custom geometry");
        };
        assert_eq!(
            paths[0].commands,
            [
                PathCommand::MoveTo(Point { x: 0.0, y: 0.0 }),
                PathCommand::LineTo(Point { x: 0.0, y: 10.0 }),
                PathCommand::LineTo(Point { x: 10.0, y: 10.0 }),
                PathCommand::LineTo(Point { x: 10.0, y: 20.0 }),
            ]
        );
    }

    #[test]
    fn zero_extent_and_unknown_connector_geometry_keep_finite_visible_results() {
        let line = r#"<a:ln><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln>"#;
        let fixture = Fixture::new(
            &format!(
                "{}{}",
                connector("straightConnector1", 0, 127_000, "", line),
                connector("futureConnector", 127_000, 254_000, "", line),
            ),
            "",
            "",
        );

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();

        let ResolvedGeometry::Custom { paths, .. } = &resolved.shapes[0].geometry else {
            panic!("zero-width connector should still evaluate: {:?}", resolved);
        };
        assert_eq!(
            paths[0].commands,
            [
                PathCommand::MoveTo(Point { x: 0.0, y: 0.0 }),
                PathCommand::LineTo(Point { x: 0.0, y: 10.0 }),
            ]
        );
        assert!(paths[0].commands.iter().all(|command| match command {
            PathCommand::MoveTo(point) | PathCommand::LineTo(point) =>
                point.x.is_finite() && point.y.is_finite(),
            _ => true,
        }));
        assert_eq!(
            resolved.shapes[1].geometry,
            ResolvedGeometry::BoundsFallback
        );
        assert_eq!(
            resolved.shapes[1].unsupported,
            Some("unknown connector preset geometry")
        );
        assert!(resolved.diagnostics.iter().any(|diagnostic| {
            diagnostic.message.contains("futureConnector")
                && diagnostic.message.contains("retained as shape bounds")
        }));
    }

    #[test]
    fn connector_without_a_direct_line_keeps_a_diagnosed_visible_default() {
        let fixture = Fixture::new(&connector("line", 127_000, 254_000, "", ""), "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let connector = &resolved.shapes[0];

        assert!(matches!(
            connector.geometry,
            ResolvedGeometry::Custom { .. }
        ));
        assert_eq!(connector.unsupported, Some("connector line style"));
        let line = connector.line.as_ref().expect("visible default line");
        assert_eq!(line.paint, Paint::Solid(Color::BLACK));
        assert_eq!(line.width, 1.0);
        assert_eq!(
            resolved.diagnostics,
            [Diagnostic {
                message: "unsupported connector line style retained as visible default".to_owned(),
            }]
        );
    }

    #[test]
    fn unknown_preset_keeps_bounds_text_and_diagnostic() {
        let shape = format!(
            "<p:sp><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>{}<a:prstGeom prst=\"futureBurst\"/></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>keep this text</a:t></a:r></a:p></p:txBody></p:sp>",
            transform(0)
        );
        let fixture = Fixture::new(&shape, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let shape = &resolved.shapes[0];
        assert_eq!(shape.geometry, ResolvedGeometry::BoundsFallback);
        assert_eq!(shape.bounds.width, 100.0 / 12_700.0);
        assert_eq!(shape.bounds.height, 100.0 / 12_700.0);
        assert_eq!(shape.unsupported, Some("unknown preset geometry"));
        let ResolvedContent::Text(text) = &shape.content else {
            panic!("unknown preset lost its text body");
        };
        assert!(matches!(
            &text.paragraphs[0].runs[0],
            ResolvedTextRun::Text { text, .. } if text == "keep this text"
        ));
        assert!(resolved.diagnostics.iter().any(|diagnostic| {
            diagnostic.message.contains("unknown preset geometry")
                && diagnostic.message.contains("futureBurst")
        }));
    }

    #[test]
    fn text_box_without_geometry_does_not_render_a_fallback_border() {
        let shape = shape_with_details(None, None, &transform(0), Some("<a:bodyPr/>"));
        let fixture = Fixture::new(&shape, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();

        assert_eq!(resolved.shapes[0].geometry, ResolvedGeometry::Rectangle);
        assert_eq!(resolved.shapes[0].unsupported, None);
        assert!(resolved.diagnostics.is_empty());
    }

    #[test]
    fn empty_paragraph_end_properties_use_the_normal_character_cascade() {
        let shape = shape_with_details(
            None,
            None,
            &format!(
                r#"{}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#,
                transform(0)
            ),
            Some(
                r#"<a:bodyPr/><a:lstStyle><a:defPPr><a:defRPr sz="1600" cap="all"><a:latin typeface="Carlito"/></a:defRPr></a:defPPr></a:lstStyle>"#,
            ),
        )
        .replace(
            "<a:p/>",
            r#"<a:p><a:pPr><a:defRPr b="1"/></a:pPr><a:endParaRPr sz="3200" i="1"/></a:p>"#,
        );
        let fixture = Fixture::new(&shape, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let ResolvedContent::Text(body) = &resolved.shapes[0].content else {
            panic!("expected resolved text")
        };
        let paragraph = &body.paragraphs[0];

        assert!(paragraph.runs.is_empty());
        assert_eq!(paragraph.end_style.font_size, Some(32.0));
        assert!(paragraph.end_style.bold);
        assert!(paragraph.end_style.italic);
        assert!(paragraph.end_style.all_caps);
        assert_eq!(
            paragraph.end_style.latin_typeface.as_deref(),
            Some("Carlito")
        );
    }

    #[test]
    fn preset_black_and_white_resolve_to_concrete_paint() {
        let black = format!(
            "{}<a:solidFill><a:prstClr val=\"black\"/></a:solidFill>",
            transform(0)
        );
        let white = format!(
            "{}<a:solidFill><a:prstClr val=\"white\"/></a:solidFill>",
            transform(100)
        );
        let fixture = Fixture::new(
            &format!(
                "{}{}",
                shape_with_details(None, None, &black, None),
                shape_with_details(None, None, &white, None)
            ),
            "",
            "",
        );

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        assert_eq!(resolved.shapes[0].fill, Some(Paint::Solid(Color::BLACK)));
        assert_eq!(resolved.shapes[1].fill, Some(Paint::Solid(Color::WHITE)));
    }

    #[test]
    fn group_coordinate_scale_changes_bounds_not_stroke_width() {
        let leaf = shape_with_details(
            None,
            None,
            r#"<a:xfrm><a:off x="12700" y="25400"/><a:ext cx="25400" cy="38100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>"#,
            None,
        );
        let group = format!(
            r#"<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm><a:off x="127000" y="254000"/><a:ext cx="254000" cy="762000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="254000"/></a:xfrm></p:grpSpPr>{leaf}</p:grpSp>"#
        );
        let fixture = Fixture::new(&group, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let shape = &resolved.shapes[0];
        assert_eq!(
            shape.bounds,
            Rect {
                x: 2.0,
                y: 6.0,
                width: 4.0,
                height: 9.0,
            }
        );
        assert_eq!(shape.line.as_ref().map(|line| line.width), Some(0.75));
        assert_eq!(
            shape.group_transform,
            Transform {
                e: 10.0,
                f: 20.0,
                ..Transform::IDENTITY
            }
        );
    }

    #[test]
    fn grouped_text_with_a_visible_stroke_keeps_path_scale_and_stroke_width_invariant() {
        let leaf = shape_with_details(
            None,
            None,
            r#"<a:xfrm><a:off x="12700" y="25400"/><a:ext cx="25400" cy="38100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>"#,
            Some(r#"<a:bodyPr/><a:p><a:r><a:t>visible outline</a:t></a:r></a:p>"#),
        );
        let group = format!(
            r#"<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm><a:off x="127000" y="254000"/><a:ext cx="254000" cy="762000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="254000"/></a:xfrm></p:grpSpPr>{leaf}</p:grpSp>"#
        );
        let fixture = Fixture::new(&group, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let shape = &resolved.shapes[0];
        assert_eq!(
            shape.bounds,
            Rect {
                x: 2.0,
                y: 6.0,
                width: 4.0,
                height: 9.0,
            }
        );
        assert_eq!(shape.line.as_ref().map(|line| line.width), Some(0.75));
        assert_eq!(
            shape.group_transform,
            Transform {
                e: 10.0,
                f: 20.0,
                ..Transform::IDENTITY
            }
        );
        assert!(matches!(shape.content, ResolvedContent::Text(_)));
    }

    #[test]
    fn nested_group_coordinate_mappings_apply_inner_before_outer() {
        let leaf = shape_with_details(
            None,
            None,
            r#"<a:xfrm><a:off x="12700" y="25400"/><a:ext cx="38100" cy="50800"/></a:xfrm>"#,
            None,
        );
        let inner = format!(
            r#"<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm><a:off x="12700" y="25400"/><a:ext cx="254000" cy="381000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm></p:grpSpPr>{leaf}</p:grpSp>"#
        );
        let outer = format!(
            r#"<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm><a:off x="127000" y="254000"/><a:ext cx="508000" cy="635000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm></p:grpSpPr>{inner}</p:grpSp>"#
        );
        let fixture = Fixture::new(&outer, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let shape = &resolved.shapes[0];
        assert_eq!(
            shape.bounds,
            Rect {
                x: 8.0,
                y: 30.0,
                width: 24.0,
                height: 60.0,
            }
        );
        assert_eq!(
            shape.group_transform,
            Transform {
                e: 14.0,
                f: 30.0,
                ..Transform::IDENTITY
            }
        );
    }

    #[test]
    fn zero_group_child_extent_is_finite_and_diagnosed_once() {
        let leaf = shape_with_details(
            None,
            None,
            r#"<a:xfrm><a:off x="76200" y="101600"/><a:ext cx="12700" cy="12700"/></a:xfrm>"#,
            None,
        );
        let group = format!(
            r#"<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm><a:off x="127000" y="254000"/><a:ext cx="254000" cy="381000"/><a:chOff x="63500" y="88900"/><a:chExt cx="0" cy="127000"/></a:xfrm></p:grpSpPr>{leaf}</p:grpSp>"#
        );
        let fixture = Fixture::new(&group, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let shape = &resolved.shapes[0];
        assert_eq!(
            shape.bounds,
            Rect {
                x: 6.0,
                y: 24.0,
                width: 1.0,
                height: 3.0,
            }
        );
        assert!(shape.bounds.x.is_finite());
        assert!(shape.bounds.y.is_finite());
        assert_eq!(
            shape.group_transform,
            Transform {
                e: 5.0,
                f: -1.0,
                ..Transform::IDENTITY
            }
        );
        assert_eq!(
            resolved
                .diagnostics
                .iter()
                .filter(|diagnostic| diagnostic.message.contains("group child extent"))
                .count(),
            1
        );
    }

    #[test]
    fn unsupported_sheared_group_mapping_is_diagnosed_without_non_finite_values() {
        let leaf = shape_with_details(None, None, &transform(0), None);
        let inner = format!(
            r#"<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm rot="2700000"><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm></p:grpSpPr>{leaf}</p:grpSp>"#
        );
        let outer = format!(
            r#"<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="254000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm></p:grpSpPr>{inner}</p:grpSp>"#
        );
        let fixture = Fixture::new(&outer, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let transform = resolved.shapes[0].group_transform;
        assert!(
            [
                transform.a,
                transform.b,
                transform.c,
                transform.d,
                transform.e,
                transform.f,
            ]
            .into_iter()
            .all(f64::is_finite)
        );
        assert!(resolved.diagnostics.iter().any(|diagnostic| {
            diagnostic.message
                == "unsupported sheared or singular group transform retained as affine fallback"
        }));
    }

    #[test]
    fn flattened_group_children_keep_sibling_order() {
        let group = format!(
            r#"<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm></p:grpSpPr>{}{}</p:grpSp>"#,
            shape_with_text(None, None, "inside one"),
            shape_with_text(None, None, "inside two")
        );
        let fixture = Fixture::new(
            &format!(
                "{}{}{}",
                shape_with_text(None, None, "before"),
                group,
                shape_with_text(None, None, "after")
            ),
            "",
            "",
        );

        assert_eq!(
            fixture
                .context()
                .flatten()
                .iter()
                .filter_map(item_text)
                .collect::<Vec<_>>(),
            ["before", "inside one", "inside two", "after"]
        );
    }

    #[test]
    fn paragraph_spacing_and_bullet_size_are_concrete() {
        let text = r#"<a:bodyPr/><a:p><a:pPr><a:lnSpc><a:spcPct val="120000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPts val="300"/></a:spcAft><a:buSzPct val="125000"/><a:buChar char="*"/></a:pPr><a:r><a:t>spaced</a:t></a:r></a:p>"#;
        let fixture = Fixture::new(
            &shape_with_details(None, None, &transform(0), Some(text)),
            "",
            "",
        );

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let ResolvedContent::Text(body) = &resolved.shapes[0].content else {
            panic!("expected text");
        };
        let paragraph = &body.paragraphs[0];
        assert_eq!(
            paragraph.line_spacing,
            Some(ResolvedTextSpacing::Percent(1.2))
        );
        assert_eq!(
            paragraph.space_before,
            Some(ResolvedTextSpacing::Points(6.0))
        );
        assert_eq!(
            paragraph.space_after,
            Some(ResolvedTextSpacing::Points(3.0))
        );
        assert!(matches!(
            paragraph.bullet,
            Some(ResolvedBullet::Character {
                size: Some(ResolvedBulletSize::Percent(1.25)),
                ..
            })
        ));
    }

    #[test]
    fn explicit_rtl_paragraph_direction_reaches_resolved_layout() {
        let text = r#"<a:bodyPr/><a:p><a:pPr rtl="1"/><a:r><a:t>123</a:t></a:r></a:p>"#;
        let fixture = Fixture::new(
            &shape_with_details(None, None, &transform(0), Some(text)),
            "",
            "",
        );

        let (resolved, directions) = fixture
            .context()
            .resolve_slide_inner((720.0, 540.0), None, None, None, None)
            .unwrap();
        let ResolvedContent::Text(body) = &resolved.shapes[0].content else {
            panic!("expected text");
        };
        assert_eq!(body.paragraphs[0].runs.len(), 1);
        assert_eq!(directions[0][0].len(), body.paragraphs.len());
        assert_eq!(directions[0][0][0], oxml_layout::TextDirection::RightToLeft);
    }

    #[test]
    fn legacy_resolved_paragraph_literal_remains_source_compatible() {
        let paragraph = ResolvedParagraph {
            level: 0,
            left_margin: 0.0,
            right_margin: 0.0,
            indent: 0.0,
            alignment: ParagraphAlignment::Left,
            line_spacing: None,
            space_before: None,
            space_after: None,
            bullet: None,
            end_style: ResolvedRunStyle::default(),
            runs: Vec::new(),
        };

        assert_eq!(paragraph, ResolvedParagraph::default());
    }

    #[test]
    fn auto_number_bullet_keeps_independently_inherited_style() {
        let text = r#"<a:bodyPr/><a:lstStyle><a:lvl1pPr><a:buClr><a:srgbClr val="123456"/></a:buClr><a:buSzPct val="125000"/><a:buFont typeface="Wingdings"/><a:buChar char="*"/></a:lvl1pPr></a:lstStyle><a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr><a:r><a:t>numbered</a:t></a:r></a:p>"#;
        let fixture = Fixture::new(
            &shape_with_details(None, None, &transform(0), Some(text)),
            "",
            "",
        );

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let ResolvedContent::Text(body) = &resolved.shapes[0].content else {
            panic!("expected text");
        };
        assert_eq!(
            body.paragraphs[0].bullet,
            Some(ResolvedBullet::AutoNumber {
                scheme: "arabicPeriod".to_owned(),
                start_at: 3,
                font: Some("Wingdings".to_owned()),
                color: Some(Color {
                    r: 0x12 as f64 / 255.0,
                    g: 0x34 as f64 / 255.0,
                    b: 0x56 as f64 / 255.0,
                    a: 1.0,
                }),
                size: Some(ResolvedBulletSize::Percent(1.25)),
            })
        );
    }

    #[test]
    fn table_cells_resolve_body_and_paragraph_properties() {
        let frame = r#"<p:graphicFrame><p:nvGraphicFramePr/><p:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid><a:gridCol w="127000"/></a:tblGrid><a:tr h="127000"><a:tc><a:txBody><a:bodyPr lIns="12700" anchor="b" vert="vert"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr marL="25400" algn="ctr"/><a:r><a:t>cell</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
        let fixture = Fixture::new(frame, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let ResolvedContent::Table(table) = &resolved.shapes[0].content else {
            panic!("expected table");
        };
        let body = table.rows[0].cells[0].text.as_ref().unwrap();
        assert_eq!(body.insets.left, 1.0);
        assert_eq!(body.anchor, ResolvedTextAnchor::Bottom);
        assert_eq!(body.vertical, TextDirection::Vertical);
        assert_eq!(body.autofit, ResolvedAutofit::None);
        assert_eq!(body.paragraphs[0].left_margin, 2.0);
        assert_eq!(
            body.paragraphs[0].alignment,
            crate::ParagraphAlignment::Center
        );
        assert!(resolved.diagnostics.iter().any(|diagnostic| {
            diagnostic.message == "table cell autofit is unsupported and was ignored"
        }));
    }

    #[test]
    fn table_style_regions_resolve_in_documented_precedence() {
        let frame = r#"<p:graphicFrame><p:nvGraphicFramePr/><p:xfrm><a:off x="0" y="0"/><a:ext cx="254000" cy="254000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" firstCol="1" bandRow="1"><a:tableStyleId>style</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="127000"/><a:gridCol w="127000"/></a:tblGrid><a:tr h="127000"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>styled</a:t></a:r></a:p></a:txBody></a:tc><a:tc/></a:tr><a:tr h="127000"><a:tc/><a:tc/></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
        let styles = oxml_drawing::table::CT_TableStyleList::from_xml(br#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="style"><a:tblStyle styleId="style" styleName="Style"><a:wholeTbl><a:tcStyle><a:solidFill><a:srgbClr val="111111"/></a:solidFill></a:tcStyle></a:wholeTbl><a:band1H><a:tcStyle><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:tcStyle></a:band1H><a:firstRow><a:tcTxStyle b="on" i="1"><a:fontRef idx="major"/><a:srgbClr val="AABBCC"/></a:tcTxStyle><a:tcStyle><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:tcStyle></a:firstRow><a:nwCell><a:tcStyle><a:solidFill><a:srgbClr val="444444"/></a:solidFill></a:tcStyle></a:nwCell></a:tblStyle></a:tblStyleLst>"#).unwrap();
        let fixture = Fixture::new(frame, "", "");

        let resolved = fixture
            .context()
            .with_table_styles(&styles)
            .resolve_slide((20.0, 20.0))
            .unwrap();
        let ResolvedContent::Table(table) = &resolved.shapes[0].content else {
            panic!("expected table");
        };

        assert_eq!(
            table.rows[0].cells[0].fill,
            Some(Paint::Solid(Color::from_hex("444444")))
        );
        assert_eq!(
            table.rows[0].cells[1].fill,
            Some(Paint::Solid(Color::from_hex("333333")))
        );
        assert_eq!(
            table.rows[1].cells[0].fill,
            Some(Paint::Solid(Color::from_hex("222222")))
        );
        let ResolvedTextRun::Text { style, .. } =
            &table.rows[0].cells[0].text.as_ref().unwrap().paragraphs[0].runs[0]
        else {
            panic!("expected styled text run")
        };
        assert!(style.bold && style.italic);
        assert_eq!(style.fill, Some(Paint::Solid(Color::from_hex("AABBCC"))));
        assert_eq!(style.latin_typeface.as_deref(), Some("Aptos Display"));
    }

    #[test]
    fn direct_table_text_and_inside_borders_override_style_defaults() {
        let frame = r#"<p:graphicFrame><p:nvGraphicFramePr/><p:xfrm><a:off x="0" y="0"/><a:ext cx="254000" cy="127000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr><a:tableStyleId>style</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="127000"/><a:gridCol w="127000"/></a:tblGrid><a:tr h="127000"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr b="0"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:latin typeface="Courier New"/></a:rPr><a:t>direct</a:t></a:r></a:p></a:txBody></a:tc><a:tc/></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
        let styles = oxml_drawing::table::CT_TableStyleList::from_xml(br#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="style"><a:tblStyle styleId="style" styleName="Style"><a:wholeTbl><a:tcTxStyle b="on"><a:fontRef idx="major"/><a:srgbClr val="FF0000"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:left><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:left><a:right><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:right><a:insideV><a:ln w="25400"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></a:insideV></a:tcBdr></a:tcStyle></a:wholeTbl></a:tblStyle></a:tblStyleLst>"#).unwrap();
        let fixture = Fixture::new(frame, "", "");

        let resolved = fixture
            .context()
            .with_table_styles(&styles)
            .resolve_slide((20.0, 10.0))
            .unwrap();
        let ResolvedContent::Table(table) = &resolved.shapes[0].content else {
            panic!("expected table");
        };
        let ResolvedTextRun::Text { style, .. } =
            &table.rows[0].cells[0].text.as_ref().unwrap().paragraphs[0].runs[0]
        else {
            panic!("expected direct run");
        };
        assert!(!style.bold);
        assert_eq!(style.fill, Some(Paint::Solid(Color::from_hex("00FF00"))));
        assert_eq!(style.latin_typeface.as_deref(), Some("Courier New"));

        let first = &table.rows[0].cells[0];
        let second = &table.rows[0].cells[1];
        assert_eq!(
            first.left.as_ref().unwrap().stroke.as_ref().unwrap().paint,
            Paint::Solid(Color::from_hex("FF0000"))
        );
        assert_eq!(
            first.right.as_ref().unwrap().stroke.as_ref().unwrap().paint,
            Paint::Solid(Color::from_hex("0000FF"))
        );
        assert_eq!(
            second.left.as_ref().unwrap().stroke.as_ref().unwrap().paint,
            Paint::Solid(Color::from_hex("0000FF"))
        );
        assert_eq!(
            second
                .right
                .as_ref()
                .unwrap()
                .stroke
                .as_ref()
                .unwrap()
                .paint,
            Paint::Solid(Color::from_hex("FF0000"))
        );
    }

    #[test]
    fn table_corners_require_both_flags_and_unsupported_fills_are_diagnosed() {
        let frame = r#"<p:graphicFrame><p:nvGraphicFramePr/><p:xfrm><a:off x="0" y="0"/><a:ext cx="254000" cy="127000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"><a:tableStyleId>style</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="127000"/><a:gridCol w="127000"/></a:tblGrid><a:tr h="127000"><a:tc/><a:tc><a:tcPr><a:lnL><a:pattFill prst="pct5"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></a:lnL><a:pattFill prst="pct5"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
        let styles = oxml_drawing::table::CT_TableStyleList::from_xml(br#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="style"><a:tblStyle styleId="style" styleName="Style"><a:wholeTbl><a:tcStyle><a:solidFill><a:srgbClr val="111111"/></a:solidFill></a:tcStyle></a:wholeTbl><a:firstRow><a:tcStyle><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:tcStyle></a:firstRow><a:nwCell><a:tcStyle><a:solidFill><a:srgbClr val="444444"/></a:solidFill></a:tcStyle></a:nwCell></a:tblStyle></a:tblStyleLst>"#).unwrap();
        let fixture = Fixture::new(frame, "", "");

        let resolved = fixture
            .context()
            .with_table_styles(&styles)
            .resolve_slide((20.0, 10.0))
            .unwrap();
        let ResolvedContent::Table(table) = &resolved.shapes[0].content else {
            panic!("expected table");
        };

        assert_eq!(
            table.rows[0].cells[0].fill,
            Some(Paint::Solid(Color::from_hex("333333")))
        );
        assert_eq!(table.rows[0].cells[1].fill, None);
        assert!(
            table.rows[0].cells[1]
                .left
                .as_ref()
                .is_some_and(|border| border.stroke.is_none())
        );
        assert!(resolved.diagnostics.iter().any(|diagnostic| {
            diagnostic.message == "unsupported table cell pattern fill was ignored"
        }));
    }

    #[test]
    fn table_cell_autofit_is_ignored_and_records_a_diagnostic() {
        let frame = r#"<p:graphicFrame><p:nvGraphicFramePr/><p:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid><a:gridCol w="127000"/></a:tblGrid><a:tr h="127000"><a:tc><a:txBody><a:bodyPr><a:spAutoFit/></a:bodyPr><a:p><a:r><a:t>cell</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
        let fixture = Fixture::new(frame, "", "");

        let resolved = fixture.context().resolve_slide((10.0, 10.0)).unwrap();
        let ResolvedContent::Table(table) = &resolved.shapes[0].content else {
            panic!("expected table");
        };

        assert_eq!(
            table.rows[0].cells[0].text.as_ref().unwrap().autofit,
            ResolvedAutofit::None
        );
        assert!(
            resolved
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message
                    == "table cell autofit is unsupported and was ignored")
        );
    }

    #[test]
    fn linear_gradient_scaled_modes_cover_non_square_bounds() {
        let gradient = |scaled: &str| {
            format!(
                r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="508000" cy="254000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="2700000"{scaled}/></a:gradFill>"#
            )
        };
        let slide = [
            shape_with_details(None, None, &gradient(""), None),
            shape_with_details(None, None, &gradient(r#" scaled="0""#), None),
            shape_with_details(None, None, &gradient(r#" scaled="1""#), None),
        ]
        .join("");
        let resolved = Fixture::new(&slide, "", "")
            .context()
            .resolve_slide((720.0, 540.0))
            .unwrap();

        for shape in &resolved.shapes[..2] {
            let Some(Paint::Linear { start, end, .. }) = shape.fill.as_ref() else {
                panic!("expected unscaled linear gradient");
            };
            assert_close(start.x, 5.0);
            assert_close(start.y, -5.0);
            assert_close(end.x, 35.0);
            assert_close(end.y, 25.0);
        }
        let Some(Paint::Linear { start, end, .. }) = resolved.shapes[2].fill.as_ref() else {
            panic!("expected scaled linear gradient");
        };
        assert_close(start.x, 0.0);
        assert_close(start.y, 0.0);
        assert_close(end.x, 40.0);
        assert_close(end.y, 20.0);
        assert!(resolved.diagnostics.is_empty());
    }

    #[test]
    fn unsupported_gradient_variants_have_distinct_diagnostics() {
        let variants = [
            (
                r#"flip="x""#,
                r#"<a:lin ang="0"/>"#,
                "horizontal gradient flip",
            ),
            (
                r#"flip="y""#,
                r#"<a:lin ang="0"/>"#,
                "vertical gradient flip",
            ),
            (
                r#"flip="xy""#,
                r#"<a:lin ang="0"/>"#,
                "horizontal and vertical gradient flip",
            ),
            (
                "",
                r#"<a:lin ang="0"/><a:tileRect l="10000"/>"#,
                "gradient tile rectangle",
            ),
            ("", r#"<a:path path="circle"/>"#, "circle path gradient"),
            ("", r#"<a:path path="rect"/>"#, "rectangle path gradient"),
            ("", r#"<a:path path="shape"/>"#, "shape path gradient"),
        ];

        for (attributes, geometry, expected) in variants {
            let details = format!(
                r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/></a:xfrm><a:gradFill {attributes}><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst>{geometry}</a:gradFill>"#
            );
            let fixture = Fixture::new(&shape_with_details(None, None, &details, None), "", "");
            let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();

            assert!(resolved.shapes[0].fill.is_none(), "{expected}");
            assert_eq!(resolved.shapes[0].unsupported, Some(expected));
            assert!(resolved.diagnostics[0].message.contains(expected));
        }
    }

    #[test]
    fn independent_gradient_is_concrete_only_without_an_effective_shape_rotation() {
        let geometry = r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#;
        let gradient = r#"<a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>"#;
        let unrotated = shape_with_details(
            None,
            None,
            &format!(
                r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/></a:xfrm>{geometry}{gradient}"#
            ),
            None,
        );
        let rotated = shape_with_details(
            None,
            None,
            &format!(
                r#"<a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/></a:xfrm>{geometry}{gradient}"#
            ),
            None,
        );
        let shapes = format!("{unrotated}{rotated}");
        let resolved = Fixture::new(&shapes, "", "")
            .context()
            .resolve_slide((720.0, 540.0))
            .unwrap();

        assert!(matches!(
            resolved.shapes[0].fill,
            Some(Paint::Linear { .. })
        ));
        assert_eq!(resolved.shapes[0].unsupported, None);
        assert_eq!(resolved.shapes[1].fill, None);
        assert_eq!(
            resolved.shapes[1].unsupported,
            Some("gradient independent of shape rotation")
        );
        assert_eq!(resolved.diagnostics.len(), 1);

        let grouped = format!(
            r#"<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm></p:grpSpPr>{unrotated}</p:grpSp>"#
        );
        let resolved = Fixture::new(&grouped, "", "")
            .context()
            .resolve_slide((720.0, 540.0))
            .unwrap();
        assert_eq!(resolved.shapes[0].fill, None);
        assert_eq!(
            resolved.shapes[0].unsupported,
            Some("gradient independent of shape rotation")
        );
    }

    #[test]
    fn theme_style_resolves_to_concrete_paint_line_and_shadow() {
        let shape = format!(
            "<p:sp><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>{}<a:effectLst><a:outerShdw blurRad=\"12700\" dist=\"25400\" dir=\"0\"><a:schemeClr val=\"accent2\"/></a:outerShdw></a:effectLst></p:spPr><p:style><a:lnRef idx=\"1\"><a:srgbClr val=\"112233\"/></a:lnRef><a:fillRef idx=\"1\"><a:srgbClr val=\"445566\"/></a:fillRef><a:effectRef idx=\"1\"><a:srgbClr val=\"778899\"/></a:effectRef><a:fontRef idx=\"minor\"><a:srgbClr val=\"000000\"/></a:fontRef></p:style></p:sp>",
            transform(0)
        );
        let fixture = Fixture::new(&shape, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let shape = &resolved.shapes[0];
        assert!(matches!(shape.fill, Some(Paint::Solid(_))));
        assert!(shape.line.is_some());
        assert!(matches!(shape.shadow, Some(Effect::OuterShadow { .. })));
    }

    #[test]
    fn unsupported_content_keeps_bounds_and_diagnostic() {
        let frame = r#"<p:graphicFrame><p:nvGraphicFramePr/><p:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="254000"/></p:xfrm><a:graphic><a:graphicData uri="urn:unsupported"><a:unknown/></a:graphicData></a:graphic></p:graphicFrame>"#;
        let fixture = Fixture::new(frame, "", "");

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();

        assert_eq!(resolved.shapes.len(), 1);
        assert_eq!(resolved.shapes[0].bounds.width, 10.0);
        assert_eq!(resolved.shapes[0].bounds.height, 20.0);
        assert_eq!(
            resolved.shapes[0].geometry,
            ResolvedGeometry::BoundsFallback
        );
        assert_eq!(
            resolved.shapes[0].unsupported,
            Some("unknown graphic frame")
        );
        assert_eq!(resolved.diagnostics.len(), 1);
    }

    #[test]
    fn ole_png_previews_use_their_producer_scope_and_graphic_frame_transform() {
        let fixture = Fixture::new(
            &ole_frame("rId7", 25_400, 9_999_999),
            &ole_frame("rId7", 12_700, 8_888_888),
            &ole_frame("rId7", 0, 7_777_777),
        );
        let media = ScopedMediaIds {
            slide: HashMap::from([("rId7".to_owned(), MediaId(1))]),
            layout: HashMap::from([("rId7".to_owned(), MediaId(2))]),
            master: HashMap::from([("rId7".to_owned(), MediaId(3))]),
            media_content_types: HashMap::from([
                (MediaId(1), "image/png".to_owned()),
                (MediaId(2), "image/png".to_owned()),
                (MediaId(3), "image/png".to_owned()),
            ]),
        };

        let resolved = fixture
            .context()
            .resolve_slide_with_media((720.0, 540.0), &media)
            .unwrap();
        let ids = resolved
            .shapes
            .iter()
            .map(|shape| match &shape.content {
                ResolvedContent::Image(image) => image.media,
                _ => panic!("OLE PNG preview did not resolve as image content"),
            })
            .collect::<Vec<_>>();
        let positions = resolved
            .shapes
            .iter()
            .map(|shape| shape.bounds.x)
            .collect::<Vec<_>>();

        assert_eq!(ids, [MediaId(3), MediaId(2), MediaId(1)]);
        assert_eq!(positions, [0.0, 1.0, 2.0]);
        assert!(resolved.shapes.iter().all(|shape| {
            shape.geometry == ResolvedGeometry::Rectangle
                && shape.unsupported == Some("embedded OLE interactivity")
        }));
        assert_eq!(
            resolved
                .diagnostics
                .iter()
                .map(|diagnostic| diagnostic.message.as_str())
                .collect::<Vec<_>>(),
            vec![
                "OLE object rendered as a static PNG preview. Embedded OLE interactivity is not rendered";
                3
            ]
        );
    }

    #[test]
    fn ole_preview_requires_embedded_resolved_png_media() {
        let children = [
            ole_frame_without_preview(0),
            ole_frame_with_blip(
                r#"r:link="https://example.invalid/preview.png""#,
                12_700,
                12_700,
            ),
            ole_frame("rIdMissing", 25_400, 25_400),
            ole_frame("rIdWmf", 38_100, 38_100),
        ]
        .join("");
        let media = ScopedMediaIds {
            slide: HashMap::from([("rIdWmf".to_owned(), MediaId(4))]),
            media_content_types: HashMap::from([(MediaId(4), "image/x-wmf".to_owned())]),
            ..ScopedMediaIds::default()
        };

        let resolved = Fixture::new(&children, "", "")
            .context()
            .resolve_slide_with_media((720.0, 540.0), &media)
            .unwrap();

        assert_eq!(resolved.shapes.len(), 4);
        assert!(resolved.shapes.iter().all(|shape| {
            shape.content == ResolvedContent::None
                && shape.geometry == ResolvedGeometry::BoundsFallback
                && shape.unsupported == Some("OLE")
        }));
        assert_eq!(
            resolved.diagnostics,
            vec![
                Diagnostic {
                    message: "unsupported OLE content retained as bounds".to_owned(),
                };
                4
            ]
        );
    }

    #[test]
    fn resolved_shapes_follow_the_flattened_order() {
        let master = shape_with_details(None, None, &transform(12_700), None);
        let layout = shape_with_details(None, None, &transform(25_400), None);
        let slide = shape_with_details(None, None, &transform(38_100), None);
        let fixture = Fixture::new(&slide, &layout, &master);

        let resolved = fixture.context().resolve_slide((720.0, 540.0)).unwrap();
        let x = resolved
            .shapes
            .iter()
            .map(|shape| shape.bounds.x)
            .collect::<Vec<_>>();
        assert_eq!(x, [1.0, 2.0, 3.0]);
    }

    #[test]
    fn all_corpus_slides_resolve_without_panics() {
        let stats = resolve_pinned_corpus();

        assert_eq!(stats.decks, EXPECTED_CORPUS_DECKS);
        assert_eq!(stats.resolved + stats.contextual_errors, stats.slides);
    }

    #[test]
    fn all_corpus_preset_geometries_evaluate_or_fallback() {
        let stats = resolve_pinned_corpus();

        assert_eq!(stats.decks, EXPECTED_CORPUS_DECKS);
        assert_eq!(stats.contextual_errors, 0, "{}", stats.errors.join("\n"));
        assert!(
            stats.preset_inputs > 0,
            "corpus exercised no preset geometry"
        );
        assert_eq!(stats.preset_errors, 0, "{}", stats.errors.join("\n"));
        assert_eq!(
            stats.preset_inputs,
            stats.preset_evaluated + stats.preset_unknown,
            "not every corpus preset produced geometry or a named unknown fallback"
        );
        assert_eq!(stats.preset_fallbacks, stats.preset_unknown);
    }

    #[derive(Default)]
    struct CorpusResolveStats {
        decks: usize,
        slides: usize,
        resolved: usize,
        contextual_errors: usize,
        theme_references: usize,
        preset_inputs: usize,
        preset_evaluated: usize,
        preset_unknown: usize,
        preset_errors: usize,
        preset_fallbacks: usize,
        errors: Vec<String>,
    }

    fn resolve_pinned_corpus() -> CorpusResolveStats {
        let corpus = corpus_dir();
        assert!(
            corpus.is_dir(),
            "the required pinned corpus is missing at {}",
            corpus.display()
        );
        let entries = include_str!("../../../scripts/pptx-corpus-manifest.tsv")
            .lines()
            .skip(1)
            .map(|line| line.split('\t').next().unwrap())
            .collect::<Vec<_>>();
        assert_eq!(entries.len(), EXPECTED_CORPUS_DECKS);

        let mut stats = CorpusResolveStats::default();
        for entry in entries {
            let path = corpus.join(entry);
            assert!(path.is_file(), "missing pinned deck {}", path.display());
            let package = OpcPackage::open(&path)
                .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
            let presentation_part = package
                .main_document_part()
                .unwrap_or_else(|| panic!("{}: no presentation part", path.display()));
            let presentation = CT_Presentation::from_xml(part(&package, &presentation_part))
                .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
            let default_text_style = presentation.default_text_style.unwrap_or_default();
            let size = presentation.slide_size.map_or((720.0, 540.0), |size| {
                (
                    super::emu_to_points(size.cx.0),
                    super::emu_to_points(size.cy.0),
                )
            });
            let presentation_rels = package
                .get_part_rels(&presentation_part)
                .unwrap_or_else(|| panic!("{}: no presentation relationships", path.display()));

            for slide_id in presentation.slide_ids {
                stats.slides += 1;
                let slide_rel = presentation_rels
                    .get_by_id(&slide_id.relationship_id)
                    .unwrap_or_else(|| {
                        panic!(
                            "{}: missing slide relationship {}",
                            path.display(),
                            slide_id.relationship_id
                        )
                    });
                assert_eq!(slide_rel.rel_type, rel_types::SLIDE, "{}", path.display());
                let slide_part =
                    OpcPackage::resolve_rel_target(&presentation_part, &slide_rel.target);
                let layout_part =
                    related_part(&package, &slide_part, rel_types::SLIDE_LAYOUT, &path);
                let master_part =
                    related_part(&package, &layout_part, rel_types::SLIDE_MASTER, &path);
                let theme_part = related_part(&package, &master_part, rel_types::THEME, &path);
                let slide = CT_Slide::from_xml(part(&package, &slide_part))
                    .unwrap_or_else(|error| panic!("{} {slide_part}: {error}", path.display()));
                let layout = CT_SlideLayout::from_xml(part(&package, &layout_part))
                    .unwrap_or_else(|error| panic!("{} {layout_part}: {error}", path.display()));
                let master = CT_SlideMaster::from_xml(part(&package, &master_part))
                    .unwrap_or_else(|error| panic!("{} {master_part}: {error}", path.display()));
                let theme = CT_OfficeStyleSheet::from_xml(part(&package, &theme_part))
                    .unwrap_or_else(|error| panic!("{} {theme_part}: {error}", path.display()));
                let color_map = effective_corpus_color_map(&master, &layout, &slide);
                let context = ResolveCtx::new(
                    &theme,
                    color_map,
                    &master,
                    &layout,
                    &slide,
                    &default_text_style,
                );
                for item in context.flatten() {
                    let FlattenedItem::Shape {
                        child: ShapeTreeChild::Shape(shape),
                        ..
                    } = item
                    else {
                        continue;
                    };
                    if shape.shape_properties.custom_geometry.is_some() {
                        continue;
                    }
                    let Some(preset) = shape.shape_properties.preset_geometry.as_ref() else {
                        continue;
                    };
                    let Some((bounds, _, _, _)) =
                        transform_values(context.effective_xfrm(shape).as_ref())
                    else {
                        continue;
                    };
                    stats.preset_inputs += 1;
                    match context.concrete_preset_geometry(preset, (bounds.width, bounds.height)) {
                        Ok(Some(_)) => stats.preset_evaluated += 1,
                        Ok(None) => stats.preset_unknown += 1,
                        Err(error) => {
                            stats.preset_errors += 1;
                            stats.errors.push(format!(
                                "{} {slide_part}: preset {}: {error}",
                                path.display(),
                                preset.preset
                            ));
                        }
                    }
                }
                let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                    context.resolve_slide(size)
                }))
                .unwrap_or_else(|_| panic!("{} {slide_part}: resolver panicked", path.display()));
                match result {
                    Ok(resolved) => {
                        let debug = format!("{resolved:?}");
                        stats.theme_references += ["schemeClr", "sysClr", "scrgbClr", "prstClr"]
                            .into_iter()
                            .filter(|marker| debug.contains(marker))
                            .count();
                        stats.preset_fallbacks += resolved
                            .shapes
                            .iter()
                            .filter(|shape| {
                                matches!(
                                    shape.unsupported,
                                    Some("unknown preset geometry" | "preset geometry evaluation")
                                )
                            })
                            .count();
                        stats.resolved += 1;
                    }
                    Err(error) => {
                        stats.contextual_errors += 1;
                        stats
                            .errors
                            .push(format!("{} {slide_part}: {error}", path.display()));
                    }
                }
            }
            stats.decks += 1;
        }
        stats
    }

    fn corpus_dir() -> PathBuf {
        std::env::var_os("RDOCX_PPTX_CORPUS_DIR")
            .map(PathBuf::from)
            .unwrap_or_else(|| workspace_root().join("corpus/pptx"))
    }

    fn workspace_root() -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .and_then(Path::parent)
            .unwrap()
            .to_path_buf()
    }

    fn part<'a>(package: &'a OpcPackage, part_name: &str) -> &'a [u8] {
        package
            .get_part(part_name)
            .unwrap_or_else(|| panic!("missing package part {part_name}"))
    }

    fn related_part(
        package: &OpcPackage,
        source_part: &str,
        relationship_type: &str,
        deck: &Path,
    ) -> String {
        let relationship = package
            .get_part_rels(source_part)
            .and_then(|relationships| relationships.get_by_type(relationship_type))
            .unwrap_or_else(|| {
                panic!(
                    "{} {source_part}: missing relationship {relationship_type}",
                    deck.display()
                )
            });
        OpcPackage::resolve_rel_target(source_part, &relationship.target)
    }

    fn effective_corpus_color_map(
        master: &CT_SlideMaster,
        layout: &CT_SlideLayout,
        slide: &CT_Slide,
    ) -> ColorMap {
        for override_value in [
            slide.color_map_override.as_ref(),
            layout.color_map_override.as_ref(),
        ]
        .into_iter()
        .flatten()
        {
            if let ColorMapOverrideKind::Override(map) = &override_value.kind {
                return map.clone();
            }
        }
        master.color_map.clone()
    }

    fn slide_xml(children: &str) -> String {
        slide_xml_with("", "", children)
    }

    fn layout_xml(children: &str) -> String {
        layout_xml_with("", "", children, "")
    }

    fn master_xml(children: &str) -> String {
        master_xml_with("", children, "")
    }

    fn slide_xml_with(attributes: &str, background: &str, children: &str) -> String {
        format!(
            "<p:sld xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\" xmlns:mc=\"{MC_NS}\" {attributes}><p:cSld>{background}{}</p:cSld></p:sld>",
            shape_tree(children)
        )
    }

    fn layout_xml_with(
        attributes: &str,
        background: &str,
        children: &str,
        header_footer: &str,
    ) -> String {
        format!(
            "<p:sldLayout xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\" xmlns:mc=\"{MC_NS}\" {attributes}><p:cSld>{background}{}</p:cSld>{header_footer}</p:sldLayout>",
            shape_tree(children)
        )
    }

    fn master_xml_with(background: &str, children: &str, header_footer: &str) -> String {
        format!(
            "<p:sldMaster xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\" xmlns:mc=\"{MC_NS}\"><p:cSld>{background}{}</p:cSld><p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>{header_footer}</p:sldMaster>",
            shape_tree(children)
        )
    }

    fn shape_tree(children: &str) -> String {
        format!("<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{children}</p:spTree>")
    }

    fn shape(ph_type: Option<&str>, idx: Option<u32>) -> String {
        shape_with_details(ph_type, idx, "", None)
    }

    fn shape_with_text(ph_type: Option<&str>, idx: Option<u32>, text: &str) -> String {
        let placeholder = if ph_type.is_none() && idx.is_none() {
            String::new()
        } else {
            let ph_type = ph_type.map_or_else(String::new, |value| format!(" type=\"{value}\""));
            let idx = idx.map_or_else(String::new, |value| format!(" idx=\"{value}\""));
            format!("<p:ph{ph_type}{idx}/>")
        };
        format!(
            "<p:sp><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"
        )
    }

    fn item_text(item: &FlattenedItem<'_>) -> Option<String> {
        let FlattenedItem::Shape {
            child: ShapeTreeChild::Shape(shape),
            ..
        } = item
        else {
            return None;
        };
        shape.text_body.as_ref().map(|body| body.plain_text())
    }

    fn background_source(fixture: &Fixture) -> BackgroundSource {
        fixture.context().effective_background().unwrap().source
    }

    #[test]
    fn same_relationship_id_resolves_hyperlink_in_its_shape_source_scope() {
        let linked_shape = |label: &str, x: i64| {
            shape_with_details(None, None, &transform(x), Some("<a:bodyPr/>"))
                .replace(
                    "<a:p/>",
                    &format!(r#"<a:p><a:r><a:rPr><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId7"/></a:rPr><a:t>{label}</a:t></a:r></a:p>"#),
                )
        };
        let fixture = Fixture::new(
            &linked_shape("slide", 200),
            &linked_shape("layout", 100),
            &linked_shape("master", 0),
        );
        let hyperlinks = ScopedHyperlinkTargets {
            slide: HashMap::from([("rId7".to_owned(), "https://slide.example".to_owned())]),
            layout: HashMap::from([("rId7".to_owned(), "https://layout.example".to_owned())]),
            master: HashMap::from([("rId7".to_owned(), "https://master.example".to_owned())]),
        };

        let resolved = fixture
            .context()
            .resolve_slide_with_resources((720.0, 540.0), &ScopedMediaIds::default(), &hyperlinks)
            .expect("resolve source-scoped hyperlinks");
        let targets = resolved
            .shapes
            .iter()
            .filter_map(|shape| match &shape.content {
                ResolvedContent::Text(body) => body.paragraphs[0].runs.first(),
                _ => None,
            })
            .filter_map(|run| match run {
                ResolvedTextRun::Text { style, .. } => style.hyperlink_url.as_deref(),
                _ => None,
            })
            .collect::<Vec<_>>();

        assert_eq!(
            targets,
            vec![
                "https://master.example",
                "https://layout.example",
                "https://slide.example"
            ]
        );
    }

    #[test]
    fn missing_hyperlink_relationship_keeps_text_and_records_diagnostic() {
        let shape = shape_with_details(None, None, &transform(0), Some("<a:bodyPr/>"))
            .replace(
                "<a:p/>",
                r#"<a:p><a:r><a:rPr><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId404"/></a:rPr><a:t>missing</a:t></a:r><a:r><a:rPr><a:hlinkClick action="ppaction://macro"/></a:rPr><a:t> action</a:t></a:r><a:r><a:rPr><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId8" action="ppaction://hlinksldjump"/></a:rPr><a:t> internal</a:t></a:r></a:p>"#,
            );
        let resolved = Fixture::new(&shape, "", "")
            .context()
            .resolve_slide_with_resources(
                (720.0, 540.0),
                &ScopedMediaIds::default(),
                &ScopedHyperlinkTargets::default(),
            )
            .expect("resolve unsupported hyperlinks without dropping text");
        let ResolvedContent::Text(body) = &resolved.shapes[0].content else {
            panic!("linked text should remain visible")
        };

        assert_eq!(
            body.paragraphs[0]
                .runs
                .iter()
                .filter_map(|run| match run {
                    ResolvedTextRun::Text { text, .. } => Some(text.as_str()),
                    _ => None,
                })
                .collect::<String>(),
            "missing action internal"
        );
        let diagnostics = resolved
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.as_str())
            .collect::<Vec<_>>();
        for expected in [
            "missing slide hyperlink relationship `rId404`",
            "unsupported slide hyperlink action `ppaction://macro`",
            "unsupported slide hyperlink action `ppaction://hlinksldjump`",
        ] {
            assert!(diagnostics.contains(&expected), "missing `{expected}`");
        }
    }

    #[test]
    fn untyped_slide_number_placeholder_uses_the_current_page_number() {
        let slide_shape = shape_with_details(None, Some(4), &transform(0), Some("<a:bodyPr/>"))
            .replace(
                "<a:p/>",
                r#"<a:p><a:fld id="{00112233-4455-6677-8899-AABBCCDDEEFF}"><a:t>stored</a:t></a:fld></a:p>"#,
            );
        let layout_shape =
            shape_with_details(Some("sldNum"), Some(4), &transform(0), Some("<a:bodyPr/>"));
        let resolved = Fixture::new(&slide_shape, &layout_shape, "")
            .context()
            .resolve_slide((720.0, 540.0))
            .expect("resolve inherited slide-number placeholder");
        let ResolvedContent::Text(body) = &resolved.shapes[0].content else {
            panic!("slide-number field should remain visible")
        };
        let ResolvedTextRun::Field { field_type, .. } = &body.paragraphs[0].runs[0] else {
            panic!("expected a field")
        };

        assert_eq!(field_type.as_deref(), Some("slidenum"));
    }

    #[test]
    fn three_dimensional_chart_uses_cached_image_and_diagnostic() {
        let alternate = chart_alternate("rIdChart", Some("rIdPreview"));
        let fixture = Fixture::new(&alternate, "", "");
        let preview = MediaId(47);
        let media = ScopedMediaIds {
            slide: HashMap::from([("rIdPreview".to_owned(), preview)]),
            media_content_types: HashMap::from([(preview, "image/png".to_owned())]),
            ..ScopedMediaIds::default()
        };
        let charts = ScopedChartResources {
            slide: HashMap::from([(
                "rIdChart".to_owned(),
                ChartResource::Parsed(Box::new(three_dimensional_chart())),
            )]),
            ..ScopedChartResources::default()
        };
        let mut fonts = FontManager::new_deterministic().unwrap();

        let resolved = fixture
            .context()
            .resolve_slide_with_chart_resources(
                (720.0, 540.0),
                &media,
                &ScopedHyperlinkTargets::default(),
                &charts,
                &mut fonts,
            )
            .unwrap();

        assert!(matches!(
            resolved.shapes[0].content,
            ResolvedContent::Image(ref image) if image.media == preview
        ));
        assert!(resolved.shapes[0].unsupported.is_some());
        assert!(resolved.diagnostics.iter().any(|diagnostic| {
            diagnostic.message.contains("unsupported chart")
                && diagnostic.message.contains("cached image")
        }));

        let unsupported_media = ScopedMediaIds {
            slide: HashMap::from([("rIdPreview".to_owned(), preview)]),
            media_content_types: HashMap::from([(preview, "image/x-wmf".to_owned())]),
            ..ScopedMediaIds::default()
        };
        let mut fonts = FontManager::new_deterministic().unwrap();
        let resolved = fixture
            .context()
            .resolve_slide_with_chart_resources(
                (720.0, 540.0),
                &unsupported_media,
                &ScopedHyperlinkTargets::default(),
                &charts,
                &mut fonts,
            )
            .unwrap();
        assert!(matches!(
            resolved.shapes[0].content,
            ResolvedContent::Group(_)
        ));
        assert!(resolved.diagnostics.iter().any(|diagnostic| {
            diagnostic.message.contains("not renderer-compatible")
                && diagnostic.message.contains("unsupported chart")
        }));
    }

    #[test]
    fn same_chart_relationship_id_is_scoped_to_its_source_part() {
        let fixture = Fixture::new(
            &chart_frame_at("rId8", 200),
            &chart_frame_at("rId8", 100),
            &chart_frame("rId8"),
        );
        let slide = ChartResource::MissingTarget("/ppt/charts/slide.xml".to_owned());
        let layout = ChartResource::MissingTarget("/ppt/charts/layout.xml".to_owned());
        let master = ChartResource::MissingTarget("/ppt/charts/master.xml".to_owned());
        let charts = ScopedChartResources {
            slide: HashMap::from([("rId8".to_owned(), slide.clone())]),
            layout: HashMap::from([("rId8".to_owned(), layout.clone())]),
            master: HashMap::from([("rId8".to_owned(), master.clone())]),
        };

        let mut fonts = FontManager::new_deterministic().unwrap();
        let resolved = fixture
            .context()
            .resolve_slide_with_chart_resources(
                (720.0, 540.0),
                &ScopedMediaIds::default(),
                &ScopedHyperlinkTargets::default(),
                &charts,
                &mut fonts,
            )
            .unwrap();
        assert_eq!(resolved.shapes.len(), 3);
        let diagnostics = resolved
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.as_str())
            .collect::<Vec<_>>();
        for (scope, target) in [
            ("master", "/ppt/charts/master.xml"),
            ("layout", "/ppt/charts/layout.xml"),
            ("slide", "/ppt/charts/slide.xml"),
        ] {
            assert!(diagnostics.iter().any(|message| {
                message.contains(scope) && message.contains("rId8") && message.contains(target)
            }));
        }
    }

    #[test]
    fn unsupported_chart_without_preview_keeps_labelled_bounds() {
        let fixture = Fixture::new(&chart_frame("rIdChart"), "", "");
        let charts = ScopedChartResources {
            slide: HashMap::from([(
                "rIdChart".to_owned(),
                ChartResource::Parsed(Box::new(three_dimensional_chart())),
            )]),
            ..ScopedChartResources::default()
        };
        let mut fonts = FontManager::new_deterministic().unwrap();
        let resolved = fixture
            .context()
            .resolve_slide_with_chart_resources(
                (720.0, 540.0),
                &ScopedMediaIds::default(),
                &ScopedHyperlinkTargets::default(),
                &charts,
                &mut fonts,
            )
            .unwrap();

        assert_eq!(
            resolved.shapes[0].geometry,
            ResolvedGeometry::BoundsFallback
        );
        let ResolvedContent::Group(group) = &resolved.shapes[0].content else {
            panic!("unsupported chart should retain a labelled group");
        };
        let mut labels = Vec::new();
        walk(&group.children, &mut |element, _| {
            if let PositionedElement::Text(run) = element {
                labels.push(run.text.clone());
            }
        });
        assert_eq!(labels, vec!["Unsupported chart"]);
    }

    #[test]
    fn missing_or_external_chart_relationship_is_contextual() {
        let children = format!(
            "{}{}{}",
            chart_frame("missingChart"),
            chart_frame_at("externalChart", 200),
            chart_frame_at("missingTarget", 400)
        );
        let fixture = Fixture::new(&children, "", "");
        let charts = ScopedChartResources {
            slide: HashMap::from([
                (
                    "externalChart".to_owned(),
                    ChartResource::External("https://example.invalid/chart.xml".to_owned()),
                ),
                (
                    "missingTarget".to_owned(),
                    ChartResource::MissingTarget("/ppt/charts/missing.xml".to_owned()),
                ),
            ]),
            ..ScopedChartResources::default()
        };
        let mut fonts = FontManager::new_deterministic().unwrap();
        let resolved = fixture
            .context()
            .resolve_slide_with_chart_resources(
                (720.0, 540.0),
                &ScopedMediaIds::default(),
                &ScopedHyperlinkTargets::default(),
                &charts,
                &mut fonts,
            )
            .unwrap();
        let diagnostics = resolved
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.as_str())
            .collect::<Vec<_>>();

        assert!(
            diagnostics
                .iter()
                .any(|message| { message.contains("slide") && message.contains("missingChart") })
        );
        assert!(diagnostics.iter().any(|message| {
            message.contains("slide")
                && message.contains("externalChart")
                && message.contains("https://example.invalid/chart.xml")
        }));
        assert!(diagnostics.iter().any(|message| {
            message.contains("slide")
                && message.contains("missingTarget")
                && message.contains("/ppt/charts/missing.xml")
        }));
    }

    fn shape_with_details(
        ph_type: Option<&str>,
        idx: Option<u32>,
        shape_properties: &str,
        body_properties: Option<&str>,
    ) -> String {
        let placeholder = if ph_type.is_none() && idx.is_none() {
            String::new()
        } else {
            let ph_type = ph_type.map_or_else(String::new, |value| format!(" type=\"{value}\""));
            let idx = idx.map_or_else(String::new, |value| format!(" idx=\"{value}\""));
            format!("<p:ph{ph_type}{idx}/>")
        };
        let text_body = body_properties.map_or_else(String::new, |body_properties| {
            format!("<p:txBody>{body_properties}<a:p/></p:txBody>")
        });
        format!(
            "<p:sp><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr><p:spPr>{shape_properties}</p:spPr>{text_body}</p:sp>"
        )
    }

    fn chart_frame(relationship_id: &str) -> String {
        chart_frame_at(relationship_id, 0)
    }

    fn chart_frame_at(relationship_id: &str, x: i64) -> String {
        format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr/><p:xfrm><a:off x="{x}" y="0"/><a:ext cx="1270000" cy="762000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{relationship_id}"/></a:graphicData></a:graphic></p:graphicFrame>"#
        )
    }

    fn chart_alternate(relationship_id: &str, preview_id: Option<&str>) -> String {
        let preview = preview_id.map_or_else(String::new, |preview_id| {
            picture(preview_id, "", "<a:stretch><a:fillRect/></a:stretch>", 999)
        });
        format!(
            r#"<mc:AlternateContent><mc:Choice Requires="c">{}</mc:Choice><mc:Fallback>{preview}</mc:Fallback></mc:AlternateContent>"#,
            chart_frame(relationship_id)
        )
    }

    fn three_dimensional_chart() -> CT_ChartSpace {
        CT_ChartSpace::from_xml(
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:bar3DChart><c:barDir val="col"/></c:bar3DChart></c:plotArea></c:chart></c:chartSpace>"#,
        )
        .unwrap()
    }

    fn connector(preset: &str, cx: i64, cy: i64, adjustments: &str, line: &str) -> String {
        format!(
            r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{preset}"><a:avLst>{adjustments}</a:avLst></a:prstGeom>{line}</p:spPr></p:cxnSp>"#
        )
    }

    fn picture(
        relationship_id: &str,
        fill_attributes: &str,
        fill_children: &str,
        x: i64,
    ) -> String {
        format!(
            r#"<p:pic><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill {fill_attributes}><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{relationship_id}"/>{fill_children}</p:blipFill><p:spPr>{}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#,
            transform(x)
        )
    }

    fn shape_picture_fill(relationship_id: &str, x: i64, text_body: Option<&str>) -> String {
        shape_with_details(
            None,
            None,
            &format!(
                r#"{}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{relationship_id}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>"#,
                transform(x)
            ),
            text_body,
        )
    }

    fn linked_picture(url: &str, x: i64) -> String {
        format!(
            r#"<p:pic><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:link="{url}"/></p:blipFill><p:spPr>{}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#,
            transform(x)
        )
    }

    fn ole_frame(relationship_id: &str, frame_x: i64, preview_x: i64) -> String {
        ole_frame_with_blip(
            &format!(r#"r:embed="{relationship_id}""#),
            frame_x,
            preview_x,
        )
    }

    fn ole_frame_with_blip(blip_attribute: &str, frame_x: i64, preview_x: i64) -> String {
        let preview = format!(
            r#"<p:pic><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" {blip_attribute}/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#,
            transform(preview_x)
        );
        ole_frame_payload(frame_x, &preview)
    }

    fn ole_frame_without_preview(frame_x: i64) -> String {
        ole_frame_payload(frame_x, "")
    }

    fn ole_frame_payload(frame_x: i64, preview: &str) -> String {
        format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr/><p:xfrm><a:off x="{frame_x}" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj><p:embed/>{preview}</p:oleObj></a:graphicData></a:graphic></p:graphicFrame>"#
        )
    }

    fn transform(x: i64) -> String {
        format!("<a:xfrm><a:off x=\"{x}\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/></a:xfrm>")
    }
}
