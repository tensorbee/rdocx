use std::collections::HashSet;
use std::fmt::Write as _;

use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64;
use oxml_layout::{
    Color, Effect, FillRule, FontData, FontId, GlyphRun, GroupElement, LayoutResult, LineCap,
    LineJoin, Paint, Path, PathCommand, PathElement, PositionedElement, Rect, Stroke, Transform,
};

/// One stable diagnostic produced while lowering a page to SVG.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SvgDiagnostic {
    /// Location in the layout result that required a fallback or was omitted.
    pub path: String,
    /// Stable description of the lossy conversion.
    pub message: String,
}

/// One self-contained SVG page and its ordered diagnostics.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SvgRenderResult {
    /// Complete UTF-8 SVG document.
    pub svg: String,
    /// Layout diagnostics followed by SVG lowering diagnostics.
    pub diagnostics: Vec<SvgDiagnostic>,
}

/// Lower one laid-out page to a self-contained SVG document.
pub fn render_page(layout: &LayoutResult, page_index: usize) -> Option<SvgRenderResult> {
    let page = layout.pages.get(page_index)?;
    let mut state = SvgState::new(
        layout,
        page.width,
        page.height,
        layout
            .diagnostics
            .iter()
            .enumerate()
            .map(|(index, diagnostic)| SvgDiagnostic {
                path: format!("layout.diagnostics[{index}]"),
                message: diagnostic.message.clone(),
            })
            .collect(),
    );

    let mut used_fonts = Vec::new();
    let mut seen_fonts = HashSet::new();
    collect_used_fonts(&page.elements, &mut seen_fonts, &mut used_fonts);
    state.emit_font_definitions(&used_fonts);

    let mut body = String::new();
    write!(
        body,
        "<rect x=\"0\" y=\"0\" width=\"{}\" height=\"{}\" fill=\"#FFFFFF\"/>",
        number(page.width),
        number(page.height)
    )
    .unwrap();
    if let Some(background) = &page.background {
        let path = format!("pages[{page_index}].background");
        if let Some(attributes) = state.paint_attributes(background, "fill", &path) {
            write!(
                body,
                "<rect x=\"0\" y=\"0\" width=\"{}\" height=\"{}\" {attributes}/>",
                number(page.width),
                number(page.height)
            )
            .unwrap();
        }
    }
    state.emit_elements(
        &page.elements,
        &format!("pages[{page_index}].elements"),
        &mut body,
        Transform::IDENTITY,
    );

    let mut svg = String::new();
    write!(
        svg,
        "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{}pt\" height=\"{}pt\" viewBox=\"0 0 {} {}\">",
        number(page.width),
        number(page.height),
        number(page.width),
        number(page.height)
    )
    .unwrap();
    if !state.defs.is_empty() {
        write!(svg, "<defs>{}</defs>", state.defs).unwrap();
    }
    svg.push_str(&body);
    svg.push_str("</svg>");

    Some(SvgRenderResult {
        svg,
        diagnostics: state.diagnostics,
    })
}

struct SvgState<'a> {
    layout: &'a LayoutResult,
    page_width: f64,
    page_height: f64,
    defs: String,
    diagnostics: Vec<SvgDiagnostic>,
    next_definition: usize,
    diagnosed_fonts: HashSet<FontId>,
}

impl<'a> SvgState<'a> {
    fn new(
        layout: &'a LayoutResult,
        page_width: f64,
        page_height: f64,
        diagnostics: Vec<SvgDiagnostic>,
    ) -> Self {
        Self {
            layout,
            page_width,
            page_height,
            defs: String::new(),
            diagnostics,
            next_definition: 0,
            diagnosed_fonts: HashSet::new(),
        }
    }

    fn definition_id(&mut self) -> String {
        let id = format!("rdocx-def-{}", self.next_definition);
        self.next_definition += 1;
        id
    }

    fn diagnose(&mut self, path: &str, message: &str) {
        self.diagnostics.push(SvgDiagnostic {
            path: path.to_owned(),
            message: message.to_owned(),
        });
    }

    fn emit_font_definitions(&mut self, used_fonts: &[FontId]) {
        if used_fonts.is_empty() {
            return;
        }

        let mut rules = String::new();
        for font_id in used_fonts {
            let Some(font) = self.layout.fonts.iter().find(|font| font.id == *font_id) else {
                continue;
            };
            let Some(data) = standalone_font_data(font) else {
                continue;
            };
            let (mime, format) = if font.data.starts_with(b"ttcf") {
                ("font/ttf", "truetype")
            } else {
                font_content_type(font)
            };

            write!(
                rules,
                "@font-face{{font-family:'rdocx-font-{}';src:url('data:{mime};base64,{}') format('{format}');font-weight:{};font-style:{}}}",
                font.id.0,
                BASE64.encode(data.as_ref()),
                if font.bold && !font.synthetic_bold { "bold" } else { "normal" },
                if font.italic && !font.synthetic_italic { "italic" } else { "normal" }
            )
            .unwrap();
            let settings = variation_settings(font);
            if !settings.is_empty() {
                write!(
                    rules,
                    "text[font-family='rdocx-font-{}']{{font-variation-settings:{}}}",
                    font.id.0, settings
                )
                .unwrap();
            }
        }
        if !rules.is_empty() {
            write!(self.defs, "<style>{rules}</style>").unwrap();
        }
    }

    fn emit_elements(
        &mut self,
        elements: &[PositionedElement],
        path: &str,
        output: &mut String,
        accumulated: Transform,
    ) {
        for (index, element) in elements.iter().enumerate() {
            let element_path = format!("{path}[{index}]");
            match element {
                PositionedElement::Text(run) => self.emit_text(run, &element_path, output, false),
                PositionedElement::MultilingualText(run) => {
                    if !run.is_valid() {
                        self.diagnose(
                            &element_path,
                            "invalid multilingual glyph positioning was omitted from SVG output",
                        );
                        continue;
                    }
                    let diagnostic_count = self.diagnostics.len();
                    self.emit_text(&run.legacy_projection(), &element_path, output, true);
                    if self.diagnostics.len() == diagnostic_count {
                        self.diagnose(
                            &element_path,
                            "multilingual shaping kept searchable text with browser-positioned glyph approximation",
                        );
                    }
                }
                PositionedElement::Line {
                    start,
                    end,
                    width,
                    color,
                    dash_pattern,
                } => {
                    write!(
                        output,
                        "<line x1=\"{}\" y1=\"{}\" x2=\"{}\" y2=\"{}\" stroke=\"{}\" stroke-opacity=\"{}\" stroke-width=\"{}\"",
                        number(start.x),
                        number(start.y),
                        number(end.x),
                        number(end.y),
                        color_hex(*color),
                        number(color.a.clamp(0.0, 1.0)),
                        number(*width)
                    )
                    .unwrap();
                    if let Some((on, off)) = dash_pattern {
                        write!(
                            output,
                            " stroke-dasharray=\"{} {}\"",
                            number(*on),
                            number(*off)
                        )
                        .unwrap();
                    }
                    output.push_str("/>");
                }
                PositionedElement::FilledRect { rect, color } => {
                    write!(
                        output,
                        "<rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" fill=\"{}\" fill-opacity=\"{}\"/>",
                        number(rect.x),
                        number(rect.y),
                        number(rect.width),
                        number(rect.height),
                        color_hex(*color),
                        number(color.a.clamp(0.0, 1.0))
                    )
                    .unwrap();
                }
                PositionedElement::Image { rect, data, .. } => {
                    self.emit_image(*rect, data, &element_path, output)
                }
                PositionedElement::LinkAnnotation { rect, url } => {
                    self.emit_link(*rect, url, &element_path, output)
                }
                PositionedElement::Path(path_element) => {
                    self.emit_path_element(path_element, &element_path, output)
                }
                PositionedElement::Group(group) => {
                    self.emit_group(group, &element_path, output, accumulated)
                }
                PositionedElement::MarkedContent { children, .. } => {
                    output.push_str("<g>");
                    self.emit_elements(
                        children,
                        &format!("{element_path}.children"),
                        output,
                        accumulated,
                    );
                    output.push_str("</g>");
                }
                _ => self.diagnose(
                    &element_path,
                    "unsupported positioned element was omitted from SVG output",
                ),
            }
        }
    }

    fn emit_text(&mut self, run: &GlyphRun, path: &str, output: &mut String, multilingual: bool) {
        if run.text.is_empty() {
            return;
        }
        self.diagnose_font_on_first_use(run.font_id);

        let font = self.layout.fonts.iter().find(|font| font.id == run.font_id);
        write!(
            output,
            "<text xml:space=\"preserve\" y=\"{}\" font-family=\"rdocx-font-{}\" font-size=\"{}\" font-weight=\"{}\" font-style=\"{}\" fill=\"{}\" fill-opacity=\"{}\"",
            number(run.origin.y),
            run.font_id.0,
            number(run.font_size),
            if run.bold && !font.is_some_and(|font| font.synthetic_bold) { "bold" } else { "normal" },
            if run.italic && !font.is_some_and(|font| font.synthetic_italic) { "italic" } else { "normal" },
            color_hex(run.color),
            number(run.color.a.clamp(0.0, 1.0))
        )
        .unwrap();

        if let Some(font) = font {
            let settings = variation_settings(font);
            if !settings.is_empty() || font.synthetic_bold || font.synthetic_italic {
                write!(
                    output,
                    " style=\"font-synthesis:none{}\"",
                    if settings.is_empty() {
                        String::new()
                    } else {
                        format!(";font-variation-settings:{settings}")
                    }
                )
                .unwrap();
            }
            if font.synthetic_bold {
                write!(output, " stroke=\"{}\" stroke-opacity=\"{}\" stroke-width=\"{}\" stroke-linecap=\"round\" stroke-linejoin=\"round\" paint-order=\"fill stroke\"", color_hex(run.color), number(run.color.a.clamp(0.0, 1.0)), run.font_size / 24.0).unwrap();
            }
            if font.synthetic_italic {
                write!(
                    output,
                    " transform=\"matrix(1 0 {} 1 {} 0)\"",
                    -0.212_556_561_670_022_1,
                    run.origin.y * 0.212_556_561_670_022_1
                )
                .unwrap();
            }
        }

        let scalar_count = run.text.chars().count();
        // Equal counts do not prove scalar ownership after GSUB. Admit the
        // scalar x-list only when the selected face's nominal cmap agrees.
        let selected_face = font.and_then(|font| validated_font_face(&font.data, font.face_index));
        let nominal_positions = selected_face.as_ref().is_some_and(|face| {
            run.text.chars().zip(&run.glyph_ids).all(|(scalar, glyph)| {
                face.glyph_index(scalar).unwrap_or(ttf_parser::GlyphId(0)).0 == *glyph
            })
        });
        if scalar_count == run.glyph_ids.len()
            && run.advances.len() == run.glyph_ids.len()
            && nominal_positions
            && run.font_size.is_finite()
            && run.font_size > 0.0
            && run.origin.x.is_finite()
            && run.origin.y.is_finite()
            && run.advances.iter().all(|advance| advance.is_finite())
        {
            let mut x = run.origin.x;
            output.push_str(" x=\"");
            for (index, advance) in run.advances.iter().enumerate() {
                if index != 0 {
                    output.push(' ');
                }
                output.push_str(&number(x));
                x += advance;
            }
            output.push('"');
        } else {
            let advance = run.advances.iter().copied().sum::<f64>().max(0.0);
            write!(
                output,
                " x=\"{}\" textLength=\"{}\" lengthAdjust=\"spacingAndGlyphs\"",
                number(run.origin.x),
                number(advance)
            )
            .unwrap();
            if selected_face.is_some() {
                self.diagnose(
                    path,
                    if multilingual {
                        "multilingual shaping kept searchable text with browser-positioned glyph approximation"
                    } else {
                        "complex shaping kept searchable text with total-advance positioning"
                    },
                );
            }
        }
        let (text, replaced_invalid_xml) = sanitize_xml_text(&run.text);
        if replaced_invalid_xml {
            self.diagnose(
                path,
                "XML-invalid text characters were replaced with U+FFFD",
            );
        }
        write!(output, ">{}</text>", escape_text(&text)).unwrap();
    }

    fn diagnose_font_on_first_use(&mut self, font_id: FontId) {
        if !self.diagnosed_fonts.insert(font_id) {
            return;
        }
        let Some(font) = self.layout.fonts.iter().find(|font| font.id == font_id) else {
            self.diagnose(
                &format!("fonts[{}]", font_id.0),
                "text references font data that is absent from the layout result",
            );
            return;
        };
        if standalone_font_data(font).is_none() {
            self.diagnose(
                &format!("fonts[{}]", font.id.0),
                "selected font face cannot be embedded in searchable SVG",
            );
        }
    }

    fn emit_image(&mut self, rect: Rect, data: &[u8], path: &str, output: &mut String) {
        let Some(mime) = image_mime(data) else {
            self.diagnose(
                path,
                "image bytes are neither PNG nor JPEG and were omitted",
            );
            return;
        };
        write!(
            output,
            "<image x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" preserveAspectRatio=\"none\" href=\"data:{mime};base64,{}\"/>",
            number(rect.x),
            number(rect.y),
            number(rect.width),
            number(rect.height),
            BASE64.encode(data)
        )
        .unwrap();
    }

    fn emit_link(&mut self, rect: Rect, url: &str, path: &str, output: &mut String) {
        if !is_safe_link(url) {
            self.diagnose(path, "active or unsupported link target was omitted");
            return;
        }
        write!(
            output,
            "<a href=\"{}\"><rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" fill=\"none\" pointer-events=\"all\"/></a>",
            escape_attribute(url),
            number(rect.x),
            number(rect.y),
            number(rect.width),
            number(rect.height)
        )
        .unwrap();
    }

    fn emit_path_element(&mut self, element: &PathElement, path: &str, output: &mut String) {
        write!(
            output,
            "<path d=\"{}\" fill-rule=\"{}\"",
            path_data(&element.path),
            fill_rule(element.path.fill_rule)
        )
        .unwrap();
        match &element.fill {
            Some(paint) => {
                if let Some(attributes) =
                    self.paint_attributes(paint, "fill", &format!("{path}.fill"))
                {
                    write!(output, " {attributes}").unwrap();
                } else {
                    output.push_str(" fill=\"none\"");
                }
            }
            None => output.push_str(" fill=\"none\""),
        }
        if let Some(stroke) = &element.stroke {
            self.emit_stroke_attributes(stroke, &format!("{path}.stroke"), output);
        }
        output.push_str("/>");
    }

    fn emit_stroke_attributes(&mut self, stroke: &Stroke, path: &str, output: &mut String) {
        if let Some(attributes) = self.paint_attributes(&stroke.paint, "stroke", path) {
            write!(
                output,
                " {attributes} stroke-width=\"{}\" stroke-linecap=\"{}\" stroke-linejoin=\"{}\"",
                number(stroke.width),
                line_cap(stroke.cap),
                line_join(stroke.join)
            )
            .unwrap();
            if let Some(dash) = &stroke.dash {
                output.push_str(" stroke-dasharray=\"");
                for (index, value) in dash.iter().enumerate() {
                    if index != 0 {
                        output.push(' ');
                    }
                    output.push_str(&number(*value));
                }
                output.push('"');
            }
        } else {
            output.push_str(" stroke=\"none\"");
        }
    }

    fn emit_group(
        &mut self,
        group: &GroupElement,
        path: &str,
        output: &mut String,
        accumulated: Transform,
    ) {
        let mut attributes = String::new();
        let mut effect_attributes = String::new();
        let mut mask_wrappers = Vec::new();
        for (index, effect) in group.effects.iter().enumerate() {
            let Effect::LinearOpacityMask {
                bounds,
                start,
                end,
                stops,
            } = effect
            else {
                continue;
            };
            if !effect.linear_opacity_mask_is_valid() {
                mask_wrappers.push("<g opacity=\"0\">".to_owned());
                self.diagnose(
                    &format!("{path}.effects[{index}]"),
                    "invalid linear opacity mask makes the group transparent",
                );
                continue;
            }
            let gradient_id = self.definition_id();
            let mask_id = self.definition_id();
            write!(self.defs, "<linearGradient id=\"{gradient_id}\" gradientUnits=\"userSpaceOnUse\" x1=\"{}\" y1=\"{}\" x2=\"{}\" y2=\"{}\" spreadMethod=\"pad\">", number(start.x), number(start.y), number(end.x), number(end.y)).unwrap();
            for stop in stops {
                write!(
                    self.defs,
                    "<stop offset=\"{}\" stop-color=\"#FFFFFF\" stop-opacity=\"{}\"/>",
                    number(stop.offset),
                    number(stop.color.a)
                )
                .unwrap();
            }
            write!(self.defs, "</linearGradient><mask id=\"{mask_id}\" maskUnits=\"userSpaceOnUse\" maskContentUnits=\"userSpaceOnUse\" x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" style=\"mask-type:alpha\"><rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" fill=\"url(#{gradient_id})\"/></mask>", number(bounds.x), number(bounds.y), number(bounds.width), number(bounds.height), number(bounds.x), number(bounds.y), number(bounds.width), number(bounds.height)).unwrap();
            mask_wrappers.push(format!("<g mask=\"url(#{mask_id})\">"));
        }

        let group_to_page = group.transform.then(accumulated);
        if !group.transform.is_identity() {
            write!(
                attributes,
                " transform=\"matrix({} {} {} {} {} {})\"",
                number(group.transform.a),
                number(group.transform.b),
                number(group.transform.c),
                number(group.transform.d),
                number(group.transform.e),
                number(group.transform.f)
            )
            .unwrap();
        }
        let clip_id = if let Some(clip) = &group.clip {
            let id = self.definition_id();
            write!(
                self.defs,
                "<clipPath id=\"{id}\" clipPathUnits=\"userSpaceOnUse\"><path d=\"{}\" clip-rule=\"{}\"/></clipPath>",
                path_data(clip),
                fill_rule(clip.fill_rule)
            )
            .unwrap();
            Some(id)
        } else {
            None
        };
        if group.opacity < 1.0 {
            write!(
                attributes,
                " opacity=\"{}\"",
                number(group.opacity.clamp(0.0, 1.0))
            )
            .unwrap();
        }
        if !group.effects.is_empty() {
            let id = self.definition_id();
            let filter_region = group
                .effects
                .iter()
                .any(|effect| matches!(effect, Effect::OuterShadow { .. }))
                .then(|| self.filter_region(group, group_to_page));
            let mut filter = String::new();
            let mut shadows = Vec::new();
            for (index, effect) in group.effects.iter().enumerate() {
                match effect {
                    Effect::OuterShadow {
                        dx,
                        dy,
                        blur,
                        color,
                    } => {
                        let blur_id = format!("{id}-blur-{index}");
                        let offset_id = format!("{id}-offset-{index}");
                        let flood_id = format!("{id}-flood-{index}");
                        let shadow_id = format!("{id}-shadow-{index}");
                        write!(
                            filter,
                            "<feGaussianBlur in=\"SourceAlpha\" stdDeviation=\"{}\" result=\"{blur_id}\"/><feOffset in=\"{blur_id}\" dx=\"{}\" dy=\"{}\" result=\"{offset_id}\"/><feFlood flood-color=\"{}\" flood-opacity=\"{}\" result=\"{flood_id}\"/><feComposite in=\"{flood_id}\" in2=\"{offset_id}\" operator=\"in\" result=\"{shadow_id}\"/>",
                            number((*blur).max(0.0) / 2.0),
                            number(*dx),
                            number(*dy),
                            color_hex(*color),
                            number(color.a.clamp(0.0, 1.0))
                        )
                        .unwrap();
                        shadows.push(shadow_id);
                        if *blur > 0.0
                            && !preserves_isotropic_blur(group_to_page)
                            && matches!(filter_region, Some(Ok(Some(_))))
                        {
                            self.diagnose(
                                &format!("{path}.effects[{index}]"),
                                "non-uniform or skewed transform makes SVG shadow blur anisotropic instead of the raster backend's average-scale isotropic blur",
                            );
                        }
                    }
                    Effect::LinearOpacityMask { .. } => {}
                    _ => self.diagnose(
                        &format!("{path}.effects[{index}]"),
                        "unsupported group effect was omitted while its children were preserved",
                    ),
                }
            }
            if !filter.is_empty() {
                filter.push_str("<feMerge>");
                for shadow in shadows {
                    write!(filter, "<feMergeNode in=\"{shadow}\"/>").unwrap();
                }
                filter.push_str("<feMergeNode in=\"SourceGraphic\"/></feMerge>");
                match filter_region {
                    Some(Ok(Some(region))) => {
                        write!(
                            self.defs,
                            "<filter id=\"{id}\" filterUnits=\"userSpaceOnUse\" primitiveUnits=\"userSpaceOnUse\" x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" color-interpolation-filters=\"sRGB\">{filter}</filter>",
                            number(region.x),
                            number(region.y),
                            number(region.width),
                            number(region.height)
                        )
                        .unwrap();
                        write!(effect_attributes, " filter=\"url(#{id})\"").unwrap();
                    }
                    Some(Ok(None)) | None => {}
                    Some(Err(())) => self.diagnose(
                        path,
                        "group effects were omitted because singular-transform source bounds could not be proven",
                    ),
                }
            }
        }

        // Preserve the old unmasked markup. With masks, their wrappers surround
        // the filtered scene so shadow and descendants are masked once together.
        if mask_wrappers.is_empty() {
            attributes.push_str(&effect_attributes);
        }
        write!(output, "<g{attributes}>").unwrap();
        for wrapper in &mask_wrappers {
            output.push_str(wrapper);
        }
        if !mask_wrappers.is_empty() && !effect_attributes.is_empty() {
            write!(output, "<g{effect_attributes}>").unwrap();
        }
        if let Some(clip_id) = clip_id {
            write!(output, "<g clip-path=\"url(#{clip_id})\">").unwrap();
        }
        self.emit_elements(
            &group.children,
            &format!("{path}.children"),
            output,
            group_to_page,
        );
        if group.clip.is_some() {
            output.push_str("</g>");
        }
        if !mask_wrappers.is_empty() && !effect_attributes.is_empty() {
            output.push_str("</g>");
        }
        for _ in &mask_wrappers {
            output.push_str("</g>");
        }
        output.push_str("</g>");
    }

    fn filter_region(
        &self,
        group: &GroupElement,
        group_to_page: Transform,
    ) -> Result<Option<Rect>, ()> {
        if let Some(region) = inverse_transform(group_to_page)
            .map(|page_to_group| {
                page_to_group.transform_rect_bbox(Rect {
                    x: 0.0,
                    y: 0.0,
                    width: self.page_width,
                    height: self.page_height,
                })
            })
            .filter(rect_is_finite)
        {
            return Ok(Some(region));
        }
        group_effect_bounds(group, &self.layout.fonts)
    }

    fn paint_attributes(&mut self, paint: &Paint, property: &str, path: &str) -> Option<String> {
        match paint {
            Paint::Solid(color) => Some(format!(
                "{property}=\"{}\" {property}-opacity=\"{}\"",
                color_hex(*color),
                number(color.a.clamp(0.0, 1.0))
            )),
            Paint::Linear {
                start,
                end,
                stops,
                extend,
            } => {
                self.diagnose_gradient_extension(*extend, path);
                let id = self.definition_id();
                let mut definition = format!(
                    "<linearGradient id=\"{id}\" gradientUnits=\"userSpaceOnUse\" x1=\"{}\" y1=\"{}\" x2=\"{}\" y2=\"{}\">",
                    number(start.x),
                    number(start.y),
                    number(end.x),
                    number(end.y)
                );
                emit_stops(stops, &mut definition);
                definition.push_str("</linearGradient>");
                self.defs.push_str(&definition);
                Some(format!("{property}=\"url(#{id})\""))
            }
            Paint::Radial {
                center,
                radius,
                focal,
                stops,
                extend,
            } => {
                self.diagnose_gradient_extension(*extend, path);
                let id = self.definition_id();
                let mut definition = format!(
                    "<radialGradient id=\"{id}\" gradientUnits=\"userSpaceOnUse\" cx=\"{}\" cy=\"{}\" r=\"{}\" fx=\"{}\" fy=\"{}\">",
                    number(center.x),
                    number(center.y),
                    number(*radius),
                    number(focal.x),
                    number(focal.y)
                );
                emit_stops(stops, &mut definition);
                definition.push_str("</radialGradient>");
                self.defs.push_str(&definition);
                Some(format!("{property}=\"url(#{id})\""))
            }
            Paint::Grid {
                bounds,
                columns,
                rows,
                colors,
            } => {
                Paint::grid(*bounds, *columns, *rows, colors.clone())?;
                let id = self.definition_id();
                let dx = bounds.width / f64::from(*columns);
                let dy = bounds.height / f64::from(*rows);
                let mut definition = format!(
                    "<pattern id=\"{id}\" patternUnits=\"userSpaceOnUse\" x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\">",
                    number(bounds.x),
                    number(bounds.y),
                    number(bounds.width),
                    number(bounds.height)
                );
                for row in 0..usize::from(*rows) {
                    let mut col = 0;
                    while col < usize::from(*columns) {
                        let color = colors[row * usize::from(*columns) + col];
                        let mut end = col + 1;
                        while end < usize::from(*columns)
                            && colors[row * usize::from(*columns) + end] == color
                        {
                            end += 1;
                        }
                        let _ = write!(
                            definition,
                            "<rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" fill=\"{}\" fill-opacity=\"{}\"/>",
                            number(col as f64 * dx),
                            number(row as f64 * dy),
                            number((end - col) as f64 * dx),
                            number(dy),
                            color_hex(color),
                            number(color.a)
                        );
                        col = end;
                    }
                }
                definition.push_str("</pattern>");
                self.defs.push_str(&definition);
                Some(format!("{property}=\"url(#{id})\""))
            }
            Paint::Tile { .. } => {
                self.diagnose(
                    path,
                    "tile paint was omitted because its carrier has no media bytes",
                );
                None
            }
        }
    }

    fn diagnose_gradient_extension(&mut self, extend: (bool, bool), path: &str) {
        if extend != (true, true) {
            self.diagnose(
                path,
                "asymmetric or disabled gradient extension uses SVG pad extension",
            );
        }
    }
}

fn collect_used_fonts(
    elements: &[PositionedElement],
    seen: &mut HashSet<FontId>,
    ordered: &mut Vec<FontId>,
) {
    for element in elements {
        match element {
            PositionedElement::Text(run) if !run.text.is_empty() => {
                if seen.insert(run.font_id) {
                    ordered.push(run.font_id);
                }
            }
            PositionedElement::MultilingualText(run) if !run.logical_text.is_empty() => {
                if seen.insert(run.font_id) {
                    ordered.push(run.font_id);
                }
            }
            PositionedElement::Group(group) => collect_used_fonts(&group.children, seen, ordered),
            PositionedElement::MarkedContent { children, .. } => {
                collect_used_fonts(children, seen, ordered)
            }
            _ => {}
        }
    }
}

fn inverse_transform(transform: Transform) -> Option<Transform> {
    let determinant = transform.a * transform.d - transform.b * transform.c;
    if !determinant.is_finite() || determinant == 0.0 {
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
    .iter()
    .all(|value| value.is_finite())
    .then_some(inverse)
}

fn preserves_isotropic_blur(transform: Transform) -> bool {
    let x_scale_squared = transform.a * transform.a + transform.b * transform.b;
    let y_scale_squared = transform.c * transform.c + transform.d * transform.d;
    let dot = transform.a * transform.c + transform.b * transform.d;
    if !x_scale_squared.is_finite() || !y_scale_squared.is_finite() || !dot.is_finite() {
        return false;
    }
    let tolerance = x_scale_squared.max(y_scale_squared).max(1.0) * 1e-12;
    (x_scale_squared - y_scale_squared).abs() <= tolerance && dot.abs() <= tolerance
}

/// Bound the selected SFNT directory before handing producer bytes to the
/// parser. RawFace's collection-index stream arithmetic assumes these bounds.
fn validated_font_face(data: &[u8], face_index: u32) -> Option<ttf_parser::Face<'_>> {
    let read_u32 = |offset: usize| {
        Some(u32::from_be_bytes(
            data.get(offset..offset.checked_add(4)?)?.try_into().ok()?,
        ))
    };
    if data.len() > u32::MAX as usize {
        return None;
    }
    let offset = if data.starts_with(b"ttcf") {
        if !matches!(read_u32(4)?, 0x0001_0000 | 0x0002_0000) {
            return None;
        }
        let count = read_u32(8)?;
        let end = 12usize.checked_add(usize::try_from(count).ok()?.checked_mul(4)?)?;
        if face_index >= count || end > data.len() || end > u32::MAX as usize {
            return None;
        }
        usize::try_from(read_u32(
            12usize.checked_add(usize::try_from(face_index).ok()?.checked_mul(4)?)?,
        )?)
        .ok()?
    } else {
        if face_index != 0 {
            return None;
        }
        0
    };
    if !matches!(read_u32(offset)?, 0x0001_0000 | 0x7472_7565 | 0x4f54_544f) {
        return None;
    }
    let count_at = offset.checked_add(4)?;
    let count = usize::from(u16::from_be_bytes(
        data.get(count_at..count_at.checked_add(2)?)?
            .try_into()
            .ok()?,
    ));
    let directory = offset.checked_add(12)?;
    let end = directory.checked_add(count.checked_mul(16)?)?;
    if end > data.len() || end > u32::MAX as usize {
        return None;
    }
    for index in 0..count {
        let entry = directory.checked_add(index.checked_mul(16)?)?;
        let start = read_u32(entry.checked_add(8)?)?;
        let length = read_u32(entry.checked_add(12)?)?;
        let end = start.checked_add(length)?;
        if usize::try_from(end).ok()? > data.len() {
            return None;
        }
    }
    ttf_parser::Face::parse(data, face_index).ok()
}

/// Extract only the selected collection face, retaining every unsigned table.
/// Standalone fonts are borrowed verbatim so ordinary SVG output is unchanged.
fn standalone_font_data(font: &FontData) -> Option<std::borrow::Cow<'_, [u8]>> {
    validated_font_face(&font.data, font.face_index)?;
    if !font.data.starts_with(b"ttcf") {
        return Some(std::borrow::Cow::Borrowed(&font.data));
    }
    let raw = ttf_parser::RawFace::parse(&font.data, font.face_index).ok()?;
    let mut tables = raw
        .table_records
        .into_iter()
        .filter(|record| record.tag != ttf_parser::Tag::from_bytes(b"DSIG"))
        .collect::<Vec<_>>();
    tables.sort_by_key(|record| record.tag);
    if tables.is_empty()
        || tables.len() > 4095
        || tables.windows(2).any(|pair| pair[0].tag == pair[1].tag)
    {
        return None;
    }
    let count = u16::try_from(tables.len()).ok()?;
    let power = 1u16 << (15 - count.leading_zeros());
    let search = power.checked_mul(16)?;
    let directory = 12usize.checked_add(tables.len().checked_mul(16)?)?;
    let total = tables.iter().try_fold(directory, |size, record| {
        size.checked_add((usize::try_from(record.length).ok()?.checked_add(3)?) & !3)
    })?;
    if total
        > font
            .data
            .len()
            .checked_add(directory)?
            .checked_add(tables.len().checked_mul(3)?)?
    {
        return None;
    }
    let face_offset_pos =
        12usize.checked_add(usize::try_from(font.face_index).ok()?.checked_mul(4)?)?;
    let face_offset = usize::try_from(u32::from_be_bytes(
        font.data
            .get(face_offset_pos..face_offset_pos.checked_add(4)?)?
            .try_into()
            .ok()?,
    ))
    .ok()?;
    let mut output = vec![0u8; total];
    output[..4].copy_from_slice(font.data.get(face_offset..face_offset.checked_add(4)?)?);
    output[4..6].copy_from_slice(&count.to_be_bytes());
    output[6..8].copy_from_slice(&search.to_be_bytes());
    output[8..10].copy_from_slice(&(power.trailing_zeros() as u16).to_be_bytes());
    output[10..12].copy_from_slice(&count.checked_mul(16)?.checked_sub(search)?.to_be_bytes());
    let mut offset = directory;
    let mut head_offset = None;
    for (index, record) in tables.iter().enumerate() {
        let start = usize::try_from(record.offset).ok()?;
        let length = usize::try_from(record.length).ok()?;
        let data = font.data.get(start..start.checked_add(length)?)?;
        let end = offset.checked_add(length)?;
        output.get_mut(offset..end)?.copy_from_slice(data);
        if record.tag == ttf_parser::Tag::from_bytes(b"head") {
            if length < 12 {
                return None;
            }
            output[offset + 8..offset + 12].fill(0);
            head_offset = Some(offset);
        }
        let padded_end = end.checked_add(3)? & !3;
        let checksum = sfnt_checksum(output.get(offset..padded_end)?);
        let row = 12 + index * 16;
        output[row..row + 4].copy_from_slice(&record.tag.to_bytes());
        output[row + 4..row + 8].copy_from_slice(&checksum.to_be_bytes());
        output[row + 8..row + 12].copy_from_slice(&u32::try_from(offset).ok()?.to_be_bytes());
        output[row + 12..row + 16].copy_from_slice(&record.length.to_be_bytes());
        offset = padded_end;
    }
    let head = head_offset?;
    let adjustment = 0xB1B0_AFBAu32.wrapping_sub(sfnt_checksum(&output));
    output[head + 8..head + 12].copy_from_slice(&adjustment.to_be_bytes());
    validated_font_face(&output, 0)?;
    Some(std::borrow::Cow::Owned(output))
}

fn sfnt_checksum(bytes: &[u8]) -> u32 {
    bytes.chunks(4).fold(0u32, |sum, chunk| {
        let mut word = [0u8; 4];
        word[..chunk.len()].copy_from_slice(chunk);
        sum.wrapping_add(u32::from_be_bytes(word))
    })
}

fn variation_settings(font: &FontData) -> String {
    // CSS weight, width and style select registered axes even when FontData
    // has no nondefault coordinates. Optical sizing can also select opsz.
    // Freeze those effective defaults before applying the selected overrides.
    let mut coordinates = Vec::new();
    if let Some(face) = validated_font_face(&font.data, font.face_index) {
        for axis in face.variation_axes() {
            let tag = axis.tag.to_bytes();
            if matches!(&tag, b"wght" | b"wdth" | b"ital" | b"slnt" | b"opsz") {
                coordinates.push((tag, axis.def_value));
            }
        }
    }
    for &(tag, value) in &font.variations {
        if let Some((_, coordinate)) = coordinates.iter_mut().find(|(axis, _)| *axis == tag) {
            *coordinate = value;
        } else {
            coordinates.push((tag, value));
        }
    }
    coordinates
        .iter()
        .map(|(tag, value)| {
            format!(
                "'{}' {}",
                String::from_utf8_lossy(tag),
                number(f64::from(*value))
            )
        })
        .collect::<Vec<_>>()
        .join(",")
}

/// Conservative painted glyph bounds in points, before ancestor transforms.
fn glyph_ink_bounds(
    run: &GlyphRun,
    fonts: &[FontData],
    x_offsets: &[f64],
    y_offsets: &[f64],
    y_advances: &[f64],
) -> Result<Option<Rect>, ()> {
    if !run.font_size.is_finite()
        || run.font_size <= 0.0
        || !run.origin.x.is_finite()
        || !run.origin.y.is_finite()
        || run.glyph_ids.len() != run.advances.len()
    {
        return Err(());
    }
    let font = fonts.iter().find(|font| font.id == run.font_id).ok_or(())?;
    let mut face = validated_font_face(&font.data, font.face_index).ok_or(())?;
    for (tag, value) in &font.variations {
        if !value.is_finite() {
            return Err(());
        }
        face.set_variation(ttf_parser::Tag::from_bytes(tag), *value)
            .ok_or(())?;
    }
    if face.units_per_em() == 0 {
        return Err(());
    }
    let scale = run.font_size / f64::from(face.units_per_em());
    let radius = if font.synthetic_bold {
        run.font_size / 48.0
    } else {
        0.0
    };
    let shear = if font.synthetic_italic {
        0.212_556_561_670_022_1
    } else {
        0.0
    };
    let affine = Transform {
        c: -shear,
        e: shear * run.origin.y,
        ..Transform::IDENTITY
    };
    let (mut x, mut y) = (run.origin.x, run.origin.y);
    let mut result = None;
    for (index, glyph) in run.glyph_ids.iter().enumerate() {
        if *glyph >= face.number_of_glyphs() {
            return Err(());
        }
        let position_x = x + x_offsets.get(index).copied().unwrap_or(0.0);
        let position_y = y - y_offsets.get(index).copied().unwrap_or(0.0);
        if let Some(bounds) = face.glyph_bounding_box(ttf_parser::GlyphId(*glyph)) {
            let rect = Rect {
                x: position_x + f64::from(bounds.x_min) * scale - radius,
                y: position_y - f64::from(bounds.y_max) * scale - radius,
                width: f64::from(i32::from(bounds.x_max) - i32::from(bounds.x_min)) * scale
                    + 2.0 * radius,
                height: f64::from(i32::from(bounds.y_max) - i32::from(bounds.y_min)) * scale
                    + 2.0 * radius,
            };
            let rect = affine.transform_rect_bbox(rect);
            if !rect_is_finite(&rect) {
                return Err(());
            }
            result = Some(result.map_or(rect, |old| union_rect(old, rect)));
        }
        x += run.advances[index];
        y -= y_advances.get(index).copied().unwrap_or(0.0);
        if !x.is_finite() || !y.is_finite() {
            return Err(());
        }
    }
    Ok(result)
}

fn group_effect_bounds(group: &GroupElement, fonts: &[FontData]) -> Result<Option<Rect>, ()> {
    let Some(mut source) = elements_bounds(&group.children, fonts)? else {
        return Ok(None);
    };
    if let Some(clip) = &group.clip {
        let Some(clip_bounds) = clip.bounds() else {
            return Ok(None);
        };
        let Some(clipped) = intersect_rect(source, clip_bounds) else {
            return Ok(None);
        };
        source = clipped;
    }
    let mut bounds = source;
    for effect in &group.effects {
        if let Effect::OuterShadow { dx, dy, blur, .. } = effect {
            let margin = blur.max(0.0) * 2.0 + 1.0;
            let shadow = Rect {
                x: source.x + dx - margin,
                y: source.y + dy - margin,
                width: source.width + margin * 2.0,
                height: source.height + margin * 2.0,
            };
            bounds = union_rect(bounds, shadow);
        }
    }
    rect_is_finite(&bounds).then_some(Some(bounds)).ok_or(())
}

fn elements_bounds(elements: &[PositionedElement], fonts: &[FontData]) -> Result<Option<Rect>, ()> {
    let mut bounds = None;
    for element in elements {
        if let Some(element) = element_bounds(element, fonts)? {
            bounds = Some(match bounds {
                Some(bounds) => union_rect(bounds, element),
                None => element,
            });
        }
    }
    Ok(bounds)
}

fn element_bounds(element: &PositionedElement, fonts: &[FontData]) -> Result<Option<Rect>, ()> {
    match element {
        PositionedElement::Text(run) if !run.text.is_empty() => {
            glyph_ink_bounds(run, fonts, &[], &[], &[])
        }
        PositionedElement::MultilingualText(run) if !run.logical_text.is_empty() => {
            if !run.is_valid() {
                return Err(());
            }
            glyph_ink_bounds(
                &run.legacy_projection(),
                fonts,
                &run.x_offsets,
                &run.y_offsets,
                &run.y_advances,
            )
        }
        PositionedElement::Line {
            start, end, width, ..
        } => {
            let margin = width.abs() / 2.0 + 1.0;
            Ok(Some(Rect {
                x: start.x.min(end.x) - margin,
                y: start.y.min(end.y) - margin,
                width: (start.x - end.x).abs() + margin * 2.0,
                height: (start.y - end.y).abs() + margin * 2.0,
            }))
        }
        PositionedElement::FilledRect { rect, .. } | PositionedElement::Image { rect, .. } => {
            Ok(Some(normal_rect(*rect)))
        }
        PositionedElement::Path(element) => Ok(element.path.bounds().map(|bounds| {
            let margin = element
                .stroke
                .as_ref()
                .map(|stroke| match stroke.join {
                    LineJoin::Miter => stroke.width.abs() * 2.0 + 1.0,
                    LineJoin::Round | LineJoin::Bevel => stroke.width.abs() / 2.0 + 1.0,
                })
                .unwrap_or(0.0);
            Rect {
                x: bounds.x - margin,
                y: bounds.y - margin,
                width: bounds.width + margin * 2.0,
                height: bounds.height + margin * 2.0,
            }
        })),
        PositionedElement::Group(group) => Ok(group_effect_bounds(group, fonts)?
            .map(|bounds| group.transform.transform_rect_bbox(bounds))),
        PositionedElement::MarkedContent { children, .. } => elements_bounds(children, fonts),
        PositionedElement::LinkAnnotation { .. }
        | PositionedElement::Text(_)
        | PositionedElement::MultilingualText(_) => Ok(None),
        _ => Err(()),
    }
}

fn normal_rect(rect: Rect) -> Rect {
    Rect {
        x: rect.x.min(rect.x + rect.width),
        y: rect.y.min(rect.y + rect.height),
        width: rect.width.abs(),
        height: rect.height.abs(),
    }
}

fn intersect_rect(left: Rect, right: Rect) -> Option<Rect> {
    let left = normal_rect(left);
    let right = normal_rect(right);
    let x = left.x.max(right.x);
    let y = left.y.max(right.y);
    let right_edge = (left.x + left.width).min(right.x + right.width);
    let bottom = (left.y + left.height).min(right.y + right.height);
    (right_edge >= x && bottom >= y).then_some(Rect {
        x,
        y,
        width: right_edge - x,
        height: bottom - y,
    })
}

fn union_rect(left: Rect, right: Rect) -> Rect {
    let left = normal_rect(left);
    let right = normal_rect(right);
    let x = left.x.min(right.x);
    let y = left.y.min(right.y);
    let right_edge = (left.x + left.width).max(right.x + right.width);
    let bottom = (left.y + left.height).max(right.y + right.height);
    Rect {
        x,
        y,
        width: right_edge - x,
        height: bottom - y,
    }
}

fn rect_is_finite(rect: &Rect) -> bool {
    [rect.x, rect.y, rect.width, rect.height]
        .iter()
        .all(|value| value.is_finite())
}

fn emit_stops(stops: &[oxml_layout::GradientStop], output: &mut String) {
    let mut normalized = stops
        .iter()
        .map(|stop| oxml_layout::GradientStop {
            offset: if stop.offset.is_nan() {
                0.0
            } else {
                stop.offset.clamp(0.0, 1.0)
            },
            color: stop.color,
        })
        .collect::<Vec<_>>();
    normalized.sort_by(|left, right| left.offset.total_cmp(&right.offset));
    let mut deduplicated: Vec<oxml_layout::GradientStop> = Vec::with_capacity(normalized.len());
    for stop in normalized {
        if let Some(previous) = deduplicated.last_mut()
            && previous.offset == stop.offset
        {
            *previous = stop;
        } else {
            deduplicated.push(stop);
        }
    }

    for stop in &deduplicated {
        write!(
            output,
            "<stop offset=\"{}\" stop-color=\"{}\" stop-opacity=\"{}\"/>",
            number(stop.offset),
            color_hex(stop.color),
            number(stop.color.a.clamp(0.0, 1.0))
        )
        .unwrap();
    }
}

fn path_data(path: &Path) -> String {
    let mut data = String::new();
    for command in &path.commands {
        if !data.is_empty() {
            data.push(' ');
        }
        match command {
            PathCommand::MoveTo(point) => {
                write!(data, "M{} {}", number(point.x), number(point.y)).unwrap()
            }
            PathCommand::LineTo(point) => {
                write!(data, "L{} {}", number(point.x), number(point.y)).unwrap()
            }
            PathCommand::CurveTo { c1, c2, to } => write!(
                data,
                "C{} {} {} {} {} {}",
                number(c1.x),
                number(c1.y),
                number(c2.x),
                number(c2.y),
                number(to.x),
                number(to.y)
            )
            .unwrap(),
            PathCommand::Close => data.push('Z'),
        }
    }
    data
}

fn number(value: f64) -> String {
    if value == 0.0 {
        return "0".to_owned();
    }
    let mut value = format!("{value:.6}");
    while value.contains('.') && value.ends_with('0') {
        value.pop();
    }
    if value.ends_with('.') {
        value.pop();
    }
    value
}

fn color_hex(color: Color) -> String {
    let channel = |value: f64| (value.clamp(0.0, 1.0) * 255.0).round() as u8;
    format!(
        "#{:02X}{:02X}{:02X}",
        channel(color.r),
        channel(color.g),
        channel(color.b)
    )
}

fn line_cap(cap: LineCap) -> &'static str {
    match cap {
        LineCap::Butt => "butt",
        LineCap::Round => "round",
        LineCap::Square => "square",
    }
}

fn line_join(join: LineJoin) -> &'static str {
    match join {
        LineJoin::Miter => "miter",
        LineJoin::Round => "round",
        LineJoin::Bevel => "bevel",
    }
}

fn fill_rule(rule: FillRule) -> &'static str {
    match rule {
        FillRule::NonZero => "nonzero",
        FillRule::EvenOdd => "evenodd",
    }
}

fn font_content_type(font: &FontData) -> (&'static str, &'static str) {
    if font.data.starts_with(b"ttcf") {
        ("font/collection", "collection")
    } else if font.data.starts_with(b"OTTO") {
        ("font/otf", "opentype")
    } else if font.data.starts_with(b"wOFF") {
        ("font/woff", "woff")
    } else if font.data.starts_with(b"wOF2") {
        ("font/woff2", "woff2")
    } else {
        ("font/ttf", "truetype")
    }
}

fn image_mime(data: &[u8]) -> Option<&'static str> {
    if data.starts_with(b"\x89PNG\r\n\x1a\n") {
        Some("image/png")
    } else if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
        Some("image/jpeg")
    } else {
        None
    }
}

fn is_safe_link(target: &str) -> bool {
    if target.is_empty()
        || target
            .chars()
            .any(|character| character.is_control() || !is_xml_character(character))
    {
        return false;
    }
    let target = target.trim();
    if target.is_empty() {
        return false;
    }
    let Some(colon) = target.find(':') else {
        return true;
    };
    if target[..colon]
        .chars()
        .any(|character| matches!(character, '/' | '?' | '#'))
    {
        return true;
    }
    let scheme = &target[..colon];
    if scheme.is_empty()
        || !scheme.chars().enumerate().all(|(index, character)| {
            character.is_ascii_alphabetic()
                || (index > 0
                    && (character.is_ascii_digit()
                        || character == '+'
                        || character == '-'
                        || character == '.'))
        })
    {
        return false;
    }
    matches!(
        scheme.to_ascii_lowercase().as_str(),
        "http" | "https" | "mailto"
    )
}

fn escape_text(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn sanitize_xml_text(value: &str) -> (String, bool) {
    let mut replaced = false;
    let text = value
        .chars()
        .map(|character| {
            if is_xml_character(character) {
                character
            } else {
                replaced = true;
                '\u{FFFD}'
            }
        })
        .collect();
    (text, replaced)
}

fn is_xml_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(character as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn escape_attribute(value: &str) -> String {
    escape_text(value)
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;
    use std::sync::Arc;

    use oxml_layout::{
        Color, Diagnostic, Effect, FillRule, FontData, FontId, FontManager, GlyphCluster, GlyphRun,
        GradientStop, GroupElement, LayoutResult, LineCap, LineJoin, MediaId, MultilingualGlyphRun,
        PageFrame, Paint, Path, PathCommand, PathElement, Point, PositionedElement, Rect, Stroke,
        TextDirection, TextScript, Transform,
    };

    use base64::Engine;

    use super::{
        BASE64, glyph_ink_bounds, image_mime, is_safe_link, preserves_isotropic_blur, render_page,
        sfnt_checksum, standalone_font_data,
    };

    const RESVG_ORACLE_VERSION: &str = "resvg 0.48.1";

    fn color(r: f64, g: f64, b: f64) -> Color {
        Color { r, g, b, a: 1.0 }
    }

    fn test_font() -> FontData {
        FontData {
            id: FontId(7),
            family: "Carlito".to_owned(),
            data: Arc::from(
                include_bytes!("../../oxml-layout/fonts/Carlito-Regular.ttf").as_slice(),
            ),
            face_index: 0,
            bold: false,
            italic: false,
            variations: vec![],
            synthetic_bold: false,
            synthetic_italic: false,
        }
    }

    fn text(value: &str, glyphs: usize) -> PositionedElement {
        let bytes = include_bytes!("../../oxml-layout/fonts/Carlito-Regular.ttf");
        let face = ttf_parser::Face::parse(bytes, 0).unwrap();
        let glyph_ids = if value.chars().count() == glyphs {
            value
                .chars()
                .map(|scalar| face.glyph_index(scalar).unwrap_or(ttf_parser::GlyphId(0)).0)
                .collect()
        } else {
            vec![1; glyphs]
        };
        PositionedElement::Text(GlyphRun {
            origin: Point { x: 12.0, y: 28.0 },
            font_id: FontId(7),
            font_size: 14.0,
            glyph_ids,
            advances: vec![8.0; glyphs],
            text: value.to_owned(),
            source: None,
            color: Color::BLACK,
            bold: false,
            italic: false,
            field_kind: None,
            field_source: None,
            note: None,
            note_reference_source: None,
            tab_aligned: None,
        })
    }

    fn layout(elements: Vec<PositionedElement>) -> LayoutResult {
        LayoutResult::new(
            vec![Arc::new(PageFrame::new(1, 120.0, 80.0, elements))],
            vec![test_font()],
            None,
            Vec::new(),
        )
    }

    #[test]
    fn fx195_svg_freezes_css_registered_defaults_and_selected_overrides() {
        // CSS normal maps wght to 400. Explicit coordinates must override it,
        // rather than assuming a missing declaration means the fvar default.
        for (bytes, default_weight, selected_weight) in [
            (
                include_bytes!("../../oxml-layout/fonts/NotoSansSC-FX058-subset.ttf").as_slice(),
                100.0,
                100.0,
            ),
            (
                include_bytes!("../../oxml-layout/fonts/NotoSansThai.ttf").as_slice(),
                400.0,
                400.0,
            ),
            (
                include_bytes!("../../oxml-layout/fonts/NotoSansThai.ttf").as_slice(),
                400.0,
                700.0,
            ),
        ] {
            let face = ttf_parser::Face::parse(bytes, 0).unwrap();
            let axis = face
                .variation_axes()
                .into_iter()
                .find(|axis| axis.tag == ttf_parser::Tag::from_bytes(b"wght"))
                .unwrap();
            assert_eq!(axis.def_value, default_weight);
            let mut font = test_font();
            font.data = Arc::from(bytes);
            font.bold = selected_weight == 700.0;
            if selected_weight != default_weight {
                font.variations.push((*b"wght", selected_weight));
            }
            let original_data = font.data.clone();
            let original_coordinates = font.variations.clone();
            let mut document = layout(Vec::new());
            document.fonts = vec![font];
            let mut state = super::SvgState::new(&document, 120.0, 80.0, Vec::new());
            state.emit_font_definitions(&[FontId(7)]);
            let PositionedElement::Text(mut run) = text("A", 1) else {
                unreachable!()
            };
            run.bold = selected_weight == 700.0;
            run.glyph_ids = vec![face.glyph_index('A').expect("source cmap control").0];
            let mut selected_face = ttf_parser::Face::parse(bytes, 0).unwrap();
            for &(tag, value) in &document.fonts[0].variations {
                selected_face
                    .set_variation(ttf_parser::Tag::from_bytes(&tag), value)
                    .unwrap();
            }
            run.advances = vec![
                f64::from(
                    selected_face
                        .glyph_hor_advance(ttf_parser::GlyphId(run.glyph_ids[0]))
                        .unwrap(),
                ) * run.font_size
                    / f64::from(selected_face.units_per_em()),
            ];
            let mut emitted = String::new();
            state.emit_text(&run, "control", &mut emitted, false);
            let xml = resvg::usvg::roxmltree::Document::parse(&emitted).unwrap();
            let node = xml.root_element();
            assert_eq!(
                node.attribute("font-weight"),
                Some(if selected_weight == 700.0 {
                    "bold"
                } else {
                    "normal"
                })
            );
            let style = node
                .attribute("style")
                .expect("explicit effective instance");
            let settings = style
                .strip_prefix("font-synthesis:none;font-variation-settings:")
                .unwrap();
            let mut explicit = std::collections::BTreeMap::new();
            for declaration in settings.split(',') {
                let (tag, value) = declaration.trim().split_once(' ').unwrap();
                assert!(
                    explicit
                        .insert(tag, value.parse::<f32>().unwrap())
                        .is_none()
                );
            }
            let mut effective_weight = if run.bold { 700.0 } else { 400.0 };
            if let Some(selected) = explicit.get("'wght'") {
                effective_weight = *selected;
            }
            assert_eq!(effective_weight, selected_weight);
            assert_eq!(explicit.get("'wght'"), Some(&selected_weight));
            assert!(
                state
                    .defs
                    .contains(&format!("font-variation-settings:{settings}"))
            );
            assert_eq!(document.fonts[0].variations, original_coordinates);
            assert_eq!(document.fonts[0].data.as_ref(), original_data.as_ref());
        }
        let static_font = test_font();
        assert!(super::variation_settings(&static_font).is_empty());
        let static_document = layout(Vec::new());
        let mut static_state = super::SvgState::new(&static_document, 120.0, 80.0, Vec::new());
        static_state.emit_font_definitions(&[FontId(7)]);
        assert_eq!(
            static_state.defs,
            format!(
                "<style>@font-face{{font-family:'rdocx-font-7';src:url('data:font/ttf;base64,{}') format('truetype');font-weight:normal;font-style:normal}}</style>",
                BASE64.encode(static_font.data.as_ref())
            )
        );
        let PositionedElement::Text(mut run) = text("A", 1) else {
            unreachable!()
        };
        let mut static_emitted = String::new();
        static_state.emit_text(&run, "static", &mut static_emitted, false);
        let static_xml = resvg::usvg::roxmltree::Document::parse(&static_emitted).unwrap();
        assert_eq!(static_xml.root_element().attribute("style"), None);

        let mut synthetic = test_font();
        synthetic.data = Arc::from(
            include_bytes!("../../oxml-layout/fonts/NotoSansSC-FX058-subset.ttf").as_slice(),
        );
        synthetic.bold = true;
        synthetic.italic = true;
        synthetic.synthetic_bold = true;
        synthetic.synthetic_italic = true;
        let face = ttf_parser::Face::parse(&synthetic.data, 0).unwrap();
        run.glyph_ids = vec![face.glyph_index('A').expect("source cmap control").0];
        run.bold = true;
        run.italic = true;
        let mut synthetic_document = layout(Vec::new());
        synthetic_document.fonts = vec![synthetic];
        let mut synthetic_state =
            super::SvgState::new(&synthetic_document, 120.0, 80.0, Vec::new());
        synthetic_state.emit_font_definitions(&[FontId(7)]);
        let mut synthetic_emitted = String::new();
        synthetic_state.emit_text(&run, "synthetic", &mut synthetic_emitted, false);
        let synthetic_xml = resvg::usvg::roxmltree::Document::parse(&synthetic_emitted).unwrap();
        let node = synthetic_xml.root_element();
        assert_eq!(node.attribute("font-weight"), Some("normal"));
        assert_eq!(node.attribute("font-style"), Some("normal"));
        assert_eq!(
            node.attribute("style"),
            Some("font-synthesis:none;font-variation-settings:'wght' 100")
        );
        assert_eq!(node.attribute("stroke-linecap"), Some("round"));
        assert_eq!(node.attribute("stroke-linejoin"), Some("round"));
        assert_eq!(node.attribute("paint-order"), Some("fill stroke"));
        assert!(
            node.attribute("transform")
                .unwrap()
                .starts_with("matrix(1 0 ")
        );
        assert!(
            synthetic_state
                .defs
                .contains("font-weight:normal;font-style:normal")
        );
        assert!(
            synthetic_state
                .defs
                .contains("font-variation-settings:'wght' 100")
        );
        assert!(synthetic_document.fonts[0].variations.is_empty());
    }

    #[test]
    fn svg_export_is_deterministic_and_emits_only_used_definitions() {
        let mut layout = layout(vec![text("Hi", 2)]);
        let mut unused = test_font();
        unused.id = FontId(8);
        layout.fonts.push(unused);
        let result = render_page(&layout, 0).unwrap();
        let again = render_page(&layout, 0).unwrap();
        assert_eq!(result, again);
        assert_eq!(result.svg.matches("@font-face").count(), 1);
        assert!(result.svg.contains("font-family=\"rdocx-font-7\""));
        assert!(!result.svg.contains("rdocx-font-8"));
        assert!(!result.svg.contains("file://"));
        assert!(render_page(&layout, 1).is_none());
    }

    #[test]
    fn svg_export_keeps_multilingual_logical_text_searchable() {
        let rich = PositionedElement::MultilingualText(MultilingualGlyphRun {
            origin: Point { x: 12.0, y: 28.0 },
            font_id: FontId(7),
            font_size: 14.0,
            glyph_ids: vec![1, 2],
            x_advances: vec![8.0, 8.0],
            y_advances: vec![0.0, 0.0],
            x_offsets: vec![1.0, 0.0],
            y_offsets: vec![0.0, 1.0],
            clusters: vec![
                GlyphCluster {
                    glyph_start: 0,
                    glyph_end: 1,
                    char_start: 1,
                    char_end: 2,
                },
                GlyphCluster {
                    glyph_start: 1,
                    glyph_end: 2,
                    char_start: 0,
                    char_end: 1,
                },
            ],
            logical_text: "אב".to_owned(),
            logical_index: 0,
            source: None,
            script: TextScript::Hebrew,
            language: Some("he".to_owned()),
            direction: TextDirection::RightToLeft,
            bidi_level: 1,
            color: Color::BLACK,
            bold: false,
            italic: false,
            field_kind: None,
            field_source: None,
            note: None,
            note_reference_source: None,
        });

        let result = render_page(&layout(vec![rich]), 0).unwrap();
        assert!(result.svg.contains(">אב</text>"));
        assert!(result.svg.contains("rdocx-font-7"));
        assert_eq!(result.diagnostics.len(), 1);
        assert_eq!(result.diagnostics[0].path, "pages[0].elements[0]");
        assert!(
            result.diagnostics[0]
                .message
                .contains("multilingual shaping kept searchable text")
        );
    }

    #[test]
    fn svg_rejects_non_finite_multilingual_positioning_without_invalid_numbers() {
        let rich = PositionedElement::MultilingualText(MultilingualGlyphRun {
            origin: Point { x: 12.0, y: 28.0 },
            font_id: FontId(7),
            font_size: 14.0,
            glyph_ids: vec![1, 2],
            x_advances: vec![f64::NAN, 8.0],
            y_advances: vec![0.0, 0.0],
            x_offsets: vec![0.0, 0.0],
            y_offsets: vec![0.0, 0.0],
            clusters: vec![
                GlyphCluster {
                    glyph_start: 0,
                    glyph_end: 1,
                    char_start: 1,
                    char_end: 2,
                },
                GlyphCluster {
                    glyph_start: 1,
                    glyph_end: 2,
                    char_start: 0,
                    char_end: 1,
                },
            ],
            logical_text: "אב".to_owned(),
            logical_index: 0,
            source: None,
            script: TextScript::Hebrew,
            language: Some("he".to_owned()),
            direction: TextDirection::RightToLeft,
            bidi_level: 1,
            color: Color::BLACK,
            bold: false,
            italic: false,
            field_kind: None,
            field_source: None,
            note: None,
            note_reference_source: None,
        });

        let result = render_page(&layout(vec![rich]), 0).unwrap();
        assert!(!result.svg.contains("NaN"), "{}", result.svg);
        assert!(!result.svg.contains(">אב</text>"));
        assert_eq!(result.diagnostics.len(), 1);
        assert!(
            result.diagnostics[0]
                .message
                .contains("invalid multilingual")
        );
    }

    #[test]
    fn svg_complex_text_stays_text_and_reports_positioning_approximation() {
        let result = render_page(&layout(vec![text("office", 5)]), 0).unwrap();
        assert!(result.svg.contains(">office</text>"));
        assert!(result.svg.contains("textLength=\"40\""));
        assert_eq!(result.diagnostics.len(), 1);
        assert_eq!(result.diagnostics[0].path, "pages[0].elements[0]");
        let mut substituted = text("office", 6);
        let PositionedElement::Text(run) = &mut substituted else {
            unreachable!()
        };
        run.glyph_ids[1] = run.glyph_ids[0];
        assert_eq!(run.text.chars().count(), run.glyph_ids.len());
        let result = render_page(&layout(vec![substituted]), 0).unwrap();
        assert!(result.svg.contains(">office</text>"));
        assert!(result.svg.contains("textLength=\"48\""));
        assert_eq!(result.diagnostics.len(), 1);
        assert!(
            result.diagnostics[0]
                .message
                .contains("total-advance positioning")
        );
    }

    #[test]
    fn svg_export_recurses_without_dropping_supported_siblings() {
        let nested = PositionedElement::Group(GroupElement {
            transform: Transform {
                e: 3.0,
                f: 4.0,
                ..Transform::IDENTITY
            },
            clip: Some(Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 40.0,
                height: 40.0,
            })),
            opacity: 0.75,
            effects: vec![Effect::OuterShadow {
                dx: 1.0,
                dy: 2.0,
                blur: 3.0,
                color: Color::BLACK,
            }],
            children: vec![PositionedElement::MarkedContent {
                structure: None,
                children: vec![PositionedElement::Group(GroupElement {
                    transform: Transform::IDENTITY,
                    clip: None,
                    opacity: 1.0,
                    effects: Vec::new(),
                    children: vec![
                        PositionedElement::FilledRect {
                            rect: Rect {
                                x: 1.0,
                                y: 2.0,
                                width: 3.0,
                                height: 4.0,
                            },
                            color: color(1.0, 0.0, 0.0),
                        },
                        PositionedElement::Path(PathElement {
                            path: Path {
                                commands: Vec::new(),
                                fill_rule: FillRule::NonZero,
                            },
                            fill: Some(Paint::Tile {
                                image: MediaId(1),
                                tile: Rect {
                                    x: 0.0,
                                    y: 0.0,
                                    width: 1.0,
                                    height: 1.0,
                                },
                                transform: Transform::IDENTITY,
                            }),
                            stroke: None,
                        }),
                        PositionedElement::FilledRect {
                            rect: Rect {
                                x: 5.0,
                                y: 6.0,
                                width: 7.0,
                                height: 8.0,
                            },
                            color: color(0.0, 0.0, 1.0),
                        },
                    ],
                })],
            }],
        });
        let result = render_page(&layout(vec![nested]), 0).unwrap();
        assert!(result.svg.contains("matrix(1 0 0 1 3 4)"));
        assert!(result.svg.contains("clip-path=\"url(#rdocx-def-0)\""));
        assert!(result.svg.contains("filter=\"url(#rdocx-def-1)\""));
        assert!(result.svg.contains("filterUnits=\"userSpaceOnUse\""));
        assert!(result.svg.contains("in=\"SourceAlpha\""));
        assert!(
            result
                .svg
                .contains("filter=\"url(#rdocx-def-1)\"><g clip-path=\"url(#rdocx-def-0)\"")
        );
        let red = result.svg.find("#FF0000").unwrap();
        let omitted = result
            .svg
            .find("<path d=\"\" fill-rule=\"nonzero\" fill=\"none\"")
            .unwrap();
        let blue = result.svg.find("#0000FF").unwrap();
        assert!(red < omitted && omitted < blue);
        assert_eq!(result.diagnostics.len(), 1);
    }

    #[test]
    fn three_group_noncommuting_transform_order_matches_flattened_pixels() {
        let outer = Transform {
            a: 1.2,
            d: 0.85,
            e: 12.0,
            f: 6.0,
            ..Transform::IDENTITY
        };
        let middle = Transform {
            a: 0.0,
            b: 1.0,
            c: -1.0,
            d: 0.0,
            e: 55.0,
            f: 0.0,
        };
        let inner = Transform {
            a: 1.0,
            b: 0.2,
            c: 0.35,
            d: 1.0,
            e: 0.0,
            f: 0.0,
        };
        let leaf = PositionedElement::FilledRect {
            rect: Rect {
                x: 5.0,
                y: 5.0,
                width: 12.0,
                height: 8.0,
            },
            color: Color::BLACK,
        };
        let nested = PositionedElement::Group(GroupElement {
            transform: outer,
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![PositionedElement::Group(GroupElement {
                transform: middle,
                clip: None,
                opacity: 1.0,
                effects: Vec::new(),
                children: vec![PositionedElement::Group(GroupElement {
                    transform: inner,
                    clip: None,
                    opacity: 1.0,
                    effects: Vec::new(),
                    children: vec![leaf.clone()],
                })],
            })],
        });
        let flattened = PositionedElement::Group(GroupElement {
            transform: inner.then(middle).then(outer),
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![leaf.clone()],
        });
        let reversed = PositionedElement::Group(GroupElement {
            transform: outer.then(middle).then(inner),
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![leaf],
        });

        let nested = render_page(&layout(vec![nested]), 0).unwrap();
        let outer_position = nested.svg.find("matrix(1.2 0 0 0.85 12 6)").unwrap();
        let middle_position = nested.svg.find("matrix(0 1 -1 0 55 0)").unwrap();
        let inner_position = nested.svg.find("matrix(1 0.2 0.35 1 0 0)").unwrap();
        assert!(outer_position < middle_position && middle_position < inner_position);
        let flattened = render_page(&layout(vec![flattened]), 0).unwrap();
        let reversed = render_page(&layout(vec![reversed]), 0).unwrap();
        let nested_pixels = rasterise_with_resvg(&nested.svg, 240, 160, &[]);
        let flattened_pixels = rasterise_with_resvg(&flattened.svg, 240, 160, &[]);
        let reversed_pixels = rasterise_with_resvg(&reversed.svg, 240, 160, &[]);
        assert_eq!(nested_pixels, flattened_pixels);
        assert_ne!(nested_pixels, reversed_pixels);
    }

    #[test]
    fn svg_export_escapes_untrusted_content_and_rejects_active_links() {
        let mut dangerous = match text("<&>", 3) {
            PositionedElement::Text(run) => run,
            _ => unreachable!(),
        };
        dangerous.text = "<&>\u{B}\"'".to_owned();
        let face = ttf_parser::Face::parse(
            include_bytes!("../../oxml-layout/fonts/Carlito-Regular.ttf"),
            0,
        )
        .unwrap();
        dangerous.glyph_ids = dangerous
            .text
            .chars()
            .map(|scalar| face.glyph_index(scalar).unwrap_or(ttf_parser::GlyphId(0)).0)
            .collect();
        assert_eq!(dangerous.text.chars().count(), 6);
        assert_eq!(dangerous.glyph_ids.len(), 6);
        dangerous.advances = vec![5.0; 6];
        let result = render_page(
            &layout(vec![
                PositionedElement::Text(dangerous),
                PositionedElement::LinkAnnotation {
                    rect: Rect {
                        x: 0.0,
                        y: 0.0,
                        width: 10.0,
                        height: 10.0,
                    },
                    url: "javascript:alert(1)\"/><script>".to_owned(),
                },
                PositionedElement::LinkAnnotation {
                    rect: Rect {
                        x: 0.0,
                        y: 0.0,
                        width: 10.0,
                        height: 10.0,
                    },
                    url: "https://example.test/\u{FFFE}".to_owned(),
                },
                PositionedElement::LinkAnnotation {
                    rect: Rect {
                        x: 0.0,
                        y: 0.0,
                        width: 10.0,
                        height: 10.0,
                    },
                    url: "https://example.test/\u{B}".to_owned(),
                },
                PositionedElement::LinkAnnotation {
                    rect: Rect {
                        x: 0.0,
                        y: 0.0,
                        width: 10.0,
                        height: 10.0,
                    },
                    url: "https://example.test/?a=1&b=2".to_owned(),
                },
            ]),
            0,
        )
        .unwrap();
        resvg::usvg::roxmltree::Document::parse(&result.svg)
            .expect("sanitized SVG remains parseable XML");
        assert!(result.svg.contains("&lt;&amp;&gt;�\"'"));
        assert!(!result.svg.contains('\u{B}'));
        assert!(!result.svg.contains("<script>"));
        assert!(!result.svg.contains("javascript:"));
        assert!(!result.svg.contains('\u{FFFE}'));
        assert!(!result.svg.contains("https://example.test/\u{B}"));
        assert!(result.svg.contains("a=1&amp;b=2"));
        assert_eq!(result.diagnostics.len(), 4);
        assert_eq!(result.diagnostics[0].path, "pages[0].elements[0]");
        assert_eq!(result.diagnostics[1].path, "pages[0].elements[1]");
        assert_eq!(result.diagnostics[2].path, "pages[0].elements[2]");
        assert_eq!(result.diagnostics[3].path, "pages[0].elements[3]");
        assert!(is_safe_link("chapter/one?label=a:b#section"));
        assert!(!is_safe_link("data:image/png;base64,AAAA"));
    }

    #[test]
    fn only_png_and_jpeg_media_and_reviewed_links_are_accepted() {
        assert_eq!(image_mime(b"\x89PNG\r\n\x1a\n"), Some("image/png"));
        assert_eq!(image_mime(&[0xFF, 0xD8, 0xFF]), Some("image/jpeg"));
        assert_eq!(image_mime(b"GIF89a"), None);
        for target in [
            "https://example.test",
            "HTTP://example.test",
            "mailto:a@b",
            "../a",
            "#x",
        ] {
            assert!(is_safe_link(target), "{target}");
        }
        for target in [
            "javascript:x",
            "data:text/html,x",
            "file:///tmp/x",
            "a\nb",
            "https://example.test/\u{B}",
            "https://example.test/\u{FFFE}",
            "https://example.test/\u{FFFF}",
            "",
        ] {
            assert!(!is_safe_link(target), "{target}");
        }
    }

    fn fx195_svg_rectangle_font() -> Vec<u8> {
        let mut head = vec![0; 54];
        head[0..4].copy_from_slice(&0x0001_0000u32.to_be_bytes());
        head[4..8].copy_from_slice(&0x0001_0000u32.to_be_bytes());
        head[12..16].copy_from_slice(&0x5f0f_3cf5u32.to_be_bytes());
        head[18..20].copy_from_slice(&1000u16.to_be_bytes());
        for (offset, value) in [(36, 50i16), (38, 0), (40, 590), (42, 700), (50, 1)] {
            head[offset..offset + 2].copy_from_slice(&value.to_be_bytes());
        }
        let mut hhea = vec![0; 36];
        hhea[0..4].copy_from_slice(&0x0001_0000u32.to_be_bytes());
        for (offset, value) in [(4, 800i16), (6, -200), (10, 600), (18, 1), (34, 2)] {
            hhea[offset..offset + 2].copy_from_slice(&value.to_be_bytes());
        }
        let mut maxp = vec![0; 32];
        maxp[0..4].copy_from_slice(&0x0001_0000u32.to_be_bytes());
        for (offset, value) in [(4, 2u16), (6, 4), (8, 1)] {
            maxp[offset..offset + 2].copy_from_slice(&value.to_be_bytes());
        }
        let mut glyf = Vec::new();
        for value in [1i16, 50, 0, 450, 700, 3, 0] {
            glyf.extend_from_slice(&value.to_be_bytes());
        }
        glyf.extend_from_slice(&[1; 4]); // Four on-curve points.
        for value in [50i16, 400, 0, -400, 0, 0, 700, 0] {
            glyf.extend_from_slice(&value.to_be_bytes());
        }
        while !glyf.len().is_multiple_of(4) {
            glyf.push(0);
        }
        let mut cmap = vec![0, 0, 0, 1, 0, 3, 0, 1, 0, 0, 0, 12];
        for value in [
            4u16, 32, 0, 4, 4, 1, 0, 65, 65535, 0, 65, 65535, 65472, 1, 0, 0,
        ] {
            cmap.extend_from_slice(&value.to_be_bytes());
        }
        let mut names = Vec::new();
        let mut strings = Vec::new();
        for (id, value) in [
            (1u16, "FX195SvgRectangle"),
            (2, "Regular"),
            (4, "FX195SvgRectangle Regular"),
            (6, "FX195SvgRectangle-Regular"),
            (256, "Axis"),
        ] {
            let text: Vec<u8> = value.encode_utf16().flat_map(u16::to_be_bytes).collect();
            for value in [3, 1, 0x409, id, text.len() as u16, strings.len() as u16] {
                names.extend_from_slice(&value.to_be_bytes());
            }
            strings.extend_from_slice(&text);
        }
        let mut name = vec![0, 0, 0, 5, 0, 66];
        name.extend(names);
        name.extend(strings);
        let mut os2 = vec![0; 78];
        for (offset, value) in [
            (2, 600i16),
            (4, 400),
            (6, 5),
            (62, 64),
            (64, 65),
            (66, 65),
            (68, 800),
            (70, -200),
            (74, 800),
            (76, 200),
        ] {
            os2[offset..offset + 2].copy_from_slice(&value.to_be_bytes());
        }
        let mut post = vec![0; 32];
        post[0..4].copy_from_slice(&0x0003_0000u32.to_be_bytes());
        let mut tables = vec![
            (*b"head", head),
            (*b"hhea", hhea),
            (*b"maxp", maxp),
            (*b"glyf", glyf),
            (*b"cmap", cmap),
            (*b"name", name),
            (*b"OS/2", os2),
            (*b"post", post),
            (*b"hmtx", vec![2, 88, 0, 0, 2, 88, 0, 50]),
            (*b"loca", vec![0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 36]),
        ];
        tables.sort_by_key(|(tag, _)| *tag);
        let count = tables.len() as u16;
        let power = 1u16 << (15 - count.leading_zeros());
        let mut font = Vec::from(0x0001_0000u32.to_be_bytes());
        for value in [
            count,
            power * 16,
            power.trailing_zeros() as u16,
            count * 16 - power * 16,
        ] {
            font.extend_from_slice(&value.to_be_bytes());
        }
        font.resize(12 + tables.len() * 16, 0);
        let mut head_offset = 0;
        for (index, (tag, mut table)) in tables.into_iter().enumerate() {
            let length = table.len();
            let offset = font.len();
            if tag == *b"head" {
                head_offset = offset;
            }
            while !table.len().is_multiple_of(4) {
                table.push(0);
            }
            let checksum = table.chunks_exact(4).fold(0u32, |sum, word| {
                sum.wrapping_add(u32::from_be_bytes(word.try_into().unwrap()))
            });
            let record = 12 + index * 16;
            font[record..record + 4].copy_from_slice(&tag);
            for (field, value) in [(4, checksum), (8, offset as u32), (12, length as u32)] {
                font[record + field..record + field + 4].copy_from_slice(&value.to_be_bytes());
            }
            font.extend(table);
        }
        let checksum = font.chunks_exact(4).fold(0u32, |sum, word| {
            sum.wrapping_add(u32::from_be_bytes(word.try_into().unwrap()))
        });
        font[head_offset + 8..head_offset + 12]
            .copy_from_slice(&0xb1b0_afbau32.wrapping_sub(checksum).to_be_bytes());
        font
    }

    #[test]
    fn fx195_variable_svg_paint_matches_independent_static_instances() {
        use sha2::{Digest, Sha256};
        use std::process::Command;
        let version = Command::new("hb-subset")
            .arg("--version")
            .output()
            .expect("required independent hb-subset14.4.0");
        assert!(version.status.success());
        assert!(
            String::from_utf8_lossy(&version.stdout).starts_with("hb-subset (HarfBuzz) 14.4.0\n")
        );
        for (script, source, value, source_hash, static_hashes) in [
            (
                "Thai",
                include_bytes!("../../oxml-layout/fonts/NotoSansThai.ttf").as_slice(),
                "ภาษาไทยยินดีต้อนรับ",
                "5a1c559bb539583c8a1fd99d1c5b9491e5e14478c9cd2bd0970d5c3096cc9ef8",
                [
                    "1101a6ce29474ff4bf2a68c092e8e68c652e8ff0d0b5093d36a81bbd3af1567e",
                    "171d20a8c3a3e7c1946618d4eebbf265bce47321d9438288079b683ab00451f2",
                ],
            ),
            (
                "Arabic",
                include_bytes!("../../oxml-layout/fonts/NotoSansArabic.ttf").as_slice(),
                "العربية مرحبا بالعالم",
                "63111b5b2e074dd48cc67692e0a2726d86ee94c1c37fe8598257b7b4e87e869e",
                [
                    "21e46030b17ddf81108d91fdf2f612674f998405d2a8837694db1d43d1db9878",
                    "2582ee0d7c72c06b675e05fffd5ad156a75a20f1509654f33a60e8760bb3d872",
                ],
            ),
            (
                "Devanagari",
                include_bytes!("../../oxml-layout/fonts/NotoSansDevanagari.ttf").as_slice(),
                "देवनागरी नमस्ते दुनिया",
                "9ce7b04f60e363d8870e5997744cf85cf69d38a4d7d129d364d92a3b14b461d7",
                [
                    "7d968955f9d3ec35b02cdb9265e7d04acc4b8e9e41fc3db7b44074018d5bbdbf",
                    "7ce34ec2d15c3f5f44c49e85e94e52ab1c8cfa336ac13507da973b5d6519d7dd",
                ],
            ),
        ] {
            assert_eq!(
                Sha256::digest(source)
                    .iter()
                    .map(|byte| format!("{byte:02x}"))
                    .collect::<String>(),
                source_hash
            );
            let mut manager = FontManager::new_with_fonts(vec![(
                "FX195 SVG variable caller".into(),
                source.to_vec(),
            )]);
            let mut paintings = Vec::new();
            let mut regular_oracle = None;
            for (index, weight) in [400, 700].into_iter().enumerate() {
                let filename = format!("NotoSans{script}-{weight}-static.ttf");
                let static_bytes =
                    if let Some(directory) = std::env::var_os("FX195_STATIC_ORACLE_DIR") {
                        std::fs::read(std::path::Path::new(&directory).join(&filename))
                            .expect("retained independently generated static oracle")
                    } else {
                        let directory = std::env::temp_dir().join(format!(
                            "fx195-svg-oracle-{}-{script}-{weight}",
                            std::process::id()
                        ));
                        std::fs::create_dir(&directory).unwrap();
                        let input = directory.join("source.ttf");
                        let output = directory.join(&filename);
                        std::fs::write(&input, source).unwrap();
                        let result = Command::new("hb-subset")
                            .arg(&input)
                            .arg(format!("--output-file={}", output.display()))
                            .args([
                                "--unicodes=*",
                                &format!("--variations=wght={weight},wdth=100"),
                                "--name-IDs=*",
                                "--name-languages=*",
                            ])
                            .output()
                            .unwrap();
                        let bytes = std::fs::read(&output);
                        std::fs::remove_dir_all(&directory).unwrap();
                        assert!(
                            result.status.success(),
                            "{}",
                            String::from_utf8_lossy(&result.stderr)
                        );
                        bytes.unwrap()
                    };
                assert_eq!(
                    Sha256::digest(&static_bytes)
                        .iter()
                        .map(|byte| format!("{byte:02x}"))
                        .collect::<String>(),
                    static_hashes[index]
                );
                let static_face = ttf_parser::Face::parse(&static_bytes, 0).unwrap();
                assert!(!static_face.is_variable());
                let id = manager
                    .resolve_font(Some("FX195 SVG variable caller"), weight == 700, false)
                    .unwrap();
                let shaped = manager.shape_text(id, value, 24.0).unwrap();
                let font = manager.font_data(id).unwrap();
                assert_eq!(font.data.as_ref(), source);
                let mut element = text(value, shaped.glyph_ids.len());
                let PositionedElement::Text(run) = &mut element else {
                    unreachable!()
                };
                run.origin = Point { x: 8.0, y: 32.0 };
                run.font_id = id;
                run.font_size = 24.0;
                run.glyph_ids = shaped.glyph_ids;
                run.advances = shaped.advances;
                run.bold = weight == 700;
                let input = LayoutResult::new(
                    vec![Arc::new(PageFrame::new(1, 300.0, 60.0, vec![element]))],
                    vec![font.clone()],
                    None,
                    Vec::new(),
                );
                let generated = render_page(&input, 0).unwrap();
                let raw_actual =
                    rasterise_with_resvg(&generated.svg, 600, 120, std::slice::from_ref(&font));
                let mut reference = font.clone();
                reference.data = Arc::from(static_bytes);
                reference.variations.clear();
                reference.bold = false;
                reference.synthetic_bold = false;
                // The expected SVG is independently authored, with the measured
                // static face loaded directly and no producer CSS/axis selection.
                let expected_svg = format!(
                    r#"<svg xmlns="http://www.w3.org/2000/svg" width="300" height="60" viewBox="0 0 300 60"><rect width="300" height="60" fill="white"/><text x="8" y="32" font-family="rdocx-font-{}" font-size="24" fill="black">{value}</text></svg>"#,
                    reference.id.0
                );
                let expected = rasterise_with_resvg(&expected_svg, 600, 120, &[reference.clone()]);
                // usvg0.48.1 scales textLength by max rather than summed glyph
                // advance inside a cluster. Only an independently measured identity
                // adjustment can be omitted for this oracle. Raw SVG is unchanged.
                let natural_advance =
                    independent_static_advance(&reference.data, value, script, weight);
                let oracle_svg =
                    identity_length_adjustment_for_usvg(&generated.svg, natural_advance);
                assert!(
                    identity_length_adjustment_for_usvg(
                        &format!(
                            r#"<text textLength="{}" lengthAdjust="spacingAndGlyphs">x</text>"#,
                            natural_advance + 1.0
                        ),
                        natural_advance,
                    )
                    .is_none(),
                    "nonidentity adjustment must never be stripped"
                );
                let actual = if let Some(svg) = &oracle_svg {
                    rasterise_with_resvg(svg, 600, 120, std::slice::from_ref(&font))
                } else {
                    assert!(
                        !generated.svg.contains(" textLength="),
                        "unqualified/nonidentity textLength requires a separate oracle"
                    );
                    raw_actual.clone()
                };
                assert!(actual.chunks_exact(4).any(|pixel| pixel[3] != 0));
                let score = luminance_ssim(&actual, &expected);
                println!(
                    "{script} weight{weight}: independent HB natural advance {natural_advance}, identity correction {}, raw SSIM {}, qualified oracle SSIM {score}",
                    oracle_svg.is_some(),
                    luminance_ssim(&raw_actual, &expected)
                );
                if script == "Devanagari"
                    && weight == 400
                    && let Some(directory) = std::env::var_os("FX195_SVG_DIAGNOSTIC_DIR")
                {
                    let directory = std::path::PathBuf::from(directory);
                    std::fs::create_dir_all(&directory).unwrap();
                    std::fs::write(directory.join("actual.svg"), &generated.svg).unwrap();
                    std::fs::write(directory.join("independent-static.svg"), &expected_svg)
                        .unwrap();
                    let without_length = generated.svg.replace(
                        &format!(
                            " textLength=\"{}\" lengthAdjust=\"spacingAndGlyphs\"",
                            super::number(shaped.width)
                        ),
                        "",
                    );
                    let natural = rasterise_with_resvg(
                        &without_length,
                        600,
                        120,
                        std::slice::from_ref(&font),
                    );
                    let same_source_plain =
                        rasterise_with_resvg(&expected_svg, 600, 120, std::slice::from_ref(&font));
                    let size = resvg::tiny_skia::IntSize::from_wh(600, 120).unwrap();
                    for (name, pixels) in [
                        ("actual", &raw_actual),
                        ("identity-corrected-oracle", &actual),
                        ("independent-static", &expected),
                        ("without-length", &natural),
                        ("same-source-plain", &same_source_plain),
                    ] {
                        let image =
                            resvg::tiny_skia::Pixmap::from_vec(pixels.clone(), size).unwrap();
                        image
                            .save_png(directory.join(format!("{name}.png")))
                            .unwrap();
                    }
                    let PositionedElement::Text(run) = &input.pages[0].elements[0] else {
                        unreachable!()
                    };
                    let face = ttf_parser::Face::parse(source, 0).unwrap();
                    let nominal = value
                        .chars()
                        .map(|scalar| face.glyph_index(scalar).unwrap_or(ttf_parser::GlyphId(0)).0)
                        .collect::<Vec<_>>();
                    std::fs::write(directory.join("actual-shape.json"), format!(
                        "{{\"script\":\"{script}\",\"weight\":{weight},\"scalarCount\":{},\"glyphIds\":{:?},\"nominalGlyphIds\":{:?},\"advances\":{:?},\"totalAdvance\":{},\"hasTextLength\":{},\"ssimActual\":{},\"ssimWithoutLength\":{},\"ssimSameSourcePlain\":{}}}\n",
                        value.chars().count(), run.glyph_ids, nominal, run.advances, shaped.width,
                        generated.svg.contains(" textLength="), luminance_ssim(&raw_actual, &expected),
                        luminance_ssim(&natural, &expected), luminance_ssim(&same_source_plain, &expected)
                    )).unwrap();
                    std::fs::write(
                        directory.join("independent-static.ttf"),
                        reference.data.as_ref(),
                    )
                    .unwrap();
                    let hb_version = Command::new("hb-shape").arg("--version").output().unwrap();
                    assert!(hb_version.status.success());
                    assert!(String::from_utf8_lossy(&hb_version.stdout).contains("14.4.0"));
                    std::fs::write(directory.join("hb-shape-version.txt"), hb_version.stdout)
                        .unwrap();
                    for (name, file) in [
                        (
                            "variable",
                            std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                                .join("../oxml-layout/fonts/NotoSansDevanagari.ttf"),
                        ),
                        (
                            "independent-static",
                            directory.join("independent-static.ttf"),
                        ),
                    ] {
                        let result = Command::new("hb-shape")
                            .args([
                                "--output-format=json",
                                "--no-glyph-names",
                                "--font-size=24000",
                            ])
                            .arg(file)
                            .arg(value)
                            .output()
                            .unwrap();
                        assert!(result.status.success());
                        std::fs::write(directory.join(format!("{name}-hb.json")), result.stdout)
                            .unwrap();
                    }
                }
                assert!(
                    score >= 0.99,
                    "{script} weight{weight}: independent static SVG SSIM {score}"
                );
                if weight == 400 {
                    regular_oracle = Some(expected);
                } else {
                    let correct_error: u64 = actual
                        .iter()
                        .zip(&expected)
                        .map(|(a, b)| u64::from(a.abs_diff(*b)))
                        .sum();
                    let default_error: u64 = actual
                        .iter()
                        .zip(regular_oracle.as_ref().unwrap())
                        .map(|(a, b)| u64::from(a.abs_diff(*b)))
                        .sum();
                    assert!(
                        default_error > 0 && correct_error.saturating_mul(4) < default_error,
                        "{script}: selected700 must paint substantially closer to independent700 than400: {correct_error}/{default_error}"
                    );
                }
                paintings.push(actual);
            }
            assert_ne!(
                paintings[0], paintings[1],
                "{script} selected700 paint must differ from400"
            );
        }
    }

    #[test]
    fn fx195_svg_synthetic_paint_matches_independent_rectangle() {
        let mut fonts = FontManager::new_with_fonts(vec![(
            "FX195SvgRectangle".into(),
            fx195_svg_rectangle_font(),
        )]);
        let plain = fonts
            .resolve_font(Some("FX195SvgRectangle"), false, false)
            .unwrap();
        let plain_shape = fonts.shape_text(plain, "A", 125.0).unwrap();
        let mut ordinary_pixels = None;
        for (bold, italic) in [(false, false), (true, false), (false, true), (true, true)] {
            let id = fonts
                .resolve_font(Some("FX195SvgRectangle"), bold, italic)
                .unwrap();
            let shaped = fonts.shape_text(id, "A", 125.0).unwrap();
            assert_eq!(shaped.advances, plain_shape.advances);
            let font = fonts.font_data(id).unwrap();
            assert_eq!(font.synthetic_bold, bold);
            assert_eq!(font.synthetic_italic, italic);
            let mut element = text("A", 1);
            let PositionedElement::Text(run) = &mut element else {
                unreachable!()
            };
            run.origin = Point { x: 12.0, y: 100.0 };
            run.font_id = id;
            run.font_size = 125.0;
            run.glyph_ids = shaped.glyph_ids;
            run.advances = shaped.advances;
            run.bold = bold;
            run.italic = italic;
            let input = LayoutResult::new(
                vec![Arc::new(PageFrame::new(1, 120.0, 120.0, vec![element]))],
                vec![font.clone()],
                None,
                Vec::new(),
            );
            let generated = render_page(&input, 0).unwrap();
            assert!(generated.diagnostics.is_empty());
            assert!(generated.svg.contains(">A</text>"));
            let actual = rasterise_with_resvg(&generated.svg, 240, 240, &[font]);
            // Authored glyf rectangle: x50..450,y0..700 at1000UPEM.
            // Independent SVG geometry applies point-space round stroke and
            // baseline-preserving tan(12deg) shear to those four known corners.
            let stroke = if bold {
                r#"stroke="black" stroke-width="5.208333333333333" stroke-linejoin="round" stroke-linecap="round""#
            } else {
                ""
            };
            let transform = if italic {
                r#"transform="matrix(1 0 -0.2125565616700221 1 21.25565616700221 0)""#
            } else {
                ""
            };
            let expected_svg = format!(
                r#"<svg xmlns="http://www.w3.org/2000/svg" width="120" height="120" viewBox="0 0 120 120"><rect width="120" height="120" fill="white"/><rect x="18.25" y="12.5" width="50" height="87.5" fill="black" {stroke} {transform}/></svg>"#
            );
            let expected = rasterise_with_resvg(&expected_svg, 240, 240, &[]);
            assert!(actual.chunks_exact(4).any(|pixel| pixel[3] != 0));
            let worst = actual
                .iter()
                .zip(&expected)
                .map(|(a, b)| a.abs_diff(*b))
                .max()
                .unwrap();
            assert!(
                worst <= 1,
                "independent SVG synthesis {bold}/{italic}: coverage delta {worst}"
            );
            if !bold && !italic {
                ordinary_pixels = Some(actual);
            } else {
                assert_ne!(actual, *ordinary_pixels.as_ref().unwrap());
            }
        }
    }

    #[test]
    fn fx195_synthesized_ink_bounds_preserve_advance_and_include_offsets() {
        let mut font = test_font();
        font.synthetic_bold = true;
        font.synthetic_italic = true;
        let face = ttf_parser::Face::parse(&font.data, 0).unwrap();
        let glyph = face.glyph_index('W').unwrap();
        let box_units = face.glyph_bounding_box(glyph).unwrap();
        let mut element = text("W", 1);
        let PositionedElement::Text(run) = &mut element else {
            unreachable!()
        };
        run.glyph_ids = vec![glyph.0];
        run.font_size = 48.0;
        run.origin = Point { x: 20.0, y: 70.0 };
        let advances_before = run.advances.clone();
        let bounds = glyph_ink_bounds(run, &[font.clone()], &[5.0], &[3.0], &[0.0])
            .unwrap()
            .unwrap();
        let scale = 48.0 / f64::from(face.units_per_em());
        let left = 25.0 + f64::from(box_units.x_min) * scale - 1.0;
        let right = 25.0 + f64::from(box_units.x_max) * scale + 1.0;
        let top = 67.0 - f64::from(box_units.y_max) * scale - 1.0;
        let bottom = 67.0 - f64::from(box_units.y_min) * scale + 1.0;
        let tangent = 12.0_f64.to_radians().tan();
        let expected_left = left + tangent * (70.0 - bottom);
        let expected_right = right + tangent * (70.0 - top);
        for (actual, expected) in [
            (bounds.x, expected_left),
            (bounds.y, top),
            (bounds.width, expected_right - expected_left),
            (bounds.height, bottom - top),
        ] {
            assert!(
                (actual - expected).abs() < 1e-10,
                "{actual} versus {expected}"
            );
        }
        assert_eq!(run.advances, advances_before);
        run.glyph_ids[0] = face.number_of_glyphs();
        assert!(glyph_ink_bounds(run, &[font], &[], &[], &[]).is_err());
    }

    #[test]
    fn fx195_selected_collection_face_is_embedded_and_painted() {
        let sources: [&[u8]; 2] = [
            include_bytes!("../../oxml-layout/fonts/Carlito-Regular.ttf"),
            include_bytes!("../../oxml-layout/fonts/Caladea-Regular.ttf"),
        ];
        let mut collection = Vec::from(&b"ttcf\0\x01\0\0\0\0\0\x02"[..]);
        collection.resize(20, 0);
        for (index, source) in sources.iter().enumerate() {
            let offset = collection.len();
            collection[12 + index * 4..16 + index * 4]
                .copy_from_slice(&(offset as u32).to_be_bytes());
            let mut face = source.to_vec();
            let count = usize::from(u16::from_be_bytes(face[4..6].try_into().unwrap()));
            for row in 0..count {
                let start = 12 + row * 16 + 8;
                let table = u32::from_be_bytes(face[start..start + 4].try_into().unwrap());
                face[start..start + 4].copy_from_slice(&(table + offset as u32).to_be_bytes());
            }
            collection.extend_from_slice(&face);
            collection.resize((collection.len() + 3) & !3, 0);
        }
        let mut font = test_font();
        font.data = Arc::from(collection);
        font.face_index = 1;
        let bytes_before = font.data.clone();
        let standalone = standalone_font_data(&font).unwrap().into_owned();
        assert_eq!(sfnt_checksum(&standalone), 0xB1B0_AFBA);
        let selected = ttf_parser::Face::parse(sources[1], 0).unwrap();
        let extracted = ttf_parser::Face::parse(&standalone, 0).unwrap();
        let first = ttf_parser::Face::parse(sources[0], 0).unwrap();
        let glyph = selected.glyph_index('W').unwrap();
        assert_eq!(extracted.glyph_index('W'), Some(glyph));
        assert_eq!(
            extracted.glyph_hor_advance(glyph),
            selected.glyph_hor_advance(glyph)
        );
        assert_ne!(
            selected.glyph_hor_advance(glyph),
            first.glyph_hor_advance(first.glyph_index('W').unwrap())
        );
        let mut element = text("W", 1);
        let PositionedElement::Text(run) = &mut element else {
            unreachable!()
        };
        run.glyph_ids = vec![glyph.0];
        run.advances = vec![
            f64::from(selected.glyph_hor_advance(glyph).unwrap()) * run.font_size
                / f64::from(selected.units_per_em()),
        ];
        let mut input = layout(vec![element]);
        input.fonts = vec![font.clone()];
        let svg = render_page(&input, 0).unwrap();
        assert!(svg.diagnostics.is_empty());
        assert!(svg.svg.contains(">W</text>"));
        let start = svg.svg.find("data:font/ttf;base64,").unwrap() + "data:font/ttf;base64,".len();
        let end = start + svg.svg[start..].find('\'').unwrap();
        assert_eq!(BASE64.decode(&svg.svg[start..end]).unwrap(), standalone);
        let mut extracted_font = font.clone();
        extracted_font.data = Arc::from(standalone);
        extracted_font.face_index = 0;
        let painted = rasterise_with_resvg(&svg.svg, 240, 160, &[extracted_font]);
        assert!(painted.chunks_exact(4).any(|pixel| pixel[3] != 0));
        let mut reference = font.clone();
        reference.data = Arc::from(sources[1]);
        reference.face_index = 0;
        let expected = rasterise_with_resvg(&svg.svg, 240, 160, &[reference]);
        assert_eq!(painted, expected);
        assert_eq!(font.data, bytes_before);
        let mut invalid = font;
        invalid.face_index = 2;
        assert!(standalone_font_data(&invalid).is_none());
        invalid.data = Arc::from(&invalid.data[..19]);
        assert!(standalone_font_data(&invalid).is_none());
    }

    #[test]
    fn invalid_collection_face_is_diagnosed_without_embedding() {
        let mut collection = test_font();
        collection.data = Arc::from(b"ttcf synthetic collection".as_slice());
        collection.face_index = 2;
        let mut layout = layout(vec![text("face", 4)]);
        layout.fonts = vec![collection];

        let result = render_page(&layout, 0).unwrap();
        assert!(!result.svg.contains("data:font/collection;base64,"));
        assert!(!result.svg.contains("@font-face"));
        assert_eq!(result.diagnostics.len(), 1);
        assert_eq!(result.diagnostics[0].path, "fonts[7]");
        assert!(result.diagnostics[0].message.contains("selected font face"));
        // Header count and table spans are checked before parser arithmetic.
        for bytes in [
            b"ttcf\0\x01\0\0\xff\xff\xff\xff".to_vec(),
            b"ttcf\0\x01\0\0\0\0\0\x01\xff\xff\xff\xff".to_vec(),
        ] {
            let mut font = test_font();
            font.data = Arc::from(bytes);
            font.face_index = u32::MAX;
            assert!(standalone_font_data(&font).is_none());
        }
        let mut invalid_table = test_font();
        let mut bytes = invalid_table.data.to_vec();
        bytes[20..24].copy_from_slice(&u32::MAX.to_be_bytes());
        bytes[24..28].copy_from_slice(&2u32.to_be_bytes());
        invalid_table.data = Arc::from(bytes);
        assert!(standalone_font_data(&invalid_table).is_none());
    }

    #[test]
    fn font_diagnostics_are_emitted_at_first_use_in_element_order() {
        let unsupported = PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            }),
            fill: Some(Paint::Tile {
                image: MediaId(1),
                tile: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0,
                },
                transform: Transform::IDENTITY,
            }),
            stroke: None,
        });
        let mut font = test_font();
        font.data = Arc::from(b"ttcf synthetic collection".as_slice());
        font.face_index = 2;
        let mut layout = layout(vec![unsupported, text("face", 4)]);
        layout.fonts = vec![font];

        let result = render_page(&layout, 0).unwrap();
        assert_eq!(result.diagnostics.len(), 2);
        assert_eq!(result.diagnostics[0].path, "pages[0].elements[0].fill");
        assert_eq!(result.diagnostics[1].path, "fonts[7]");
        assert!(result.diagnostics[1].message.contains("selected font face"));
    }

    #[test]
    fn multiple_outer_shadows_are_independent_and_do_not_clip_to_object_bounds() {
        let group = PositionedElement::Group(GroupElement {
            transform: Transform::IDENTITY,
            clip: None,
            opacity: 1.0,
            effects: vec![
                Effect::OuterShadow {
                    dx: 200.0,
                    dy: -150.0,
                    blur: 80.0,
                    color: Color::BLACK,
                },
                Effect::OuterShadow {
                    dx: -240.0,
                    dy: 175.0,
                    blur: 60.0,
                    color: color(1.0, 0.0, 0.0),
                },
            ],
            children: vec![PositionedElement::FilledRect {
                rect: Rect {
                    x: 1.0,
                    y: 1.0,
                    width: 1.0,
                    height: 20.0,
                },
                color: Color::BLACK,
            }],
        });

        let result = render_page(&layout(vec![group]), 0).unwrap();
        assert_eq!(result.svg.matches("in=\"SourceAlpha\"").count(), 2);
        assert_eq!(
            result
                .svg
                .matches("<feMergeNode in=\"rdocx-def-0-shadow-")
                .count(),
            2
        );
        assert_eq!(
            result
                .svg
                .matches("<feMergeNode in=\"SourceGraphic\"")
                .count(),
            1
        );
        assert!(
            result
                .svg
                .contains("x=\"0\" y=\"0\" width=\"120\" height=\"80\"")
        );
        assert!(!result.svg.contains("-50%"));
        assert!(!result.svg.contains("1000000"));
    }

    #[test]
    fn non_uniform_shadow_blur_is_diagnosed_without_dropping_output() {
        let group = PositionedElement::Group(GroupElement {
            transform: Transform {
                a: 4.0,
                d: 0.25,
                ..Transform::IDENTITY
            },
            clip: None,
            opacity: 1.0,
            effects: vec![Effect::OuterShadow {
                dx: 2.0,
                dy: 2.0,
                blur: 4.0,
                color: Color::BLACK,
            }],
            children: vec![PositionedElement::FilledRect {
                rect: Rect {
                    x: 2.0,
                    y: 8.0,
                    width: 10.0,
                    height: 20.0,
                },
                color: Color::BLACK,
            }],
        });

        let result = render_page(&layout(vec![group]), 0).unwrap();
        assert!(result.svg.contains("matrix(4 0 0 0.25 0 0)"));
        assert!(result.svg.contains("<filter "));
        assert!(result.svg.contains("<rect x=\"2\" y=\"8\""));
        assert_eq!(result.diagnostics.len(), 1);
        assert_eq!(
            result.diagnostics[0].path,
            "pages[0].elements[0].effects[0]"
        );
        assert_eq!(
            result.diagnostics[0].message,
            "non-uniform or skewed transform makes SVG shadow blur anisotropic instead of the raster backend's average-scale isotropic blur"
        );
        assert!(preserves_isotropic_blur(Transform::rotate_about(
            37.0, 0.0, 0.0
        )));
        assert!(!preserves_isotropic_blur(Transform {
            a: 1.0,
            c: 0.25,
            ..Transform::IDENTITY
        }));
    }

    #[test]
    fn filter_region_tracks_the_page_through_large_group_coordinates() {
        let group = PositionedElement::Group(GroupElement {
            transform: Transform {
                e: -2_000_000.0,
                ..Transform::IDENTITY
            },
            clip: None,
            opacity: 1.0,
            effects: vec![Effect::OuterShadow {
                dx: 2.0,
                dy: 2.0,
                blur: 2.0,
                color: Color::BLACK,
            }],
            children: vec![PositionedElement::FilledRect {
                rect: Rect {
                    x: 2_000_000.0,
                    y: 1.0,
                    width: 20.0,
                    height: 20.0,
                },
                color: Color::BLACK,
            }],
        });

        let result = render_page(&layout(vec![group]), 0).unwrap();
        assert!(
            result
                .svg
                .contains("x=\"2000000\" y=\"0\" width=\"120\" height=\"80\"")
        );
        assert!(!result.svg.contains("x=\"-1000000\""));
    }

    #[test]
    fn singular_group_filter_region_uses_source_and_effect_bounds() {
        let group = PositionedElement::Group(GroupElement {
            transform: Transform {
                a: 0.0,
                e: 10.0,
                ..Transform::IDENTITY
            },
            clip: None,
            opacity: 1.0,
            effects: vec![Effect::OuterShadow {
                dx: 50.0,
                dy: 0.0,
                blur: 2.0,
                color: Color::BLACK,
            }],
            children: vec![PositionedElement::FilledRect {
                rect: Rect {
                    x: 2_000_000.0,
                    y: 10.0,
                    width: 20.0,
                    height: 20.0,
                },
                color: Color::BLACK,
            }],
        });

        let result = render_page(&layout(vec![group]), 0).unwrap();
        assert!(
            result
                .svg
                .contains("x=\"2000000\" y=\"5\" width=\"75\" height=\"30\"")
        );
    }

    #[test]
    fn singular_filter_bounds_do_not_clip_mitered_stroke_pixels() {
        let mut stroke = Stroke::new(Paint::Solid(Color::BLACK), 10.0);
        stroke.join = LineJoin::Miter;
        let miter = PositionedElement::Path(PathElement {
            path: Path {
                commands: vec![
                    PathCommand::MoveTo(Point { x: 60.0, y: 50.6 }),
                    PathCommand::LineTo(Point { x: 20.0, y: 40.0 }),
                    PathCommand::LineTo(Point { x: 60.0, y: 29.4 }),
                ],
                fill_rule: FillRule::NonZero,
            },
            fill: None,
            stroke: Some(stroke),
        });
        let group = PositionedElement::Group(GroupElement {
            transform: Transform {
                a: 1.0,
                b: 0.0,
                c: 0.0,
                d: 0.0,
                e: 20.0,
                f: 40.0,
            },
            clip: None,
            opacity: 1.0,
            effects: vec![Effect::OuterShadow {
                dx: 2.0,
                dy: 0.0,
                blur: 2.0,
                color: Color::BLACK,
            }],
            children: vec![miter],
        });
        let layout = layout(vec![group]);

        let result = render_page(&layout, 0).unwrap();
        assert!(result.svg.contains("matrix(1 0 0 0 20 40)"));
        let renderable_svg = result
            .svg
            .replace("matrix(1 0 0 0 20 40)", "matrix(1 0 0 1 20 0)");
        let document = resvg::usvg::roxmltree::Document::parse(&renderable_svg).unwrap();
        let filter = document
            .descendants()
            .find(|node| node.has_tag_name("filter"))
            .unwrap();
        let x = filter.attribute("x").unwrap().parse::<f64>().unwrap();
        let y = filter.attribute("y").unwrap().parse::<f64>().unwrap();
        let width = filter.attribute("width").unwrap().parse::<f64>().unwrap();
        let height = filter.attribute("height").unwrap().parse::<f64>().unwrap();
        let original_tag_start = renderable_svg.find("<filter ").unwrap();
        let original_tag_end =
            renderable_svg[original_tag_start..].find('>').unwrap() + original_tag_start + 1;
        let original_tag = &renderable_svg[original_tag_start..original_tag_end];
        let expansion = layout.pages[0].width.max(layout.pages[0].height) * 4.0;
        let expanded_tag = original_tag
            .replace(
                &format!("x=\"{}\"", super::number(x)),
                &format!("x=\"{}\"", super::number(x - expansion)),
            )
            .replace(
                &format!("y=\"{}\"", super::number(y)),
                &format!("y=\"{}\"", super::number(y - expansion)),
            )
            .replace(
                &format!("width=\"{}\"", super::number(width)),
                &format!("width=\"{}\"", super::number(width + expansion * 2.0)),
            )
            .replace(
                &format!("height=\"{}\"", super::number(height)),
                &format!("height=\"{}\"", super::number(height + expansion * 2.0)),
            );
        let expanded_svg = format!(
            "{}{}{}",
            &renderable_svg[..original_tag_start],
            expanded_tag,
            &renderable_svg[original_tag_end..]
        );
        let rendered = rasterise_with_resvg(&renderable_svg, 250, 167, &layout.fonts);
        let expanded = rasterise_with_resvg(&expanded_svg, 250, 167, &layout.fonts);
        assert_eq!(rendered, expanded);
        assert!(
            rendered
                .chunks_exact(4)
                .any(|pixel| pixel[..3] != [255, 255, 255])
        );
    }

    #[test]
    fn singular_group_with_text_omits_only_the_unprovable_effect() {
        let mut unprovable = text("kept", 4);
        let PositionedElement::Text(run) = &mut unprovable else {
            unreachable!()
        };
        run.glyph_ids[0] = u16::MAX;
        assert!(glyph_ink_bounds(run, &[test_font()], &[], &[], &[]).is_err());
        let group = PositionedElement::Group(GroupElement {
            transform: Transform {
                d: 0.0,
                ..Transform::IDENTITY
            },
            clip: None,
            opacity: 1.0,
            effects: vec![Effect::OuterShadow {
                dx: 2.0,
                dy: 2.0,
                blur: 2.0,
                color: Color::BLACK,
            }],
            children: vec![
                unprovable,
                PositionedElement::FilledRect {
                    rect: Rect {
                        x: 50.0,
                        y: 10.0,
                        width: 15.0,
                        height: 20.0,
                    },
                    color: color(1.0, 0.0, 0.0),
                },
            ],
        });
        let layout = layout(vec![group]);

        let result = render_page(&layout, 0).unwrap();
        assert!(result.svg.contains(">kept</text>"));
        assert!(result.svg.contains("<rect x=\"50\" y=\"10\""));
        assert!(!result.svg.contains("<filter "));
        assert_eq!(result.diagnostics.len(), 2);
        assert_eq!(result.diagnostics[0].path, "pages[0].elements[0]");
        assert_eq!(
            result.diagnostics[0].message,
            "group effects were omitted because singular-transform source bounds could not be proven"
        );
        assert_eq!(
            result.diagnostics[1].path,
            "pages[0].elements[0].children[0]"
        );
        assert_eq!(
            result.diagnostics[1].message,
            "complex shaping kept searchable text with total-advance positioning"
        );
    }

    #[test]
    fn svg_geometry_preserves_gradient_stroke_background_and_dash_contracts() {
        let mut stroke = Stroke::new(
            Paint::Solid(Color {
                r: 0.0,
                g: 0.0,
                b: 1.0,
                a: 0.5,
            }),
            2.0,
        );
        stroke.cap = LineCap::Round;
        stroke.join = LineJoin::Bevel;
        stroke.dash = Some(vec![3.0, 4.0]);
        let mut layout = layout(vec![PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 10.0,
                y: 12.0,
                width: 30.0,
                height: 20.0,
            }),
            fill: Some(Paint::linear(
                Point { x: 10.0, y: 12.0 },
                Point { x: 40.0, y: 32.0 },
                vec![
                    GradientStop {
                        offset: 0.0,
                        color: Color::BLACK,
                    },
                    GradientStop {
                        offset: 1.0,
                        color: Color::WHITE,
                    },
                ],
                (true, true),
            )),
            stroke: Some(stroke),
        })]);
        layout.pages[0] = Arc::new({
            let mut page = PageFrame::new(1, 120.0, 80.0, layout.pages[0].elements.clone());
            page.background = Some(Paint::radial(
                Point { x: 60.0, y: 40.0 },
                40.0,
                Point { x: 55.0, y: 35.0 },
                vec![
                    GradientStop {
                        offset: 0.0,
                        color: Color::WHITE,
                    },
                    GradientStop {
                        offset: 1.0,
                        color: color(0.5, 0.5, 0.5),
                    },
                ],
                (true, true),
            ));
            page
        });

        let result = render_page(&layout, 0).unwrap();
        assert!(result.svg.contains("<radialGradient id=\"rdocx-def-0\""));
        assert!(result.svg.contains("<linearGradient id=\"rdocx-def-1\""));
        assert!(
            result
                .svg
                .contains("stroke=\"#0000FF\" stroke-opacity=\"0.5\"")
        );
        assert!(result.svg.contains("stroke-linecap=\"round\""));
        assert!(result.svg.contains("stroke-linejoin=\"bevel\""));
        assert!(result.svg.contains("stroke-dasharray=\"3 4\""));
        assert!(result.svg.contains("fill-rule=\"nonzero\""));
        assert!(result.diagnostics.is_empty());
    }

    #[test]
    fn gradient_stops_match_shared_backend_normalization() {
        let gradient = PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 40.0,
                height: 20.0,
            }),
            fill: Some(Paint::linear(
                Point { x: 0.0, y: 0.0 },
                Point { x: 40.0, y: 0.0 },
                vec![
                    GradientStop {
                        offset: 2.0,
                        color: color(0.0, 0.0, 1.0),
                    },
                    GradientStop {
                        offset: f64::NAN,
                        color: color(1.0, 0.0, 0.0),
                    },
                    GradientStop {
                        offset: -1.0,
                        color: color(0.0, 1.0, 0.0),
                    },
                    GradientStop {
                        offset: 0.5,
                        color: Color::WHITE,
                    },
                    GradientStop {
                        offset: 1.0,
                        color: color(1.0, 1.0, 0.0),
                    },
                ],
                (true, true),
            )),
            stroke: None,
        });

        let result = render_page(&layout(vec![gradient]), 0).unwrap();
        assert!(!result.svg.contains("NaN"));
        assert_eq!(result.svg.matches("<stop ").count(), 3);
        let green = result
            .svg
            .find("offset=\"0\" stop-color=\"#00FF00\"")
            .unwrap();
        let white = result
            .svg
            .find("offset=\"0.5\" stop-color=\"#FFFFFF\"")
            .unwrap();
        let yellow = result
            .svg
            .find("offset=\"1\" stop-color=\"#FFFF00\"")
            .unwrap();
        assert!(green < white && white < yellow);
    }

    #[test]
    fn layout_diagnostics_precede_lowering_diagnostics() {
        let mut layout = layout(vec![PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 1.0,
                height: 1.0,
            }),
            fill: Some(Paint::Tile {
                image: MediaId(9),
                tile: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0,
                },
                transform: Transform::IDENTITY,
            }),
            stroke: None,
        })]);
        layout.diagnostics.push(Diagnostic {
            message: "layout approximation".to_owned(),
        });
        let result = render_page(&layout, 0).unwrap();
        assert_eq!(result.diagnostics.len(), 2);
        assert_eq!(result.diagnostics[0].path, "layout.diagnostics[0]");
        assert_eq!(result.diagnostics[0].message, "layout approximation");
        assert_eq!(result.diagnostics[1].path, "pages[0].elements[0].fill");
    }

    #[test]
    fn non_pad_gradient_extension_is_diagnosed() {
        let result = render_page(
            &layout(vec![PositionedElement::Path(PathElement {
                path: Path::rect(Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 10.0,
                    height: 10.0,
                }),
                fill: Some(Paint::linear(
                    Point { x: 0.0, y: 0.0 },
                    Point { x: 10.0, y: 0.0 },
                    vec![
                        GradientStop {
                            offset: 0.0,
                            color: Color::BLACK,
                        },
                        GradientStop {
                            offset: 1.0,
                            color: Color::WHITE,
                        },
                    ],
                    (false, true),
                )),
                stroke: None,
            })]),
            0,
        )
        .unwrap();
        assert_eq!(result.diagnostics.len(), 1);
        assert_eq!(result.diagnostics[0].path, "pages[0].elements[0].fill");
    }

    fn luminance_ssim(left: &[u8], right: &[u8]) -> f64 {
        assert_eq!(left.len(), right.len());
        let left = left
            .chunks_exact(4)
            .map(|pixel| {
                0.2126 * f64::from(pixel[0])
                    + 0.7152 * f64::from(pixel[1])
                    + 0.0722 * f64::from(pixel[2])
            })
            .collect::<Vec<_>>();
        let right = right
            .chunks_exact(4)
            .map(|pixel| {
                0.2126 * f64::from(pixel[0])
                    + 0.7152 * f64::from(pixel[1])
                    + 0.0722 * f64::from(pixel[2])
            })
            .collect::<Vec<_>>();
        let count = left.len() as f64;
        let mean_left = left.iter().sum::<f64>() / count;
        let mean_right = right.iter().sum::<f64>() / count;
        let mut variance_left = 0.0;
        let mut variance_right = 0.0;
        let mut covariance = 0.0;
        for (left, right) in left.iter().zip(&right) {
            variance_left += (left - mean_left).powi(2);
            variance_right += (right - mean_right).powi(2);
            covariance += (left - mean_left) * (right - mean_right);
        }
        let divisor = (count - 1.0).max(1.0);
        variance_left /= divisor;
        variance_right /= divisor;
        covariance /= divisor;
        let c1 = (0.01_f64 * 255.0).powi(2);
        let c2 = (0.03_f64 * 255.0).powi(2);
        ((2.0 * mean_left * mean_right + c1) * (2.0 * covariance + c2))
            / ((mean_left.powi(2) + mean_right.powi(2) + c1)
                * (variance_left + variance_right + c2))
    }

    // The production formatter rounds to six decimals. Half a final decimal
    // plus floating summation noise is the entire allowed identity tolerance.
    fn identity_length_adjustment_for_usvg(svg: &str, natural_advance: f64) -> Option<String> {
        if !natural_advance.is_finite()
            || natural_advance <= 0.0
            || svg.matches(" textLength=\"").count() != 1
        {
            return None;
        }
        let start = svg.find(" textLength=\"")?;
        let value_start = start + " textLength=\"".len();
        let value_end = value_start + svg[value_start..].find('"')?;
        let target: f64 = svg[value_start..value_end].parse().ok()?;
        let suffix = "\" lengthAdjust=\"spacingAndGlyphs\"";
        if !target.is_finite()
            || (target - natural_advance).abs() > 0.000_000_500_1
            || !svg[value_end..].starts_with(suffix)
        {
            return None;
        }
        let mut corrected = svg.to_owned();
        corrected.replace_range(start..value_end + suffix.len(), "");
        Some(corrected)
    }

    fn independent_static_advance(bytes: &[u8], text: &str, script: &str, weight: u32) -> f64 {
        let version = std::process::Command::new("hb-shape")
            .arg("--version")
            .output()
            .unwrap();
        assert!(version.status.success());
        assert!(
            String::from_utf8_lossy(&version.stdout).starts_with("hb-shape (HarfBuzz) 14.4.0\n")
        );
        let directory = std::env::temp_dir().join(format!(
            "fx195-svg-hb-width-{}-{script}-{weight}",
            std::process::id()
        ));
        std::fs::create_dir(&directory).unwrap();
        let file = directory.join("static.ttf");
        std::fs::write(&file, bytes).unwrap();
        let result = std::process::Command::new("hb-shape")
            .args([
                "--output-format=json",
                "--no-glyph-names",
                "--font-size=24000",
            ])
            .arg(&file)
            .arg(text)
            .output();
        std::fs::remove_file(&file).unwrap();
        std::fs::remove_dir(&directory).unwrap();
        let result = result.unwrap();
        assert!(
            result.status.success(),
            "independent static width oracle failed"
        );
        let glyphs: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
        let glyphs = glyphs.as_array().unwrap();
        assert!(!glyphs.is_empty());
        glyphs
            .iter()
            .map(|glyph| glyph["ax"].as_i64().unwrap() as f64 / 1000.0)
            .sum()
    }

    fn rasterise_with_resvg(svg: &str, width: u32, height: u32, fonts: &[FontData]) -> Vec<u8> {
        let mut options = resvg::usvg::Options::default();
        assert!(options.fontdb.faces().next().is_none());
        let mut exact_faces = HashMap::new();
        for font in fonts {
            let ids = options
                .fontdb_mut()
                .load_font_source(resvg::usvg::fontdb::Source::Binary(Arc::new(
                    font.data.to_vec(),
                )));
            let id = ids
                .get(font.face_index as usize)
                .copied()
                .expect("layout font face is loadable by the resvg oracle");
            exact_faces.insert(format!("rdocx-font-{}", font.id.0), id);
        }
        options.font_resolver = resvg::usvg::FontResolver {
            select_font: Box::new(move |font, _| {
                font.families().iter().find_map(|family| {
                    let rendered = family.to_string();
                    exact_faces.get(rendered.trim_matches('"')).copied()
                })
            }),
            select_fallback: Box::new(|_, _, _| None),
        };
        let tree = resvg::usvg::Tree::from_str(svg, &options).expect("resvg parses generated SVG");
        let mut pixmap = resvg::tiny_skia::Pixmap::new(width, height).unwrap();
        let scale_x = width as f32 / tree.size().width();
        let scale_y = height as f32 / tree.size().height();
        resvg::render(
            &tree,
            resvg::tiny_skia::Transform::from_scale(scale_x, scale_y),
            &mut pixmap.as_mut(),
        );
        pixmap.data().to_vec()
    }

    fn golden_text() -> (PositionedElement, FontData) {
        let bytes = include_bytes!("../../oxml-layout/fonts/Carlito-Regular.ttf").to_vec();
        let mut fonts = FontManager::new_with_fonts(vec![("Golden".to_owned(), bytes)]);
        let font_id = fonts.resolve_font(Some("Golden"), false, false).unwrap();
        let shaped = fonts.shape_text(font_id, "SVG TEXT", 15.0).unwrap();
        assert_eq!(shaped.glyph_ids.len(), "SVG TEXT".chars().count());
        let font = fonts.font_data(font_id).unwrap();
        (
            PositionedElement::Text(GlyphRun {
                origin: Point { x: 8.0, y: 22.0 },
                font_id,
                font_size: 15.0,
                glyph_ids: shaped.glyph_ids,
                advances: shaped.advances,
                text: "SVG TEXT".to_owned(),
                source: None,
                color: Color::BLACK,
                bold: false,
                italic: false,
                field_kind: None,
                field_source: None,
                note: None,
                note_reference_source: None,
                tab_aligned: None,
            }),
            font,
        )
    }

    fn golden_png() -> Vec<u8> {
        let mut pixmap = resvg::tiny_skia::Pixmap::new(16, 12).unwrap();
        for (index, pixel) in pixmap.data_mut().chunks_exact_mut(4).enumerate() {
            let x = index % 16;
            let y = index / 16;
            let dark = (x / 4 + y / 3) % 2 == 0;
            pixel.copy_from_slice(if dark {
                &[24, 72, 160, 255]
            } else {
                &[240, 184, 40, 255]
            });
        }
        pixmap.encode_png().unwrap()
    }

    fn representative_golden_layout() -> LayoutResult {
        let (text, font) = golden_text();
        let gradient = PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 6.0,
                y: 32.0,
                width: 132.0,
                height: 34.0,
            }),
            fill: Some(Paint::linear(
                Point { x: 6.0, y: 32.0 },
                Point { x: 138.0, y: 66.0 },
                vec![
                    GradientStop {
                        offset: 1.0,
                        color: color(0.05, 0.2, 0.55),
                    },
                    GradientStop {
                        offset: 0.0,
                        color: color(0.9, 0.25, 0.1),
                    },
                    GradientStop {
                        offset: 0.55,
                        color: color(0.95, 0.8, 0.15),
                    },
                ],
                (true, true),
            )),
            stroke: Some(Stroke::new(Paint::Solid(Color::BLACK), 1.0)),
        });
        let skewed_content = PositionedElement::Group(GroupElement {
            transform: Transform {
                b: 0.015,
                c: 0.025,
                ..Transform::IDENTITY
            },
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![PositionedElement::Path(PathElement {
                path: Path::rect(Rect {
                    x: -8.0,
                    y: 4.0,
                    width: 125.0,
                    height: 34.0,
                }),
                fill: Some(Paint::Solid(color(0.1, 0.65, 0.35))),
                stroke: Some(Stroke::new(Paint::Solid(Color::BLACK), 2.0)),
            })],
        });
        let angle = 3.0_f64.to_radians();
        let (sin, cos) = angle.sin_cos();
        let clipped_content = PositionedElement::Group(GroupElement {
            transform: Transform {
                a: cos,
                b: sin,
                c: -sin,
                d: cos,
                e: 5.0,
                f: 3.0,
            },
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![skewed_content],
        });
        let effected_group = PositionedElement::Group(GroupElement {
            transform: Transform {
                a: 1.03,
                d: 0.97,
                e: 9.0,
                f: 82.0,
                ..Transform::IDENTITY
            },
            clip: Some(Path::rect(Rect {
                x: 5.0,
                y: 3.0,
                width: 105.0,
                height: 43.0,
            })),
            opacity: 0.82,
            effects: vec![Effect::OuterShadow {
                dx: 4.0,
                dy: 4.0,
                blur: 1.0,
                color: Color {
                    a: 0.55,
                    ..Color::BLACK
                },
            }],
            children: vec![PositionedElement::MarkedContent {
                structure: None,
                children: vec![clipped_content],
            }],
        });
        let fallback = PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 130.0,
                y: 130.0,
                width: 4.0,
                height: 4.0,
            }),
            fill: Some(Paint::Tile {
                image: MediaId(99),
                tile: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 2.0,
                    height: 2.0,
                },
                transform: Transform::IDENTITY,
            }),
            stroke: None,
        });
        LayoutResult::new(
            vec![Arc::new(PageFrame::new(
                1,
                144.0,
                144.0,
                vec![
                    gradient,
                    text,
                    PositionedElement::Image {
                        rect: Rect {
                            x: 102.0,
                            y: 7.0,
                            width: 34.0,
                            height: 20.0,
                        },
                        content_type: "image/png".to_owned(),
                        media_id: MediaId(5),
                        data: golden_png(),
                    },
                    effected_group,
                    PositionedElement::LinkAnnotation {
                        rect: Rect {
                            x: 8.0,
                            y: 5.0,
                            width: 80.0,
                            height: 20.0,
                        },
                        url: "https://example.test/golden".to_owned(),
                    },
                    fallback,
                ],
            ))],
            vec![font],
            None,
            Vec::new(),
        )
    }

    #[test]
    fn fx205_svg_colour_grid_preserves_origin_alpha_and_repeated_source_cells() {
        for (origin, samples) in [
            (
                (8.0, 8.0),
                [
                    [255, 0, 0],
                    [127, 127, 255],
                    [0, 255, 0],
                    [191, 191, 191],
                    [255, 0, 0],
                    [0, 255, 0],
                ],
            ),
            (
                (-3.0, -5.0),
                [
                    [191, 191, 191],
                    [0, 255, 0],
                    [127, 127, 255],
                    [255, 0, 0],
                    [191, 191, 191],
                    [127, 127, 255],
                ],
            ),
        ] {
            let paint = Paint::grid(
                Rect {
                    x: origin.0,
                    y: origin.1,
                    width: 8.0,
                    height: 8.0,
                },
                2,
                2,
                vec![
                    color(1.0, 0.0, 0.0),
                    Color {
                        r: 0.0,
                        g: 0.0,
                        b: 1.0,
                        a: 0.5,
                    },
                    color(0.0, 1.0, 0.0),
                    Color {
                        r: 0.0,
                        g: 0.0,
                        b: 0.0,
                        a: 0.25,
                    },
                ],
            )
            .unwrap();
            let scene = layout(vec![PositionedElement::Path(PathElement {
                path: Path::rect(Rect {
                    x: 8.0,
                    y: 8.0,
                    width: 32.0,
                    height: 32.0,
                }),
                fill: Some(paint),
                stroke: None,
            })]);
            let result = render_page(&scene, 0).unwrap();
            assert!(result.diagnostics.is_empty());
            let pattern = format!(
                "patternUnits=\"userSpaceOnUse\" x=\"{}\" y=\"{}\" width=\"8\" height=\"8\"",
                origin.0, origin.1
            );
            assert!(result.svg.contains(&pattern));
            let png = oxml_pdf::render_page_to_png(&scene, 0, 72.0).unwrap();
            let native = resvg::tiny_skia::Pixmap::decode_png(&png).unwrap();
            let svg =
                rasterise_with_resvg(&result.svg, native.width(), native.height(), &scene.fonts);
            // Interior samples derive directly from the four source cells and white
            // background. Repeat positions also decide the nonzero pattern origin.
            for ((x, y), expected) in [(9, 9), (13, 9), (9, 13), (13, 13), (17, 9), (25, 13)]
                .into_iter()
                .zip(samples)
            {
                let offset = ((y * native.width() + x) * 4) as usize;
                for pixels in [native.data(), svg.as_slice()] {
                    for channel in 0..3 {
                        assert!(
                            (i16::from(pixels[offset + channel]) - expected[channel]).abs() <= 1
                        );
                    }
                    assert_eq!(pixels[offset + 3], 255);
                }
            }
            let score = luminance_ssim(native.data(), &svg);
            assert!(score >= 0.99, "Grid SVG/raster SSIM {score}");
            let shifted_pattern = format!(
                "patternUnits=\"userSpaceOnUse\" x=\"{}\" y=\"{}\" width=\"8\" height=\"8\"",
                origin.0 + 1.0,
                origin.1
            );
            let shifted = result.svg.replacen(&pattern, &shifted_pattern, 1);
            assert_ne!(shifted, result.svg);
            let shifted =
                rasterise_with_resvg(&shifted, native.width(), native.height(), &scene.fonts);
            let score = luminance_ssim(native.data(), &shifted);
            assert!(score < 0.99, "One-point Grid origin shift SSIM {score}");
        }
    }

    #[test]
    fn svg_page_rasterises_like_the_png_backend() {
        assert_eq!(RESVG_ORACLE_VERSION, "resvg 0.48.1");
        let layout = representative_golden_layout();
        let result = render_page(&layout, 0).unwrap();
        assert_eq!(result.diagnostics.len(), 2);
        assert_eq!(
            result.diagnostics[0].path,
            "pages[0].elements[3].effects[0]"
        );
        assert_eq!(result.diagnostics[1].path, "pages[0].elements[5].fill");
        assert!(result.svg.contains("<text "));
        assert!(result.svg.contains("data:image/png;base64,"));
        assert!(result.svg.contains("<linearGradient "));
        assert_eq!(result.svg.matches("transform=\"matrix(").count(), 3);
        assert!(result.svg.contains("clip-path=\"url(#"));
        assert!(result.svg.contains(" opacity=\"0.82\""));
        assert!(result.svg.contains("<filter "));
        assert!(result.svg.contains("<a href="));
        assert!(result.svg.contains("<g><g"));
        let svg = result.svg;
        let png = oxml_pdf::render_page_to_png(&layout, 0, 150.0).unwrap();
        let png = resvg::tiny_skia::Pixmap::decode_png(&png).unwrap();
        assert_eq!((png.width(), png.height()), (300, 300));
        let rasterised = rasterise_with_resvg(&svg, png.width(), png.height(), &layout.fonts);
        let score = luminance_ssim(png.data(), &rasterised);
        assert!(score >= 0.99, "{RESVG_ORACLE_VERSION}: SSIM {score}");

        let perturbed = svg.replacen("viewBox=\"0 0 ", "viewBox=\"1 0 ", 1);
        let perturbed = rasterise_with_resvg(&perturbed, png.width(), png.height(), &layout.fonts);
        let perturbed_score = luminance_ssim(png.data(), &perturbed);
        assert!(
            perturbed_score < 0.99,
            "one-point perturbation unexpectedly passed at SSIM {perturbed_score}"
        );
    }
    #[test]
    #[ignore = "mandatory F-X195 installed-artifact acceptance stage, execute exact named test with --ignored after installed104 qualification"]
    fn fx195_installed_svg_styles_paint_authenticated_independent_contours() {
        use sha2::{Digest, Sha256};
        use std::collections::BTreeSet;
        use std::path::Path;

        fn digest(bytes: &[u8]) -> String {
            Sha256::digest(bytes)
                .iter()
                .map(|byte| format!("{byte:02x}"))
                .collect()
        }
        fn checked(path: &str, expected: &str) -> Vec<u8> {
            let bytes = std::fs::read(path).expect("authenticated artifact is readable");
            assert_eq!(digest(&bytes), expected, "artifact changed: {path}");
            bytes
        }
        fn string<'a>(row: &'a serde_json::Value, key: &str) -> &'a str {
            row[key].as_str().expect("required admitted string")
        }
        fn number(row: &serde_json::Value, key: &str) -> u32 {
            u32::try_from(row[key].as_u64().expect("required admitted integer")).unwrap()
        }
        fn crop(bytes: &[u8], width: u32, roi: [u32; 4]) -> Vec<u8> {
            let mut result = Vec::new();
            for y in roi[1]..roi[3] {
                let start = ((y * width + roi[0]) * 4) as usize;
                let end = ((y * width + roi[2]) * 4) as usize;
                result.extend_from_slice(&bytes[start..end]);
            }
            result
        }
        fn variable_owners(svg: &str) -> Vec<(std::collections::BTreeMap<String, String>, String)> {
            use quick_xml::events::Event;
            use quick_xml::name::ResolveResult;
            let mut reader = quick_xml::reader::NsReader::from_str(svg);
            let mut owners = Vec::new();
            let mut current = None;
            loop {
                let (namespace, event) = reader.read_resolved_event().unwrap();
                let svg_namespace = matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == b"http://www.w3.org/2000/svg");
                match event {
                    Event::Start(element) if element.local_name().as_ref() == b"text" => {
                        assert!(svg_namespace && current.is_none());
                        let attributes = element
                            .attributes()
                            .map(|attribute| {
                                let attribute = attribute.unwrap();
                                (
                                    std::str::from_utf8(attribute.key.as_ref())
                                        .unwrap()
                                        .to_owned(),
                                    attribute
                                        .decoded_and_normalized_value(
                                            quick_xml::XmlVersion::Implicit1_0,
                                            reader.decoder(),
                                        )
                                        .unwrap()
                                        .into_owned(),
                                )
                            })
                            .collect::<std::collections::BTreeMap<_, _>>();
                        current = Some((attributes, String::new()));
                    }
                    Event::Text(text) => {
                        if let Some((_, contents)) = current.as_mut() {
                            contents.push_str(&text.decode().unwrap());
                        }
                    }
                    Event::End(element) if element.local_name().as_ref() == b"text" => {
                        assert!(svg_namespace);
                        owners.push(current.take().expect("actual opened text owner"));
                    }
                    Event::Eof => break,
                    _ => assert!(
                        current.is_none(),
                        "authenticated positioned Thai owners have no nested or synthetic payload"
                    ),
                }
            }
            assert!(current.is_none());
            owners
        }
        // ROOT executes this named stage serially after installed tests. Python
        // never invokes Cargo. Missing admission is a failure, never a skip.
        assert_eq!(std::env::var("F_X195_ROOT_SERIAL_SVG_STAGE").unwrap(), "1");
        let admission_path = std::env::var("F_X195_INSTALLED_SVG_ADMISSION").unwrap();
        let admission_sha = std::env::var("F_X195_INSTALLED_SVG_ADMISSION_SHA").unwrap();
        let admission: serde_json::Value =
            serde_json::from_slice(&checked(&admission_path, &admission_sha)).unwrap();
        assert_eq!(admission["installedActualPassed"].as_u64(), Some(104));
        assert_eq!(
            admission["pairedFreshWheelsAuthenticated"].as_bool(),
            Some(true)
        );
        assert_eq!(
            admission["rootActualTerminalAuthenticated"].as_bool(),
            Some(true)
        );
        checked(
            string(&admission, "installedReceipt"),
            string(&admission, "installedReceiptSHA"),
        );
        let wheels = admission["wheels"].as_array().unwrap();
        assert_eq!(wheels.len(), 2);
        for wheel in wheels {
            checked(string(wheel, "path"), string(wheel, "sha256"));
        }
        let sources: serde_json::Value = serde_json::from_slice(&checked(
            string(&admission, "sourceInputs"),
            string(&admission, "sourceInputsSHA"),
        ))
        .unwrap();
        assert_eq!(sources.as_object().unwrap().len(), 317);
        assert_eq!(
            admission["sourceModesAuthenticatedByRoot"].as_bool(),
            Some(true)
        );
        let root = Path::new(string(&admission, "sourceRoot"));
        for (relative, guard) in sources.as_object().unwrap() {
            let path = root.join(relative);
            let bytes = checked(path.to_str().unwrap(), string(guard, "sha256"));
            assert_eq!(bytes.len() as u64, guard["size"].as_u64().unwrap());
            // Native source modes are authenticated by ROOT before admission.
            // Byte/size guards here remain portable across native targets.
        }
        let required: BTreeSet<String> = [
            "genuine-italic",
            "ital-axis",
            "slnt-axis",
            "synthetic-bold",
            "synthetic-italic",
            "synthetic-combined",
            "variable-400",
            "variable-700",
        ]
        .into_iter()
        .map(str::to_owned)
        .collect();
        let mut observed = BTreeSet::new();
        let rows = admission["artifacts"].as_array().unwrap();
        assert_eq!(rows.len(), required.len());
        for row in rows {
            let artifact: serde_json::Value =
                serde_json::from_slice(&checked(string(row, "receipt"), string(row, "receiptSHA")))
                    .unwrap();
            let case = string(&artifact, "case");
            assert!(
                observed.insert(case.to_owned()),
                "duplicate installed owner"
            );
            assert!(required.contains(case));
            assert_eq!(artifact["threshold"].as_f64(), Some(0.99));
            let svg_bytes = checked(string(&artifact, "svg"), string(&artifact, "svgSHA"));
            let svg = std::str::from_utf8(&svg_bytes).unwrap();
            let family = string(&artifact, "fontFamily");
            let font_id: u32 = family.strip_prefix("rdocx-font-").unwrap().parse().unwrap();
            let definition = format!("@font-face{{font-family:'{family}';src:url('data:");
            let tail = &svg[svg.find(&definition).unwrap() + definition.len()..];
            let data_tail = &tail[tail.find(";base64,").unwrap() + ";base64,".len()..];
            let embedded = BASE64
                .decode(&data_tail[..data_tail.find('\'').unwrap()])
                .unwrap();
            assert_eq!(digest(&embedded), string(&artifact, "fontSHA"));
            assert_eq!(artifact["fontIndex"].as_u64(), Some(0));
            let selected = if let Some(path) = artifact["oracleFont"].as_str() {
                // usvg does not apply variable CSS coordinates. This resolver
                // adapter is admitted only with independently selected physical
                // static bytes plus a verified actual emitted CSS selector.
                let selector = artifact["selector"].as_str().unwrap_or_else(|| {
                    match artifact["axis"].as_str().unwrap() {
                        "ital" => "'ital' 1",
                        "slnt" => "'slnt' -12",
                        _ => panic!("unknown axis"),
                    }
                });
                if case == "variable-400" || case == "variable-700" {
                    let weight = if case == "variable-400" { 400.0 } else { 700.0 };
                    let face = ttf_parser::Face::parse(&embedded, 0).unwrap();
                    let defaults = face
                        .variation_axes()
                        .into_iter()
                        .map(|axis| {
                            (
                                std::str::from_utf8(&axis.tag.to_bytes())
                                    .unwrap()
                                    .to_owned(),
                                f64::from(axis.def_value),
                            )
                        })
                        .collect::<std::collections::BTreeMap<_, _>>();
                    assert_eq!(
                        defaults,
                        std::collections::BTreeMap::from([
                            ("wght".to_owned(), 400.0),
                            ("wdth".to_owned(), 100.0)
                        ])
                    );
                    assert_eq!(
                        artifact["authenticatedFvarDefaults"],
                        serde_json::json!({"wght": 400.0, "wdth": 100.0})
                    );
                    let owners = variable_owners(svg);
                    assert_eq!(owners.len(), 5);
                    assert_eq!(
                        owners
                            .iter()
                            .map(|(_, text)| text.as_str())
                            .collect::<Vec<_>>(),
                        ["ภาษา", "ไทยยิน", "ดี", "ต้อน", "รับ"]
                    );
                    let declared = artifact["owners"].as_array().unwrap();
                    assert_eq!(declared.len(), owners.len());
                    for ((attributes, text), proof) in owners.iter().zip(declared) {
                        assert_eq!(text, string(proof, "text"));
                        assert_eq!(attributes["font-family"], family);
                        assert_eq!(string(proof, "family"), family);
                        assert_eq!(string(proof, "physicalFontSHA"), digest(&embedded));
                        assert_eq!(attributes["font-style"], "normal");
                        assert_eq!(
                            attributes["font-weight"],
                            if weight == 700.0 { "bold" } else { "normal" }
                        );
                        for refused in ["stroke", "transform", "filter", "textLength"] {
                            assert!(!attributes.contains_key(refused));
                        }
                        assert_eq!(
                            attributes["x"].split_whitespace().count(),
                            text.chars().count()
                        );
                        attributes["y"].parse::<f64>().unwrap();
                        let mut effective = defaults.clone();
                        let mut explicit = std::collections::BTreeMap::new();
                        if let Some(style) = attributes.get("style") {
                            let payload = style
                                .strip_prefix("font-synthesis:none;font-variation-settings:")
                                .unwrap();
                            for declaration in payload.split(',') {
                                let (tag, value) = declaration.trim().split_once(' ').unwrap();
                                let tag = tag
                                    .strip_prefix('\'')
                                    .and_then(|tag| tag.strip_suffix('\''))
                                    .unwrap();
                                assert!(tag == "wght" || tag == "wdth");
                                assert!(
                                    explicit
                                        .insert(tag.to_owned(), value.parse::<f64>().unwrap())
                                        .is_none()
                                );
                            }
                            assert!(!explicit.is_empty());
                            effective.extend(explicit.clone());
                        } else {
                            assert_eq!(weight, 400.0);
                            assert_eq!(proof["implicitDefault"].as_bool(), Some(true));
                        }
                        assert_eq!(
                            effective,
                            std::collections::BTreeMap::from([
                                ("wght".to_owned(), weight),
                                ("wdth".to_owned(), 100.0)
                            ])
                        );
                        if weight == 700.0 {
                            assert_eq!(explicit.get("wght"), Some(&700.0));
                        }
                        assert_eq!(
                            proof["effectiveCoordinates"],
                            serde_json::json!({"wght": weight, "wdth": 100.0})
                        );
                    }
                } else {
                    assert!(svg.contains(selector));
                }
                checked(path, string(&artifact, "oracleFontSHA"))
            } else {
                embedded
            };
            let font = FontData {
                id: FontId(font_id),
                family: family.to_owned(),
                data: Arc::from(selected),
                face_index: 0,
                bold: false,
                italic: false,
                variations: vec![],
                synthetic_bold: false,
                synthetic_italic: false,
            };
            let width = number(&artifact, "width");
            let height = number(&artifact, "height");
            let actual = rasterise_with_resvg(svg, width, height, std::slice::from_ref(&font));
            let expected = if artifact["expectedSVG"].is_string() {
                let bytes = checked(
                    string(&artifact, "expectedSVG"),
                    string(&artifact, "expectedSVGSHA"),
                );
                // Independent hb-subset static source, already proven to match
                // actual source HB positions and contours, supplies expected.
                rasterise_with_resvg(std::str::from_utf8(&bytes).unwrap(), width, height, &[font])
            } else {
                let bytes = checked(
                    string(&artifact, "expected"),
                    string(&artifact, "expectedSHA"),
                );
                let png = resvg::tiny_skia::Pixmap::decode_png(&bytes).unwrap();
                assert_eq!((png.width(), png.height()), (width, height));
                png.data().to_vec()
            };
            assert_eq!(actual.len(), expected.len());
            let references: Vec<(&str, Vec<u8>)> = [
                ("quadratic", "expectedQuadratic", "expectedQuadraticSHA"),
                ("cubic", "expectedCubic", "expectedCubicSHA"),
            ]
            .into_iter()
            .map(|(label, path, hash)| {
                let bytes = checked(string(&artifact, path), string(&artifact, hash));
                let image = resvg::tiny_skia::Pixmap::decode_png(&bytes).unwrap();
                assert_eq!((image.width(), image.height()), (width, height));
                (label, image.data().to_vec())
            })
            .collect();
            let roi = if let Some(values) = artifact["roi"].as_array() {
                assert_eq!(values.len(), 4);
                std::array::from_fn(|i| u32::try_from(values[i].as_u64().unwrap()).unwrap())
            } else {
                // ROI derives only from independent expected ink. No actual
                // image registration or actual-mask bounds can hide a shift.
                let mut bounds = [width, height, 0, 0];
                for y in 0..height {
                    for x in 0..width {
                        let pixel = ((y * width + x) * 4) as usize;
                        if expected[pixel..pixel + 3]
                            .iter()
                            .any(|channel| *channel < 200)
                        {
                            bounds[0] = bounds[0].min(x);
                            bounds[1] = bounds[1].min(y);
                            bounds[2] = bounds[2].max(x + 1);
                            bounds[3] = bounds[3].max(y + 1);
                        }
                    }
                }
                [
                    bounds[0].saturating_sub(3),
                    bounds[1].saturating_sub(3),
                    (bounds[2] + 3).min(width),
                    (bounds[3] + 3).min(height),
                ]
            };
            assert!(roi[0] < roi[2] && roi[1] < roi[3] && roi[2] <= width && roi[3] <= height);
            let score = luminance_ssim(&crop(&actual, width, roi), &crop(&expected, width, roi));
            assert!(
                score >= 0.99,
                "installed SVG {case} original static paint SSIM {score}"
            );
            for (curve, reference) in &references {
                assert_eq!(actual.len(), reference.len());
                let score =
                    luminance_ssim(&crop(&actual, width, roi), &crop(reference, width, roi));
                assert!(
                    score >= 0.99,
                    "installed SVG {case} {curve} independent paint SSIM {score}"
                );
                eprintln!("authenticated installed SVG case={case} curve={curve} SSIM={score}");
            }
        }
        assert_eq!(observed, required);
        for (relative, guard) in sources.as_object().unwrap() {
            let path = root.join(relative);
            let bytes = checked(path.to_str().unwrap(), string(guard, "sha256"));
            assert_eq!(bytes.len() as u64, guard["size"].as_u64().unwrap());
            // Native source modes are authenticated by ROOT before admission.
            // Byte/size guards here remain portable across native targets.
        }
    }

    #[test]
    fn linear_opacity_mask_svg_keeps_one_text_owner_and_multiplies_masks() {
        let mask = Effect::LinearOpacityMask {
            bounds: Rect {
                x: 0.0,
                y: 0.0,
                width: 100.0,
                height: 100.0,
            },
            start: Point { x: 0.0, y: 0.0 },
            end: Point { x: 100.0, y: 0.0 },
            stops: vec![
                GradientStop {
                    offset: 0.0,
                    color: Color::WHITE,
                },
                GradientStop {
                    offset: 1.0,
                    color: Color {
                        a: 0.0,
                        ..Color::BLACK
                    },
                },
            ],
        };
        let mut g = GroupElement {
            transform: Transform {
                e: 5.0,
                f: 7.0,
                ..Transform::IDENTITY
            },
            clip: Some(Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 90.0,
                height: 90.0,
            })),
            opacity: 0.5,
            effects: vec![
                mask.clone(),
                mask,
                Effect::OuterShadow {
                    dx: 2.0,
                    dy: 3.0,
                    blur: 0.0,
                    color: Color::BLACK,
                },
            ],
            children: vec![text("Once", 4)],
        };
        let output = render_page(&layout(vec![PositionedElement::Group(g.clone())]), 0).unwrap();
        assert_eq!(output.svg.matches(">Once</text>").count(), 1);
        assert_eq!(output.svg.matches("<mask ").count(), 2);
        assert_eq!(output.svg.matches("<linearGradient ").count(), 2);
        assert!(output.svg.contains("mask-type:alpha"));
        assert!(output.svg.contains("stop-opacity=\"0\""));
        assert!(output.svg.find(" mask=\"").unwrap() < output.svg.find(" filter=\"").unwrap());
        if let Effect::LinearOpacityMask { bounds, .. } = &mut g.effects[0] {
            bounds.width = f64::NAN;
        }
        let invalid = render_page(&layout(vec![PositionedElement::Group(g)]), 0).unwrap();
        assert!(invalid.svg.contains("<g opacity=\"0\">"));
        assert!(!invalid.svg.contains("NaN"));
        assert_eq!(invalid.svg.matches(">Once</text>").count(), 1);
    }

    #[test]
    fn linear_opacity_mask_svg_raster_agree_on_colored_background_and_overlap() {
        let mask = Effect::LinearOpacityMask {
            bounds: Rect {
                x: 0.0,
                y: 0.0,
                width: 24.0,
                height: 16.0,
            },
            start: Point { x: 0.0, y: 0.0 },
            end: Point { x: 24.0, y: 0.0 },
            stops: vec![
                GradientStop {
                    offset: 0.0,
                    color: Color::WHITE,
                },
                GradientStop {
                    offset: 1.0,
                    color: Color {
                        a: 0.0,
                        ..Color::BLACK
                    },
                },
            ],
        };
        let rect = PositionedElement::FilledRect {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: 24.0,
                height: 16.0,
            },
            color: Color::BLACK,
        };
        let g = PositionedElement::Group(GroupElement {
            transform: Transform {
                a: 1.5,
                d: 1.5,
                e: 4.0,
                f: 5.0,
                ..Transform::IDENTITY
            },
            clip: None,
            opacity: 0.5,
            effects: vec![mask.clone(), mask],
            children: vec![rect.clone(), rect],
        });
        let mut scene = layout(vec![g]);
        std::sync::Arc::make_mut(&mut scene.pages[0]).background =
            Some(Paint::Solid(Color::from_hex("2080C0")));
        let svg = render_page(&scene, 0).unwrap();
        let png = oxml_pdf::render_page_to_png(&scene, 0, 72.0).unwrap();
        let native = resvg::tiny_skia::Pixmap::decode_png(&png).unwrap();
        let independent =
            rasterise_with_resvg(&svg.svg, native.width(), native.height(), &scene.fonts);
        for (x, y) in [(5, 10), (12, 10), (22, 10), (35, 10), (50, 10)] {
            let i = (y * native.width() as usize + x) * 4;
            for channel in 0..3 {
                assert!(
                    (i16::from(native.data()[i + channel]) - i16::from(independent[i + channel]))
                        .abs()
                        <= 3,
                    "{x},{y},{channel}"
                );
            }
        }
    }
}
