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

pub(crate) fn render_page(layout: &LayoutResult, page_index: usize) -> Option<SvgRenderResult> {
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
            let (mime, format) = font_content_type(font);
            write!(
                rules,
                "@font-face{{font-family:'rdocx-font-{}';src:url('data:{mime};base64,{}') format('{format}');font-weight:{};font-style:{}}}",
                font.id.0,
                BASE64.encode(font.data.as_ref()),
                if font.bold { "bold" } else { "normal" },
                if font.italic { "italic" } else { "normal" }
            )
            .unwrap();
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
                PositionedElement::Text(run) => self.emit_text(run, &element_path, output),
                PositionedElement::MultilingualText(run) => {
                    if !run.is_valid() {
                        self.diagnose(
                            &element_path,
                            "invalid multilingual glyph positioning was omitted from SVG output",
                        );
                        continue;
                    }
                    let diagnostic_count = self.diagnostics.len();
                    self.emit_text(&run.legacy_projection(), &element_path, output);
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

    fn emit_text(&mut self, run: &GlyphRun, path: &str, output: &mut String) {
        if run.text.is_empty() {
            return;
        }
        self.diagnose_font_on_first_use(run.font_id);

        write!(
            output,
            "<text xml:space=\"preserve\" y=\"{}\" font-family=\"rdocx-font-{}\" font-size=\"{}\" font-weight=\"{}\" font-style=\"{}\" fill=\"{}\" fill-opacity=\"{}\"",
            number(run.origin.y),
            run.font_id.0,
            number(run.font_size),
            if run.bold { "bold" } else { "normal" },
            if run.italic { "italic" } else { "normal" },
            color_hex(run.color),
            number(run.color.a.clamp(0.0, 1.0))
        )
        .unwrap();

        let scalar_count = run.text.chars().count();
        if scalar_count == run.glyph_ids.len() && run.advances.len() == run.glyph_ids.len() {
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
            self.diagnose(
                path,
                "complex shaping kept searchable text with total-advance positioning",
            );
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
        if font.face_index != 0 {
            self.diagnose(
                &format!("fonts[{}]", font.id.0),
                "font collection face index cannot be selected by SVG and uses the default face",
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
                        write!(attributes, " filter=\"url(#{id})\"").unwrap();
                    }
                    Some(Ok(None)) | None => {}
                    Some(Err(())) => self.diagnose(
                        path,
                        "group effects were omitted because singular-transform source bounds could not be proven",
                    ),
                }
            }
        }

        write!(output, "<g{attributes}>").unwrap();
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
        group_effect_bounds(group)
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

fn group_effect_bounds(group: &GroupElement) -> Result<Option<Rect>, ()> {
    let Some(mut source) = elements_bounds(&group.children)? else {
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

fn elements_bounds(elements: &[PositionedElement]) -> Result<Option<Rect>, ()> {
    let mut bounds = None;
    for element in elements {
        if let Some(element) = element_bounds(element)? {
            bounds = Some(match bounds {
                Some(bounds) => union_rect(bounds, element),
                None => element,
            });
        }
    }
    Ok(bounds)
}

fn element_bounds(element: &PositionedElement) -> Result<Option<Rect>, ()> {
    match element {
        PositionedElement::Text(run) if !run.text.is_empty() => Err(()),
        PositionedElement::MultilingualText(run) if !run.logical_text.is_empty() => Err(()),
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
        PositionedElement::Group(group) => {
            Ok(group_effect_bounds(group)?
                .map(|bounds| group.transform.transform_rect_bbox(bounds)))
        }
        PositionedElement::MarkedContent { children, .. } => elements_bounds(children),
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

    use super::{image_mime, is_safe_link, preserves_isotropic_blur, render_page};

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
        }
    }

    fn text(value: &str, glyphs: usize) -> PositionedElement {
        PositionedElement::Text(GlyphRun {
            origin: Point { x: 12.0, y: 28.0 },
            font_id: FontId(7),
            font_size: 14.0,
            glyph_ids: vec![1; glyphs],
            advances: vec![8.0; glyphs],
            text: value.to_owned(),
            source: None,
            color: Color::BLACK,
            bold: false,
            italic: false,
            field_kind: None,
            field_source: None,
            note: None,
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
        dangerous.glyph_ids = vec![1; 6];
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

    #[test]
    fn font_collection_face_identity_is_diagnosed() {
        let mut collection = test_font();
        collection.data = Arc::from(b"ttcf synthetic collection".as_slice());
        collection.face_index = 2;
        let mut layout = layout(vec![text("face", 4)]);
        layout.fonts = vec![collection];

        let result = render_page(&layout, 0).unwrap();
        assert!(result.svg.contains("data:font/collection;base64,"));
        assert!(result.svg.contains("format('collection')"));
        assert_eq!(result.diagnostics.len(), 1);
        assert_eq!(result.diagnostics[0].path, "fonts[7]");
        assert!(result.diagnostics[0].message.contains("face index"));
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
        assert!(result.diagnostics[1].message.contains("face index"));
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
                text("kept", 4),
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
        assert_eq!(result.diagnostics.len(), 1);
        assert_eq!(result.diagnostics[0].path, "pages[0].elements[0]");
        assert_eq!(
            result.diagnostics[0].message,
            "group effects were omitted because singular-transform source bounds could not be proven"
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
}
