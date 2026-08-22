//! PDF document writer: assembles pages, fonts, images, metadata, and outlines.

use std::collections::{BTreeMap, BTreeSet, HashMap};

use oxml_layout::{
    Color, FillRule, FontId, GradientStop, LayoutResult, LineCap, LineJoin, Paint, Path,
    PathCommand, PathElement, PositionedElement, Stroke, Transform, walk,
};
use pdf_writer::types::{
    ActionType, AnnotationType, CidFontType, ColorSpaceOperand, FontFlags, FunctionShadingType,
    LineCapStyle, LineJoinStyle, SystemInfo,
};
use pdf_writer::{Content, Filter, Finish, Name, Pdf, Rect, Ref, Str, TextStr};

use crate::font::{self, PreparedFont};
use crate::image;

struct AlphaEntry {
    key: u32,
    alpha: f32,
    name: String,
    reference: Ref,
}

struct AlphaStates {
    entries: Vec<AlphaEntry>,
    by_key: HashMap<u32, usize>,
}

impl AlphaStates {
    fn new(
        pages: &[impl std::borrow::Borrow<oxml_layout::PageFrame>],
        alloc: &mut impl FnMut() -> Ref,
    ) -> Self {
        let mut keys = BTreeSet::new();
        for page in pages {
            collect_alpha_keys(&page.borrow().elements, &mut keys);
        }

        let entries: Vec<AlphaEntry> = keys
            .into_iter()
            .enumerate()
            .map(|(index, key)| AlphaEntry {
                key,
                alpha: f32::from_bits(key),
                name: format!("GS{index}"),
                reference: alloc(),
            })
            .collect();
        let by_key = entries
            .iter()
            .enumerate()
            .map(|(index, entry)| (entry.key, index))
            .collect();
        Self { entries, by_key }
    }

    fn entry(&self, alpha: f64) -> Option<&AlphaEntry> {
        let key = alpha_key(alpha)?;
        self.by_key.get(&key).map(|index| &self.entries[*index])
    }

    fn page_entries(&self, page: &oxml_layout::PageFrame) -> Vec<&AlphaEntry> {
        let mut keys = BTreeSet::new();
        collect_alpha_keys(&page.elements, &mut keys);
        keys.into_iter()
            .filter_map(|key| self.by_key.get(&key).map(|index| &self.entries[*index]))
            .collect()
    }
}

fn alpha_key(alpha: f64) -> Option<u32> {
    let normalized = if alpha.is_finite() {
        alpha.clamp(0.0, 1.0) as f32
    } else {
        1.0
    };
    let normalized = if normalized == 0.0 { 0.0 } else { normalized };
    (normalized < 1.0).then(|| normalized.to_bits())
}

fn collect_alpha_keys(elements: &[PositionedElement], keys: &mut BTreeSet<u32>) {
    for element in elements {
        match element {
            PositionedElement::Text(run) => insert_alpha(keys, run.color.a),
            PositionedElement::Line { color, .. } | PositionedElement::FilledRect { color, .. } => {
                insert_alpha(keys, color.a)
            }
            PositionedElement::Path(path) => {
                if let Some(Paint::Solid(color)) = &path.fill {
                    insert_alpha(keys, color.a);
                }
                if let Some(Stroke {
                    paint: Paint::Solid(color),
                    ..
                }) = &path.stroke
                {
                    insert_alpha(keys, color.a);
                }
            }
            PositionedElement::Group(group) => {
                insert_alpha(keys, group.opacity);
                collect_alpha_keys(&group.children, keys);
            }
            _ => {}
        }
    }
}

fn insert_alpha(keys: &mut BTreeSet<u32>, alpha: f64) {
    if let Some(key) = alpha_key(alpha) {
        keys.insert(key);
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum GradientTarget {
    Fill,
    Stroke,
}

enum GradientGeometry {
    Linear([f32; 4]),
    Radial([f32; 6]),
}

struct GradientEntry {
    page_idx: usize,
    name: String,
    pattern_ref: Ref,
    shading_ref: Ref,
    stitching_ref: Ref,
    interval_refs: Vec<Ref>,
    matrix: Transform,
    geometry: GradientGeometry,
    stops: Vec<GradientStop>,
    extend: [bool; 2],
}

struct GradientRegistry {
    entries: Vec<GradientEntry>,
    by_occurrence: HashMap<(usize, usize, GradientTarget), usize>,
}

impl GradientRegistry {
    fn new(
        pages: &[impl std::borrow::Borrow<oxml_layout::PageFrame>],
        alloc: &mut impl FnMut() -> Ref,
    ) -> Self {
        let mut registry = Self {
            entries: Vec::new(),
            by_occurrence: HashMap::new(),
        };

        for (page_idx, page) in pages.iter().enumerate() {
            let page = page.borrow();
            let mut leaf_idx = 0;
            let page_flip = Transform {
                a: 1.0,
                b: 0.0,
                c: 0.0,
                d: -1.0,
                e: 0.0,
                f: page.height,
            };
            walk(&page.elements, &mut |element, transform| {
                let current_idx = leaf_idx;
                leaf_idx += 1;
                let PositionedElement::Path(path) = element else {
                    return;
                };
                if let Some(fill) = &path.fill {
                    registry.collect(
                        fill,
                        page_idx,
                        current_idx,
                        GradientTarget::Fill,
                        transform.then(page_flip),
                        alloc,
                    );
                }
                if let Some(stroke) = &path.stroke {
                    registry.collect(
                        &stroke.paint,
                        page_idx,
                        current_idx,
                        GradientTarget::Stroke,
                        transform.then(page_flip),
                        alloc,
                    );
                }
            });
        }

        registry
    }

    fn collect(
        &mut self,
        paint: &Paint,
        page_idx: usize,
        leaf_idx: usize,
        target: GradientTarget,
        matrix: Transform,
        alloc: &mut impl FnMut() -> Ref,
    ) {
        let (geometry, stops, extend) = match paint {
            Paint::Linear {
                start,
                end,
                stops,
                extend,
            } => (
                GradientGeometry::Linear([
                    start.x as f32,
                    start.y as f32,
                    end.x as f32,
                    end.y as f32,
                ]),
                stops,
                *extend,
            ),
            Paint::Radial {
                center,
                radius,
                focal,
                stops,
                extend,
            } => (
                GradientGeometry::Radial([
                    focal.x as f32,
                    focal.y as f32,
                    0.0,
                    center.x as f32,
                    center.y as f32,
                    finite_f32(*radius).max(0.0),
                ]),
                stops,
                *extend,
            ),
            Paint::Solid(_) | Paint::Tile { .. } => return,
        };

        let Some(stops) = normalize_gradient_stops(stops) else {
            return;
        };
        let stops = complete_gradient_domain(stops);
        let index = self.entries.len();
        let interval_refs = (1..stops.len()).map(|_| alloc()).collect();
        self.entries.push(GradientEntry {
            page_idx,
            name: format!("P{index}"),
            pattern_ref: alloc(),
            shading_ref: alloc(),
            stitching_ref: alloc(),
            interval_refs,
            matrix,
            geometry,
            stops,
            extend: [extend.0, extend.1],
        });
        self.by_occurrence
            .insert((page_idx, leaf_idx, target), index);
    }

    fn entry(
        &self,
        page_idx: usize,
        leaf_idx: usize,
        target: GradientTarget,
    ) -> Option<&GradientEntry> {
        self.by_occurrence
            .get(&(page_idx, leaf_idx, target))
            .map(|index| &self.entries[*index])
    }

    fn page_entries(&self, page_idx: usize) -> impl Iterator<Item = &GradientEntry> {
        self.entries
            .iter()
            .filter(move |entry| entry.page_idx == page_idx)
    }
}

fn normalize_gradient_stops(stops: &[GradientStop]) -> Option<Vec<GradientStop>> {
    let mut normalized: Vec<GradientStop> = stops
        .iter()
        .map(|stop| GradientStop {
            offset: normalized_offset(stop.offset),
            color: composite_stop_alpha(stop.color),
        })
        .collect();
    normalized.sort_by(|left, right| left.offset.total_cmp(&right.offset));

    let mut deduplicated: Vec<GradientStop> = Vec::with_capacity(normalized.len());
    for stop in normalized {
        if let Some(previous) = deduplicated.last_mut()
            && previous.offset == stop.offset
        {
            *previous = stop;
        } else {
            deduplicated.push(stop);
        }
    }

    match deduplicated.as_slice() {
        [] => None,
        [stop] => Some(vec![
            GradientStop {
                offset: 0.0,
                color: stop.color,
            },
            GradientStop {
                offset: 1.0,
                color: stop.color,
            },
        ]),
        _ => Some(deduplicated),
    }
}

fn complete_gradient_domain(mut stops: Vec<GradientStop>) -> Vec<GradientStop> {
    let first = stops[0];
    if first.offset > 0.0 {
        stops.insert(
            0,
            GradientStop {
                offset: 0.0,
                color: first.color,
            },
        );
    }

    let last = stops[stops.len() - 1];
    if last.offset < 1.0 {
        stops.push(GradientStop {
            offset: 1.0,
            color: last.color,
        });
    }
    stops
}

fn normalized_offset(offset: f64) -> f64 {
    if offset.is_nan() {
        0.0
    } else {
        offset.clamp(0.0, 1.0)
    }
}

fn composite_stop_alpha(color: Color) -> Color {
    let alpha = if color.a.is_finite() {
        color.a.clamp(0.0, 1.0)
    } else {
        1.0
    };
    Color {
        r: color.r * alpha + (1.0 - alpha),
        g: color.g * alpha + (1.0 - alpha),
        b: color.b * alpha + (1.0 - alpha),
        a: 1.0,
    }
}

fn finite_f32(value: f64) -> f32 {
    if value.is_finite() { value as f32 } else { 0.0 }
}

/// Write a complete PDF document from layout results.
pub(crate) fn write_pdf(layout: &LayoutResult) -> Vec<u8> {
    let mut pdf = Pdf::new();

    // ── Reference ID allocation ──────────────────────────────────────
    let mut next_id = 1;
    let mut alloc = || {
        let r = Ref::new(next_id);
        next_id += 1;
        r
    };

    let catalog_id = alloc();
    let page_tree_id = alloc();
    let info_id = alloc();
    let outline_root_id = alloc();

    // Pre-allocate page IDs
    let page_ids: Vec<Ref> = layout.pages.iter().map(|_| alloc()).collect();
    let content_ids: Vec<Ref> = layout.pages.iter().map(|_| alloc()).collect();

    // ── Font preparation ─────────────────────────────────────────────
    let mut glyph_usage = font::collect_glyph_usage(layout);
    // Ordered by font ID. Both maps are iterated to write the font objects and
    // the page /Font dictionaries, so a hashed order wrote the same document
    // differently on every run.
    let mut prepared_fonts: BTreeMap<FontId, PreparedFont> = BTreeMap::new();
    let mut font_refs: BTreeMap<FontId, (Ref, Ref, Ref, Ref, Ref)> = BTreeMap::new(); // type0, cid, descriptor, stream, cmap

    for fd in &layout.fonts {
        if let Some(usage) = glyph_usage.get_mut(&fd.id)
            && let Some(prepared) = font::prepare_font(fd, usage)
        {
            let type0_ref = alloc();
            let cid_ref = alloc();
            let descriptor_ref = alloc();
            let stream_ref = alloc();
            let cmap_ref = alloc();
            font_refs.insert(
                fd.id,
                (type0_ref, cid_ref, descriptor_ref, stream_ref, cmap_ref),
            );
            prepared_fonts.insert(fd.id, prepared);
        }
    }

    // ── Image collection ─────────────────────────────────────────────
    // Collect all unique images across pages and allocate refs.
    struct ImageEntry {
        decoded: image::DecodedImage,
        xobject_ref: Ref,
        smask_ref: Option<Ref>,
    }

    let mut image_entries: Vec<ImageEntry> = Vec::new();
    // Map from (page_index, element_index) to image_entries index
    let mut image_map: HashMap<(usize, usize), usize> = HashMap::new();

    for (page_idx, page) in layout.pages.iter().enumerate() {
        let mut elem_idx = 0;
        walk(&page.elements, &mut |element, _| {
            let current_idx = elem_idx;
            elem_idx += 1;
            if let PositionedElement::Image {
                data, content_type, ..
            } = element
            {
                if data.is_empty() {
                    return;
                }
                if let Some(decoded) = image::decode_image(data, content_type) {
                    let xobject_ref = alloc();
                    let smask_ref = if decoded.alpha.is_some() {
                        Some(alloc())
                    } else {
                        None
                    };
                    let idx = image_entries.len();
                    image_entries.push(ImageEntry {
                        decoded,
                        xobject_ref,
                        smask_ref,
                    });
                    image_map.insert((page_idx, current_idx), idx);
                }
            }
        });
    }

    // ── Annotation refs ──────────────────────────────────────────────
    let mut annotation_refs: HashMap<(usize, usize), Ref> = HashMap::new();
    for (page_idx, page) in layout.pages.iter().enumerate() {
        let mut elem_idx = 0;
        walk(&page.elements, &mut |element, _| {
            let current_idx = elem_idx;
            elem_idx += 1;
            if let PositionedElement::LinkAnnotation { .. } = element {
                annotation_refs.insert((page_idx, current_idx), alloc());
            }
        });
    }

    let alpha_states = AlphaStates::new(&layout.pages, &mut alloc);
    let gradients = GradientRegistry::new(&layout.pages, &mut alloc);

    // ── Outline item refs ────────────────────────────────────────────
    let outline_item_ids: Vec<Ref> = layout.outlines.iter().map(|_| alloc()).collect();

    // ── Write catalog ────────────────────────────────────────────────
    {
        let mut cat = pdf.catalog(catalog_id);
        cat.pages(page_tree_id);
        if !layout.outlines.is_empty() {
            cat.outlines(outline_root_id);
        }
    }

    // ── Write page tree ──────────────────────────────────────────────
    {
        let mut pages = pdf.pages(page_tree_id);
        pages.count(layout.pages.len() as i32);
        pages.kids(page_ids.iter().copied());
    }

    // ── Write document info ──────────────────────────────────────────
    if let Some(meta) = &layout.metadata {
        let mut info = pdf.document_info(info_id);
        if let Some(title) = &meta.title {
            info.title(TextStr(title));
        }
        if let Some(author) = &meta.author {
            info.author(TextStr(author));
        }
        if let Some(subject) = &meta.subject {
            info.subject(TextStr(subject));
        }
        if let Some(keywords) = &meta.keywords {
            info.keywords(TextStr(keywords));
        }
        if let Some(creator) = &meta.creator {
            info.creator(TextStr(creator));
        }
        info.producer(TextStr("rdocx-pdf"));
    }

    // ── Write fonts ──────────────────────────────────────────────────
    for (font_id, prepared) in &prepared_fonts {
        let (type0_ref, cid_ref, descriptor_ref, stream_ref, cmap_ref) = font_refs[font_id];

        let base_name = sanitize_font_name(
            &prepared.font_data.family,
            prepared.font_data.bold,
            prepared.font_data.italic,
        );

        // Type0 font (composite font)
        let mut type0 = pdf.type0_font(type0_ref);
        type0
            .base_font(Name(base_name.as_bytes()))
            .encoding_predefined(Name(b"Identity-H"))
            .descendant_font(cid_ref)
            .to_unicode(cmap_ref);
        type0.finish();

        // CID font
        let mut cid = pdf.cid_font(cid_ref);
        cid.subtype(CidFontType::Type2);
        cid.base_font(Name(base_name.as_bytes()));
        cid.system_info(SystemInfo {
            registry: Str(b"Adobe"),
            ordering: Str(b"Identity"),
            supplement: 0,
        });
        cid.font_descriptor(descriptor_ref);
        cid.default_width(0.0);
        cid.cid_to_gid_map_predefined(Name(b"Identity"));

        // Write widths
        if !prepared.widths.is_empty() {
            let mut widths = cid.widths();
            // Group consecutive glyph IDs for the W array
            let mut i = 0;
            while i < prepared.widths.len() {
                let start_gid = prepared.widths[i].0;
                let mut consecutive_widths = vec![prepared.widths[i].1 as f32];
                let mut j = i + 1;
                while j < prepared.widths.len()
                    && prepared.widths[j].0 == start_gid + (j - i) as u16
                {
                    consecutive_widths.push(prepared.widths[j].1 as f32);
                    j += 1;
                }
                widths.consecutive(start_gid, consecutive_widths.iter().copied());
                i = j;
            }
            widths.finish();
        }
        cid.finish();

        // Font descriptor
        if let Some(metrics) = font::get_font_metrics(&prepared.font_data) {
            let mut desc = pdf.font_descriptor(descriptor_ref);
            desc.name(Name(base_name.as_bytes()));

            let mut flags = FontFlags::empty();
            if prepared.font_data.italic {
                flags |= FontFlags::ITALIC;
            }
            // Most text fonts are not fixed pitch and not symbolic
            flags |= FontFlags::NON_SYMBOLIC;
            desc.flags(flags);

            desc.bbox(Rect::new(
                metrics.bbox[0] as f32,
                metrics.bbox[1] as f32,
                metrics.bbox[2] as f32,
                metrics.bbox[3] as f32,
            ));
            desc.italic_angle(metrics.italic_angle as f32);
            desc.ascent(metrics.ascent as f32);
            desc.descent(metrics.descent as f32);
            desc.cap_height(metrics.cap_height as f32);
            desc.stem_v(metrics.stem_v as f32);
            desc.font_file2(stream_ref);
            desc.finish();
        }

        // Font stream (subset TrueType data)
        let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&prepared.subset_bytes, 6);
        let mut stream = pdf.stream(stream_ref, &compressed);
        stream.filter(Filter::FlateDecode);
        stream.pair(Name(b"Length1"), prepared.subset_bytes.len() as i32);
        stream.finish();

        // ToUnicode CMap stream
        let compressed_cmap = miniz_oxide::deflate::compress_to_vec_zlib(&prepared.cmap_bytes, 6);
        let mut cmap_stream = pdf.stream(cmap_ref, &compressed_cmap);
        cmap_stream.filter(Filter::FlateDecode);
        cmap_stream.finish();
    }

    // ── Write image XObjects ─────────────────────────────────────────
    for entry in &image_entries {
        let dec = &entry.decoded;

        if dec.is_jpeg {
            // JPEG pass-through
            let mut img = pdf.image_xobject(entry.xobject_ref, &dec.data);
            img.filter(Filter::DctDecode);
            img.width(dec.width as i32);
            img.height(dec.height as i32);
            set_color_space(&mut img, dec.color_space);
            img.bits_per_component(8);
            img.finish();
        } else {
            // Raw pixel data, compress with Deflate
            let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&dec.data, 6);
            let mut img = pdf.image_xobject(entry.xobject_ref, &compressed);
            img.filter(Filter::FlateDecode);
            img.width(dec.width as i32);
            img.height(dec.height as i32);
            set_color_space(&mut img, dec.color_space);
            img.bits_per_component(8);
            if let Some(smask_ref) = entry.smask_ref {
                img.s_mask(smask_ref);
            }
            img.finish();

            // Write soft mask (alpha channel) if present
            if let (Some(alpha), Some(smask_ref)) = (&dec.alpha, entry.smask_ref) {
                let compressed_alpha = miniz_oxide::deflate::compress_to_vec_zlib(alpha, 6);
                let mut mask = pdf.image_xobject(smask_ref, &compressed_alpha);
                mask.filter(Filter::FlateDecode);
                mask.width(dec.width as i32);
                mask.height(dec.height as i32);
                mask.color_space().device_gray();
                mask.bits_per_component(8);
                mask.finish();
            }
        }
    }

    for entry in &alpha_states.entries {
        let mut state = pdf.ext_graphics(entry.reference);
        state.stroking_alpha(entry.alpha);
        state.non_stroking_alpha(entry.alpha);
        state.finish();
    }

    write_gradient_objects(&mut pdf, &gradients);

    // ── Write pages ──────────────────────────────────────────────────
    for (page_idx, page) in layout.pages.iter().enumerate() {
        let page_id = page_ids[page_idx];
        let content_id = content_ids[page_idx];

        // Build content stream
        let content_bytes = build_page_content(
            page_idx,
            page,
            &prepared_fonts,
            &font_refs,
            &image_map,
            &alpha_states,
            &gradients,
        );

        // Compress and write content stream
        let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&content_bytes, 6);
        let mut stream = pdf.stream(content_id, &compressed);
        stream.filter(Filter::FlateDecode);
        stream.finish();

        // Write page dictionary
        let mut page_dict = pdf.page(page_id);
        page_dict.parent(page_tree_id);
        page_dict.media_box(Rect::new(0.0, 0.0, page.width as f32, page.height as f32));
        page_dict.contents(content_id);

        // Resources: fonts + images
        let mut resources = page_dict.resources();

        // Font resources
        if !prepared_fonts.is_empty() {
            let mut font_dict = resources.fonts();
            for (font_id, (type0_ref, _, _, _, _)) in &font_refs {
                let font_name = format!("F{}", font_id.0);
                font_dict.pair(Name(font_name.as_bytes()), *type0_ref);
            }
            font_dict.finish();
        }

        // Image XObject resources
        // Sorted by element index. `image_map` is hashed, and its iteration
        // order reached the /XObject dictionary, so the same page named its
        // images in a different order on every run.
        let mut page_images: Vec<(usize, usize)> = image_map
            .iter()
            .filter(|((pi, _), _)| *pi == page_idx)
            .map(|((_, ei), img_idx)| (*ei, *img_idx))
            .collect();
        page_images.sort_unstable();
        let page_image_names: Vec<(String, Ref)> = page_images
            .into_iter()
            .map(|(ei, img_idx)| {
                (
                    format!("Im{}_{}", page_idx, ei),
                    image_entries[img_idx].xobject_ref,
                )
            })
            .collect();

        if !page_image_names.is_empty() {
            let mut xobjects = resources.x_objects();
            for (name, xobj_ref) in &page_image_names {
                xobjects.pair(Name(name.as_bytes()), *xobj_ref);
            }
            xobjects.finish();
        }

        let page_alpha_states = alpha_states.page_entries(page);
        if !page_alpha_states.is_empty() {
            let mut states = resources.ext_g_states();
            for entry in page_alpha_states {
                states.pair(Name(entry.name.as_bytes()), entry.reference);
            }
            states.finish();
        }

        let page_gradients: Vec<_> = gradients.page_entries(page_idx).collect();
        if !page_gradients.is_empty() {
            let mut patterns = resources.patterns();
            for entry in page_gradients {
                patterns.pair(Name(entry.name.as_bytes()), entry.pattern_ref);
            }
            patterns.finish();
        }

        resources.finish();

        // Annotations (hyperlinks)
        let mut page_annotations = Vec::new();
        let mut elem_idx = 0;
        walk(&page.elements, &mut |element, _| {
            let current_idx = elem_idx;
            elem_idx += 1;
            if matches!(element, PositionedElement::LinkAnnotation { .. })
                && let Some(reference) = annotation_refs.get(&(page_idx, current_idx))
            {
                page_annotations.push(*reference);
            }
        });

        if !page_annotations.is_empty() {
            page_dict.annotations(page_annotations.iter().copied());
        }

        page_dict.finish();
    }

    // ── Write link annotations ───────────────────────────────────────
    for (page_idx, page) in layout.pages.iter().enumerate() {
        let mut elem_idx = 0;
        walk(&page.elements, &mut |element, transform| {
            let current_idx = elem_idx;
            elem_idx += 1;
            if let PositionedElement::LinkAnnotation { rect, url } = element
                && let Some(&annot_ref) = annotation_refs.get(&(page_idx, current_idx))
            {
                let rect = transform.transform_rect_bbox(*rect);
                let page_height = page.height;
                let mut annot = pdf.annotation(annot_ref);
                annot.subtype(AnnotationType::Link);
                annot.rect(Rect::new(
                    rect.x as f32,
                    (page_height - rect.y - rect.height) as f32,
                    (rect.x + rect.width) as f32,
                    (page_height - rect.y) as f32,
                ));
                annot.border(0.0, 0.0, 0.0, None);
                annot
                    .action()
                    .action_type(ActionType::Uri)
                    .uri(Str(url.as_bytes()));
                annot.finish();
            }
        });
    }

    // ── Write outlines/bookmarks ─────────────────────────────────────
    if !layout.outlines.is_empty() {
        write_outlines(
            &mut pdf,
            outline_root_id,
            &layout.outlines,
            &outline_item_ids,
            &page_ids,
            layout,
        );
    }

    pdf.finish()
}

fn write_gradient_objects(pdf: &mut Pdf, gradients: &GradientRegistry) {
    for entry in &gradients.entries {
        let mut pattern = pdf.shading_pattern(entry.pattern_ref);
        pattern.shading_ref(entry.shading_ref);
        pattern.matrix([
            entry.matrix.a as f32,
            entry.matrix.b as f32,
            entry.matrix.c as f32,
            entry.matrix.d as f32,
            entry.matrix.e as f32,
            entry.matrix.f as f32,
        ]);
        pattern.finish();

        let mut shading = pdf.function_shading(entry.shading_ref);
        match entry.geometry {
            GradientGeometry::Linear(coords) => {
                shading.shading_type(FunctionShadingType::Axial);
                shading.coords(coords);
            }
            GradientGeometry::Radial(coords) => {
                shading.shading_type(FunctionShadingType::Radial);
                shading.coords(coords);
            }
        }
        shading.color_space().device_rgb();
        shading
            .insert(Name(b"Domain"))
            .array()
            .items([0.0_f32, 1.0_f32]);
        shading.extend(entry.extend);
        shading.function(entry.stitching_ref);
        shading.finish();

        let mut stitching = pdf.stitching_function(entry.stitching_ref);
        stitching.domain([0.0, 1.0]);
        stitching.functions(entry.interval_refs.iter().copied());
        stitching.bounds(
            entry.stops[1..entry.stops.len() - 1]
                .iter()
                .map(|stop| stop.offset as f32),
        );
        stitching.encode(entry.interval_refs.iter().flat_map(|_| [0.0_f32, 1.0_f32]));
        stitching.finish();

        for ((start, end), reference) in entry
            .stops
            .windows(2)
            .map(|pair| (&pair[0], &pair[1]))
            .zip(entry.interval_refs.iter())
        {
            let mut function = pdf.exponential_function(*reference);
            function.domain([0.0, 1.0]);
            function.c0([
                start.color.r as f32,
                start.color.g as f32,
                start.color.b as f32,
            ]);
            function.c1([end.color.r as f32, end.color.g as f32, end.color.b as f32]);
            function.n(1.0);
            function.finish();
        }
    }
}

/// Build the content stream for a single page.
fn build_page_content(
    page_idx: usize,
    page: &oxml_layout::PageFrame,
    prepared_fonts: &BTreeMap<FontId, PreparedFont>,
    font_refs: &BTreeMap<FontId, (Ref, Ref, Ref, Ref, Ref)>,
    image_map: &HashMap<(usize, usize), usize>,
    alpha_states: &AlphaStates,
    gradients: &GradientRegistry,
) -> Vec<u8> {
    let mut content = Content::new();
    let page_height = page.height as f32;
    content.save_state();
    content.transform([1.0, 0.0, 0.0, -1.0, 0.0, page_height]);

    let mut state = EmitState {
        page_idx,
        prepared_fonts,
        font_refs,
        image_map,
        alpha_states,
        gradients,
        leaf_index: 0,
    };
    emit_elements(&mut content, &page.elements, &mut state);

    content.restore_state();
    content.finish().to_vec()
}

struct EmitState<'a> {
    page_idx: usize,
    prepared_fonts: &'a BTreeMap<FontId, PreparedFont>,
    font_refs: &'a BTreeMap<FontId, (Ref, Ref, Ref, Ref, Ref)>,
    image_map: &'a HashMap<(usize, usize), usize>,
    alpha_states: &'a AlphaStates,
    gradients: &'a GradientRegistry,
    leaf_index: usize,
}

fn emit_elements(content: &mut Content, elements: &[PositionedElement], state: &mut EmitState<'_>) {
    for element in elements {
        let elem_idx = state.leaf_index;
        if !matches!(element, PositionedElement::Group(_)) {
            state.leaf_index += 1;
        }
        match element {
            PositionedElement::Text(run) => {
                if let Some(prepared) = state.prepared_fonts.get(&run.font_id)
                    && state.font_refs.contains_key(&run.font_id)
                {
                    let font_name = format!("F{}", run.font_id.0);

                    content.save_state();

                    // Set text color
                    content.set_fill_rgb(
                        run.color.r as f32,
                        run.color.g as f32,
                        run.color.b as f32,
                    );
                    apply_alpha(content, run.color.a, state.alpha_states);

                    content.begin_text();
                    content.set_font(Name(font_name.as_bytes()), run.font_size as f32);

                    // Cancel the page flip so glyphs remain upright.
                    content.set_text_matrix([
                        1.0,
                        0.0,
                        0.0,
                        -1.0,
                        run.origin.x as f32,
                        run.origin.y as f32,
                    ]);

                    // Remap glyph IDs and emit with TJ operator
                    emit_glyphs(
                        content,
                        &run.glyph_ids,
                        &run.advances,
                        run.font_size,
                        &prepared.remapper,
                        &prepared.widths,
                    );

                    content.end_text();
                    content.restore_state();
                }
            }
            PositionedElement::Line {
                start,
                end,
                width,
                color,
                dash_pattern,
            } => {
                content.save_state();
                content.set_stroke_rgb(color.r as f32, color.g as f32, color.b as f32);
                apply_alpha(content, color.a, state.alpha_states);
                content.set_line_width(*width as f32);
                if let Some((on, off)) = dash_pattern {
                    content.set_dash_pattern([*on as f32, *off as f32], 0.0);
                }
                content.move_to(start.x as f32, start.y as f32);
                content.line_to(end.x as f32, end.y as f32);
                content.stroke();
                content.restore_state();
            }
            PositionedElement::FilledRect { rect, color } => {
                content.save_state();
                content.set_fill_rgb(color.r as f32, color.g as f32, color.b as f32);
                apply_alpha(content, color.a, state.alpha_states);
                content.rect(
                    rect.x as f32,
                    rect.y as f32,
                    rect.width as f32,
                    rect.height as f32,
                );
                content.fill_nonzero();
                content.restore_state();
            }
            PositionedElement::Image { rect, data, .. } => {
                if !data.is_empty() && state.image_map.contains_key(&(state.page_idx, elem_idx)) {
                    let img_name = format!("Im{}_{}", state.page_idx, elem_idx);

                    content.save_state();
                    // Cancel the page flip while preserving the image rectangle.
                    content.transform([
                        rect.width as f32,
                        0.0,
                        0.0,
                        -rect.height as f32,
                        rect.x as f32,
                        (rect.y + rect.height) as f32,
                    ]);
                    content.x_object(Name(img_name.as_bytes()));
                    content.restore_state();
                }
            }
            PositionedElement::LinkAnnotation { .. } => {
                // Link annotations are written separately, not in the content stream
            }
            PositionedElement::Path(path) => {
                emit_path(
                    content,
                    path,
                    state.page_idx,
                    elem_idx,
                    state.alpha_states,
                    state.gradients,
                );
            }
            PositionedElement::Group(group) => {
                content.save_state();
                content.transform([
                    group.transform.a as f32,
                    group.transform.b as f32,
                    group.transform.c as f32,
                    group.transform.d as f32,
                    group.transform.e as f32,
                    group.transform.f as f32,
                ]);
                if let Some(clip) = &group.clip {
                    emit_path_geometry(content, clip);
                    match clip.fill_rule {
                        FillRule::NonZero => content.clip_nonzero(),
                        FillRule::EvenOdd => content.clip_even_odd(),
                    };
                    content.end_path();
                }
                apply_alpha(content, group.opacity, state.alpha_states);
                emit_elements(content, &group.children, state);
                content.restore_state();
            }
            _ => {
                // Later oxml-layout elements remain unsupported until assigned.
            }
        }
    }
}

fn apply_alpha(content: &mut Content, alpha: f64, states: &AlphaStates) {
    if let Some(entry) = states.entry(alpha) {
        content.set_parameters(Name(entry.name.as_bytes()));
    }
}

enum ResolvedPaint<'a> {
    Solid(Color),
    Gradient(&'a GradientEntry),
}

fn emit_path(
    content: &mut Content,
    element: &PathElement,
    page_idx: usize,
    leaf_idx: usize,
    alpha_states: &AlphaStates,
    gradients: &GradientRegistry,
) {
    let fill = element.fill.as_ref().and_then(|paint| {
        resolve_paint(paint, page_idx, leaf_idx, GradientTarget::Fill, gradients)
    });
    let stroke = element.stroke.as_ref().and_then(|stroke| {
        resolve_paint(
            &stroke.paint,
            page_idx,
            leaf_idx,
            GradientTarget::Stroke,
            gradients,
        )
        .map(|paint| (stroke, paint))
    });

    if fill.is_none() && stroke.is_none() {
        return;
    }

    content.save_state();

    match (fill, stroke) {
        (Some(ResolvedPaint::Solid(fill)), Some((stroke, ResolvedPaint::Solid(stroke_color))))
            if alpha_key(fill.a) == alpha_key(stroke_color.a) =>
        {
            set_fill_color(content, fill, alpha_states);
            emit_stroke_style(content, stroke, stroke_color);
            emit_path_geometry(content, &element.path);
            match element.path.fill_rule {
                FillRule::NonZero => content.fill_nonzero_and_stroke(),
                FillRule::EvenOdd => content.fill_even_odd_and_stroke(),
            };
        }
        (Some(fill), Some((stroke, stroke_paint))) => {
            content.save_state();
            set_fill_paint(content, fill, alpha_states);
            emit_path_geometry(content, &element.path);
            match element.path.fill_rule {
                FillRule::NonZero => content.fill_nonzero(),
                FillRule::EvenOdd => content.fill_even_odd(),
            };
            content.restore_state();
            content.save_state();
            emit_stroke_paint(content, stroke, stroke_paint, alpha_states);
            emit_path_geometry(content, &element.path);
            content.stroke();
            content.restore_state();
        }
        (Some(fill), None) => {
            set_fill_paint(content, fill, alpha_states);
            emit_path_geometry(content, &element.path);
            match element.path.fill_rule {
                FillRule::NonZero => content.fill_nonzero(),
                FillRule::EvenOdd => content.fill_even_odd(),
            };
        }
        (None, Some((stroke, stroke_paint))) => {
            emit_stroke_paint(content, stroke, stroke_paint, alpha_states);
            emit_path_geometry(content, &element.path);
            content.stroke();
        }
        (None, None) => {}
    }

    content.restore_state();
}

fn resolve_paint<'a>(
    paint: &Paint,
    page_idx: usize,
    leaf_idx: usize,
    target: GradientTarget,
    gradients: &'a GradientRegistry,
) -> Option<ResolvedPaint<'a>> {
    match paint {
        Paint::Solid(color) => Some(ResolvedPaint::Solid(*color)),
        Paint::Linear { .. } | Paint::Radial { .. } => gradients
            .entry(page_idx, leaf_idx, target)
            .map(ResolvedPaint::Gradient),
        Paint::Tile { .. } => None,
    }
}

fn set_fill_paint(content: &mut Content, paint: ResolvedPaint<'_>, alpha_states: &AlphaStates) {
    match paint {
        ResolvedPaint::Solid(color) => set_fill_color(content, color, alpha_states),
        ResolvedPaint::Gradient(entry) => {
            content.set_fill_color_space(ColorSpaceOperand::Pattern);
            content.set_fill_pattern([], Name(entry.name.as_bytes()));
        }
    }
}

fn emit_stroke_paint(
    content: &mut Content,
    stroke: &Stroke,
    paint: ResolvedPaint<'_>,
    alpha_states: &AlphaStates,
) {
    match paint {
        ResolvedPaint::Solid(color) => {
            emit_stroke_style(content, stroke, color);
            apply_alpha(content, color.a, alpha_states);
        }
        ResolvedPaint::Gradient(entry) => {
            emit_stroke_geometry_style(content, stroke);
            content.set_stroke_color_space(ColorSpaceOperand::Pattern);
            content.set_stroke_pattern([], Name(entry.name.as_bytes()));
        }
    }
}

fn set_fill_color(content: &mut Content, color: oxml_layout::Color, alpha_states: &AlphaStates) {
    content.set_fill_rgb(color.r as f32, color.g as f32, color.b as f32);
    apply_alpha(content, color.a, alpha_states);
}

fn emit_stroke_style(content: &mut Content, stroke: &Stroke, color: oxml_layout::Color) {
    emit_stroke_geometry_style(content, stroke);
    content.set_stroke_rgb(color.r as f32, color.g as f32, color.b as f32);
}

fn emit_stroke_geometry_style(content: &mut Content, stroke: &Stroke) {
    let width = if stroke.width.is_finite() {
        stroke.width.max(0.0) as f32
    } else {
        0.0
    };

    content.set_line_width(width);
    content.set_line_cap(match stroke.cap {
        LineCap::Butt => LineCapStyle::ButtCap,
        LineCap::Round => LineCapStyle::RoundCap,
        LineCap::Square => LineCapStyle::ProjectingSquareCap,
    });
    content.set_line_join(match stroke.join {
        LineJoin::Miter => LineJoinStyle::MiterJoin,
        LineJoin::Round => LineJoinStyle::RoundJoin,
        LineJoin::Bevel => LineJoinStyle::BevelJoin,
    });
    content.set_miter_limit(10.0);
    if let Some(dash) = &stroke.dash {
        content.set_dash_pattern(dash.iter().map(|value| *value as f32), 0.0);
    }
}

fn emit_path_geometry(content: &mut Content, path: &Path) {
    for command in &path.commands {
        match command {
            PathCommand::MoveTo(point) => {
                content.move_to(point.x as f32, point.y as f32);
            }
            PathCommand::LineTo(point) => {
                content.line_to(point.x as f32, point.y as f32);
            }
            PathCommand::CurveTo { c1, c2, to } => {
                content.cubic_to(
                    c1.x as f32,
                    c1.y as f32,
                    c2.x as f32,
                    c2.y as f32,
                    to.x as f32,
                    to.y as f32,
                );
            }
            PathCommand::Close => {
                content.close_path();
            }
        }
    }
}

/// Emit glyph IDs using the TJ operator with per-glyph positioning.
///
/// The TJ operator alternates between glyph strings and numeric adjustments.
/// After showing a glyph, PDF auto-advances by the glyph's declared width
/// (from the CID font's W table, in thousandths of text space).
/// The adjustment value is *subtracted* from the current position.
///
/// To achieve the shaped advance: adjustment = declared_width - (advance_pt / font_size * 1000)
fn emit_glyphs(
    content: &mut Content,
    glyph_ids: &[u16],
    advances: &[f64],
    font_size: f64,
    remapper: &subsetter::GlyphRemapper,
    widths: &[(u16, f64)],
) {
    if glyph_ids.is_empty() {
        return;
    }

    // Build a lookup from new_gid → declared width (in 1000ths of em)
    let width_map: HashMap<u16, f64> = widths.iter().copied().collect();

    let mut positioned = content.show_positioned();
    let mut items = positioned.items();

    for (i, &gid) in glyph_ids.iter().enumerate() {
        let new_gid = remapper.get(gid).unwrap_or(0);

        // Encode glyph ID as 2-byte big-endian
        let bytes = new_gid.to_be_bytes();
        items.show(Str(&bytes));

        // Calculate adjustment for next glyph
        if i + 1 < glyph_ids.len() {
            let advance_pt = advances[i];
            // declared_width is the width PDF will auto-advance by (in 1000ths of em)
            // default_width is 0.0, so glyphs not in W table get no auto-advance
            let declared_width = width_map.get(&new_gid).copied().unwrap_or(0.0);
            // We want net advance = advance_pt
            // net = (declared_width - adjustment) / 1000 * font_size = advance_pt
            // => adjustment = declared_width - advance_pt / font_size * 1000
            let adjustment = (declared_width - advance_pt / font_size * 1000.0) as f32;
            items.adjust(adjustment);
        }
    }

    items.finish();
    positioned.finish();
}

/// Write the outline tree (bookmarks) with hierarchical structure.
///
/// H2 entries become children of the preceding H1, H3 of the preceding H2, etc.
fn write_outlines(
    pdf: &mut Pdf,
    root_id: Ref,
    outlines: &[oxml_layout::OutlineEntry],
    item_ids: &[Ref],
    page_ids: &[Ref],
    layout: &LayoutResult,
) {
    if outlines.is_empty() {
        return;
    }

    // Build parent/children/prev/next relationships based on heading levels.
    // parent_idx[i] = index of this item's parent, or None if top-level
    // children[i] = indices of children of item i
    let n = outlines.len();
    let mut parent_idx: Vec<Option<usize>> = vec![None; n];
    let mut children: Vec<Vec<usize>> = vec![Vec::new(); n];

    // Stack tracks (index, level) of open ancestors
    let mut stack: Vec<(usize, u32)> = Vec::new();

    for (i, entry) in outlines.iter().enumerate() {
        // Pop entries that are at the same or deeper level
        while let Some(&(_, lvl)) = stack.last() {
            if lvl >= entry.level {
                stack.pop();
            } else {
                break;
            }
        }

        if let Some(&(parent, _)) = stack.last() {
            parent_idx[i] = Some(parent);
            children[parent].push(i);
        }
        // else: top-level child of root

        stack.push((i, entry.level));
    }

    // Identify top-level items (children of root)
    let top_level: Vec<usize> = (0..n).filter(|i| parent_idx[*i].is_none()).collect();

    // Write outline root
    if let (Some(&first), Some(&last)) = (top_level.first(), top_level.last()) {
        let mut root = pdf.outline(root_id);
        root.first(item_ids[first]);
        root.last(item_ids[last]);
        root.count(n as i32); // total count including descendants
        root.finish();
    }

    // Write each outline item
    for (i, entry) in outlines.iter().enumerate() {
        let mut item = pdf.outline_item(item_ids[i]);
        item.title(TextStr(&entry.title));

        // Parent
        let actual_parent = match parent_idx[i] {
            Some(pi) => item_ids[pi],
            None => root_id,
        };
        item.parent(actual_parent);

        // Siblings (prev/next among items sharing the same parent)
        let siblings: &[usize] = if let Some(pi) = parent_idx[i] {
            &children[pi]
        } else {
            &top_level
        };
        if let Some(pos) = siblings.iter().position(|&x| x == i) {
            if pos > 0 {
                item.prev(item_ids[siblings[pos - 1]]);
            }
            if pos + 1 < siblings.len() {
                item.next(item_ids[siblings[pos + 1]]);
            }
        }

        // First/last child
        if !children[i].is_empty() {
            item.first(item_ids[children[i][0]]);
            item.last(item_ids[*children[i].last().unwrap()]);
            item.count(children[i].len() as i32);
        }

        // Page destination
        let page_idx = entry.page_index.min(page_ids.len().saturating_sub(1));
        if page_idx < page_ids.len() {
            let page_height = layout
                .pages
                .get(page_idx)
                .map(|p| p.height)
                .unwrap_or(792.0);
            let pdf_y = page_height - entry.y_position;
            item.dest()
                .page(page_ids[page_idx])
                .xyz(0.0, pdf_y as f32, None);
        }

        item.finish();
    }
}

/// Set color space on an image XObject.
fn set_color_space(img: &mut pdf_writer::writers::ImageXObject<'_>, cs: &str) {
    match cs {
        "DeviceGray" => {
            img.color_space().device_gray();
        }
        "DeviceCMYK" => {
            img.color_space().device_cmyk();
        }
        _ => {
            img.color_space().device_rgb();
        }
    }
}

/// Create a sanitized PostScript font name.
fn sanitize_font_name(family: &str, bold: bool, italic: bool) -> String {
    let mut name = family
        .chars()
        .filter(|c| c.is_ascii_alphanumeric() || *c == '-')
        .collect::<String>();

    if name.is_empty() {
        name = "Font".to_string();
    }

    if bold && italic {
        name.push_str("-BoldItalic");
    } else if bold {
        name.push_str("-Bold");
    } else if italic {
        name.push_str("-Italic");
    }

    name
}

#[cfg(test)]
mod tests {
    use super::*;
    use oxml_layout::{
        Color, Effect, FillRule, FontData, GlyphRun, GradientStop, GroupElement, LineCap, LineJoin,
        MediaId, PageFrame, Paint, Path, PathCommand, PathElement, Point, Rect, Stroke, Transform,
    };

    fn page_with(elements: Vec<PositionedElement>) -> PageFrame {
        PageFrame::new(1, 612.0, 792.0, elements)
    }

    fn content_for(elements: Vec<PositionedElement>) -> String {
        let page = page_with(elements);
        let mut next_ref = 100;
        let alpha_states = AlphaStates::new(std::slice::from_ref(&page), &mut || {
            let reference = Ref::new(next_ref);
            next_ref += 1;
            reference
        });
        let gradients = GradientRegistry::new(std::slice::from_ref(&page), &mut || {
            let reference = Ref::new(next_ref);
            next_ref += 1;
            reference
        });
        String::from_utf8(build_page_content(
            0,
            &page,
            &BTreeMap::new(),
            &BTreeMap::new(),
            &HashMap::new(),
            &alpha_states,
            &gradients,
        ))
        .expect("PDF content operators are ASCII")
    }

    fn path_element(
        fill_rule: FillRule,
        fill: Option<Paint>,
        stroke: Option<Stroke>,
    ) -> PositionedElement {
        PositionedElement::Path(PathElement {
            path: Path {
                commands: vec![
                    PathCommand::MoveTo(Point { x: 10.0, y: 20.0 }),
                    PathCommand::LineTo(Point { x: 30.0, y: 40.0 }),
                    PathCommand::CurveTo {
                        c1: Point { x: 50.0, y: 60.0 },
                        c2: Point { x: 70.0, y: 80.0 },
                        to: Point { x: 90.0, y: 100.0 },
                    },
                    PathCommand::Close,
                ],
                fill_rule,
            },
            fill,
            stroke,
        })
    }

    fn solid_stroke() -> Stroke {
        Stroke {
            paint: Paint::Solid(Color::BLACK),
            width: 2.0,
            cap: LineCap::Round,
            join: LineJoin::Bevel,
            dash: Some(vec![3.0, 4.0]),
        }
    }

    fn has_operator(content: &str, operator: &str) -> bool {
        content.lines().any(|line| line == operator)
    }

    #[test]
    fn path_fill_only_emits_f() {
        let content = content_for(vec![path_element(
            FillRule::NonZero,
            Some(Paint::Solid(Color::BLACK)),
            None,
        )]);

        assert!(has_operator(&content, "f"), "{content}");
    }

    #[test]
    fn path_stroke_only_emits_s() {
        let content = content_for(vec![path_element(
            FillRule::NonZero,
            None,
            Some(solid_stroke()),
        )]);

        assert!(has_operator(&content, "S"), "{content}");
    }

    #[test]
    fn path_fill_and_stroke_emit_b() {
        let content = content_for(vec![path_element(
            FillRule::NonZero,
            Some(Paint::Solid(Color::BLACK)),
            Some(solid_stroke()),
        )]);

        assert!(has_operator(&content, "B"), "{content}");
        assert_eq!(content.lines().filter(|line| *line == "q").count(), 2);
        assert_eq!(content.lines().filter(|line| *line == "Q").count(), 2);
    }

    #[test]
    fn even_odd_paths_use_starred_operators() {
        let fill = content_for(vec![path_element(
            FillRule::EvenOdd,
            Some(Paint::Solid(Color::BLACK)),
            None,
        )]);
        let both = content_for(vec![path_element(
            FillRule::EvenOdd,
            Some(Paint::Solid(Color::BLACK)),
            Some(solid_stroke()),
        )]);

        assert!(has_operator(&fill, "f*"), "{fill}");
        assert!(has_operator(&both, "B*"), "{both}");
    }

    #[test]
    fn path_geometry_preserves_command_order() {
        let content = content_for(vec![path_element(
            FillRule::NonZero,
            Some(Paint::Solid(Color::BLACK)),
            None,
        )]);

        assert!(
            content.contains("10 20 m\n30 40 l\n50 60 70 80 90 100 c\nh\nf"),
            "{content}"
        );
    }

    #[test]
    fn stroke_state_maps_cap_join_miter_and_dash() {
        let content = content_for(vec![path_element(
            FillRule::NonZero,
            None,
            Some(solid_stroke()),
        )]);

        assert!(content.contains("2 w"), "{content}");
        assert!(content.contains("1 J"), "{content}");
        assert!(content.contains("2 j"), "{content}");
        assert!(content.contains("10 M"), "{content}");
        assert!(content.contains("[3 4] 0 d"), "{content}");
    }

    fn gradient_stops() -> Vec<GradientStop> {
        vec![
            GradientStop {
                offset: 0.0,
                color: Color {
                    r: 1.0,
                    g: 0.0,
                    b: 0.0,
                    a: 1.0,
                },
            },
            GradientStop {
                offset: 1.0,
                color: Color {
                    r: 0.0,
                    g: 0.0,
                    b: 1.0,
                    a: 1.0,
                },
            },
        ]
    }

    fn linear_gradient() -> Paint {
        Paint::Linear {
            start: Point { x: 10.0, y: 20.0 },
            end: Point { x: 90.0, y: 20.0 },
            stops: gradient_stops(),
            extend: (true, false),
        }
    }

    #[test]
    fn linear_gradient_writes_pattern_shading_and_stitching_resources() {
        let pdf = pdf_for(vec![path_element(
            FillRule::NonZero,
            Some(linear_gradient()),
            None,
        )]);

        assert!(pdf.contains("/PatternType 2"), "{pdf}");
        assert!(pdf.contains("/ShadingType 2"), "{pdf}");
        assert!(pdf.contains("/FunctionType 3"), "{pdf}");
        assert!(pdf.contains("/FunctionType 2"), "{pdf}");
    }

    #[test]
    fn radial_gradient_writes_type_three_shading() {
        let pdf = pdf_for(vec![path_element(
            FillRule::NonZero,
            Some(Paint::Radial {
                center: Point { x: 50.0, y: 50.0 },
                radius: 40.0,
                focal: Point { x: 40.0, y: 45.0 },
                stops: gradient_stops(),
                extend: (false, true),
            }),
            None,
        )]);

        assert!(pdf.contains("/ShadingType 3"), "{pdf}");
        assert!(pdf.contains("/Coords [40 45 0 50 50 40]"), "{pdf}");
        assert!(pdf.contains("/Extend [false true]"), "{pdf}");
    }

    #[test]
    fn gradient_stops_are_sorted_deduplicated_and_clamped() {
        let stops = vec![
            GradientStop {
                offset: 2.0,
                color: Color::WHITE,
            },
            GradientStop {
                offset: -1.0,
                color: Color::BLACK,
            },
            GradientStop {
                offset: 0.5,
                color: Color::BLACK,
            },
            GradientStop {
                offset: 0.5,
                color: Color::WHITE,
            },
        ];
        let normalized = normalize_gradient_stops(&stops).expect("non-empty gradient");
        assert_eq!(
            normalized
                .iter()
                .map(|stop| stop.offset)
                .collect::<Vec<_>>(),
            vec![0.0, 0.5, 1.0]
        );
        assert_eq!(normalized[1].color, Color::WHITE);
        let pdf = pdf_for(vec![path_element(
            FillRule::NonZero,
            Some(Paint::Linear {
                start: Point { x: 0.0, y: 0.0 },
                end: Point { x: 100.0, y: 0.0 },
                stops,
                extend: (false, false),
            }),
            None,
        )]);

        assert!(pdf.contains("/Domain [0 1]"), "{pdf}");
        assert!(pdf.contains("/Bounds [0.5]"), "{pdf}");
        assert!(pdf.contains("/Encode [0 1 0 1]"), "{pdf}");
    }

    #[test]
    fn partial_range_gradient_holds_endpoint_colours_over_full_axis() {
        let red = Color {
            r: 1.0,
            g: 0.0,
            b: 0.0,
            a: 1.0,
        };
        let blue = Color {
            r: 0.0,
            g: 0.0,
            b: 1.0,
            a: 1.0,
        };
        let pdf = pdf_for(vec![path_element(
            FillRule::NonZero,
            Some(Paint::Linear {
                start: Point { x: 0.0, y: 0.0 },
                end: Point { x: 100.0, y: 0.0 },
                stops: vec![
                    GradientStop {
                        offset: 0.25,
                        color: red,
                    },
                    GradientStop {
                        offset: 0.75,
                        color: blue,
                    },
                ],
                extend: (false, false),
            }),
            None,
        )]);

        assert!(pdf.contains("/Domain [0 1]"), "{pdf}");
        assert!(pdf.contains("/Bounds [0.25 0.75]"), "{pdf}");
        assert!(pdf.contains("/Encode [0 1 0 1 0 1]"), "{pdf}");
        assert_eq!(pdf.matches("/FunctionType 2").count(), 3, "{pdf}");
        assert!(pdf.contains("/C0 [1 0 0]\n  /C1 [1 0 0]"), "{pdf}");
        assert!(pdf.contains("/C0 [0 0 1]\n  /C1 [0 0 1]"), "{pdf}");
    }

    #[test]
    fn gradient_fill_and_solid_stroke_both_render() {
        let content = content_for(vec![path_element(
            FillRule::NonZero,
            Some(linear_gradient()),
            Some(solid_stroke()),
        )]);

        assert!(content.contains("/Pattern cs"), "{content}");
        assert!(has_operator(&content, "f"), "{content}");
        assert!(has_operator(&content, "S"), "{content}");
    }

    #[test]
    fn gradient_stroke_uses_pattern_stroke_operators() {
        let content = content_for(vec![path_element(
            FillRule::NonZero,
            None,
            Some(Stroke {
                paint: linear_gradient(),
                ..solid_stroke()
            }),
        )]);

        assert!(content.contains("/Pattern CS"), "{content}");
        assert!(content.contains("/P0 SCN"), "{content}");
        assert!(has_operator(&content, "S"), "{content}");
    }

    #[test]
    fn page_patterns_include_only_gradients_used_on_that_page() {
        let first = page_with(vec![path_element(
            FillRule::NonZero,
            Some(linear_gradient()),
            None,
        )]);
        let second = PageFrame::new(
            2,
            612.0,
            792.0,
            vec![path_element(
                FillRule::NonZero,
                Some(Paint::Radial {
                    center: Point { x: 50.0, y: 50.0 },
                    radius: 40.0,
                    focal: Point { x: 50.0, y: 50.0 },
                    stops: gradient_stops(),
                    extend: (false, false),
                }),
                None,
            )],
        );
        let layout = LayoutResult::new(vec![first, second], Vec::new(), None, Vec::new());
        let pdf = String::from_utf8_lossy(&write_pdf(&layout)).into_owned();

        let pattern_resources: Vec<_> = pdf
            .split("/Pattern <<")
            .skip(1)
            .map(|tail| tail.split(">>").next().expect("pattern dictionary end"))
            .collect();
        assert_eq!(pattern_resources.len(), 2, "{pdf}");
        assert!(pattern_resources[0].contains("/P0"), "{pdf}");
        assert!(!pattern_resources[0].contains("/P1"), "{pdf}");
        assert!(pattern_resources[1].contains("/P1"), "{pdf}");
        assert!(!pattern_resources[1].contains("/P0"), "{pdf}");
    }

    #[test]
    fn rotated_linear_gradient_renders_with_its_axis_rotated() {
        let gradient_path = PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 20.0,
                y: 40.0,
                width: 80.0,
                height: 40.0,
            }),
            fill: Some(Paint::Linear {
                start: Point { x: 20.0, y: 60.0 },
                end: Point { x: 100.0, y: 60.0 },
                stops: gradient_stops(),
                extend: (false, false),
            }),
            stroke: None,
        });
        let rotated = group(
            Transform::rotate_about(90.0, 60.0, 60.0),
            vec![gradient_path],
        );
        let pdf = write_pdf(&LayoutResult::new(
            vec![PageFrame::new(1, 120.0, 120.0, vec![rotated])],
            Vec::new(),
            None,
            Vec::new(),
        ));
        let base = std::env::temp_dir().join(format!("oxml-pdf-f043-{}", std::process::id()));
        let pdf_path = base.with_extension("pdf");
        let png_path = base.with_extension("png");
        std::fs::write(&pdf_path, pdf).expect("write PDF fixture");
        let version = std::process::Command::new("pdftoppm")
            .arg("-v")
            .output()
            .expect("pdftoppm is required for the F-043 gate");
        let version = String::from_utf8_lossy(&version.stderr);
        assert!(version.starts_with("pdftoppm version 26.01.0"), "{version}");
        let status = std::process::Command::new("pdftoppm")
            .args(["-png", "-r", "72", "-f", "1", "-singlefile"])
            .arg(&pdf_path)
            .arg(&base)
            .status()
            .expect("rasterise F-043 fixture");
        assert!(status.success());
        let pixmap = tiny_skia::Pixmap::load_png(&png_path).expect("decode raster fixture");
        let top = pixmap.pixel(60, 30).expect("top sample");
        let bottom = pixmap.pixel(60, 90).expect("bottom sample");
        assert_eq!(
            (top.red(), top.green(), top.blue(), top.alpha()),
            (223, 0, 32, 255)
        );
        assert_eq!(
            (bottom.red(), bottom.green(), bottom.blue(), bottom.alpha()),
            (32, 0, 223, 255)
        );
        let _ = std::fs::remove_file(pdf_path);
        let _ = std::fs::remove_file(png_path);
    }

    fn alpha_rect(alpha: f64) -> PositionedElement {
        PositionedElement::FilledRect {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            color: Color {
                r: 0.0,
                g: 0.0,
                b: 0.0,
                a: alpha,
            },
        }
    }

    fn pdf_for(elements: Vec<PositionedElement>) -> String {
        let layout = LayoutResult::new(vec![page_with(elements)], Vec::new(), None, Vec::new());
        String::from_utf8_lossy(&write_pdf(&layout)).into_owned()
    }

    #[test]
    fn same_alpha_reuses_one_extgstate() {
        let pdf = pdf_for(vec![alpha_rect(0.5), alpha_rect(0.5)]);

        assert!(pdf.contains("/ExtGState"), "{pdf}");
        assert_eq!(pdf.matches("/CA 0.5").count(), 1, "{pdf}");
        assert_eq!(pdf.matches("/ca 0.5").count(), 1, "{pdf}");
    }

    #[test]
    fn distinct_alpha_values_create_distinct_states() {
        let pdf = pdf_for(vec![alpha_rect(0.5), alpha_rect(0.25)]);

        assert_eq!(pdf.matches("/CA ").count(), 2, "{pdf}");
        assert_eq!(pdf.matches("/ca ").count(), 2, "{pdf}");
    }

    #[test]
    fn opaque_elements_emit_no_alpha_state() {
        let pdf = pdf_for(vec![alpha_rect(1.0)]);

        assert!(!pdf.contains("/ExtGState"), "{pdf}");
        assert!(!pdf.contains("/CA "), "{pdf}");
        assert!(!pdf.contains("/ca "), "{pdf}");
    }

    #[test]
    fn alpha_state_sets_fill_and_stroke_constants() {
        let content = content_for(vec![alpha_rect(0.5)]);
        let pdf = pdf_for(vec![alpha_rect(0.5)]);

        assert!(content.contains("/GS0 gs"), "{content}");
        assert!(pdf.contains("/CA 0.5"), "{pdf}");
        assert!(pdf.contains("/ca 0.5"), "{pdf}");
    }

    #[test]
    fn solid_path_alpha_uses_shared_registry() {
        let half_black = Color {
            r: 0.0,
            g: 0.0,
            b: 0.0,
            a: 0.5,
        };
        let path = path_element(
            FillRule::NonZero,
            Some(Paint::Solid(half_black)),
            Some(Stroke {
                paint: Paint::Solid(half_black),
                ..solid_stroke()
            }),
        );
        let content = content_for(vec![alpha_rect(0.5), path]);
        let pdf = pdf_for(vec![
            alpha_rect(0.5),
            path_element(
                FillRule::NonZero,
                Some(Paint::Solid(half_black)),
                Some(Stroke {
                    paint: Paint::Solid(half_black),
                    ..solid_stroke()
                }),
            ),
        ]);

        assert_eq!(pdf.matches("/CA 0.5").count(), 1, "{pdf}");
        assert_eq!(content.matches("/GS0 gs").count(), 2, "{content}");
        assert!(has_operator(&content, "B"), "{content}");
    }

    #[test]
    fn different_fill_and_stroke_alpha_repeat_path_geometry() {
        let fill = Color {
            r: 0.0,
            g: 0.0,
            b: 0.0,
            a: 0.25,
        };
        let stroke = Color { a: 0.5, ..fill };
        let content = content_for(vec![path_element(
            FillRule::NonZero,
            Some(Paint::Solid(fill)),
            Some(Stroke {
                paint: Paint::Solid(stroke),
                ..solid_stroke()
            }),
        )]);

        assert_eq!(content.matches("10 20 m").count(), 2, "{content}");
        assert!(has_operator(&content, "f"), "{content}");
        assert!(has_operator(&content, "S"), "{content}");
        assert!(!has_operator(&content, "B"), "{content}");
    }

    #[test]
    fn opaque_stroke_restores_after_transparent_fill() {
        let fill = Color {
            r: 0.0,
            g: 0.0,
            b: 0.0,
            a: 0.25,
        };
        let content = content_for(vec![path_element(
            FillRule::NonZero,
            Some(Paint::Solid(fill)),
            Some(Stroke {
                paint: Paint::Solid(Color::BLACK),
                ..solid_stroke()
            }),
        )]);

        assert_eq!(content.matches("/GS0 gs").count(), 1, "{content}");
        assert_eq!(content.lines().filter(|line| *line == "q").count(), 4);
        assert_eq!(content.lines().filter(|line| *line == "Q").count(), 4);
    }

    fn group(transform: Transform, children: Vec<PositionedElement>) -> PositionedElement {
        PositionedElement::Group(GroupElement {
            transform,
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children,
        })
    }

    #[test]
    fn three_deep_groups_balance_graphics_state() {
        let content = content_for(vec![group(
            Transform::IDENTITY,
            vec![group(
                Transform::IDENTITY,
                vec![group(Transform::IDENTITY, vec![alpha_rect(1.0)])],
            )],
        )]);

        assert_eq!(
            content.lines().filter(|line| *line == "q").count(),
            content.lines().filter(|line| *line == "Q").count(),
            "{content}"
        );
        assert_eq!(content.lines().filter(|line| *line == "q").count(), 5);
    }

    #[test]
    fn group_emits_transform_before_children() {
        let transform = Transform {
            a: 1.0,
            b: 2.0,
            c: 3.0,
            d: 4.0,
            e: 5.0,
            f: 6.0,
        };
        let content = content_for(vec![group(transform, vec![alpha_rect(1.0)])]);

        let matrix = content.find("1 2 3 4 5 6 cm").expect("group matrix");
        let child = content.find("0 0 10 10 re").expect("child rectangle");
        assert!(matrix < child, "{content}");
    }

    #[test]
    fn group_clip_uses_declared_fill_rule() {
        let mut nonzero = match group(Transform::IDENTITY, Vec::new()) {
            PositionedElement::Group(group) => group,
            _ => unreachable!(),
        };
        nonzero.clip = Some(Path::rect(Rect {
            x: 0.0,
            y: 0.0,
            width: 10.0,
            height: 10.0,
        }));
        let mut even_odd = nonzero.clone();
        even_odd.clip.as_mut().expect("clip").fill_rule = FillRule::EvenOdd;
        let content = content_for(vec![
            PositionedElement::Group(nonzero),
            PositionedElement::Group(even_odd),
        ]);

        assert!(content.contains("W\nn"), "{content}");
        assert!(content.contains("W*\nn"), "{content}");
    }

    #[test]
    fn group_opacity_uses_registered_graphics_state() {
        let mut transparent = match group(Transform::IDENTITY, vec![alpha_rect(1.0)]) {
            PositionedElement::Group(group) => group,
            _ => unreachable!(),
        };
        transparent.opacity = 0.5;
        let content = content_for(vec![PositionedElement::Group(transparent)]);

        let alpha = content.find("/GS0 gs").expect("group opacity state");
        let child = content.find("0 0 10 10 re").expect("child rectangle");
        assert!(alpha < child, "{content}");
    }

    #[test]
    fn group_effects_do_not_change_pdf_output_yet() {
        let plain = group(Transform::IDENTITY, vec![alpha_rect(1.0)]);
        let mut with_effect = match plain.clone() {
            PositionedElement::Group(group) => group,
            _ => unreachable!(),
        };
        with_effect.effects.push(Effect::OuterShadow {
            dx: 1.0,
            dy: 2.0,
            blur: 3.0,
            color: Color::BLACK,
        });

        assert_eq!(
            content_for(vec![plain]),
            content_for(vec![PositionedElement::Group(with_effect)])
        );
    }

    #[test]
    fn grouped_text_is_included_in_font_subsetting() {
        let font_id = FontId(42);
        let text = PositionedElement::Text(GlyphRun {
            origin: Point { x: 0.0, y: 0.0 },
            font_id,
            font_size: 12.0,
            glyph_ids: vec![7],
            advances: vec![6.0],
            text: "A".to_owned(),
            source: None,
            color: Color::BLACK,
            bold: false,
            italic: false,
            field_kind: None,
            note: None,
        });
        let layout = LayoutResult::new(
            vec![page_with(vec![group(Transform::IDENTITY, vec![text])])],
            Vec::new(),
            None,
            Vec::new(),
        );

        assert!(font::collect_glyph_usage(&layout).contains_key(&font_id));
    }

    #[test]
    fn grouped_image_registers_and_uses_xobject() {
        let png = tiny_skia::Pixmap::new(1, 1)
            .expect("one-pixel pixmap")
            .encode_png()
            .expect("encode fixture");
        let image = PositionedElement::Image {
            rect: Rect {
                x: 1.0,
                y: 2.0,
                width: 3.0,
                height: 4.0,
            },
            data: png,
            content_type: "image/png".to_owned(),
            media_id: MediaId(9),
        };
        let pdf = pdf_for(vec![group(Transform::IDENTITY, vec![image])]);

        assert!(pdf.contains("/Subtype /Image"), "{pdf}");
        assert!(pdf.contains("/Im0_0"), "{pdf}");
    }

    #[test]
    fn grouped_link_emits_transformed_annotation() {
        let transform = Transform {
            e: 10.0,
            f: 20.0,
            ..Transform::IDENTITY
        };
        let link = PositionedElement::LinkAnnotation {
            rect: Rect {
                x: 1.0,
                y: 2.0,
                width: 3.0,
                height: 4.0,
            },
            url: "https://example.com".to_owned(),
        };
        let pdf = pdf_for(vec![group(transform, vec![link])]);

        assert!(pdf.contains("/Annots ["), "{pdf}");
        assert!(pdf.contains("/Rect [11 766 14 770]"), "{pdf}");
    }

    #[test]
    fn nested_leaf_ordinals_match_content_order() {
        let png = tiny_skia::Pixmap::new(1, 1)
            .expect("one-pixel pixmap")
            .encode_png()
            .expect("encode fixture");
        let image = PositionedElement::Image {
            rect: Rect {
                x: 1.0,
                y: 2.0,
                width: 3.0,
                height: 4.0,
            },
            data: png,
            content_type: "image/png".to_owned(),
            media_id: MediaId(10),
        };
        let elements = vec![alpha_rect(1.0), group(Transform::IDENTITY, vec![image])];
        let pdf = pdf_for(elements.clone());
        let page = page_with(elements);
        let image_map = HashMap::from([((0, 1), 0)]);
        let mut next_ref = 100;
        let alpha_states = AlphaStates::new(std::slice::from_ref(&page), &mut || {
            let reference = Ref::new(next_ref);
            next_ref += 1;
            reference
        });
        let gradients = GradientRegistry::new(std::slice::from_ref(&page), &mut || {
            let reference = Ref::new(next_ref);
            next_ref += 1;
            reference
        });
        let content = String::from_utf8(build_page_content(
            0,
            &page,
            &BTreeMap::new(),
            &BTreeMap::new(),
            &image_map,
            &alpha_states,
            &gradients,
        ))
        .expect("PDF content operators are ASCII");

        assert!(pdf.contains("/Im0_1"), "{pdf}");
        assert!(content.contains("/Im0_1 Do"), "{content}");
    }

    #[test]
    fn top_level_collection_output_remains_stable() {
        let link = PositionedElement::LinkAnnotation {
            rect: Rect {
                x: 72.0,
                y: 100.0,
                width: 100.0,
                height: 15.0,
            },
            url: "https://example.com".to_owned(),
        };
        let pdf = pdf_for(vec![link]);

        assert!(pdf.contains("/Rect [72 677 172 692]"), "{pdf}");
        assert_eq!(pdf.matches("/Subtype /Link").count(), 1, "{pdf}");
    }

    #[test]
    fn page_content_uses_one_global_flip() {
        let content = content_for(Vec::new());

        assert_eq!(content.matches("1 0 0 -1 0 792 cm").count(), 1);
        assert!(content.starts_with("q\n1 0 0 -1 0 792 cm\n"));
        assert!(content.ends_with('Q'));
    }

    #[test]
    fn text_and_images_cancel_the_outer_flip() {
        let font_id = FontId(7);
        let mut remapper = subsetter::GlyphRemapper::new();
        remapper.remap(1);
        let prepared = PreparedFont {
            font_data: FontData {
                id: font_id,
                family: "Test".to_owned(),
                data: Vec::new().into(),
                face_index: 0,
                bold: false,
                italic: false,
            },
            subset_bytes: Vec::new(),
            remapper,
            cmap_bytes: Vec::new(),
            widths: vec![(0, 500.0)],
        };
        let elements = vec![
            PositionedElement::Text(GlyphRun {
                origin: Point { x: 30.0, y: 40.0 },
                font_id,
                font_size: 12.0,
                glyph_ids: vec![1],
                advances: vec![6.0],
                text: "A".to_owned(),
                source: None,
                color: Color::BLACK,
                bold: false,
                italic: false,
                field_kind: None,
                note: None,
            }),
            PositionedElement::Image {
                rect: Rect {
                    x: 50.0,
                    y: 60.0,
                    width: 70.0,
                    height: 80.0,
                },
                data: vec![1],
                content_type: "image/png".to_owned(),
                media_id: MediaId(1),
            },
        ];
        let page = page_with(elements);
        let prepared_fonts = BTreeMap::from([(font_id, prepared)]);
        let font_refs = BTreeMap::from([(
            font_id,
            (
                Ref::new(1),
                Ref::new(2),
                Ref::new(3),
                Ref::new(4),
                Ref::new(5),
            ),
        )]);
        let image_map = HashMap::from([((0, 1), 0)]);
        let mut next_ref = 100;
        let alpha_states = AlphaStates::new(std::slice::from_ref(&page), &mut || {
            let reference = Ref::new(next_ref);
            next_ref += 1;
            reference
        });
        let gradients = GradientRegistry::new(std::slice::from_ref(&page), &mut || {
            let reference = Ref::new(next_ref);
            next_ref += 1;
            reference
        });
        let content = String::from_utf8(build_page_content(
            0,
            &page,
            &prepared_fonts,
            &font_refs,
            &image_map,
            &alpha_states,
            &gradients,
        ))
        .expect("PDF content operators are ASCII");

        assert!(content.contains("1 0 0 -1 30 40 Tm"));
        assert!(content.contains("70 0 0 -80 50 140 cm"));
    }

    #[test]
    fn lines_and_rectangles_use_top_left_coordinates() {
        let content = content_for(vec![
            PositionedElement::Line {
                start: Point { x: 72.0, y: 73.0 },
                end: Point { x: 540.0, y: 74.0 },
                width: 1.0,
                color: Color::BLACK,
                dash_pattern: None,
            },
            PositionedElement::FilledRect {
                rect: Rect {
                    x: 75.0,
                    y: 101.0,
                    width: 468.0,
                    height: 20.0,
                },
                color: Color::WHITE,
            },
        ]);

        assert!(content.contains("72 73 m"));
        assert!(content.contains("540 74 l"));
        assert!(content.contains("75 101 468 20 re"));
    }

    #[test]
    fn annotations_remain_outside_the_content_transform() {
        let layout = LayoutResult::new(
            vec![page_with(vec![PositionedElement::LinkAnnotation {
                rect: Rect {
                    x: 72.0,
                    y: 100.0,
                    width: 100.0,
                    height: 15.0,
                },
                url: "https://example.com".to_owned(),
            }])],
            Vec::new(),
            None,
            Vec::new(),
        );
        let pdf = String::from_utf8_lossy(&write_pdf(&layout)).into_owned();

        assert!(pdf.contains("/Rect [72 677 172 692]"), "{pdf}");
    }
}
