//! PDF document writer: assembles pages, fonts, images, metadata, and outlines.

use std::collections::HashMap;

use pdf_writer::types::{ActionType, AnnotationType, CidFontType, FontFlags, SystemInfo};
use pdf_writer::{Content, Filter, Finish, Name, Pdf, Rect, Ref, Str, TextStr};
use rdocx_layout::{FontId, LayoutResult, PositionedElement};

use crate::font::{self, PreparedFont};
use crate::image;

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
    let mut prepared_fonts: HashMap<FontId, PreparedFont> = HashMap::new();
    let mut font_refs: HashMap<FontId, (Ref, Ref, Ref, Ref, Ref)> = HashMap::new(); // type0, cid, descriptor, stream, cmap

    for fd in &layout.fonts {
        if let Some(usage) = glyph_usage.get_mut(&fd.id) {
            if let Some(prepared) = font::prepare_font(fd, usage) {
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
        for (elem_idx, element) in page.elements.iter().enumerate() {
            if let PositionedElement::Image {
                data, content_type, ..
            } = element
            {
                if data.is_empty() {
                    continue;
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
                    image_map.insert((page_idx, elem_idx), idx);
                }
            }
        }
    }

    // ── Annotation refs ──────────────────────────────────────────────
    let mut annotation_refs: HashMap<(usize, usize), Ref> = HashMap::new();
    for (page_idx, page) in layout.pages.iter().enumerate() {
        for (elem_idx, element) in page.elements.iter().enumerate() {
            if let PositionedElement::LinkAnnotation { .. } = element {
                annotation_refs.insert((page_idx, elem_idx), alloc());
            }
        }
    }

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

    // ── Write pages ──────────────────────────────────────────────────
    for (page_idx, page) in layout.pages.iter().enumerate() {
        let page_id = page_ids[page_idx];
        let content_id = content_ids[page_idx];

        // Build content stream
        let content_bytes =
            build_page_content(page_idx, page, &prepared_fonts, &font_refs, &image_map);

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
        let page_image_names: Vec<(String, Ref)> = image_map
            .iter()
            .filter(|((pi, _), _)| *pi == page_idx)
            .map(|((_, ei), img_idx)| {
                (
                    format!("Im{}_{}", page_idx, ei),
                    image_entries[*img_idx].xobject_ref,
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

        resources.finish();

        // Annotations (hyperlinks)
        let page_annotations: Vec<Ref> = annotation_refs
            .iter()
            .filter(|((pi, _), _)| *pi == page_idx)
            .map(|(_, annot_ref)| *annot_ref)
            .collect();

        if !page_annotations.is_empty() {
            page_dict.annotations(page_annotations.iter().copied());
        }

        page_dict.finish();
    }

    // ── Write link annotations ───────────────────────────────────────
    for (page_idx, page) in layout.pages.iter().enumerate() {
        for (elem_idx, element) in page.elements.iter().enumerate() {
            if let PositionedElement::LinkAnnotation { rect, url } = element {
                if let Some(&annot_ref) = annotation_refs.get(&(page_idx, elem_idx)) {
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
            }
        }
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

/// Build the content stream for a single page.
fn build_page_content(
    page_idx: usize,
    page: &rdocx_layout::PageFrame,
    prepared_fonts: &HashMap<FontId, PreparedFont>,
    font_refs: &HashMap<FontId, (Ref, Ref, Ref, Ref, Ref)>,
    image_map: &HashMap<(usize, usize), usize>,
) -> Vec<u8> {
    let mut content = Content::new();
    let page_height = page.height as f32;

    for (elem_idx, element) in page.elements.iter().enumerate() {
        match element {
            PositionedElement::Text(run) => {
                if let Some(prepared) = prepared_fonts.get(&run.font_id) {
                    if font_refs.contains_key(&run.font_id) {
                        let font_name = format!("F{}", run.font_id.0);

                        content.save_state();

                        // Set text color
                        content.set_fill_rgb(
                            run.color.r as f32,
                            run.color.g as f32,
                            run.color.b as f32,
                        );

                        content.begin_text();
                        content.set_font(Name(font_name.as_bytes()), run.font_size as f32);

                        // PDF coordinate system: origin at bottom-left
                        let pdf_y = page_height - run.origin.y as f32;
                        content.set_text_matrix([1.0, 0.0, 0.0, 1.0, run.origin.x as f32, pdf_y]);

                        // Remap glyph IDs and emit with TJ operator
                        emit_glyphs(
                            &mut content,
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
                content.set_line_width(*width as f32);
                if let Some((on, off)) = dash_pattern {
                    content.set_dash_pattern([*on as f32, *off as f32], 0.0);
                }
                content.move_to(start.x as f32, page_height - start.y as f32);
                content.line_to(end.x as f32, page_height - end.y as f32);
                content.stroke();
                content.restore_state();
            }
            PositionedElement::FilledRect { rect, color } => {
                content.save_state();
                content.set_fill_rgb(color.r as f32, color.g as f32, color.b as f32);
                content.rect(
                    rect.x as f32,
                    page_height - rect.y as f32 - rect.height as f32,
                    rect.width as f32,
                    rect.height as f32,
                );
                content.fill_nonzero();
                content.restore_state();
            }
            PositionedElement::Image { rect, data, .. } => {
                if !data.is_empty() && image_map.contains_key(&(page_idx, elem_idx)) {
                    let img_name = format!("Im{}_{}", page_idx, elem_idx);

                    content.save_state();
                    // Image transformation matrix: scale and position
                    // PDF images are 1x1 unit, we need to scale to rect size
                    // and translate to position (bottom-left origin)
                    let pdf_y = page_height - rect.y as f32 - rect.height as f32;
                    content.transform([
                        rect.width as f32,
                        0.0,
                        0.0,
                        rect.height as f32,
                        rect.x as f32,
                        pdf_y,
                    ]);
                    content.x_object(Name(img_name.as_bytes()));
                    content.restore_state();
                }
            }
            PositionedElement::LinkAnnotation { .. } => {
                // Link annotations are written separately, not in the content stream
            }
        }
    }

    content.finish().to_vec()
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
    outlines: &[rdocx_layout::output::OutlineEntry],
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
