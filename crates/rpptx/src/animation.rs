use std::io::{self, Seek, SeekFrom, Write};
use std::num::NonZeroU16;
use std::sync::Arc;

use gif::{DisposalMethod, Encoder as GifEncoder, Frame as GifFrame, Repeat};
use jpeg_encoder::{ColorType as JpegColorType, Encoder as JpegEncoder};
use oxml_layout::{Diagnostic, FontData, LayoutResult, PageFrame};
use tiny_skia::{Pixmap, PixmapPaint, Transform};

use super::{
    MediaFallbackPolicy, PreparedRenderAssembly, Presentation, Result, TimelinePosition,
    TimelineRequest, prepare_render_context, render_failure, render_prepared_timeline_request,
};

const MAX_FRAME_RATE: u16 = 100;
const MAX_DIMENSION_PX: u32 = i16::MAX as u32;
const MAX_FRAME_PIXELS: u64 = 16_777_216;
const MAX_FRAMES: u64 = 10_000;
const MAX_TOTAL_PIXELS: u64 = 268_435_456;
const MAX_OUTPUT_BYTES: u64 = 1_073_741_824;

struct CappedBuffer {
    bytes: Vec<u8>,
    position: usize,
    cap: usize,
}

impl CappedBuffer {
    fn new(cap: usize) -> Self {
        Self {
            bytes: Vec::new(),
            position: 0,
            cap,
        }
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }

    fn len(&self) -> usize {
        self.bytes.len()
    }
}

impl Write for CappedBuffer {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let end = self
            .position
            .checked_add(buffer.len())
            .ok_or_else(output_cap_error)?;
        if end > self.cap {
            return Err(output_cap_error());
        }
        if self.position > self.bytes.len() {
            self.bytes.resize(self.position, 0);
        }
        if end > self.bytes.len() {
            self.bytes.resize(end, 0);
        }
        self.bytes[self.position..end].copy_from_slice(buffer);
        self.position = end;
        Ok(buffer.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Seek for CappedBuffer {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let base = match position {
            SeekFrom::Start(offset) => {
                let target = usize::try_from(offset).map_err(|_| output_cap_error())?;
                if target > self.cap {
                    return Err(output_cap_error());
                }
                self.position = target;
                return Ok(offset);
            }
            SeekFrom::End(_) => self.bytes.len() as i128,
            SeekFrom::Current(_) => self.position as i128,
        };
        let offset = match position {
            SeekFrom::End(offset) | SeekFrom::Current(offset) => i128::from(offset),
            SeekFrom::Start(_) => unreachable!("start returned above"),
        };
        let target = base.checked_add(offset).ok_or_else(output_cap_error)?;
        if target < 0 || target > self.cap as i128 {
            return Err(output_cap_error());
        }
        self.position = target as usize;
        Ok(self.position as u64)
    }
}

fn output_cap_error() -> io::Error {
    io::Error::other("animation output byte cap exceeded")
}

/// The outgoing slide supplied while a segment's incoming transition runs.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum AnimationTransition {
    None,
    FromSlide(usize),
}

/// Loop metadata written to an animated GIF.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum GifLoopBehavior {
    Once,
    Infinite,
    TotalPlays(NonZeroU16),
}

/// Deterministic animated output container and its bounded codec options.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum AnimationFormat {
    Gif { loop_behavior: GifLoopBehavior },
    MotionJpegAvi { quality: u8 },
}

/// One slide-local interval sampled at a fixed click state.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct AnimationSegment {
    pub slide_index: usize,
    pub duration_ms: u64,
    pub click_count: u32,
    pub transition: AnimationTransition,
}

/// Bounds, format, and media policy for one animation export.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct AnimationExportOptions {
    pub frame_rate: u16,
    pub width_px: u32,
    pub height_px: u32,
    pub format: AnimationFormat,
    pub media_fallback: MediaFallbackPolicy,
}

/// One deterministic animated container and its exact sampling evidence.
#[derive(Clone, Debug)]
pub struct DeterministicAnimation {
    pub bytes: Vec<u8>,
    pub frame_timestamps_ms: Vec<u64>,
    pub diagnostics: Vec<Diagnostic>,
}

#[derive(Clone, Copy)]
struct FrameSample {
    slide_index: usize,
    local_timestamp_ms: u64,
    output_timestamp_ms: u64,
    click_count: u32,
    outgoing_slide_index: Option<usize>,
}

struct PreparedAnimationAssembly {
    assembly: PreparedRenderAssembly,
    slide_count: usize,
    fallback: MediaFallbackPolicy,
}

impl PreparedAnimationAssembly {
    fn render_with_diagnostics(
        &mut self,
        sample: FrameSample,
    ) -> Result<(PageFrame, Vec<Diagnostic>)> {
        render_prepared_timeline_request(
            &mut self.assembly,
            TimelineRequest {
                slide_index: sample.slide_index,
                outgoing_slide_index: sample.outgoing_slide_index,
                position: TimelinePosition {
                    elapsed_ms: sample.local_timestamp_ms,
                    click_count: sample.click_count,
                },
                fallback_policy: Some(self.fallback),
            },
            self.slide_count,
            true,
        )
        .map(|(frame, _)| (frame.page, frame.diagnostics))
    }
}

impl Presentation {
    /// Samples explicit slide segments and writes a bounded deterministic animation.
    pub fn export_animation_deterministic(
        &self,
        segments: &[AnimationSegment],
        options: AnimationExportOptions,
    ) -> Result<DeterministicAnimation> {
        let samples = validate_and_sample(segments, options, self.slides.len())?;
        let package = self.staged_package(false)?;
        let mut prepared = PreparedAnimationAssembly {
            assembly: prepare_render_context(&package, true, &[])?,
            slide_count: self.slides.len(),
            fallback: options.media_fallback,
        };
        match options.format {
            AnimationFormat::Gif { loop_behavior } => {
                encode_gif(&mut prepared, &samples, options, loop_behavior)
            }
            AnimationFormat::MotionJpegAvi { quality } => {
                encode_motion_jpeg_avi(&mut prepared, &samples, options, quality)
            }
        }
    }
}

fn validate_and_sample(
    segments: &[AnimationSegment],
    options: AnimationExportOptions,
    slide_count: usize,
) -> Result<Vec<FrameSample>> {
    if segments.is_empty() {
        return Err(render_failure("animation requires at least one segment"));
    }
    if options.frame_rate == 0 || options.frame_rate > MAX_FRAME_RATE {
        return Err(render_failure(format!(
            "animation frame rate must be in 1 through {MAX_FRAME_RATE}"
        )));
    }
    if options.width_px == 0
        || options.height_px == 0
        || options.width_px > MAX_DIMENSION_PX
        || options.height_px > MAX_DIMENSION_PX
    {
        return Err(render_failure(format!(
            "animation dimensions must be in 1 through {MAX_DIMENSION_PX} pixels"
        )));
    }
    if let AnimationFormat::MotionJpegAvi { quality } = options.format
        && !(1..=100).contains(&quality)
    {
        return Err(render_failure(
            "Motion JPEG quality must be in 1 through 100",
        ));
    }
    let pixels = u64::from(options.width_px)
        .checked_mul(u64::from(options.height_px))
        .ok_or_else(|| render_failure("animation pixel count overflow"))?;
    if pixels > MAX_FRAME_PIXELS {
        return Err(render_failure(format!(
            "animation frame exceeds the {MAX_FRAME_PIXELS} pixel cap"
        )));
    }

    let mut samples = Vec::new();
    let mut output_base_ms = 0_u64;
    for segment in segments {
        if segment.slide_index >= slide_count {
            return Err(super::Error::UnknownSlideIndex {
                index: segment.slide_index,
                slide_count,
            });
        }
        if let AnimationTransition::FromSlide(index) = segment.transition
            && index >= slide_count
        {
            return Err(super::Error::UnknownSlideIndex { index, slide_count });
        }
        if segment.duration_ms == 0 {
            return Err(render_failure(
                "animation segment duration must be positive",
            ));
        }
        let scaled = u128::from(segment.duration_ms) * u128::from(options.frame_rate);
        let frame_count = scaled.div_ceil(1000);
        let frame_count = u64::try_from(frame_count)
            .map_err(|_| render_failure("animation frame count overflow"))?;
        let new_total = u64::try_from(samples.len())
            .ok()
            .and_then(|count| count.checked_add(frame_count))
            .ok_or_else(|| render_failure("animation frame count overflow"))?;
        if new_total > MAX_FRAMES {
            return Err(render_failure(format!(
                "animation exceeds the {MAX_FRAMES} frame cap"
            )));
        }
        for frame_index in 0..frame_count {
            let local_timestamp_ms = frame_index
                .checked_mul(1000)
                .ok_or_else(|| render_failure("animation timestamp overflow"))?
                / u64::from(options.frame_rate);
            if local_timestamp_ms >= segment.duration_ms {
                return Err(render_failure(
                    "animation sampling crossed a segment duration",
                ));
            }
            let output_timestamp_ms = output_base_ms
                .checked_add(local_timestamp_ms)
                .ok_or_else(|| render_failure("animation timestamp overflow"))?;
            samples.push(FrameSample {
                slide_index: segment.slide_index,
                local_timestamp_ms,
                output_timestamp_ms,
                click_count: segment.click_count,
                outgoing_slide_index: match segment.transition {
                    AnimationTransition::None => None,
                    AnimationTransition::FromSlide(index) => Some(index),
                },
            });
        }
        output_base_ms = output_base_ms
            .checked_add(segment.duration_ms)
            .ok_or_else(|| render_failure("animation duration overflow"))?;
    }
    let total_pixels = pixels
        .checked_mul(samples.len() as u64)
        .ok_or_else(|| render_failure("animation total pixel count overflow"))?;
    if total_pixels > MAX_TOTAL_PIXELS {
        return Err(render_failure(format!(
            "animation exceeds the {MAX_TOTAL_PIXELS} sampled-pixel cap"
        )));
    }
    let estimated_output = total_pixels
        .checked_mul(4)
        .and_then(|bytes| bytes.checked_add(samples.len() as u64 * 32 + 4096))
        .ok_or_else(|| render_failure("animation output size overflow"))?;
    if estimated_output > MAX_OUTPUT_BYTES {
        return Err(render_failure(format!(
            "animation exceeds the {MAX_OUTPUT_BYTES} byte output cap"
        )));
    }
    Ok(samples)
}

fn encode_gif(
    prepared: &mut PreparedAnimationAssembly,
    samples: &[FrameSample],
    options: AnimationExportOptions,
    loop_behavior: GifLoopBehavior,
) -> Result<DeterministicAnimation> {
    let width = u16::try_from(options.width_px)
        .map_err(|_| render_failure("GIF width cannot be represented"))?;
    let height = u16::try_from(options.height_px)
        .map_err(|_| render_failure("GIF height cannot be represented"))?;
    let mut output = CappedBuffer::new(MAX_OUTPUT_BYTES as usize);
    let mut timestamps = Vec::with_capacity(samples.len());
    let mut diagnostics = Vec::new();
    {
        let mut encoder = GifEncoder::new(&mut output, width, height, &[])
            .map_err(|error| render_failure(format!("GIF header encoding failed: {error}")))?;
        match loop_behavior {
            GifLoopBehavior::Once => {}
            GifLoopBehavior::Infinite => encoder
                .set_repeat(Repeat::Infinite)
                .map_err(|error| render_failure(format!("GIF loop encoding failed: {error}")))?,
            GifLoopBehavior::TotalPlays(plays) if plays.get() > 1 => encoder
                .set_repeat(Repeat::Finite(plays.get() - 1))
                .map_err(|error| render_failure(format!("GIF loop encoding failed: {error}")))?,
            GifLoopBehavior::TotalPlays(_) => {}
        }
        for (index, sample) in samples.iter().copied().enumerate() {
            let (page, mut frame_diagnostics) = prepared.render_with_diagnostics(sample)?;
            diagnostics.append(&mut frame_diagnostics);
            let mut rgba = rasterize_opaque_exact(
                &page,
                prepared.assembly.font_manager.all_font_data(),
                options.width_px,
                options.height_px,
            )?;
            let mut frame = GifFrame::from_rgba_speed(width, height, &mut rgba, 10);
            frame.delay = gif_delay(index as u64, options.frame_rate);
            frame.dispose = DisposalMethod::Keep;
            encoder
                .write_frame(&frame)
                .map_err(|error| render_failure(format!("GIF frame encoding failed: {error}")))?;
            timestamps.push(sample.output_timestamp_ms);
        }
    }
    Ok(DeterministicAnimation {
        bytes: output.into_inner(),
        frame_timestamps_ms: timestamps,
        diagnostics,
    })
}

fn gif_delay(frame_index: u64, frame_rate: u16) -> u16 {
    let rate = u64::from(frame_rate);
    let start = frame_index * 100 / rate;
    let end = (frame_index + 1) * 100 / rate;
    u16::try_from(end - start).expect("frame rate cap keeps GIF delay representable")
}

fn encode_motion_jpeg_avi(
    prepared: &mut PreparedAnimationAssembly,
    samples: &[FrameSample],
    options: AnimationExportOptions,
    quality: u8,
) -> Result<DeterministicAnimation> {
    let width = u16::try_from(options.width_px)
        .map_err(|_| render_failure("Motion JPEG width cannot be represented"))?;
    let height = u16::try_from(options.height_px)
        .map_err(|_| render_failure("Motion JPEG height cannot be represented"))?;
    let mut output = CappedBuffer::new(MAX_OUTPUT_BYTES as usize);
    let mut index_entries = Vec::with_capacity(samples.len());
    let mut timestamps = Vec::with_capacity(samples.len());
    let mut diagnostics = Vec::new();
    let mut max_payload = 0_u32;
    let header = write_avi_header(
        &mut output,
        options.width_px,
        options.height_px,
        options.frame_rate,
        samples.len(),
    )?;
    for sample in samples.iter().copied() {
        let (page, mut frame_diagnostics) = prepared.render_with_diagnostics(sample)?;
        diagnostics.append(&mut frame_diagnostics);
        let rgba = rasterize_opaque_exact(
            &page,
            prepared.assembly.font_manager.all_font_data(),
            options.width_px,
            options.height_px,
        )?;
        let mut rgb = Vec::with_capacity(rgba.len() / 4 * 3);
        for pixel in rgba.chunks_exact(4) {
            rgb.extend_from_slice(&pixel[..3]);
        }
        let offset = u32::try_from(output.len() - header.movi_type_position)
            .map_err(|_| render_failure("Motion JPEG AVI index offset overflow"))?;
        output
            .write_all(b"00dc")
            .map_err(|error| render_failure(format!("AVI frame header failed: {error}")))?;
        let size_position = output.len();
        write_u32(&mut output, 0)?;
        let payload_start = output.len();
        JpegEncoder::new(&mut output, quality)
            .encode(&rgb, width, height, JpegColorType::Rgb)
            .map_err(|error| render_failure(format!("JPEG frame encoding failed: {error}")))?;
        let payload_size = u32::try_from(output.len() - payload_start)
            .map_err(|_| render_failure("Motion JPEG frame exceeds AVI chunk limits"))?;
        max_payload = max_payload.max(payload_size);
        patch_u32(&mut output, size_position, payload_size)?;
        if !payload_size.is_multiple_of(2) {
            output
                .write_all(&[0])
                .map_err(|error| render_failure(format!("AVI frame padding failed: {error}")))?;
        }
        index_entries.push((offset, payload_size));
        timestamps.push(sample.output_timestamp_ms);
    }
    finish_avi(&mut output, header, &index_entries, max_payload)?;
    Ok(DeterministicAnimation {
        bytes: output.into_inner(),
        frame_timestamps_ms: timestamps,
        diagnostics,
    })
}

fn rasterize_opaque_exact(
    page: &PageFrame,
    fonts: Vec<FontData>,
    width: u32,
    height: u32,
) -> Result<Vec<u8>> {
    if !page.width.is_finite() || page.width <= 0.0 {
        return Err(render_failure("animation page width is invalid"));
    }
    let layout = LayoutResult::new(vec![Arc::new(page.clone())], fonts, None, Vec::new());
    let dpi = 72.0 * f64::from(width) / page.width;
    let png = oxml_pdf::render_page_to_png(&layout, 0, dpi)
        .ok_or_else(|| render_failure("animation rasterization failed"))?;
    let source = Pixmap::decode_png(&png)
        .map_err(|error| render_failure(format!("animation raster decode failed: {error}")))?;
    let mut destination = Pixmap::new(width, height)
        .ok_or_else(|| render_failure("animation raster dimensions are invalid"))?;
    destination.fill(tiny_skia::Color::WHITE);
    destination.draw_pixmap(
        0,
        0,
        source.as_ref(),
        &PixmapPaint::default(),
        Transform::from_scale(
            width as f32 / source.width() as f32,
            height as f32 / source.height() as f32,
        ),
        None,
    );
    let mut rgba = Vec::with_capacity(width as usize * height as usize * 4);
    for pixel in destination.pixels() {
        let alpha = pixel.alpha();
        rgba.push(pixel.red().saturating_add(255 - alpha));
        rgba.push(pixel.green().saturating_add(255 - alpha));
        rgba.push(pixel.blue().saturating_add(255 - alpha));
        rgba.push(255);
    }
    Ok(rgba)
}

#[derive(Clone, Copy)]
struct AviHeaderPositions {
    riff_size: usize,
    avih_max_payload: usize,
    strh_max_payload: usize,
    movi_size: usize,
    movi_type_position: usize,
}

fn write_avi_header(
    output: &mut CappedBuffer,
    width: u32,
    height: u32,
    frame_rate: u16,
    frame_count: usize,
) -> Result<AviHeaderPositions> {
    let total_frames = u32::try_from(frame_count)
        .map_err(|_| render_failure("Motion JPEG frame count exceeds AVI limits"))?;
    write_bytes(output, b"RIFF")?;
    let riff_size = output.len();
    write_u32(output, 0)?;
    write_bytes(output, b"AVI ")?;

    write_bytes(output, b"LIST")?;
    let hdrl_size = output.len();
    write_u32(output, 0)?;
    write_bytes(output, b"hdrl")?;

    write_bytes(output, b"avih")?;
    write_u32(output, 56)?;
    write_u32(output, 1_000_000 / u32::from(frame_rate))?;
    write_u32(output, 0)?;
    write_u32(output, 0)?;
    write_u32(output, 0x10)?;
    write_u32(output, total_frames)?;
    write_u32(output, 0)?;
    write_u32(output, 1)?;
    let avih_max_payload = output.len();
    write_u32(output, 0)?;
    write_u32(output, width)?;
    write_u32(output, height)?;
    write_bytes(output, &[0; 16])?;

    write_bytes(output, b"LIST")?;
    let strl_size = output.len();
    write_u32(output, 0)?;
    write_bytes(output, b"strl")?;

    write_bytes(output, b"strh")?;
    write_u32(output, 56)?;
    write_bytes(output, b"vidsMJPG")?;
    write_u32(output, 0)?;
    write_u16(output, 0)?;
    write_u16(output, 0)?;
    write_u32(output, 0)?;
    write_u32(output, 1)?;
    write_u32(output, u32::from(frame_rate))?;
    write_u32(output, 0)?;
    write_u32(output, total_frames)?;
    let strh_max_payload = output.len();
    write_u32(output, 0)?;
    write_u32(output, u32::MAX)?;
    write_u32(output, 0)?;
    write_i16(output, 0)?;
    write_i16(output, 0)?;
    write_i16(output, width as i16)?;
    write_i16(output, height as i16)?;

    write_bytes(output, b"strf")?;
    write_u32(output, 40)?;
    write_u32(output, 40)?;
    write_i32(output, width as i32)?;
    write_i32(output, height as i32)?;
    write_u16(output, 1)?;
    write_u16(output, 24)?;
    write_bytes(output, b"MJPG")?;
    let uncompressed_size = width
        .checked_mul(height)
        .and_then(|pixels| pixels.checked_mul(3))
        .ok_or_else(|| render_failure("Motion JPEG bitmap size overflow"))?;
    write_u32(output, uncompressed_size)?;
    write_bytes(output, &[0; 16])?;
    let strl_value = list_size(output.len(), strl_size)?;
    patch_u32(output, strl_size, strl_value)?;
    let hdrl_value = list_size(output.len(), hdrl_size)?;
    patch_u32(output, hdrl_size, hdrl_value)?;

    write_bytes(output, b"LIST")?;
    let movi_size = output.len();
    write_u32(output, 0)?;
    let movi_type_position = output.len();
    write_bytes(output, b"movi")?;
    Ok(AviHeaderPositions {
        riff_size,
        avih_max_payload,
        strh_max_payload,
        movi_size,
        movi_type_position,
    })
}

fn finish_avi(
    output: &mut CappedBuffer,
    header: AviHeaderPositions,
    index_entries: &[(u32, u32)],
    max_payload: u32,
) -> Result<()> {
    let movi_value = list_size(output.len(), header.movi_size)?;
    patch_u32(output, header.movi_size, movi_value)?;
    write_bytes(output, b"idx1")?;
    let index_size = index_entries
        .len()
        .checked_mul(16)
        .and_then(|size| u32::try_from(size).ok())
        .ok_or_else(|| render_failure("AVI index size overflow"))?;
    write_u32(output, index_size)?;
    for &(offset, size) in index_entries {
        write_bytes(output, b"00dc")?;
        write_u32(output, 0x10)?;
        write_u32(output, offset)?;
        write_u32(output, size)?;
    }
    patch_u32(output, header.avih_max_payload, max_payload)?;
    patch_u32(output, header.strh_max_payload, max_payload)?;
    let riff_size = u32::try_from(output.len() - 8)
        .map_err(|_| render_failure("Motion JPEG AVI exceeds RIFF limits"))?;
    patch_u32(output, header.riff_size, riff_size)
}

fn list_size(end: usize, size_position: usize) -> Result<u32> {
    end.checked_sub(size_position + 4)
        .and_then(|size| u32::try_from(size).ok())
        .ok_or_else(|| render_failure("AVI list size overflow"))
}

fn patch_u32(output: &mut CappedBuffer, position: usize, value: u32) -> Result<()> {
    let end = output.len();
    output
        .seek(SeekFrom::Start(position as u64))
        .and_then(|_| output.write_all(&value.to_le_bytes()))
        .and_then(|_| output.seek(SeekFrom::Start(end as u64)))
        .map(|_| ())
        .map_err(|error| render_failure(format!("AVI header patch failed: {error}")))
}

fn write_bytes(output: &mut CappedBuffer, bytes: &[u8]) -> Result<()> {
    output
        .write_all(bytes)
        .map_err(|error| render_failure(format!("AVI encoding failed: {error}")))
}

fn write_u16(output: &mut CappedBuffer, value: u16) -> Result<()> {
    write_bytes(output, &value.to_le_bytes())
}

fn write_i16(output: &mut CappedBuffer, value: i16) -> Result<()> {
    write_bytes(output, &value.to_le_bytes())
}

fn write_u32(output: &mut CappedBuffer, value: u32) -> Result<()> {
    write_bytes(output, &value.to_le_bytes())
}

fn write_i32(output: &mut CappedBuffer, value: i32) -> Result<()> {
    write_bytes(output, &value.to_le_bytes())
}

#[cfg(test)]
mod tests {
    use super::{
        AnimationExportOptions, AnimationFormat, AnimationSegment, AnimationTransition,
        CappedBuffer, GifLoopBehavior, MediaFallbackPolicy, gif_delay, validate_and_sample,
    };
    use crate::{
        CT_TextCharacterProperties, EmbeddedMediaInput, Emu, MediaKind, MediaPlaybackSettings,
        MediaPoster, MediaSourceInput, Presentation, TextFont, animation_preparation_counts,
        max_resolved_sample_retention_count, prepare_render_context,
        reset_resolved_sample_retention_count,
    };
    use gif::{Encoder as GifEncoder, Frame as GifFrame};
    use jpeg_encoder::{ColorType as JpegColorType, Encoder as JpegEncoder};
    use oxml_layout::{PositionedElement, walk};

    #[test]
    fn sampling_uses_integer_millisecond_timestamps_without_crossing_segment_duration() {
        let samples = validate_and_sample(
            &[
                AnimationSegment {
                    slide_index: 0,
                    duration_ms: 101,
                    click_count: 3,
                    transition: AnimationTransition::None,
                },
                AnimationSegment {
                    slide_index: 1,
                    duration_ms: 50,
                    click_count: 4,
                    transition: AnimationTransition::FromSlide(0),
                },
            ],
            AnimationExportOptions {
                frame_rate: 30,
                width_px: 16,
                height_px: 9,
                format: AnimationFormat::Gif {
                    loop_behavior: GifLoopBehavior::Once,
                },
                media_fallback: MediaFallbackPolicy::PosterFrame,
            },
            2,
        )
        .unwrap();
        assert_eq!(samples.len(), 6);
        assert_eq!(
            samples
                .iter()
                .map(|sample| sample.local_timestamp_ms)
                .collect::<Vec<_>>(),
            vec![0, 33, 66, 100, 0, 33]
        );
        assert_eq!(samples[3].click_count, 3);
        assert_eq!(samples[4].click_count, 4);
        assert_eq!(samples[4].output_timestamp_ms, 101);
        assert_eq!(samples[4].outgoing_slide_index, Some(0));
    }

    #[test]
    fn gif_delays_distribute_centisecond_error_without_duration_drift() {
        let delays = (0..6).map(|index| gif_delay(index, 30)).collect::<Vec<_>>();
        assert_eq!(delays, vec![3, 3, 4, 3, 3, 4]);
        assert_eq!(delays.into_iter().sum::<u16>(), 20);
    }

    #[test]
    fn capped_codec_writers_reject_mid_encode_without_exceeding_the_cap() {
        let mut gif_output = CappedBuffer::new(64);
        let mut gif_pixels = (0..64 * 64)
            .flat_map(|index| {
                let value = index as u8;
                [value, value.wrapping_mul(31), value.wrapping_mul(73), 255]
            })
            .collect::<Vec<_>>();
        let gif_result = (|| {
            let mut encoder = GifEncoder::new(&mut gif_output, 64, 64, &[])?;
            encoder.write_frame(&GifFrame::from_rgba_speed(64, 64, &mut gif_pixels, 10))
        })();
        assert!(gif_result.is_err());
        assert!(gif_output.len() <= 64);

        let mut jpeg_output = CappedBuffer::new(128);
        let rgb = (0..64 * 64)
            .flat_map(|index| {
                let value = index as u8;
                [value, value.wrapping_mul(17), value.wrapping_mul(47)]
            })
            .collect::<Vec<_>>();
        let jpeg_result =
            JpegEncoder::new(&mut jpeg_output, 80).encode(&rgb, 64, 64, JpegColorType::Rgb);
        assert!(jpeg_result.is_err());
        assert!(jpeg_output.len() <= 128);
    }

    #[test]
    fn motion_jpeg_avi_indexes_every_frame_at_the_declared_rate_and_dimensions() {
        let mut presentation = Presentation::new().unwrap();
        presentation.add_slide(0).unwrap();
        let exported = presentation
            .export_animation_deterministic(
                &[AnimationSegment {
                    slide_index: 0,
                    duration_ms: 200,
                    click_count: 0,
                    transition: AnimationTransition::None,
                }],
                AnimationExportOptions {
                    frame_rate: 10,
                    width_px: 32,
                    height_px: 18,
                    format: AnimationFormat::MotionJpegAvi { quality: 80 },
                    media_fallback: MediaFallbackPolicy::PosterFrame,
                },
            )
            .unwrap();
        let bytes = exported.bytes;
        let read_u32 = |offset| u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap());
        assert_eq!(&bytes[..4], b"RIFF");
        assert_eq!(read_u32(4) as usize, bytes.len() - 8);
        assert_eq!(&bytes[8..12], b"AVI ");
        assert_eq!(&bytes[12..16], b"LIST");
        assert_eq!(&bytes[20..24], b"hdrl");
        assert_eq!(&bytes[24..28], b"avih");
        assert_eq!(read_u32(48), 2);
        assert_eq!(read_u32(64), 32);
        assert_eq!(read_u32(68), 18);
        assert_eq!(&bytes[108..116], b"vidsMJPG");
        assert_eq!(read_u32(128), 1);
        assert_eq!(read_u32(132), 10);
        assert_eq!(read_u32(140), 2);
        assert_eq!(read_u32(176), 32);
        assert_eq!(read_u32(180), 18);
        assert_eq!(&bytes[188..192], b"MJPG");
        assert_eq!(&bytes[212..216], b"LIST");
        assert_eq!(&bytes[220..224], b"movi");

        let mut cursor = 224;
        let mut frames = Vec::new();
        for _ in 0..2 {
            assert_eq!(&bytes[cursor..cursor + 4], b"00dc");
            let size = read_u32(cursor + 4);
            frames.push(((cursor - 220) as u32, size));
            cursor += 8 + size as usize + size as usize % 2;
        }
        assert_eq!(&bytes[cursor..cursor + 4], b"idx1");
        assert_eq!(read_u32(cursor + 4), 32);
        for (index, &(offset, size)) in frames.iter().enumerate() {
            let entry = cursor + 8 + index * 16;
            assert_eq!(&bytes[entry..entry + 4], b"00dc");
            assert_eq!(read_u32(entry + 4), 0x10);
            assert_eq!(read_u32(entry + 8), offset);
            assert_eq!(read_u32(entry + 12), size);
        }
    }

    #[test]
    fn fifty_frame_export_prepares_package_resolver_media_layout_and_fonts_once() {
        let mut presentation = Presentation::new().unwrap();
        presentation.add_slide(0).unwrap();
        let before = animation_preparation_counts();
        reset_resolved_sample_retention_count();
        let exported = presentation
            .export_animation_deterministic(
                &[AnimationSegment {
                    slide_index: 0,
                    duration_ms: 500,
                    click_count: 0,
                    transition: AnimationTransition::None,
                }],
                AnimationExportOptions {
                    frame_rate: 100,
                    width_px: 16,
                    height_px: 9,
                    format: AnimationFormat::Gif {
                        loop_behavior: GifLoopBehavior::Once,
                    },
                    media_fallback: MediaFallbackPolicy::PosterFrame,
                },
            )
            .unwrap();
        assert_eq!(exported.frame_timestamps_ms.len(), 50);
        assert_eq!(exported.frame_timestamps_ms[0], 0);
        assert_eq!(exported.frame_timestamps_ms[49], 490);
        let after = animation_preparation_counts();
        assert_eq!(after.0 - before.0, 1, "package preparation");
        assert_eq!(after.1 - before.1, 1, "resolver preparation");
        assert_eq!(after.2 - before.2, 1, "media preparation");
        assert_eq!(after.3 - before.3, 1, "layout preparation");
        assert_eq!(after.4 - before.4, 1, "font preparation");
        assert_eq!(max_resolved_sample_retention_count(), 1);
    }

    #[test]
    fn mixed_font_samples_and_dynamic_media_label_share_one_font_identity() {
        let mut presentation = Presentation::new().unwrap();
        presentation.add_slide(0).unwrap();
        presentation.add_slide(0).unwrap();
        for (slide_index, text, family) in [
            (0, "CaladeaWAVE", "Caladea"),
            (1, "Liberationmmmm", "Liberation Sans"),
        ] {
            let mut slide = presentation.slide_mut(slide_index).unwrap();
            let mut textbox = slide
                .add_textbox(Emu(4_000_000), Emu(500_000), Emu(4_500_000), Emu(900_000))
                .unwrap();
            textbox.set_text(text).unwrap();
            let mut frame = textbox.text_frame().unwrap();
            let mut paragraph = frame.paragraph_mut(0).unwrap();
            let mut run = paragraph.run_mut(0).unwrap();
            let mut properties = CT_TextCharacterProperties::default();
            properties.font_size = Some(3_200);
            run.set_properties(properties);
            run.set_font(Some(TextFont::new(family).unwrap()));
        }
        let poster = vec![
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x78,
            0xda, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00, 0xf7, 0x03, 0x41,
            0x43, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
        ];
        presentation
            .add_media(
                1,
                MediaKind::Video,
                MediaSourceInput::Embedded(EmbeddedMediaInput {
                    bytes: b"opaque-video",
                    filename: "video.bin",
                    content_type: "video/x-test-opaque",
                }),
                MediaPoster {
                    bytes: &poster,
                    filename: "poster.png",
                },
                Emu(914_400),
                Emu(1_371_600),
                Emu(2_743_200),
                Emu(1_828_800),
                MediaPlaybackSettings::default(),
            )
            .unwrap();

        let package = presentation.staged_package(false).unwrap();
        let mut prepared = super::PreparedAnimationAssembly {
            assembly: prepare_render_context(&package, true, &[]).unwrap(),
            slide_count: 2,
            fallback: MediaFallbackPolicy::DeterministicPlaceholder,
        };
        let mut identities = Vec::new();
        for (slide_index, expected_texts) in [
            (0, &["CaladeaWAVE"][..]),
            (1, &["Liberationmmmm", "Video"][..]),
        ] {
            let (page, _) = prepared
                .render_with_diagnostics(super::FrameSample {
                    slide_index,
                    local_timestamp_ms: 0,
                    output_timestamp_ms: 0,
                    click_count: 0,
                    outgoing_slide_index: None,
                })
                .unwrap();
            let fonts = prepared.assembly.font_manager.all_font_data();
            walk(&page.elements, &mut |element, _| {
                let (text, font_id) = match element {
                    PositionedElement::Text(run) => (run.text.as_str(), run.font_id),
                    PositionedElement::MultilingualText(run) => {
                        (run.logical_text.as_str(), run.font_id)
                    }
                    _ => return,
                };
                if expected_texts.contains(&text) {
                    let family = fonts
                        .iter()
                        .find(|font| font.id == font_id)
                        .map(|font| font.family.clone())
                        .expect("every sampled glyph id must exist in the retained font table");
                    identities.push((text.to_owned(), family));
                }
            });
        }
        assert_eq!(
            identities,
            [
                ("CaladeaWAVE".to_owned(), "Caladea".to_owned()),
                ("Liberationmmmm".to_owned(), "Liberation Sans".to_owned(),),
                ("Video".to_owned(), "Carlito".to_owned()),
            ]
        );
    }

    #[test]
    fn invalid_animation_requests_fail_before_rendering_or_allocating_output() {
        let gif = AnimationFormat::Gif {
            loop_behavior: GifLoopBehavior::Once,
        };
        let valid = AnimationExportOptions {
            frame_rate: 10,
            width_px: 16,
            height_px: 9,
            format: gif,
            media_fallback: MediaFallbackPolicy::PosterFrame,
        };
        let segment = AnimationSegment {
            slide_index: 0,
            duration_ms: 100,
            click_count: 0,
            transition: AnimationTransition::None,
        };
        assert!(validate_and_sample(&[], valid, 1).is_err());
        assert!(
            validate_and_sample(
                &[AnimationSegment {
                    duration_ms: 0,
                    ..segment
                }],
                valid,
                1,
            )
            .is_err()
        );
        for invalid in [
            AnimationExportOptions {
                frame_rate: 0,
                ..valid
            },
            AnimationExportOptions {
                width_px: 0,
                ..valid
            },
            AnimationExportOptions {
                height_px: 0,
                ..valid
            },
            AnimationExportOptions {
                width_px: 32_768,
                ..valid
            },
            AnimationExportOptions {
                width_px: 5_000,
                height_px: 5_000,
                ..valid
            },
            AnimationExportOptions {
                format: AnimationFormat::MotionJpegAvi { quality: 0 },
                ..valid
            },
        ] {
            assert!(validate_and_sample(&[segment], invalid, 1).is_err());
        }
        assert!(
            validate_and_sample(
                &[AnimationSegment {
                    slide_index: 1,
                    ..segment
                }],
                valid,
                1,
            )
            .is_err()
        );
        assert!(
            validate_and_sample(
                &[AnimationSegment {
                    transition: AnimationTransition::FromSlide(1),
                    ..segment
                }],
                valid,
                1,
            )
            .is_err()
        );
        assert!(
            validate_and_sample(
                &[AnimationSegment {
                    duration_ms: u64::MAX,
                    ..segment
                }],
                valid,
                1,
            )
            .is_err()
        );
    }
}
