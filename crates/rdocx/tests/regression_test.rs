//! Regression tests for previously-fixed defects.
//!
//! Each test names the failure it locks down, so a reintroduction is obvious
//! from the test name alone rather than from a diff.

use std::alloc::{GlobalAlloc, Layout, System};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::{Duration, Instant};

use rdocx::{
    BarcodeField, BarcodeKind, BodyContentRef, BodyItemRef, BreakKind, CellItemRef, CellRef,
    ChartData, ChartKind, Document, FieldDateTime, FieldEvaluationContext, FieldOutcome,
    HyperlinkItemRef, HyperlinkRef, Length, ListLevel, MailMergeControl, MailMergeData,
    MailMergeFormattedText, MailMergeImage, MailMergeRecord, MailMergeValue, ParagraphItemRef,
    ParagraphRef, RasterFormat, RasterOptions, RasterOutput, RenderOptions, RevisionView,
    RunItemRef, RunPosition, RunRange, RunRef, StyleBuilder, TableRef, TcField, TocEntrySelection,
    TocField, TocRebuildReport, UnsupportedXmlRef,
};
use rdocx_oxml::CT_Document;
use rdocx_oxml::document::{BodyContent, CT_Body};
use rdocx_oxml::namespace::W_NS;
use rdocx_oxml::text::CT_R;

struct MeasuredAllocator;

#[repr(C)]
struct AllocationHeader {
    generation: usize,
}

static ACTIVE_ALLOCATION_GENERATION: AtomicUsize = AtomicUsize::new(0);
static NEXT_ALLOCATION_GENERATION: AtomicUsize = AtomicUsize::new(1);
static LIVE_ALLOCATION_BYTES: AtomicUsize = AtomicUsize::new(0);
static PEAK_ALLOCATION_BYTES: AtomicUsize = AtomicUsize::new(0);

fn measured_allocation_layout(layout: Layout) -> Option<(Layout, usize)> {
    Layout::new::<AllocationHeader>()
        .extend(layout)
        .ok()
        .map(|(combined, offset)| (combined.pad_to_align(), offset))
}

fn record_allocation(size: usize) {
    let live = LIVE_ALLOCATION_BYTES.fetch_add(size, Ordering::Relaxed) + size;
    let mut peak = PEAK_ALLOCATION_BYTES.load(Ordering::Relaxed);
    while live > peak {
        match PEAK_ALLOCATION_BYTES.compare_exchange_weak(
            peak,
            live,
            Ordering::Relaxed,
            Ordering::Relaxed,
        ) {
            Ok(_) => break,
            Err(observed) => peak = observed,
        }
    }
}

unsafe impl GlobalAlloc for MeasuredAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let Some((system_layout, offset)) = measured_allocation_layout(layout) else {
            return std::ptr::null_mut();
        };
        // SAFETY: `system_layout` is valid and is deallocated with the same layout below.
        let base = unsafe { System.alloc(system_layout) };
        if base.is_null() {
            return base;
        }
        let generation = ACTIVE_ALLOCATION_GENERATION.load(Ordering::Relaxed);
        // SAFETY: the combined layout reserves an aligned header at `base`.
        unsafe {
            base.cast::<AllocationHeader>()
                .write(AllocationHeader { generation })
        };
        if generation != 0 {
            record_allocation(layout.size());
        }
        // SAFETY: `offset` is the payload offset returned for the combined layout.
        unsafe { base.add(offset) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        let (system_layout, offset) = measured_allocation_layout(layout)
            .expect("an allocated layout must remain representable during deallocation");
        // SAFETY: `pointer` was returned at this exact offset by `alloc` above.
        let base = unsafe { pointer.sub(offset) };
        // SAFETY: `base` points to the initialized header for this allocation.
        let generation = unsafe { base.cast::<AllocationHeader>().read().generation };
        if generation != 0 && generation == ACTIVE_ALLOCATION_GENERATION.load(Ordering::Relaxed) {
            LIVE_ALLOCATION_BYTES.fetch_sub(layout.size(), Ordering::Relaxed);
        }
        // SAFETY: `base` and `system_layout` exactly match the allocation above.
        unsafe { System.dealloc(base, system_layout) };
    }
}

#[global_allocator]
static TEST_ALLOCATOR: MeasuredAllocator = MeasuredAllocator;

struct AllocationMeasurement {
    generation: usize,
    started: Instant,
}

impl AllocationMeasurement {
    fn start() -> Self {
        let generation = NEXT_ALLOCATION_GENERATION.fetch_add(1, Ordering::Relaxed);
        assert_ne!(generation, 0, "allocation measurement generation wrapped");
        LIVE_ALLOCATION_BYTES.store(0, Ordering::Relaxed);
        PEAK_ALLOCATION_BYTES.store(0, Ordering::Relaxed);
        ACTIVE_ALLOCATION_GENERATION
            .compare_exchange(0, generation, Ordering::AcqRel, Ordering::Relaxed)
            .expect("only one allocation measurement may run at a time");
        Self {
            generation,
            started: Instant::now(),
        }
    }

    fn finish(self) -> (Duration, usize) {
        let elapsed = self.started.elapsed();
        let peak = PEAK_ALLOCATION_BYTES.load(Ordering::Relaxed);
        ACTIVE_ALLOCATION_GENERATION
            .compare_exchange(self.generation, 0, Ordering::AcqRel, Ordering::Relaxed)
            .expect("allocation measurement generation changed unexpectedly");
        std::mem::forget(self);
        (elapsed, peak)
    }
}

impl Drop for AllocationMeasurement {
    fn drop(&mut self) {
        let _ = ACTIVE_ALLOCATION_GENERATION.compare_exchange(
            self.generation,
            0,
            Ordering::AcqRel,
            Ordering::Relaxed,
        );
    }
}

fn measure_allocations<T>(operation: impl FnOnce() -> T) -> (T, Duration, usize) {
    let measurement = AllocationMeasurement::start();
    let output = operation();
    let (elapsed, peak) = measurement.finish();
    (output, elapsed, peak)
}

#[test]
#[ignore = "release-mode large-document performance gate"]
fn a_thousand_page_document_paginates_and_renders_within_the_declared_limits() {
    const PAGE_COUNT: usize = 1_000;
    const MIB: usize = 1024 * 1024;
    const MAX_LAYOUT_PEAK_BYTES: usize = 64 * MIB;
    const MAX_PDF_PEAK_BYTES: usize = 16 * MIB;
    const MIN_LAYOUT_PAGES_PER_SECOND: f64 = 250.0;
    const MIN_PDF_PAGES_PER_SECOND: f64 = 1_000.0;

    let mut document = Document::new();
    document.add_paragraph("Performance page 1");
    for page in 2..=PAGE_COUNT {
        document
            .add_paragraph(&format!("Performance page {page}"))
            .page_break_before(true);
    }

    let (layout, layout_elapsed, layout_peak) = measure_allocations(|| {
        document
            .layout_deterministic()
            .expect("deterministic thousand-page layout")
    });
    assert_eq!(layout.layout.pages.len(), PAGE_COUNT);

    let (pdf, pdf_elapsed, pdf_peak) =
        measure_allocations(|| oxml_pdf::render_to_pdf(&layout.layout));
    assert!(pdf.starts_with(b"%PDF"));
    assert!(pdf.len() > PAGE_COUNT);

    let layout_rate = PAGE_COUNT as f64 / layout_elapsed.as_secs_f64();
    let pdf_rate = PAGE_COUNT as f64 / pdf_elapsed.as_secs_f64();
    eprintln!(
        "F-201 calibration: pages={PAGE_COUNT}, layout_seconds={:.3}, \
         layout_pages_per_second={layout_rate:.1}, layout_peak_mib={:.2}, \
         pdf_seconds={:.3}, pdf_pages_per_second={pdf_rate:.1}, pdf_peak_mib={:.2}",
        layout_elapsed.as_secs_f64(),
        layout_peak as f64 / MIB as f64,
        pdf_elapsed.as_secs_f64(),
        pdf_peak as f64 / MIB as f64,
    );
    assert!(
        layout_peak > 0,
        "layout allocation accounting must be active"
    );
    assert!(pdf_peak > 0, "PDF allocation accounting must be active");
    assert!(
        layout_peak <= MAX_LAYOUT_PEAK_BYTES,
        "layout peak {:.2} MiB exceeds {:.2} MiB",
        layout_peak as f64 / MIB as f64,
        MAX_LAYOUT_PEAK_BYTES as f64 / MIB as f64,
    );
    assert!(
        pdf_peak <= MAX_PDF_PEAK_BYTES,
        "PDF peak {:.2} MiB exceeds {:.2} MiB",
        pdf_peak as f64 / MIB as f64,
        MAX_PDF_PEAK_BYTES as f64 / MIB as f64,
    );
    assert!(
        layout_rate >= MIN_LAYOUT_PAGES_PER_SECOND,
        "layout rate {layout_rate:.1} pages/s is below {MIN_LAYOUT_PAGES_PER_SECOND:.1} pages/s",
    );
    assert!(
        pdf_rate >= MIN_PDF_PAGES_PER_SECOND,
        "PDF rate {pdf_rate:.1} pages/s is below {MIN_PDF_PAGES_PER_SECOND:.1} pages/s",
    );
}

// F-X075_BENCHMARK_MANIFEST_BEGIN
const EXPECTED_FX075_HARNESS_MANIFEST: &str = "e3fcfc1ac4332c54d3a4cf52ed6243c177b6d057";

fn fx075_git_output(
    workspace: &std::path::Path,
    arguments: &[&str],
    stdin: Option<&[u8]>,
) -> Vec<u8> {
    use std::io::Write as _;
    use std::process::Stdio;

    let mut command = std::process::Command::new("git");
    command
        .args(arguments)
        .current_dir(workspace)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    if stdin.is_some() {
        command.stdin(Stdio::piped());
    }
    let mut child = command
        .spawn()
        .expect("git must inspect the measured source");
    if let Some(stdin) = stdin {
        child
            .stdin
            .take()
            .expect("piped git stdin")
            .write_all(stdin)
            .expect("source manifest reaches git");
    }
    let output = child
        .wait_with_output()
        .expect("git must finish inspecting the measured source");
    assert!(
        output.status.success(),
        "git source inspection failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    output.stdout
}

fn fx075_git_hash(workspace: &std::path::Path, bytes: &[u8]) -> String {
    let output = fx075_git_output(workspace, &["hash-object", "--stdin"], Some(bytes));
    String::from_utf8(output)
        .expect("git object identity is UTF-8")
        .trim()
        .to_owned()
}

fn fx075_source_manifests(workspace: &std::path::Path) -> (String, String, String) {
    const BENCHMARK_PATH: &str = "crates/rdocx/tests/regression_test.rs";
    const BEGIN: &str = "// F-X075_BENCHMARK_MANIFEST_BEGIN";
    const END: &str = "// F-X075_BENCHMARK_MANIFEST_END";

    let untracked = fx075_git_output(
        workspace,
        &[
            "ls-files",
            "--others",
            "--exclude-standard",
            "--",
            "Cargo.toml",
            "Cargo.lock",
            "crates",
        ],
        None,
    );
    assert!(
        untracked.is_empty(),
        "measured crate graph contains untracked source: {}",
        String::from_utf8_lossy(&untracked)
    );

    let tracked = fx075_git_output(
        workspace,
        &["ls-files", "-z", "--", "Cargo.toml", "Cargo.lock", "crates"],
        None,
    );
    let mut paths = tracked
        .split(|byte| *byte == 0)
        .filter(|path| !path.is_empty())
        .map(|path| String::from_utf8(path.to_vec()).expect("tracked path is UTF-8"))
        .filter(|path| path != BENCHMARK_PATH)
        .collect::<Vec<_>>();
    paths.sort_unstable();
    let path_input = format!("{}\n", paths.join("\n"));
    let object_ids = fx075_git_output(
        workspace,
        &["hash-object", "--stdin-paths"],
        Some(path_input.as_bytes()),
    );
    let object_ids = String::from_utf8(object_ids).expect("source identities are UTF-8");
    let object_ids = object_ids.lines().collect::<Vec<_>>();
    assert_eq!(object_ids.len(), paths.len());
    let production_manifest = paths
        .iter()
        .zip(object_ids)
        .map(|(path, object_id)| format!("{object_id} {path}\n"))
        .collect::<String>();

    let benchmark_source = std::fs::read_to_string(workspace.join(BENCHMARK_PATH))
        .expect("benchmark source must be readable");
    let begin = benchmark_source
        .find(BEGIN)
        .expect("benchmark manifest begin marker");
    let end = benchmark_source
        .rfind(END)
        .expect("benchmark manifest end marker")
        + END.len();
    let surrounding = format!(
        "{}<F-X075 benchmark harness>\n{}",
        &benchmark_source[..begin],
        &benchmark_source[end..]
    );
    let mut harness = benchmark_source[begin..end].to_owned();
    let self_pin_prefix = ["const EXPECTED_FX075_", "HARNESS_MANIFEST: &str = "].concat();
    let candidates = harness.match_indices(&self_pin_prefix).collect::<Vec<_>>();
    assert_eq!(
        candidates.len(),
        1,
        "benchmark harness must contain exactly one self-pin declaration"
    );
    let literal_start = candidates[0].0 + self_pin_prefix.len();
    let expected_literal = format!("\"{EXPECTED_FX075_HARNESS_MANIFEST}\"");
    let literal_end = literal_start + expected_literal.len();
    assert_eq!(
        harness.get(literal_start..literal_end),
        Some(expected_literal.as_str()),
        "benchmark self-pin must contain its exact expected literal"
    );
    assert_eq!(
        harness.as_bytes().get(literal_end),
        Some(&b';'),
        "benchmark self-pin must use the exact declaration shape"
    );
    harness.replace_range(literal_start..literal_end, "\"<self>\"");

    (
        fx075_git_hash(workspace, production_manifest.as_bytes()),
        fx075_git_hash(workspace, surrounding.as_bytes()),
        fx075_git_hash(workspace, harness.as_bytes()),
    )
}

#[test]
#[ignore = "F-X075 interleaved release-mode performance evidence"]
fn issue_67_page_spanning_prose_release_measurement() {
    const V0_11_1_CHECKOUT_HEAD: &str = "5a850ce9ae6c31f8365594ed2970193266f8b2a6";
    const REGRESSION_CHECKOUT_HEAD: &str = "0582da0a38886f5ceeb65ab9afcd0797f6fa14b0";
    const CURRENT_PRODUCTION_MANIFEST: &str = "5744c802fc8096683faf175c29b9c6c359617bb1";
    const V0_11_1_PRODUCTION_MANIFEST: &str = "bd56901eb9d692e6eb5c2e6f8b33d26abe14f910";
    const REGRESSION_PRODUCTION_MANIFEST: &str = "32f3644b56f848d5d9a231c28dc6a072185e9bb0";
    const CURRENT_SURROUNDING_TEST_MANIFEST: &str = "06b08d8729dc02019f96cf9f911a1f3d5cdf3df3";
    const V0_11_1_SURROUNDING_TEST_MANIFEST: &str = "c72fe72a92befb7d4a79573835c53fd83fe5201c";
    const REGRESSION_SURROUNDING_TEST_MANIFEST: &str = "06b08d8729dc02019f96cf9f911a1f3d5cdf3df3";

    let checkout = std::env::var("RDOCX_FX075_CHECKOUT")
        .expect("RDOCX_FX075_CHECKOUT must identify current, v0.11.1, or 0582da0");
    let (expected_head, expected_production, expected_surrounding) = match checkout.as_str() {
        "current" => (
            None,
            CURRENT_PRODUCTION_MANIFEST,
            CURRENT_SURROUNDING_TEST_MANIFEST,
        ),
        "v0.11.1" => (
            Some(V0_11_1_CHECKOUT_HEAD),
            V0_11_1_PRODUCTION_MANIFEST,
            V0_11_1_SURROUNDING_TEST_MANIFEST,
        ),
        "0582da0" => (
            Some(REGRESSION_CHECKOUT_HEAD),
            REGRESSION_PRODUCTION_MANIFEST,
            REGRESSION_SURROUNDING_TEST_MANIFEST,
        ),
        _ => panic!("unsupported F-X075 checkout identity {checkout:?}"),
    };
    let workspace = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let actual_head = fx075_git_output(
        &workspace,
        &["rev-parse", "--verify", "HEAD^{commit}"],
        None,
    );
    let actual_head = String::from_utf8(actual_head).expect("checkout identity is UTF-8");
    let actual_head = actual_head.trim();
    if let Some(expected_head) = expected_head {
        assert_eq!(
            actual_head, expected_head,
            "F-X075 {checkout} measurement must use its pinned reference commit"
        );
    }
    let (production_manifest, surrounding_manifest, harness_manifest) =
        fx075_source_manifests(&workspace);
    eprintln!(
        "F-X075 source identity: checkout={checkout}, head={actual_head}, production_manifest={production_manifest}, surrounding_test_manifest={surrounding_manifest}, harness_manifest={harness_manifest}"
    );
    assert_eq!(
        production_manifest, expected_production,
        "F-X075 {checkout} production source differs from the reviewed manifest"
    );
    assert_eq!(
        surrounding_manifest, expected_surrounding,
        "F-X075 {checkout} non-harness test source differs from the reviewed manifest"
    );
    assert_eq!(
        harness_manifest, EXPECTED_FX075_HARNESS_MANIFEST,
        "F-X075 benchmark harness differs from the reviewed injection"
    );

    let paragraph_count = std::env::var("RDOCX_FX075_PARAGRAPHS")
        .ok()
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(175);
    let mode = std::env::var("RDOCX_FX075_MODE").unwrap_or_else(|_| "native".to_owned());
    assert!(matches!(paragraph_count, 175 | 700));
    assert!(matches!(mode.as_str(), "native" | "bundled-fallback"));

    let paragraph_text = |index: usize, suffix: &str| {
        format!(
            "Paragraph {index}: the quick brown fox jumps over the lazy dog, pack my box \
             with five dozen liquor jugs, and a mixed sentence that keeps going. \
             Sphinx of black quartz, judge my vow across line breaks and pages. \
             Waltz, bad nymph, for quick jigs vex. Glib jocks quiz nymph to vex dwarf. \
             Bright vixens jump, dozy fowl quack.{suffix}"
        )
    };
    let mut document = Document::new();
    for index in 0..paragraph_count {
        document.add_paragraph(&paragraph_text(index, ""));
    }
    let layout = |document: &Document| match mode.as_str() {
        "native" => document.layout().map(|result| result.layout.pages.len()),
        "bundled-fallback" => document
            .layout_with_fonts_aliases_and_bundled_fallback(&[], &[])
            .map(|result| result.layout.pages.len()),
        _ => unreachable!("mode validated above"),
    };
    let pages = layout(&document).expect("prime deterministic Issue 67 layout");
    if paragraph_count == 175 {
        assert_eq!(pages, 16);
    }

    let edit_index = paragraph_count / 2;
    let mut samples = Vec::with_capacity(10);
    for revision in 1..=10 {
        document
            .paragraph_mut(edit_index)
            .expect("middle paragraph")
            .run_mut(0)
            .expect("middle paragraph text run")
            .set_text(&paragraph_text(edit_index, &"x".repeat(revision)));
        let started = Instant::now();
        assert_eq!(
            layout(&document).expect("warm deterministic Issue 67 layout"),
            pages
        );
        samples.push(started.elapsed());
    }
    samples.sort_unstable();
    let median = samples[samples.len() / 2];
    let milliseconds = samples
        .iter()
        .map(|sample| format!("{:.3}", sample.as_secs_f64() * 1_000.0))
        .collect::<Vec<_>>();
    eprintln!(
        "F-X075 release measurement: checkout={checkout}, head={actual_head}, paragraphs={paragraph_count}, mode={mode}, pages={pages}, median_ms={:.3}, sorted_ms=[{}]",
        median.as_secs_f64() * 1_000.0,
        milliseconds.join(", ")
    );
}
// F-X075_BENCHMARK_MANIFEST_END

#[test]
fn editing_one_paragraph_of_a_thousand_page_document_rebuilds_at_most_two_pages() {
    fn thousand_page_document() -> Document {
        let mut document = Document::new();
        for page in 0..1_000 {
            document
                .add_paragraph(&format!("Incremental page {}", page + 1))
                .page_break_before(page > 0);
        }
        document
    }

    let mut warm_document = thousand_page_document();
    let initial = warm_document
        .layout_with_fonts_and_bundled_fallback(&[])
        .expect("initial deterministic thousand-page layout");
    assert_eq!(initial.layout.pages.len(), 1_000);

    warm_document
        .paragraph_mut(499)
        .expect("paragraph 500")
        .run_mut(0)
        .expect("paragraph text run")
        .set_text("Incremental page 500 changed");
    let warm = warm_document
        .layout_with_fonts_and_bundled_fallback(&[])
        .expect("warm deterministic thousand-page layout");

    let mut fresh_document = thousand_page_document();
    fresh_document
        .paragraph_mut(499)
        .expect("paragraph 500")
        .run_mut(0)
        .expect("paragraph text run")
        .set_text("Incremental page 500 changed");
    let fresh = fresh_document
        .layout_with_fonts_and_bundled_fallback(&[])
        .expect("fresh deterministic thousand-page layout");

    assert_eq!(warm.layout.pages.len(), 1_000);
    assert_eq!(format!("{warm:?}"), format!("{fresh:?}"));
    let retained_pages = warm
        .layout
        .pages
        .iter()
        .zip(&initial.layout.pages)
        .filter(|(current, retained)| Arc::ptr_eq(current, retained))
        .count();
    assert!(
        retained_pages >= 998,
        "expected at least 998 retained page frames, got {retained_pages}"
    );
}

#[test]
fn issue_53_related_stories_keep_the_700_paragraph_facade_workload_bounded() {
    fn issue_53_document() -> Document {
        let mut seed = Document::new();
        let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(
            seed.to_bytes().expect("serialize seed document"),
        ))
        .expect("open seed package");
        let (header_id, footer_id) = {
            let relationships = package.get_or_create_part_rels("/word/document.xml");
            let header_id =
                relationships.add(oxml_opc::relationship::rel_types::HEADER, "header1.xml");
            let footer_id =
                relationships.add(oxml_opc::relationship::rel_types::FOOTER, "footer1.xml");
            relationships.add(
                oxml_opc::relationship::rel_types::FOOTNOTES,
                "footnotes.xml",
            );
            (header_id, footer_id)
        };
        package.set_part(
            "/word/header1.xml",
            br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Issue 53 header</w:t></w:r></w:p></w:hdr>"#.to_vec(),
        );
        package.set_part(
            "/word/footer1.xml",
            br#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>"#.to_vec(),
        );
        package.set_part(
            "/word/footnotes.xml",
            br#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>Issue 53 footnote</w:t></w:r></w:p></w:footnote></w:footnotes>"#.to_vec(),
        );
        package.content_types.add_override(
            "/word/header1.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        );
        package.content_types.add_override(
            "/word/footer1.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        );
        package.content_types.add_override(
            "/word/footnotes.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        );

        let mut body = String::new();
        for index in 0..700 {
            let note = if index == 20 {
                r#"<w:footnoteReference w:id="1"/>"#
            } else {
                ""
            };
            body.push_str(&format!(
                "<w:p><w:r><w:t>Issue 53 paragraph {index:03}</w:t>{note}</w:r></w:p>"
            ));
        }
        package.set_part(
            "/word/document.xml",
            format!(
                r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}<w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/><w:footerReference w:type="default" r:id="{footer_id}"/></w:sectPr></w:body></w:document>"#
            )
            .into_bytes(),
        );
        let mut bytes = std::io::Cursor::new(Vec::new());
        package
            .write_to(&mut bytes)
            .expect("serialize Issue 53 package");
        Document::from_bytes(bytes.get_ref()).expect("open Issue 53 document")
    }

    let mut warm_document = issue_53_document();
    let initial = warm_document
        .layout_with_fonts_and_bundled_fallback(&[])
        .expect("initial Issue 53 layout");
    assert!(initial.layout.pages.len() > 2);
    warm_document
        .paragraph_mut(349)
        .expect("paragraph 350")
        .run_mut(0)
        .expect("paragraph text run")
        .set_text("Issue 53 paragraph 350 changed");
    let warm = warm_document
        .layout_with_fonts_and_bundled_fallback(&[])
        .expect("warm Issue 53 layout");

    let mut fresh_document = issue_53_document();
    fresh_document
        .paragraph_mut(349)
        .expect("fresh paragraph 350")
        .run_mut(0)
        .expect("fresh paragraph text run")
        .set_text("Issue 53 paragraph 350 changed");
    let fresh = fresh_document
        .layout_with_fonts_and_bundled_fallback(&[])
        .expect("fresh Issue 53 layout");

    assert_eq!(format!("{warm:?}"), format!("{fresh:?}"));
    let retained_pages = warm
        .layout
        .pages
        .iter()
        .zip(&initial.layout.pages)
        .filter(|(current, retained)| Arc::ptr_eq(current, retained))
        .count();
    assert!(
        retained_pages >= warm.layout.pages.len().saturating_sub(2),
        "expected bounded Issue 53 page work, retained {retained_pages} of {} pages",
        warm.layout.pages.len()
    );
}

#[test]
fn unsupported_html_css_is_diagnosed_without_dropping_supported_siblings() {
    let parsed = Document::from_html(
        "<!doctype html><html><head><link rel='STYLESHEET' href='external.css'><script>head()</script><style>@media print { p { color: red } } p:hover { color: blue } p { color: rgb(1, 2, 3); made-up: yes }</style></head><body><p>before<blink style='unknown-value: yes'>kept</blink>after<a href='https://example.invalid'>link</a><img alt='picture'><script>ignored()</script><input value='field'><iframe>frame</iframe></p><form>form text</form><p><b><i>repaired</b></p></body></html>",
    )
    .expect("recoverable unsupported HTML");
    assert_eq!(
        parsed.document.paragraph(0).unwrap().text(),
        "beforekeptafterlinkpicturefield"
    );
    assert_eq!(parsed.document.paragraph(1).unwrap().text(), "form text");
    assert_eq!(parsed.document.paragraph(2).unwrap().text(), "repaired");
    assert!(parsed.diagnostics.iter().any(|diagnostic| {
        diagnostic.property.as_deref() == Some("unknown-value")
            && diagnostic.location == "html/body/p[1]/blink[1]"
    }));
    assert!(parsed.diagnostics.iter().any(|diagnostic| {
        diagnostic.location == "html/body/p[1]/img[1]" && diagnostic.message.contains("image")
    }));
    for expected in [
        "CSS at-rule",
        "CSS selector",
        "CSS color",
        "CSS property",
        "external HTML stylesheet",
        "link target",
        "script",
        "form semantics",
        "form control",
        "iframe",
        "parser repair",
    ] {
        assert!(
            parsed
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains(expected)),
            "missing diagnostic containing {expected}"
        );
    }
}

#[test]
fn html_import_projects_nested_lists_and_spanned_tables() {
    let parsed = Document::from_html(
        "<ol start='3'><li>one<ul><li>nested</li></ul></li></ol><ol><li>restart</li></ol><table><caption>Caption</caption><tr><td rowspan='2'>a</td><td colspan='2'>b</td></tr><tr><td>c<p>d</p><ul><li>f</li></ul></td><td>e</td></tr></table>",
    )
    .expect("supported lists and table spans");
    let first = parsed.document.paragraph(0).unwrap().numbering().unwrap();
    let nested = parsed.document.paragraph(1).unwrap().numbering().unwrap();
    let restarted = parsed.document.paragraph(2).unwrap().numbering().unwrap();
    assert_eq!(first.1, 0);
    assert_eq!(nested, (first.0, 1));
    assert_ne!(restarted.0, first.0);
    assert_eq!(parsed.document.numbering_is_bullet(first.0), Some(false));
    assert_eq!(parsed.document.numbering_is_bullet(nested.0), Some(false));
    assert_eq!(parsed.document.paragraph(3).unwrap().text(), "Caption");
    assert!(
        parsed
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message.contains("table caption"))
    );
    assert_eq!(parsed.document.table_count(), 1);
    let table = parsed.document.table(0).unwrap();
    assert_eq!(table.cell(0, 0).unwrap().text(), "a");
    assert_eq!(table.cell(0, 1).unwrap().text(), "b");
    assert_eq!(table.cell(0, 0).unwrap().grid_span(), None);
    assert!(matches!(
        table.cell(0, 0).unwrap().v_merge(),
        Some(rdocx_oxml::table::VMerge::Restart)
    ));
    assert_eq!(table.cell(0, 1).unwrap().grid_span(), Some(2));
    assert!(matches!(
        table.cell(1, 0).unwrap().v_merge(),
        Some(rdocx_oxml::table::VMerge::Continue)
    ));
    assert_eq!(table.cell(1, 1).unwrap().paragraph_count(), 3);
    assert_eq!(table.cell(1, 1).unwrap().text(), "c\nd\nf");
    assert_eq!(
        table
            .cell(1, 1)
            .unwrap()
            .paragraph(2)
            .unwrap()
            .numbering()
            .unwrap()
            .1,
        0
    );
    let mut imported = parsed.document;
    let bytes = imported.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let numbering = std::str::from_utf8(package.get_part("/word/numbering.xml").unwrap()).unwrap();
    assert!(numbering.contains(r#"<w:start w:val="3"/>"#));

    let mut nine = String::new();
    let mut tags = Vec::new();
    for level in 0..9 {
        let tag = if level % 2 == 0 { "ul" } else { "ol" };
        tags.push(tag);
        nine.push_str(&format!("<{tag}><li>"));
        nine.push_str(&level.to_string());
    }
    for tag in tags.into_iter().rev() {
        nine.push_str(&format!("</li></{tag}>"));
    }
    let nine_html = nine;
    let nine = Document::from_html(&nine_html).expect("nine list levels");
    for level in 0..9 {
        assert_eq!(
            nine.document
                .paragraph(level)
                .unwrap()
                .numbering()
                .unwrap()
                .1,
            level as u32
        );
    }
    let ten = format!("<ul><li>outer{nine_html}</li></ul>");
    assert!(Document::from_html(&ten).is_err());

    let nested_table =
        Document::from_html("<ul><li>item<table><tr><td>nested table</td></tr></table></li></ul>")
            .expect("unsupported nested table is recoverable");
    assert!(nested_table.diagnostics.iter().any(|diagnostic| {
        diagnostic.location.contains("table")
            && diagnostic.message.contains("dropped nested HTML table")
    }));
}

fn compatibility_page_elements(
    elements: &[oxml_layout::PositionedElement],
) -> Vec<&oxml_layout::PositionedElement> {
    fn collect<'a>(
        elements: &'a [oxml_layout::PositionedElement],
        output: &mut Vec<&'a oxml_layout::PositionedElement>,
    ) {
        for element in elements {
            match element {
                oxml_layout::PositionedElement::MarkedContent { children, .. } => {
                    collect(children, output);
                }
                oxml_layout::PositionedElement::Group(group) => {
                    collect(&group.children, output);
                }
                other => output.push(other),
            }
        }
    }

    let mut output = Vec::new();
    collect(elements, &mut output);
    output
}

fn clear_compatibility_sources(elements: &mut [oxml_layout::PositionedElement]) {
    for element in elements {
        match element {
            oxml_layout::PositionedElement::Text(run) => run.source = None,
            oxml_layout::PositionedElement::MarkedContent { children, .. } => {
                clear_compatibility_sources(children);
            }
            oxml_layout::PositionedElement::Group(group) => {
                clear_compatibility_sources(&mut group.children);
            }
            _ => {}
        }
    }
}

fn remove_compatibility_empty_text(elements: &mut Vec<oxml_layout::PositionedElement>) {
    for element in elements.iter_mut() {
        match element {
            oxml_layout::PositionedElement::MarkedContent { children, .. } => {
                remove_compatibility_empty_text(children);
            }
            oxml_layout::PositionedElement::Group(group) => {
                remove_compatibility_empty_text(&mut group.children);
            }
            _ => {}
        }
    }
    elements.retain(
        |element| !matches!(element, oxml_layout::PositionedElement::Text(run) if run.text.is_empty()),
    );
}

fn document_with_settings(settings_xml: &[u8], target: &str) -> Document {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let part_name = oxml_opc::OpcPackage::resolve_rel_target("/word/document.xml", target);
    package.set_part(&part_name, settings_xml.to_vec());
    package.content_types.add_override(
        &part_name,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    );
    package
        .get_or_create_part_rels("/word/document.xml")
        .add(oxml_opc::relationship::rel_types::SETTINGS, target);
    let mut output = std::io::Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();
    Document::from_bytes(output.get_ref()).unwrap()
}

#[test]
fn each_document_protection_mode_is_reported_with_its_recorded_hash() {
    for (mode, expected) in [
        ("readOnly", rdocx::ProtectionMode::ReadOnly),
        ("comments", rdocx::ProtectionMode::Comments),
        ("trackedChanges", rdocx::ProtectionMode::TrackedChanges),
        ("forms", rdocx::ProtectionMode::Forms),
    ] {
        let settings = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:documentProtection q:edit="{mode}" q:enforcement="1" q:cryptProviderType="rsaAES" q:cryptAlgorithmClass="hash" q:cryptAlgorithmType="typeAny" q:cryptAlgorithmSid="14" q:cryptSpinCount="100000" q:hash="HASH-{mode}" q:salt="SALT-{mode}"/></q:settings>"#
        );
        let mut document = document_with_settings(settings.as_bytes(), "settings.xml");
        let protection = document.document_protection().unwrap();
        assert_eq!(protection.mode, expected);
        assert_eq!(protection.enforcement, Some(true));
        assert_eq!(protection.algorithm_sid, Some(14));
        assert_eq!(protection.spin_count, Some(100_000));
        assert_eq!(
            protection.hash.as_deref(),
            Some(format!("HASH-{mode}").as_str())
        );
        assert_eq!(
            protection.salt.as_deref(),
            Some(format!("SALT-{mode}").as_str())
        );

        let saved = document.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        assert_eq!(
            package.get_part("/word/settings.xml").unwrap(),
            settings.as_bytes()
        );
    }
}

#[test]
fn malformed_document_protection_remains_opaque_and_unreported() {
    for attributes in [
        r#"q:edit="unsupported" q:cryptSpinCount="100000""#,
        r#"q:edit="forms" q:cryptSpinCount="many""#,
        r#"q:edit="forms" q:cryptAlgorithmSid="SHA-512""#,
        r#"q:edit="forms" q:cryptAlgorithmClass="future""#,
    ] {
        let settings = format!(
            r#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:p="urn:producer"><p:before/><q:documentProtection {attributes}/><p:after/></q:settings>"#
        );
        let mut document = document_with_settings(settings.as_bytes(), "settings.xml");
        assert!(document.document_protection().is_none());

        let saved = document.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        assert_eq!(
            package.get_part("/word/settings.xml").unwrap(),
            settings.as_bytes()
        );
    }
}

const CUSTOM_XML_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const CUSTOM_XML_PROPS_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
const WORD_REVISION_ORACLE: &str = "Microsoft Word 16.104 build 16.104.25121423";
const WORD_FIELD_ORACLE: &str = "Microsoft Word 16.104 build 16.104.25121423";
const WORD_DENSE_FORM_ORACLE: &str = "Microsoft Word 16.104 build 16.104.25121423";
const WORD_COMPARISON_ORACLE: &str = "Microsoft Word 16.104 build 16.104.25121423";
const WORD_COMPARISON_ENVIRONMENT: &str = "locale=en-US; normalization=revision-records-v1";
const WORD_FIELD_ORACLE_ENVIRONMENT: &str =
    "locale=en-US; calendar=Gregorian; decimal=.; grouping=,; timezone=UTC";
const WORD_FIELD_ORACLE_INPUT: &str = "F-161-readable-field-matrix-v1";
const EXTENDED_WORD_FIELD_ORACLE_INPUT: &str = "F-231-readable-field-matrix-v1";
const DYNAMIC_TOC_WORD_ORACLE_INPUT: &str = "F-232-dynamic-toc-rebuild-v1";

fn document_with_field_parts(
    document_xml: &str,
    settings_xml: Option<&str>,
    custom_properties_xml: Option<&str>,
) -> Document {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    package.set_part("/word/document.xml", document_xml.as_bytes().to_vec());
    if let Some(settings_xml) = settings_xml {
        package.set_part("/word/settings.xml", settings_xml.as_bytes().to_vec());
        package.content_types.add_override(
            "/word/settings.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        );
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(oxml_opc::relationship::rel_types::SETTINGS, "settings.xml");
    }
    if let Some(custom_properties_xml) = custom_properties_xml {
        package.set_part(
            "/metadata/producer-properties.xml",
            custom_properties_xml.as_bytes().to_vec(),
        );
        package.content_types.add_override(
            "/metadata/producer-properties.xml",
            oxml_opc::content_types::CUSTOM_PROPERTIES,
        );
        package.package_rels.add(
            oxml_opc::relationship::rel_types::CUSTOM_PROPERTIES,
            "metadata/producer-properties.xml",
        );
    }
    let mut output = std::io::Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();
    Document::from_bytes(output.get_ref()).unwrap()
}

fn toc_entry_signatures(xml: &str) -> Vec<(String, String, Option<String>, Option<String>)> {
    let document = CT_Document::from_xml(xml.as_bytes()).unwrap();
    document
        .body
        .content
        .iter()
        .filter_map(|content| {
            let BodyContent::Paragraph(paragraph) = content else {
                return None;
            };
            let style = paragraph
                .properties
                .as_ref()
                .and_then(|properties| properties.style_id.as_deref())?;
            style.starts_with("TOC").then(|| {
                let mut display = paragraph.text();
                let mut page_target = None;
                for run in paragraph.runs() {
                    for content in &run.content {
                        if let rdocx_oxml::text::RunContent::Field(field) = content
                            && field.instruction.name == "PAGEREF"
                        {
                            display.push_str(&field.cached_result);
                            page_target =
                                field.instruction.arguments.first().and_then(|argument| {
                                    let rdocx_oxml::text::FieldArgument::Text(value) = argument
                                    else {
                                        return None;
                                    };
                                    Some(value.clone())
                                });
                        }
                    }
                }
                (
                    display,
                    style.to_owned(),
                    paragraph
                        .hyperlinks
                        .first()
                        .and_then(|hyperlink| hyperlink.anchor.clone()),
                    page_target,
                )
            })
        })
        .collect()
}

#[test]
fn every_supported_field_matches_the_pinned_word_result() {
    assert_eq!(
        WORD_FIELD_ORACLE,
        "Microsoft Word 16.104 build 16.104.25121423"
    );
    assert_eq!(
        WORD_FIELD_ORACLE_ENVIRONMENT,
        "locale=en-US; calendar=Gregorian; decimal=.; grouping=,; timezone=UTC"
    );
    assert_eq!(WORD_FIELD_ORACLE_INPUT, "F-161-readable-field-matrix-v1");

    let body = r#"
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="7" w:name="destination"/><w:r><w:t>Introduction</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>
        <w:p><w:fldSimple w:instr="IF &quot;10&quot; &gt;= &quot;2&quot; &quot;yes&quot; &quot;no&quot;"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="REF destination"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="PAGEREF destination"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="SEQ Figure"><w:r><w:t>0</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="DOCPROPERTY Title"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="DOCVARIABLE Region"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="STYLEREF &quot;Heading 1&quot;"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="INCLUDETEXT &quot;chapter.docx&quot;"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="DATE \@ &quot;MMMM d, yyyy&quot;"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="TIME \@ &quot;h:mm AM/PM&quot;"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="FILENAME"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="AUTHOR"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD Name \* Upper"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="NUMPAGES"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>
    "#;
    let settings = r#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:docVars><q:docVar q:name="Region" q:val="West"/></q:docVars></q:settings>"#;
    let mut document = document_with_field_parts(&wrap_word_body(body), Some(settings), None);
    document.set_title("Oracle title");
    document.set_author("Ada Lovelace");
    let context = FieldEvaluationContext {
        now: Some(FieldDateTime {
            year: 2025,
            month: 12,
            day: 14,
            hour: 21,
            minute: 7,
            second: 5,
        }),
        file_name: Some("report.docx".to_owned()),
        file_path: Some("/templates/report.docx".to_owned()),
        merge_fields: BTreeMap::from([("Name".to_owned(), "Grace".to_owned())]),
        included_text: BTreeMap::from([("chapter.docx".to_owned(), "Included chapter".to_owned())]),
        merge_record_number: None,
        merge_sequence_number: None,
    };
    let actual = document
        .evaluate_fields(&context)
        .unwrap()
        .into_iter()
        .map(|evaluation| evaluation.outcome)
        .collect::<Vec<_>>();
    let expected = vec![
        FieldOutcome::Resolved("yes".to_owned()),
        FieldOutcome::Resolved("Introduction".to_owned()),
        FieldOutcome::DeferredPagination,
        FieldOutcome::Resolved("1".to_owned()),
        FieldOutcome::Resolved("Oracle title".to_owned()),
        FieldOutcome::Resolved("West".to_owned()),
        FieldOutcome::Resolved("Introduction".to_owned()),
        FieldOutcome::Resolved("Included chapter".to_owned()),
        FieldOutcome::Resolved("December 14, 2025".to_owned()),
        FieldOutcome::Resolved("9:07 PM".to_owned()),
        FieldOutcome::Resolved("report.docx".to_owned()),
        FieldOutcome::Resolved("Ada Lovelace".to_owned()),
        FieldOutcome::Resolved("GRACE".to_owned()),
        FieldOutcome::DeferredPagination,
        FieldOutcome::DeferredPagination,
    ];
    assert_eq!(actual, expected, "{WORD_FIELD_ORACLE_INPUT}");
}

#[test]
fn extended_field_families_match_the_pinned_word_result() {
    assert_eq!(
        WORD_FIELD_ORACLE,
        "Microsoft Word 16.104 build 16.104.25121423"
    );
    assert_eq!(
        WORD_FIELD_ORACLE_ENVIRONMENT,
        "locale=en-US; calendar=Gregorian; decimal=.; grouping=,; timezone=UTC"
    );
    assert_eq!(
        EXTENDED_WORD_FIELD_ORACLE_INPUT,
        "F-231-readable-field-matrix-v1"
    );

    let body = r#"
        <w:p><w:fldSimple w:instr="= 50% + 0.5"><w:r><w:t>stored formula</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="TOC \o &quot;1-3&quot; \h"><w:r><w:t>stored toc</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="TC &quot;Alpha&quot; \l 2"><w:r><w:t>stored tc</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="NEXT"><w:r><w:t>stored next</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="NEXTIF &quot;A&quot; = &quot;A&quot;"><w:r><w:t>stored nextif</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="SKIPIF &quot;A&quot; = &quot;A&quot;"><w:r><w:t>stored skipif</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEREC"><w:r><w:t>stored record</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGESEQ"><w:r><w:t>stored sequence</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="DISPLAYBARCODE 0123456789012 EAN13 \x \p STD"><w:r><w:t>stored display barcode</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEBARCODE Code QR \q 3"><w:r><w:t>stored merge barcode</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="= 0.1 + 0.2"><w:r><w:t>stored decimal formula</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="TOC \o &quot;1-3&quot; \s chapter \d &quot;:&quot;"><w:r><w:t>stored sequenced toc</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="TOC \p &quot;ab&quot;"><w:r><w:t>stored invalid toc</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="TC &quot;Entry&quot; unexpected"><w:r><w:t>stored invalid tc</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="DISPLAYBARCODE value QR \s 9"><w:r><w:t>stored invalid barcode</w:t></w:r></w:fldSimple></w:p>
    "#;
    let document = document_with_field_parts(&wrap_word_body(body), None, None);
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([("Code".to_owned(), "https://example.test/42".to_owned())]),
        merge_record_number: Some(10),
        merge_sequence_number: Some(7),
        ..Default::default()
    };
    assert_eq!(
        document.evaluate_fields(&context).unwrap()[0].outcome,
        FieldOutcome::Resolved("1".to_owned()),
    );

    let actual = document
        .evaluate_fields(&context)
        .unwrap()
        .into_iter()
        .map(|evaluation| evaluation.outcome)
        .collect::<Vec<_>>();
    assert_eq!(
        actual,
        [
            FieldOutcome::Resolved("1".to_owned()),
            FieldOutcome::TableOfContents(TocField {
                heading_levels: Some((1, 3)),
                custom_styles: Vec::new(),
                entries: TocEntrySelection::None,
                sequence_identifier: None,
                bookmark: None,
                hyperlink: true,
                use_outline_levels: false,
                omit_page_number_levels: None,
                page_number_separator: None,
                entry_page_separator: None,
            }),
            FieldOutcome::TableOfContentsEntry(TcField {
                entry: "Alpha".to_owned(),
                level: 2,
                table_identifier: None,
                omit_page_number: false,
            }),
            FieldOutcome::MailMergeControl(MailMergeControl::NextRecord { record_number: 11 }),
            FieldOutcome::MailMergeControl(MailMergeControl::NextRecordIf {
                condition: true,
                record_number: 12,
            }),
            FieldOutcome::MailMergeControl(MailMergeControl::SkipRecordIf {
                condition: true,
                record_number: 12,
            }),
            FieldOutcome::MailMergeControl(MailMergeControl::RecordNumber(12)),
            FieldOutcome::MailMergeControl(MailMergeControl::SequenceNumber(7)),
            FieldOutcome::Barcode(BarcodeField {
                value: "0123456789012".to_owned(),
                kind: BarcodeKind::Ean13,
                height: None,
                scale: None,
                error_correction: None,
                point_of_sale_style: Some(rdocx::BarcodePointOfSaleStyle::Standard),
                case_style: None,
                fix_check_digit: true,
                rotation: None,
                foreground_color: None,
                background_color: None,
                display_text: false,
                add_start_stop: false,
            }),
            FieldOutcome::Barcode(BarcodeField {
                value: "https://example.test/42".to_owned(),
                kind: BarcodeKind::Qr,
                height: None,
                scale: None,
                error_correction: Some(3),
                point_of_sale_style: None,
                case_style: None,
                fix_check_digit: false,
                rotation: None,
                foreground_color: None,
                background_color: None,
                display_text: false,
                add_start_stop: false,
            }),
            FieldOutcome::Resolved("0.3".to_owned()),
            FieldOutcome::TableOfContents(TocField {
                heading_levels: Some((1, 3)),
                custom_styles: Vec::new(),
                entries: TocEntrySelection::None,
                sequence_identifier: Some("chapter".to_owned()),
                bookmark: None,
                hyperlink: false,
                use_outline_levels: false,
                omit_page_number_levels: None,
                page_number_separator: None,
                entry_page_separator: Some(":".to_owned()),
            }),
            FieldOutcome::KeepStored {
                diagnostic: "TOC page-number separator must contain exactly one character, stored display retained".to_owned(),
            },
            FieldOutcome::KeepStored {
                diagnostic: "field TC requires 1 positional operands, stored display retained".to_owned(),
            },
            FieldOutcome::KeepStored {
                diagnostic: "DISPLAYBARCODE scale must be from 10 through 1000, stored display retained".to_owned(),
            },
        ],
        "{EXTENDED_WORD_FIELD_ORACLE_INPUT}"
    );
}

#[test]
fn extended_field_updates_preserve_instruction_and_result_scaffolding() {
    let body = concat!(
        r#"<w:p xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:fldSimple q:instr="DISPLAYBARCODE 0123456789012 EAN13" q:dirty="0" data-token="kept"><q:r><q:rPr><q:b/></q:rPr><q:t>stored barcode</q:t><q:inside/></q:r></q:fldSimple></w:p>"#,
        r#"<w:p>"#,
        r#"<w:r><w:rPr><w:i/></w:rPr><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>"#,
        r#"<w:r><w:instrText xml:space="preserve">= </w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
        r#"<w:r><w:instrText>MERGEFIELD Amount</w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/><w:t>stored amount</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        r#"<w:r><w:instrText xml:space="preserve"> + 2</w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
        r#"<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>stored formula</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        r#"<w:p><w:fldSimple w:instr="TOC" w:dirty="0"><w:r><w:t>stored toc</w:t></w:r></w:fldSimple></w:p>"#,
        r#"<w:p><w:fldSimple w:instr="TC &quot;Entry&quot; \l 2" w:dirty="0"><w:r><w:t>stored tc</w:t></w:r></w:fldSimple></w:p>"#,
        r#"<w:p><w:fldSimple w:instr="NEXT" w:dirty="0"><w:r><w:t>stored next</w:t></w:r></w:fldSimple></w:p>"#,
        r#"<w:p><w:fldSimple w:instr="PAGE" w:dirty="0"><w:r><w:t>stored page</w:t></w:r></w:fldSimple></w:p>"#,
        r#"<w:p><w:fldSimple w:instr="DATE" w:dirty="0"><w:r><w:t>stored date</w:t></w:r></w:fldSimple></w:p>"#,
        r#"<w:p><w:fldSimple w:instr="UNKNOWN" w:dirty="0"><w:r><w:t>stored unknown</w:t></w:r></w:fldSimple></w:p>"#,
    );
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([("Amount".to_owned(), "3".to_owned())]),
        merge_record_number: Some(4),
        merge_sequence_number: Some(2),
        ..Default::default()
    };
    assert_eq!(document.update_fields(&context).unwrap(), 9);
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(xml.contains("DISPLAYBARCODE 0123456789012 EAN13"), "{xml}");
    assert!(xml.contains("stored barcode"), "{xml}");
    assert!(xml.contains("data-token=\"kept\""), "{xml}");
    assert!(xml.contains("<q:b/>"), "{xml}");
    assert!(xml.contains("<q:inside/>"), "{xml}");
    assert!(xml.contains("w:dirty=\"1\""), "{xml}");
    assert_eq!(xml.matches("w:dirty=\"1\"").count(), 7, "{xml}");
    for (instruction, cache) in [
        ("TOC", "stored toc"),
        ("TC &quot;Entry&quot; \\l 2", "stored tc"),
        ("NEXT", "stored next"),
        ("PAGE", "stored page"),
        ("DATE", "stored date"),
        ("UNKNOWN", "stored unknown"),
    ] {
        assert!(xml.contains(instruction), "missing {instruction}: {xml}");
        assert!(xml.contains(cache), "missing {cache}: {xml}");
    }
    for preserved in ["MERGEFIELD Amount", "<w:i/>", "<w:u w:val=\"single\"/>"] {
        assert!(xml.contains(preserved), "missing {preserved}: {xml}");
    }

    let reopened = Document::from_bytes(&saved).unwrap();
    let evaluations = reopened.evaluate_fields(&context).unwrap();
    assert_eq!(evaluations[1].cached_result, "5");
    assert_eq!(evaluations[2].cached_result, "3");
    assert_eq!(
        evaluations[1].outcome,
        FieldOutcome::Resolved("5".to_owned())
    );
    assert!(matches!(evaluations[0].outcome, FieldOutcome::Barcode(_)));
    assert!(matches!(
        evaluations[3].outcome,
        FieldOutcome::TableOfContents(_)
    ));
    assert!(matches!(
        evaluations[4].outcome,
        FieldOutcome::TableOfContentsEntry(_)
    ));
    assert_eq!(
        evaluations[5].outcome,
        FieldOutcome::MailMergeControl(MailMergeControl::NextRecord { record_number: 5 })
    );
    assert_eq!(evaluations[6].outcome, FieldOutcome::DeferredPagination);
    assert_eq!(
        evaluations[7].outcome,
        FieldOutcome::KeepStored {
            diagnostic: "DATE requires an explicit date and time, stored display retained"
                .to_owned()
        }
    );
    assert_eq!(
        evaluations[8].outcome,
        FieldOutcome::KeepStored {
            diagnostic: "field UNKNOWN is unsupported, stored display retained".to_owned()
        }
    );
}

#[test]
fn dynamic_toc_rebuild_matches_the_pinned_word_update() {
    assert_eq!(
        WORD_FIELD_ORACLE,
        "Microsoft Word 16.104 build 16.104.25121423"
    );
    assert_eq!(
        WORD_FIELD_ORACLE_ENVIRONMENT,
        "locale=en-US; calendar=Gregorian; decimal=.; grouping=,; timezone=UTC"
    );
    assert_eq!(
        DYNAMIC_TOC_WORD_ORACLE_INPUT,
        "F-232-dynamic-toc-rebuild-v1"
    );

    let body = r#"
        <w:p data-owner="keep"><w:r data-begin="keep"><w:fldChar w:fldCharType="begin"/></w:r><w:r data-code="keep"><w:instrText xml:space="preserve"> TOC \o "1-2" \t "Appendix,2" \u \f C \h </w:instrText></w:r><w:r data-separate="keep"><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>stale entry</w:t></w:r></w:p>
        <w:p data-end="keep"><w:r data-end-run="keep"><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p>
        <w:tbl><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Appendix"/></w:pPr><w:r><w:t>Appendix A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
        <w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>Outline item</w:t></w:r></w:p>
        <w:p><w:fldSimple w:instr="TC &quot;Manual entry&quot; \f C \l 2"><w:r><w:t>hidden marker</w:t></w:r></w:fldSimple></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    let report: TocRebuildReport = document.rebuild_toc().unwrap();
    assert_eq!(
        report,
        TocRebuildReport {
            entry_count: 4,
            bookmark_count: 4,
            diagnostic_count: 0,
        }
    );

    let xml = document_xml(&mut document);
    assert!(!xml.contains("stale entry"), "{xml}");
    assert_eq!(
        toc_entry_signatures(&xml),
        [
            (
                "Overview\t1".to_owned(),
                "TOC1".to_owned(),
                Some("_Toc1".to_owned()),
                Some("_Toc1".to_owned())
            ),
            (
                "Appendix A\t1".to_owned(),
                "TOC2".to_owned(),
                Some("_Toc2".to_owned()),
                Some("_Toc2".to_owned())
            ),
            (
                "Outline item\t1".to_owned(),
                "TOC2".to_owned(),
                Some("_Toc3".to_owned()),
                Some("_Toc3".to_owned())
            ),
            (
                "Manual entry\t1".to_owned(),
                "TOC2".to_owned(),
                Some("_Toc4".to_owned()),
                Some("_Toc4".to_owned())
            ),
        ]
    );
    assert_eq!(xml.matches("<w:pStyle w:val=\"TOC1\"").count(), 1, "{xml}");
    assert_eq!(xml.matches("<w:pStyle w:val=\"TOC2\"").count(), 3, "{xml}");
    assert_eq!(xml.matches("w:bookmarkStart").count(), 4, "{xml}");
    assert_eq!(xml.matches("w:hyperlink").count(), 8, "{xml}");
    assert_eq!(xml.matches("PAGEREF _Toc").count(), 4, "{xml}");
    assert!(xml.contains("TOC \\o \"1-2\""), "{xml}");
    assert!(xml.contains("w:fldCharType=\"end\""), "{xml}");
    let toc_style = xml.find("<w:pStyle w:val=\"TOC1\"").unwrap();
    let toc_tabs = xml[toc_style..].find("<w:tabs>").unwrap() + toc_style;
    let toc_link = xml[toc_tabs..].find("<w:hyperlink").unwrap() + toc_tabs;
    assert!(toc_style < toc_tabs && toc_tabs < toc_link, "{xml}");
    assert!(
        xml.contains("<w:fldChar w:fldCharType=\"begin\"/>"),
        "{xml}"
    );

    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1" \s chapter \d ":" \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old first</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \t "Appendix,2" \p </w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DOCVARIABLE TocSeparator</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>|</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old second</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="10" w:name="existingTarget"/><w:r><w:t>Chapter title</w:t></w:r><w:bookmarkEnd w:id="10"/></w:p>
        <w:p><w:pPr><w:pStyle w:val="Appendix"/></w:pPr><w:r><w:t>Appendix B</w:t></w:r></w:p>
    "#;
    let settings = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docVars><w:docVar w:name="TocSeparator" w:val="|"/></w:docVars></w:settings>"#;
    let mut document = document_with_field_parts(&wrap_word_body(body), Some(settings), None);
    assert_eq!(
        document.rebuild_toc().unwrap(),
        TocRebuildReport {
            entry_count: 2,
            bookmark_count: 1,
            diagnostic_count: 0,
        }
    );
    let xml = document_xml(&mut document);
    assert!(!xml.contains("old first"), "{xml}");
    assert!(!xml.contains("old second"), "{xml}");
    assert!(xml.contains("w:anchor=\"existingTarget\""), "{xml}");
    assert_eq!(xml.matches("w:name=\"existingTarget\"").count(), 1, "{xml}");
    assert_eq!(xml.matches("w:bookmarkStart").count(), 2, "{xml}");
    assert!(xml.contains("<w:t>:</w:t>"), "{xml}");
    assert!(xml.contains("<w:t xml:space=\"preserve\">|</w:t>"), "{xml}");
    assert_eq!(
        toc_entry_signatures(&xml),
        [
            (
                "Chapter title\t1:1".to_owned(),
                "TOC1".to_owned(),
                Some("existingTarget".to_owned()),
                Some("existingTarget".to_owned())
            ),
            (
                "Appendix B|1".to_owned(),
                "TOC2".to_owned(),
                None,
                Some("_Toc1".to_owned())
            ),
        ]
    );
}

#[test]
fn toc_rebuild_uses_final_deterministic_page_targets() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-1" \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old page</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>First page heading</w:t></w:r></w:p>
        <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Page boundary</w:t></w:r></w:p>
        <w:tbl><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Second page heading</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 2);
    let xml = document_xml(&mut document);
    assert_eq!(
        toc_entry_signatures(&xml),
        [
            (
                "First page heading\t1".to_owned(),
                "TOC1".to_owned(),
                Some("_Toc1".to_owned()),
                Some("_Toc1".to_owned())
            ),
            (
                "Second page heading\t2".to_owned(),
                "TOC1".to_owned(),
                Some("_Toc2".to_owned()),
                Some("_Toc2".to_owned())
            ),
        ]
    );

    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let layout = reopened.layout_deterministic().unwrap();
    assert!(layout.layout.pages.len() >= 2);
}

#[test]
fn toc_rebuild_preserves_unowned_field_and_package_xml() {
    let body = r#"
        <w:p><w:r><q:fldChar xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" q:fldCharType="begin"/></w:r><w:r><q:instrText xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xml:space="preserve"> TOC \o "1-1" </q:instrText></w:r><w:r><q:fldChar xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" q:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>stale cache</w:t></w:r></w:p>
        <w:p><w:r><q:fldChar xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" q:fldCharType="end"/></w:r></w:p>
        <w:p data-source="keep"><w:pPr><w:pStyle w:val="Heading1"/><producer:opaque xmlns:producer="urn:producer" producer:value="keep"/></w:pPr><w:r><w:t>Owned heading</w:t></w:r></w:p>
        <w:p><w:fldSimple w:instr="IF 1 = 1 &quot;yes&quot; &quot;no&quot;" producer:token="keep" xmlns:producer="urn:producer"><w:r><w:t>unowned cache</w:t></w:r></w:fldSimple></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let before = document.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(before)).unwrap();
    package.set_part("/custom/unrelated.bin", b"producer bytes".to_vec());
    let mut output = std::io::Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();
    let mut document = Document::from_bytes(output.get_ref()).unwrap();

    document.rebuild_toc().unwrap();
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    assert_eq!(
        package.get_part("/custom/unrelated.bin").unwrap(),
        b"producer bytes"
    );
    let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    for preserved in [
        "<q:fldChar xmlns:q=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" q:fldCharType=\"begin\"/>",
        "<q:instrText xmlns:q=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xml:space=\"preserve\">",
        "TOC \\o \"1-1\"",
        "producer:opaque",
        "producer:value=\"keep\"",
        "producer:token=\"keep\"",
        "unowned cache",
    ] {
        assert!(xml.contains(preserved), "missing {preserved}: {xml}");
    }
}

#[test]
fn toc_rebuild_rejects_ambiguous_or_malformed_sources_atomically() {
    for body in [
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p><w:bookmarkStart w:id="7" w:name="duplicate"/><w:bookmarkStart w:id="8" w:name="duplicate"/><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r><w:bookmarkEnd w:id="8"/><w:bookmarkEnd w:id="7"/></w:p>"#,
    ] {
        let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
        let before = document_xml(&mut document);
        let error = document.rebuild_toc().unwrap_err();
        assert!(
            error.to_string().contains("table of contents"),
            "unexpected error: {error}"
        );
        assert_eq!(document_xml(&mut document), before);
    }

    let unsupported = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \z</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>stored unsupported cache</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(unsupported), None, None);
    let before = document_xml(&mut document);
    assert_eq!(
        document.rebuild_toc().unwrap(),
        TocRebuildReport {
            entry_count: 0,
            bookmark_count: 0,
            diagnostic_count: 1,
        }
    );
    assert_eq!(document_xml(&mut document), before);

    let simple = r#"<w:p><w:fldSimple w:instr="TOC"><w:r><w:t>stored simple cache</w:t></w:r></w:fldSimple></w:p>"#;
    let mut document = document_with_field_parts(&wrap_word_body(simple), None, None);
    let before = document_xml(&mut document);
    assert_eq!(document.rebuild_toc().unwrap().diagnostic_count, 1);
    assert_eq!(document_xml(&mut document), before);

    let mut no_toc = Document::new();
    no_toc.add_paragraph("Heading").style("Heading1");
    assert_eq!(no_toc.rebuild_toc().unwrap(), TocRebuildReport::default());
}

#[test]
fn toc_rebuild_confines_collision_safe_page_substitution_to_owned_results() {
    let token = "__RDOCX_TOC_PAGE_0_0__";
    let body = format!(
        r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>{token}</w:t></w:r></w:p>
        <w:p><w:r><producer:opaque xmlns:producer="urn:producer" producer:value="{token}"/></w:r></w:p>
    "#
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    let xml = document_xml(&mut document);
    assert_eq!(xml.matches(token).count(), 3, "{xml}");
    assert!(
        xml.contains(&format!("producer:value=\"{token}\"")),
        "{xml}"
    );
}

#[test]
fn toc_rebuild_ignores_foreign_ancestors_and_indirect_field_markers() {
    for body in [
        r#"<x:tbl xmlns:x="urn:producer"><x:tr><x:tc><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p><w:r><w:t>foreign cache</w:t></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></x:tc></x:tr></x:tbl>"#,
        r#"<x:wrapper xmlns:x="urn:producer"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p><w:r><w:t>nested body cache</w:t></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></x:wrapper>"#,
        r#"<w:p><w:r><x:wrapper xmlns:x="urn:producer"><w:fldChar w:fldCharType="begin"/></x:wrapper></w:r><w:r><w:instrText>TOC</w:instrText></w:r><w:r><x:wrapper xmlns:x="urn:producer"><w:fldChar w:fldCharType="separate"/></x:wrapper></w:r></w:p><w:p><w:r><w:t>opaque cache</w:t></w:r></w:p><w:p><w:r><x:wrapper xmlns:x="urn:producer"><w:fldChar w:fldCharType="end"/></x:wrapper></w:r></w:p>"#,
        r#"<w:p><x:wrapper xmlns:x="urn:producer"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></x:wrapper></w:p><w:p><w:r><w:t>opaque wrapped cache</w:t></w:r></w:p><w:p><x:wrapper xmlns:x="urn:producer"><w:r><w:fldChar w:fldCharType="end"/></w:r></x:wrapper></w:p>"#,
        r#"<w:p><x:wrapper xmlns:x="urn:producer"><w:fldSimple w:instr="TOC"><w:r><w:t>opaque simple cache</w:t></w:r></w:fldSimple></x:wrapper></w:p>"#,
        r#"<w:p><w:sdtContent><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:sdtContent></w:p><w:p><w:r><w:t>opaque malformed control cache</w:t></w:r></w:p><w:p><w:sdtContent><w:r><w:fldChar w:fldCharType="end"/></w:r></w:sdtContent></w:p>"#,
        r#"<w:p><w:ins><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:ins></w:p><w:p><w:r><w:t>opaque malformed revision cache</w:t></w:r></w:p><w:p><w:ins><w:r><w:fldChar w:fldCharType="end"/></w:r></w:ins></w:p>"#,
    ] {
        let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
        let before = document_xml(&mut document);
        assert_eq!(document.rebuild_toc().unwrap(), TocRebuildReport::default());
        assert_eq!(document_xml(&mut document), before);
    }
}

#[test]
fn toc_rebuild_rejects_nested_toc_ownership_before_any_edit() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>inner cache</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let before = document_xml(&mut document);

    let error = document.rebuild_toc().unwrap_err();
    assert!(
        error.to_string().contains("nested table of contents"),
        "{error}"
    );
    assert_eq!(document_xml(&mut document), before);
}

#[test]
fn toc_bookmark_allocation_uses_the_final_id_lazily() {
    let toc = r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p><w:r><w:t>old</w:t></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#;
    let reused = format!(
        r#"{toc}<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="{}" w:name="existing"/><w:r><w:t>Heading</w:t></w:r><w:bookmarkEnd w:id="{}"/></w:p>"#,
        i32::MAX,
        i32::MAX
    );
    let mut document = document_with_field_parts(&wrap_word_body(&reused), None, None);
    assert_eq!(document.rebuild_toc().unwrap().bookmark_count, 0);
    assert!(document_xml(&mut document).contains("PAGEREF existing"));

    let allocated = format!(
        r#"{toc}<w:p><w:bookmarkStart w:id="{}" w:name="prior"/><w:r><w:t>Prior</w:t></w:r><w:bookmarkEnd w:id="{}"/></w:p><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>"#,
        i32::MAX - 1,
        i32::MAX - 1
    );
    let mut document = document_with_field_parts(&wrap_word_body(&allocated), None, None);
    assert_eq!(document.rebuild_toc().unwrap().bookmark_count, 1);
    assert!(
        document_xml(&mut document).contains(&format!("w:id=\"{}\" w:name=\"_Toc1\"", i32::MAX))
    );
}

#[test]
fn toc_rebuild_reuses_the_first_whole_paragraph_bookmark_repeatably() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1" \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="11" w:name="outer"/><w:bookmarkStart w:id="12" w:name="inner"/><w:r><w:t>Heading</w:t></w:r><w:bookmarkEnd w:id="12"/><w:bookmarkEnd w:id="11"/></w:p>
    "#;

    for _ in 0..32 {
        let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
        assert_eq!(document.rebuild_toc().unwrap().bookmark_count, 0);
        assert_eq!(
            toc_entry_signatures(&document_xml(&mut document)),
            [(
                "Heading\t1".to_owned(),
                "TOC1".to_owned(),
                Some("outer".to_owned()),
                Some("outer".to_owned())
            )]
        );
    }
}

#[test]
fn toc_rebuild_handles_the_largest_outline_level_without_panicking() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \u</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:outlineLvl w:val="4294967295"/></w:pPr><w:r><w:t>Out of range</w:t></w:r></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| document.rebuild_toc()));
    let report = result
        .expect("outline conversion stays inside the Result boundary")
        .unwrap();
    assert_eq!(report.entry_count, 0);
    assert!(!document_xml(&mut document).contains("old"));
}

#[test]
fn toc_rebuild_discovers_the_accepted_revision_projection() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1" \f C \s chapter \d ":" \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:ins w:id="20" w:author="Ada"><w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:ins></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:ins w:id="21" w:author="Ada"><w:r><w:t>Inserted heading</w:t></w:r></w:ins></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Direct </w:t></w:r><w:ins w:id="22" w:author="Ada"><w:r><w:t>and inserted</w:t></w:r></w:ins><w:r><w:t> tail</w:t></w:r></w:p>
        <w:p><w:ins w:id="23" w:author="Ada"><w:fldSimple w:instr="TC &quot;Inserted TC&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple><w:r><w:t>anchor</w:t></w:r></w:ins></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 3);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document)),
        [
            (
                "Inserted heading\t1:1".to_owned(),
                "TOC1".to_owned(),
                Some("_Toc1".to_owned()),
                Some("_Toc1".to_owned())
            ),
            (
                "Direct and inserted tail\t1:1".to_owned(),
                "TOC1".to_owned(),
                Some("_Toc2".to_owned()),
                Some("_Toc2".to_owned())
            ),
            (
                "Inserted TC\t1:1".to_owned(),
                "TOC1".to_owned(),
                Some("_Toc3".to_owned()),
                Some("_Toc3".to_owned())
            ),
        ]
    );
}

#[test]
fn toc_rebuild_does_not_reuse_a_bookmark_after_an_inline_control_prefix() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1" \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>Prefix </w:t></w:r></w:sdtContent></w:sdt><w:bookmarkStart w:id="11" w:name="partial"/><w:r><w:t>Heading</w:t></w:r><w:bookmarkEnd w:id="11"/></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().bookmark_count, 1);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document)),
        [(
            "Prefix Heading\t1".to_owned(),
            "TOC1".to_owned(),
            Some("_Toc1".to_owned()),
            Some("_Toc1".to_owned())
        )]
    );
}

#[test]
fn toc_rebuild_ignores_paragraphs_below_an_untyped_block_chain() {
    let body = r#"
        <w:sdtContent>
          <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
          <w:p><w:r><w:t>opaque cache</w:t></w:r></w:p>
          <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        </w:sdtContent>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Visible heading</w:t></w:r></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let before = document_xml(&mut document);

    assert_eq!(document.rebuild_toc().unwrap(), TocRebuildReport::default());
    assert_eq!(document_xml(&mut document), before);
}

#[test]
fn toc_rebuild_ignores_revision_content_beyond_the_typed_depth_limit() {
    let wrappers = (1..=33)
        .map(|id| format!(r#"<w:ins w:id="{id}" w:author="Ada">"#))
        .collect::<String>();
    let body = format!(
        r#"
        <w:p>{wrappers}<w:r><w:fldChar w:fldCharType="begin"/></w:r>{}</w:ins><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>opaque cache</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Visible heading</w:t></w:r></w:p>
        "#,
        "</w:ins>".repeat(32)
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);
    let before = document_xml(&mut document);

    assert_eq!(document.rebuild_toc().unwrap(), TocRebuildReport::default());
    assert_eq!(document_xml(&mut document), before);
}

#[test]
fn toc_rebuild_balances_each_supported_inline_owner_around_the_instruction() {
    for (open, close) in [
        (r#"<w:ins w:id="31" w:author="Ada">"#, "</w:ins>"),
        (r#"<w:hyperlink w:anchor="existing">"#, "</w:hyperlink>"),
        ("<w:sdt><w:sdtPr/><w:sdtContent>", "</w:sdtContent></w:sdt>"),
    ] {
        let body = format!(
            r#"
            <w:p>{open}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1" \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>{close}</w:p>
            <w:p><w:r><w:t>old</w:t></w:r></w:p>
            <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>
            "#
        );
        let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

        assert_eq!(document.rebuild_toc().unwrap().entry_count, 1, "{open}");
        assert_eq!(
            toc_entry_signatures(&document_xml(&mut document)),
            [(
                "Heading\t1".to_owned(),
                "TOC1".to_owned(),
                Some("_Toc1".to_owned()),
                Some("_Toc1".to_owned())
            )],
            "{open}",
        );
    }
}

#[test]
fn toc_rebuild_keeps_control_and_revision_fields_in_serialized_order() {
    let sequence = r#"<w:ins w:id="41" w:author="Ada"><w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:ins>"#;
    let entry = r#"<w:sdt><w:sdtContent><w:ins w:id="42" w:author="Ada"><w:fldSimple w:instr="TC &quot;Control entry&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:ins></w:sdtContent></w:sdt>"#;
    for (source, expected) in [
        (
            format!("{sequence}{entry}"),
            "Control entry\t1:1".to_owned(),
        ),
        (format!("{entry}{sequence}"), "Control entry\t1".to_owned()),
    ] {
        let body = format!(
            r#"
            <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C \s chapter \d ":"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
            <w:p><w:r><w:t>old</w:t></w:r></w:p>
            <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
            <w:p>{source}</w:p>
            "#
        );
        let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

        assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
        assert_eq!(
            toc_entry_signatures(&document_xml(&mut document))[0].0,
            expected,
        );
    }
}

#[test]
fn toc_rebuild_rejects_a_bookmark_reversed_at_one_run_boundary() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \b Scope \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkEnd w:id="7"/><w:sdt><w:sdtContent><w:r><w:t>Prefix </w:t></w:r></w:sdtContent></w:sdt><w:bookmarkStart w:id="7" w:name="Scope"/><w:r><w:t>Heading</w:t></w:r></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let before = document_xml(&mut document);

    let error = document.rebuild_toc().unwrap_err();
    assert!(error.to_string().contains("reversed bookmark"), "{error}");
    assert_eq!(document_xml(&mut document), before);
}

#[test]
fn toc_rebuild_diagnoses_simple_tocs_in_accepted_revisions() {
    for wrapper in ["ins", "moveTo"] {
        let body = format!(
            r#"<w:p><w:{wrapper} w:id="51" w:author="Ada"><w:fldSimple w:instr="TOC"><w:r><w:t>stored cache</w:t></w:r></w:fldSimple></w:{wrapper}></w:p>"#
        );
        let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);
        let before = document_xml(&mut document);

        assert_eq!(
            document.rebuild_toc().unwrap(),
            TocRebuildReport {
                entry_count: 0,
                bookmark_count: 0,
                diagnostic_count: 1,
            },
            "{wrapper}",
        );
        assert_eq!(document_xml(&mut document), before);
    }
}

#[test]
fn toc_rebuild_ignores_a_second_content_container_in_one_control() {
    let body = r#"
        <w:sdt><w:sdtContent><w:p><w:r><w:t>typed content</w:t></w:r></w:p></w:sdtContent>
          <w:sdtContent>
            <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
            <w:p><w:r><w:t>opaque cache</w:t></w:r></w:p>
            <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
          </w:sdtContent>
        </w:sdt>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Visible heading</w:t></w:r></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let before = document_xml(&mut document);

    assert_eq!(document.rebuild_toc().unwrap(), TocRebuildReport::default());
    assert_eq!(document_xml(&mut document), before);
}

#[test]
fn toc_rebuild_counts_property_revisions_toward_the_depth_limit() {
    let wrappers = (1..=32)
        .map(|id| format!(r#"<w:ins w:id="{id}" w:author="Ada">"#))
        .collect::<String>();
    let body = format!(
        r#"
        <w:p>{wrappers}<w:r><w:rPr><w:rPrChange w:id="40" w:author="Ada"><w:rPr/></w:rPrChange></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>{}<w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>opaque cache</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Visible heading</w:t></w:r></w:p>
        "#,
        "</w:ins>".repeat(32)
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);
    let before = document_xml(&mut document);

    assert_eq!(document.rebuild_toc().unwrap(), TocRebuildReport::default());
    assert_eq!(document_xml(&mut document), before);
}

#[test]
fn toc_rebuild_carries_inherited_aliases_into_wrapped_instruction_projection() {
    let body = r#"
        <w:p><q:ins xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" q:id="61" q:author="Ada"><q:r><q:fldChar q:fldCharType="begin"/></q:r><q:r><q:instrText>TOC \o "1-1" \h</q:instrText></q:r><q:r><q:fldChar q:fldCharType="separate"/></q:r></q:ins></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    let xml = document_xml(&mut document);
    assert!(xml.contains(
        r#"<q:ins xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#
    ));
    assert_eq!(toc_entry_signatures(&xml)[0].0, "Heading\t1");
}

#[test]
fn toc_rebuild_projects_accepted_revisions_inside_content_controls() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1" \f C \s chapter \d ":"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:sdt><w:sdtContent><w:ins w:id="71" w:author="Ada"><w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:ins></w:sdtContent></w:sdt></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:sdt><w:sdtContent><w:ins w:id="72" w:author="Ada"><w:r><w:t>Controlled heading</w:t></w:r></w:ins></w:sdtContent></w:sdt></w:p>
        <w:p><w:sdt><w:sdtContent><w:ins w:id="73" w:author="Ada"><w:fldSimple w:instr="TC &quot;Controlled TC&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:ins></w:sdtContent></w:sdt></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 2);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document))
            .into_iter()
            .map(|entry| entry.0)
            .collect::<Vec<_>>(),
        ["Controlled heading\t1:1", "Controlled TC\t1:1"]
    );

    for wrapper in ["ins", "moveTo"] {
        let body = format!(
            r#"<w:p><w:sdt><w:sdtContent><w:{wrapper} w:id="74" w:author="Ada"><w:fldSimple w:instr="TOC"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:{wrapper}></w:sdtContent></w:sdt></w:p>"#
        );
        let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);
        assert_eq!(document.rebuild_toc().unwrap().diagnostic_count, 1);
    }
}

#[test]
fn toc_rebuild_orders_revision_only_hyperlinks_by_their_owner_position() {
    let sequence = r#"<w:hyperlink w:anchor="sequence"><w:ins w:id="81" w:author="Ada"><w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:ins></w:hyperlink>"#;
    let entry = r#"<w:sdt><w:sdtContent><w:ins w:id="82" w:author="Ada"><w:fldSimple w:instr="TC &quot;Hyperlink order&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:ins></w:sdtContent></w:sdt>"#;
    let body = format!(
        r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C \s chapter \d ":"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p>{sequence}{entry}</w:p>
        "#
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document))[0].0,
        "Hyperlink order\t1:1"
    );
}

#[test]
fn toc_rebuild_layouts_inline_control_text_before_resolving_later_pages() {
    let long_title = "wrapped ".repeat(900);
    let body = format!(
        r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:sdt><w:sdtContent><w:r><w:t>{long_title}</w:t></w:r></w:sdtContent></w:sdt></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Later heading</w:t></w:r></w:p>
        "#
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 2);
    let entries = toc_entry_signatures(&document_xml(&mut document));
    assert_eq!(entries[0].0, format!("{long_title}\t3"));
    assert_eq!(entries[1].0, "Later heading\t6");
}

#[test]
fn toc_rebuild_injects_namespaces_after_the_real_start_tag_close() {
    let body = format!(
        r#"
        <w:p xmlns:producer="urn:producer"><q:ins xmlns:q="{W_NS}" q:id="91" q:author="Ada">
          <q:r xmlns:q="{W_NS}" producer:value=">"><q:fldChar q:fldCharType="begin"/></q:r>
          <q:r xmlns:q="{W_NS}" producer:value=">"><q:instrText>TOC \o "1-1"</q:instrText></q:r>
          <q:r xmlns:q="{W_NS}" producer:value=">"><q:fldChar q:fldCharType="separate"/></q:r>
        </q:ins></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Alias heading</w:t></w:r></w:p>
        "#
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document))[0].0,
        "Alias heading\t1"
    );
}

#[test]
fn toc_rebuild_orders_terminal_hyperlink_revisions_before_following_controls() {
    let sequence = r#"<w:hyperlink w:anchor="sequence"><w:r><w:t>lead</w:t></w:r><w:ins w:id="92" w:author="Ada"><w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:ins></w:hyperlink>"#;
    let entry = r#"<w:sdt><w:sdtContent><w:ins w:id="93" w:author="Ada"><w:fldSimple w:instr="TC &quot;Terminal order&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:ins></w:sdtContent></w:sdt>"#;
    let body = format!(
        r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C \s chapter \d ":"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p>{sequence}{entry}</w:p>
        "#
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document))[0].0,
        "Terminal order\t1:1"
    );
}

#[test]
fn toc_rebuild_filters_same_paragraph_sources_by_exact_bookmark_position() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C \b Scope</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p>
          <w:fldSimple w:instr="TC &quot;Before&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
          <w:bookmarkStart w:id="93" w:name="Scope"/>
          <w:fldSimple w:instr="TC &quot;Inside&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
          <w:bookmarkEnd w:id="93"/>
          <w:fldSimple w:instr="TC &quot;After&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
        </w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    let entries = toc_entry_signatures(&document_xml(&mut document));
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, "Inside\t1");
}

#[test]
fn toc_rebuild_excludes_only_sources_inside_the_owned_result_positions() {
    let body = r#"
        <w:p>
          <w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>1</w:t></w:r></w:fldSimple>
          <w:fldSimple w:instr="TC &quot;Before&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C \s Chapter \d ":"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:fldSimple w:instr="TC &quot;Owned start&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
        </w:p>
        <w:p><w:fldSimple w:instr="TC &quot;Owned middle&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:p>
        <w:p>
          <w:fldSimple w:instr="TC &quot;Owned end&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
          <w:fldSimple w:instr="TC &quot;After&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
        </w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 2);
    let entries = toc_entry_signatures(&document_xml(&mut document));
    assert_eq!(entries[0].0, "Before\t1:1");
    assert_eq!(entries[1].0, "After\t1:1");
}

#[test]
fn toc_rebuild_keeps_exact_positions_inside_one_accepted_insertion() {
    let body = r#"
        <w:p><w:ins w:id="96" w:author="Ada">
          <w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>1</w:t></w:r></w:fldSimple>
          <w:fldSimple w:instr="TC &quot;Before&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C \s Chapter \d ":"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
        </w:ins></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:ins w:id="98" w:author="Ada">
          <w:fldSimple w:instr="TC &quot;Owned&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
          <w:fldSimple w:instr="TC &quot;After&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
        </w:ins></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 2);
    let entries = toc_entry_signatures(&document_xml(&mut document));
    assert_eq!(entries[0].0, "Before");
    assert_eq!(entries[1].0, "After");
}

#[test]
fn toc_rebuild_keeps_exact_positions_inside_one_inline_control() {
    let body = r#"
        <w:p><w:sdt><w:sdtContent><w:ins w:id="97" w:author="Ada">
          <w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>1</w:t></w:r></w:fldSimple>
          <w:fldSimple w:instr="TC &quot;Before control&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C \s Chapter \d ":"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
        </w:ins></w:sdtContent></w:sdt></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:sdt><w:sdtContent><w:ins w:id="99" w:author="Ada">
          <w:fldSimple w:instr="TC &quot;Owned control&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
          <w:fldSimple w:instr="TC &quot;After control&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
        </w:ins></w:sdtContent></w:sdt></w:p>
        <w:p><w:fldSimple w:instr="TC &quot;Control page&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 3);
    let entries = toc_entry_signatures(&document_xml(&mut document));
    assert_eq!(entries[0].0, "Before control");
    assert_eq!(entries[1].0, "After control");
    assert_eq!(entries[2].0, "Control page\t1:1");
}

#[test]
fn toc_rebuild_uses_real_owner_slots_after_raw_paragraph_children() {
    let inner = r#"
        <w:fldSimple w:instr="TC &quot;Before owner&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
    "#;
    let owners = [
        format!(r#"<w:ins w:id="101" w:author="Ada">{inner}</w:ins>"#),
        format!(
            r#"<w:sdt><w:sdtContent><w:ins w:id="102" w:author="Ada">{inner}</w:ins></w:sdtContent></w:sdt>"#
        ),
    ];
    for owner in owners {
        let body = format!(
            r#"
            <w:p><producer:raw xmlns:producer="urn:producer"/>{owner}</w:p>
            <w:p><w:r><w:t>old</w:t></w:r></w:p>
            <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
            "#
        );
        let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

        assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
        assert_eq!(
            toc_entry_signatures(&document_xml(&mut document))[0].0,
            "Before owner"
        );
    }
}

#[test]
fn toc_rebuild_counts_comments_and_processing_instructions_in_owner_slots() {
    for raw in ["<!--producer comment-->", "<?producer keep?>"] {
        let body = format!(
            r#"
            <w:p>{raw}<w:ins w:id="105" w:author="Ada">
              <w:fldSimple w:instr="TC &quot;Before retained raw&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
            </w:ins></w:p>
            <w:p><w:r><w:t>old</w:t></w:r></w:p>
            <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
            "#
        );
        let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

        assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
        assert_eq!(
            toc_entry_signatures(&document_xml(&mut document))[0].0,
            "Before retained raw"
        );
    }
}

#[test]
fn toc_rebuild_excludes_post_separator_source_in_terminal_hyperlink_revision() {
    let body = r#"
        <w:p><w:hyperlink w:anchor="terminal"><w:r><w:t>lead</w:t></w:r><w:ins w:id="103" w:author="Ada">
          <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:fldSimple w:instr="TC &quot;Owned terminal&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
        </w:ins></w:hyperlink></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r><w:fldSimple w:instr="TC &quot;After terminal&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document))[0].0,
        "After terminal"
    );
}

#[test]
fn toc_rebuild_does_not_count_raw_hyperlink_simple_fields_as_runs() {
    let body = r#"
        <w:p><w:hyperlink w:anchor="raw-field">
          <w:fldSimple w:instr="TC &quot;Opaque hyperlink field&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:ins w:id="104" w:author="Ada"><w:fldSimple w:instr="TC &quot;Owned after raw&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:ins>
        </w:hyperlink></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r><w:fldSimple w:instr="TC &quot;After raw&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document))[0].0,
        "After raw"
    );
}

#[test]
fn toc_rebuild_does_not_count_malformed_direct_simple_fields_as_runs() {
    for malformed in [
        r#"<w:fldSimple><w:r><w:t>missing</w:t></w:r></w:fldSimple>"#,
        r#"<w:fldSimple w:instr="   "><w:r><w:t>empty</w:t></w:r></w:fldSimple>"#,
    ] {
        let body = format!(
            r#"
            <w:p>{malformed}
              <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:fldSimple w:instr="TC &quot;Owned malformed&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
            </w:p>
            <w:p><w:r><w:t>old</w:t></w:r></w:p>
            <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r><w:fldSimple w:instr="TC &quot;After malformed&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:p>
            "#
        );
        let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

        assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
        assert_eq!(
            toc_entry_signatures(&document_xml(&mut document))[0].0,
            "After malformed"
        );
    }
}

#[test]
fn toc_rebuild_excludes_revision_source_before_a_direct_hyperlink_end_run() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:hyperlink w:anchor="end"><w:ins w:id="106" w:author="Ada"><w:fldSimple w:instr="TC &quot;Owned hyperlink end&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:ins><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink><w:fldSimple w:instr="TC &quot;After hyperlink end&quot; \f C \l 1 \n"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document))[0].0,
        "After hyperlink end"
    );
}

#[test]
fn toc_rebuild_heading_title_excludes_old_result_on_the_end_boundary() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old middle</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>old boundary</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>Real heading</w:t></w:r></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    let xml = document_xml(&mut document);
    let entries = toc_entry_signatures(&xml);
    assert_eq!(entries[0].0, "Real heading\t1");
    assert_eq!(entries[0].2.as_deref(), Some("_Toc1"));
    assert!(!xml.contains("old middle"), "{xml}");
    assert!(!xml.contains("old boundary"), "{xml}");
    CT_Document::from_xml(xml.as_bytes()).expect("replacement XML remains structurally valid");
    let heading_style = xml.find(r#"w:val="Heading1""#).unwrap();
    let end_marker = xml[heading_style..]
        .find(r#"w:fldCharType="end""#)
        .map(|offset| heading_style + offset)
        .unwrap();
    let target_start = xml[end_marker..]
        .find(r#"w:name="_Toc1""#)
        .map(|offset| end_marker + offset)
        .unwrap();
    let heading_text = xml[target_start..]
        .find("Real heading")
        .map(|offset| target_start + offset)
        .unwrap();
    assert!(
        heading_style < end_marker && end_marker < target_start && target_start < heading_text,
        "{xml}"
    );
    let target = document
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("_Toc1"))
        .expect("generated TOC target exists");
    assert_eq!(target.text(), "Real heading");

    let bytes = document.to_bytes().unwrap();
    let mut reopened = Document::from_bytes(&bytes).unwrap();
    let reopened_xml = document_xml(&mut reopened);
    assert!(!reopened_xml.contains("old middle"), "{reopened_xml}");
    assert!(!reopened_xml.contains("old boundary"), "{reopened_xml}");
    let target = reopened
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("_Toc1"))
        .expect("generated TOC target survives reopening");
    assert_eq!(target.text(), "Real heading");
}

#[test]
fn toc_rebuild_preserves_an_empty_aliased_end_paragraph_property_element() {
    let exact_properties = r#"<q:pPr xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:producer="urn:producer" producer:keep="exact"/>"#;
    let body = format!(
        r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old cache</w:t></w:r></w:p>
        <w:p>{exact_properties}<w:r><w:t>old boundary</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>
        "#
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    let bytes = document.to_bytes().unwrap();
    let mut reopened = Document::from_bytes(&bytes).unwrap();
    let xml = document_xml(&mut reopened);
    assert!(xml.contains(exact_properties), "{xml}");
    assert!(!xml.contains("old cache"), "{xml}");
    assert!(!xml.contains("old boundary"), "{xml}");
}

#[test]
fn toc_rebuild_preserves_the_complete_end_marker_control_prefix() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old cache</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>old boundary</w:t></w:r><w:sdt producer:root="keep"><w:sdtPr producer:properties="keep"><w:alias w:val="End control"/><producer:before producer:value="one"/><w:tag w:val="end-tag"/><w:id w:val="77"/><w:dataBinding w:storeItemID="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" w:xpath="/root/value" w:prefixMappings="xmlns:n='urn:value'" producer:binding="keep"/><w:text/><producer:after producer:value="two"/></w:sdtPr><w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr><w:sdtContent producer:content="keep"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:sdtContent></w:sdt><w:r><w:t>Real heading</w:t></w:r></w:p>
    "#;
    let xml = wrap_word_body(body).replacen(
        "<w:document ",
        r#"<w:document xmlns:producer="urn:producer" "#,
        1,
    );
    let mut document = document_with_field_parts(&xml, None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    let bytes = document.to_bytes().unwrap();
    let mut reopened = Document::from_bytes(&bytes).unwrap();
    let xml = document_xml(&mut reopened);
    for preserved in [
        r#"producer:root="keep""#,
        r#"producer:properties="keep""#,
        r#"w:alias w:val="End control""#,
        r#"producer:before producer:value="one""#,
        r#"w:tag w:val="end-tag""#,
        r#"w:id w:val="77""#,
        r#"w:storeItemID="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}""#,
        r#"w:xpath="/root/value""#,
        r#"producer:binding="keep""#,
        "<w:text/>",
        r#"producer:after producer:value="two""#,
        "<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>",
        r#"producer:content="keep""#,
    ] {
        assert!(xml.contains(preserved), "missing {preserved}: {xml}");
    }
    let properties = xml.find("<w:sdtPr").unwrap();
    let content = xml.find("<w:sdtContent").unwrap();
    let end_marker = xml.find(r#"w:fldCharType="end""#).unwrap();
    assert!(properties < content && content < end_marker, "{xml}");
    assert!(!xml.contains("old cache"), "{xml}");
    assert!(!xml.contains("old boundary"), "{xml}");
}

#[test]
fn toc_rebuild_targets_post_end_text_inside_each_accepted_wrapper() {
    for (label, wrapped_end) in [
        (
            "hyperlink",
            r#"<w:hyperlink w:anchor="source"><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>Inside hyperlink</w:t></w:r></w:hyperlink>"#,
        ),
        (
            "revision",
            r#"<w:ins w:id="81" w:author="Ada"><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>Inside revision</w:t></w:r></w:ins>"#,
        ),
        (
            "control",
            r#"<w:sdt><w:sdtPr><w:tag w:val="source"/></w:sdtPr><w:sdtContent><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>Inside control</w:t></w:r></w:sdtContent></w:sdt>"#,
        ),
    ] {
        let expected = format!("Inside {label}");
        let body = format!(
            r#"
            <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
            <w:p><w:r><w:t>old cache</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>old boundary</w:t></w:r>{wrapped_end}</w:p>
            "#
        );
        let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

        assert_eq!(document.rebuild_toc().unwrap().entry_count, 1, "{label}");
        let xml = document_xml(&mut document);
        assert_eq!(
            toc_entry_signatures(&xml),
            [(
                format!("{expected}\t1"),
                "TOC1".to_owned(),
                Some("_Toc1".to_owned()),
                Some("_Toc1".to_owned())
            )],
            "{label}"
        );
        CT_Document::from_xml(xml.as_bytes()).expect("rebuilt wrapper remains structurally valid");
        let end_marker = xml.find(r#"w:fldCharType="end""#).unwrap();
        let target_start = xml.find(r#"w:name="_Toc1""#).unwrap();
        let heading_text = xml[target_start..]
            .find(&expected)
            .map(|offset| target_start + offset)
            .unwrap();
        assert!(
            end_marker < target_start && target_start < heading_text,
            "{label}: {xml}"
        );
        let target_end = xml[target_start..]
            .find("<w:bookmarkEnd")
            .map(|offset| target_start + offset)
            .unwrap();
        let target_xml = &xml[target_start..target_end];
        assert!(
            target_xml.contains(&format!("<w:t>{expected}</w:t>")),
            "{label}: {target_xml}"
        );
        assert_eq!(
            target_xml.matches("<w:t").count(),
            1,
            "{label}: {target_xml}"
        );
        assert!(
            !target_xml.contains("old boundary"),
            "{label}: {target_xml}"
        );

        let bytes = document.to_bytes().unwrap();
        let mut reopened = Document::from_bytes(&bytes).unwrap();
        let target = reopened
            .bookmarks()
            .into_iter()
            .find(|bookmark| bookmark.name() == Some("_Toc1"))
            .expect("wrapper-local generated target remains matched after reopen");
        assert_eq!(target.text(), expected, "{label}");
        let layout = reopened.layout_deterministic().unwrap();
        let mut target_pages = 0usize;
        for page in &layout.layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| {
                if matches!(
                    element,
                    oxml_layout::PositionedElement::Text(run)
                        if matches!(run.field_kind, Some(oxml_layout::FieldKind::TargetPage(_)))
                ) {
                    target_pages += 1;
                }
            });
        }
        assert_eq!(target_pages, 1, "{label}");
        assert_eq!(reopened.rebuild_toc().unwrap().entry_count, 1, "{label}");
    }
}

#[test]
fn toc_rebuild_repairs_wrapper_local_bookmarks_for_every_target_policy() {
    for (label, switches, expected_target, expected_page) in [
        ("link-page", r#"\h"#, Some("_Toc1"), Some("_Toc1")),
        ("link-only", r#"\h \n &quot;1-1&quot;"#, Some("_Toc1"), None),
        ("target-free", r#"\n &quot;1-1&quot;"#, None, None),
    ] {
        let body = format!(
            r#"
            <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; {switches}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
            <w:p><w:bookmarkStart w:id="91" w:name="originalTarget"/><w:r><w:t>old cache</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:ins w:id="92" w:author="Ada"><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>Repaired heading</w:t></w:r><w:bookmarkEnd w:id="91"/></w:ins></w:p>
            <w:p><w:hyperlink w:anchor="originalTarget"><w:r><w:t>existing reference</w:t></w:r></w:hyperlink></w:p>
            "#
        );
        let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

        let report = document.rebuild_toc().unwrap();
        assert_eq!(report.entry_count, 1, "{label}");
        let bytes = document.to_bytes().unwrap();
        let mut reopened = Document::from_bytes(&bytes).unwrap();
        let xml = document_xml(&mut reopened);
        assert_eq!(
            toc_entry_signatures(&xml),
            [(
                if expected_page.is_some() {
                    "Repaired heading\t1".to_owned()
                } else {
                    "Repaired heading".to_owned()
                },
                "TOC1".to_owned(),
                expected_target.map(str::to_owned),
                expected_page.map(str::to_owned),
            )],
            "{label}"
        );
        let original = reopened
            .bookmarks()
            .into_iter()
            .find(|bookmark| bookmark.name() == Some("originalTarget"))
            .expect("wrapper-local repaired target remains matched after reopen");
        assert_eq!(original.text(), "Repaired heading", "{label}");
        assert!(xml.contains(r#"w:anchor="originalTarget""#), "{label}");
        if expected_page.is_some() {
            reopened.layout_deterministic().unwrap();
        }
        assert_eq!(reopened.rebuild_toc().unwrap().entry_count, 1, "{label}");
    }
}

#[test]
fn wrapper_local_bookmark_ranges_use_the_same_accepted_boundaries_as_text() {
    let body = r#"
        <w:p><w:ins w:id="1" w:author="Ada"><w:r><w:t>revision before</w:t></w:r><w:bookmarkStart w:id="11" w:name="revisionRange"/><w:r><w:t>revision inside</w:t></w:r></w:ins><w:bookmarkEnd w:id="11"/><w:r><w:t>revision after</w:t></w:r></w:p>
        <w:p><w:sdt><w:sdtContent><w:bookmarkStart w:id="12" w:name="controlRange"/><w:r><w:t>control inside</w:t></w:r><w:bookmarkEnd w:id="12"/><w:r><w:t>control after</w:t></w:r></w:sdtContent></w:sdt></w:p>
        <w:p><w:sdt><w:sdtContent><w:bookmarkEnd w:id="13"/><w:r><w:t>reversed</w:t></w:r><w:bookmarkStart w:id="13" w:name="reversedRange"/></w:sdtContent></w:sdt></w:p>
        <w:p><w:bookmarkEnd w:id="14"/><w:bookmarkStart w:id="14" w:name="sameBoundaryReversed"/><w:bookmarkStart w:id="15" w:name="sameBoundaryEmpty"/><w:bookmarkEnd w:id="15"/><w:r><w:t>after empty ranges</w:t></w:r></w:p>
    "#;
    let document = document_with_field_parts(&wrap_word_body(body), None, None);
    let bookmarks = document.bookmarks();
    let revision = bookmarks
        .iter()
        .find(|bookmark| bookmark.name() == Some("revisionRange"))
        .expect("revision bookmark");
    assert_eq!(revision.text(), "revision inside");
    assert_eq!(
        revision.range(),
        Some(RunRange {
            start: RunPosition {
                body_index: 0,
                run_index: 1,
            },
            end: RunPosition {
                body_index: 0,
                run_index: 2,
            },
        })
    );
    let control = bookmarks
        .iter()
        .find(|bookmark| bookmark.name() == Some("controlRange"))
        .expect("control bookmark");
    assert_eq!(control.text(), "control inside");
    assert_eq!(
        control.range(),
        Some(RunRange {
            start: RunPosition {
                body_index: 1,
                run_index: 0,
            },
            end: RunPosition {
                body_index: 1,
                run_index: 1,
            },
        })
    );
    let reversed = bookmarks
        .iter()
        .find(|bookmark| bookmark.name() == Some("reversedRange"))
        .expect("reversed bookmark report");
    assert!(reversed.range().is_none());
    assert_eq!(
        reversed.issue(),
        Some("bookmark id 13 ends before it starts")
    );
    let same_boundary_reversed = bookmarks
        .iter()
        .find(|bookmark| bookmark.name() == Some("sameBoundaryReversed"))
        .expect("same-boundary reversed bookmark report");
    assert!(same_boundary_reversed.range().is_none());
    assert_eq!(
        same_boundary_reversed.issue(),
        Some("bookmark id 14 ends before it starts")
    );
    let same_boundary_empty = bookmarks
        .iter()
        .find(|bookmark| bookmark.name() == Some("sameBoundaryEmpty"))
        .expect("same-boundary empty bookmark");
    assert_eq!(same_boundary_empty.text(), "");
    assert_eq!(
        same_boundary_empty.range(),
        Some(RunRange {
            start: RunPosition {
                body_index: 3,
                run_index: 0,
            },
            end: RunPosition {
                body_index: 3,
                run_index: 0,
            },
        })
    );
}

#[test]
fn ancestor_only_bookmark_alias_survives_live_comment_and_bookmark_mutations() {
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{word}" xmlns:q="{word}"><w:body>
        <w:p><w:fldSimple w:instr=" REF aliasTarget "><w:r><w:t>stale ref</w:t></w:r></w:fldSimple><w:r><w:t>/</w:t></w:r><w:fldSimple w:instr=" PAGEREF aliasTarget "><w:r><w:t>stale page</w:t></w:r></w:fldSimple></w:p>
        <w:p><q:bookmarkStart q:id="41" q:name="aliasTarget"/><w:r><w:t>live target</w:t></w:r><q:bookmarkEnd q:id="41"/></w:p>
        </w:body></w:document>"#,
        word = rdocx_oxml::namespace::W_NS,
    );
    let mut document = document_with_field_parts(&xml, None, None);
    let target_range = RunRange {
        start: RunPosition {
            body_index: 1,
            run_index: 0,
        },
        end: RunPosition {
            body_index: 1,
            run_index: 1,
        },
    };

    document
        .add_comment(target_range, "Ada", Some("A"), "review")
        .unwrap();
    document.add_bookmark("secondTarget", target_range).unwrap();
    let bookmarks = document.bookmarks();
    for name in ["aliasTarget", "secondTarget"] {
        let bookmark = bookmarks
            .iter()
            .find(|bookmark| bookmark.name() == Some(name))
            .unwrap_or_else(|| panic!("missing live {name}"));
        assert_eq!(bookmark.text(), "live target", "{name}");
        assert!(bookmark.issue().is_none(), "{name}: {:?}", bookmark.issue());
    }

    let layout = document.layout_deterministic().unwrap();
    let mut page_targets = 0usize;
    let mut text = String::new();
    for page in &layout.layout.pages {
        oxml_layout::walk(&page.elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Text(run) = element {
                text.push_str(&run.text);
                if matches!(run.field_kind, Some(oxml_layout::FieldKind::TargetPage(_))) {
                    page_targets += 1;
                }
            }
        });
    }
    assert!(text.contains("live target"), "{text}");
    assert!(!text.contains("stale ref"), "{text}");
    assert!(!text.contains("stale page"), "{text}");
    assert_eq!(page_targets, 1);
}

#[test]
fn toc_sources_ignore_block_children_of_an_inline_content_control() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C \s Chapter \d ":"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:sdt><w:sdtContent>
          <w:p><w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>7</w:t></w:r></w:fldSimple><w:fldSimple w:instr="TC &quot;Opaque block&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:p>
          <w:ins w:id="71" w:author="Ada"><w:fldSimple w:instr="TC &quot;Visible inline&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple></w:ins>
        </w:sdtContent></w:sdt></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document))[0].0,
        "Visible inline\t1"
    );
}

#[test]
fn toc_heading_title_ignores_block_children_of_an_inline_content_control() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \n</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:sdt><w:sdtContent><w:p><w:r><w:t>Opaque block</w:t></w:r></w:p><w:r><w:t>Visible heading</w:t></w:r></w:sdtContent></w:sdt></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document))[0].0,
        "Visible heading"
    );
}

#[test]
fn toc_bookmark_scope_uses_nested_accepted_run_boundaries() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \f C \b Scope \s Chapter \d ":"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:ins w:id="21" w:author="Ada">
          <w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>7</w:t></w:r></w:fldSimple>
          <w:fldSimple w:instr="TC &quot;Before&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
          <w:bookmarkStart w:id="22" w:name="Scope"/>
          <w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>8</w:t></w:r></w:fldSimple>
          <w:fldSimple w:instr="TC &quot;Inside&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
          <w:bookmarkEnd w:id="22"/>
          <w:fldSimple w:instr="TC &quot;After&quot; \f C \l 1"><w:r><w:t>hidden</w:t></w:r></w:fldSimple>
        </w:ins></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    let entries = toc_entry_signatures(&document_xml(&mut document));
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, "Inside\t2:1");
}

#[test]
fn toc_bookmark_scope_excludes_a_heading_before_a_nested_start_boundary() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \b HeadingScope</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:sdt><w:sdtContent><w:r><w:t>Excluded heading</w:t></w:r><w:bookmarkStart w:id="31" w:name="HeadingScope"/></w:sdtContent></w:sdt></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:sdt><w:sdtContent><w:r><w:t>Included heading</w:t></w:r><w:bookmarkEnd w:id="31"/></w:sdtContent></w:sdt></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    let entries = toc_entry_signatures(&document_xml(&mut document));
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, "Included heading\t1");
}

#[test]
fn toc_block_control_targets_at_every_main_story_level_resolve_exact_pages() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>
        <w:sdt><w:sdtContent><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Body control</w:t></w:r></w:p></w:sdtContent></w:sdt>
        <w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>
        <w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:sdt><w:sdtContent><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Table control</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt></w:tbl>
        <w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>
        <w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:sdt><w:sdtContent><w:tc><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Row control</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt></w:tr></w:tbl>
        <w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>
        <w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:sdt><w:sdtContent><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Cell control</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 4);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document)),
        [
            (
                "Body control\t2".to_owned(),
                "TOC1".to_owned(),
                Some("_Toc1".to_owned()),
                Some("_Toc1".to_owned()),
            ),
            (
                "Table control\t3".to_owned(),
                "TOC1".to_owned(),
                Some("_Toc2".to_owned()),
                Some("_Toc2".to_owned()),
            ),
            (
                "Row control\t4".to_owned(),
                "TOC1".to_owned(),
                Some("_Toc3".to_owned()),
                Some("_Toc3".to_owned()),
            ),
            (
                "Cell control\t5".to_owned(),
                "TOC1".to_owned(),
                Some("_Toc4".to_owned()),
                Some("_Toc4".to_owned()),
            ),
        ]
    );
    let targets = document
        .bookmarks()
        .into_iter()
        .filter(|bookmark| bookmark.name().is_some_and(|name| name.starts_with("_Toc")))
        .collect::<Vec<_>>();
    assert_eq!(
        targets
            .iter()
            .map(|bookmark| (bookmark.name().unwrap(), bookmark.text()))
            .collect::<Vec<_>>(),
        [
            ("_Toc1", "Body control"),
            ("_Toc2", "Table control"),
            ("_Toc3", "Row control"),
            ("_Toc4", "Cell control"),
        ]
    );
    let layout = document.layout_deterministic().unwrap();
    let mut resolved = Vec::new();
    for page in &layout.layout.pages {
        oxml_layout::walk(&page.elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Text(run) = element
                && let Some(oxml_layout::FieldKind::TargetPage(target)) = run.field_kind
            {
                resolved.push((
                    layout.page_reference_name(target).unwrap().to_owned(),
                    run.text.clone(),
                ));
            }
        });
    }
    assert_eq!(
        resolved,
        [
            ("_Toc1".to_owned(), "2".to_owned()),
            ("_Toc2".to_owned(), "3".to_owned()),
            ("_Toc3".to_owned(), "4".to_owned()),
            ("_Toc4".to_owned(), "5".to_owned()),
        ]
    );
    assert_eq!(document.rebuild_toc().unwrap().entry_count, 4);
}

#[test]
fn toc_rebuild_rejects_an_unresolved_page_target_atomically() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Unlaid continuation</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let before = document.to_bytes().unwrap();

    let error = document.rebuild_toc().unwrap_err();
    assert_eq!(
        error.to_string(),
        "table of contents page target _Toc1 was not resolved"
    );
    assert_eq!(document.to_bytes().unwrap(), before);
}

fn invalid_block_control_source(label: &str, id: i32) -> String {
    format!(
        r#"<w:p data-case="{label}"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="{id}" w:name="invalid_{label}"/><w:fldSimple w:instr=" SEQ Chapter "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" TC &quot;TC {label}&quot; \f C \l 1 "/><w:r><w:t>Heading {label}</w:t></w:r><w:bookmarkEnd w:id="{id}"/></w:p>"#
    )
}

fn invalid_block_control_owner_matrix() -> String {
    let body_row = invalid_block_control_source("body_row", 201);
    let body_cell = invalid_block_control_source("body_cell", 202);
    let table_paragraph = invalid_block_control_source("table_paragraph", 203);
    let table_table = invalid_block_control_source("table_table", 204);
    let table_cell = invalid_block_control_source("table_cell", 205);
    let row_paragraph = invalid_block_control_source("row_paragraph", 206);
    let row_table = invalid_block_control_source("row_table", 207);
    let row_row = invalid_block_control_source("row_row", 208);
    let cell_row = invalid_block_control_source("cell_row", 209);
    let cell_cell = invalid_block_control_source("cell_cell", 210);
    format!(
        r#"
        <w:sdt><w:sdtContent><w:tr><w:tc>{body_row}</w:tc></w:tr><w:tc>{body_cell}</w:tc><w:p/></w:sdtContent></w:sdt>
        <w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:sdt><w:sdtContent>{table_paragraph}<w:tbl><w:tr><w:tc>{table_table}</w:tc></w:tr></w:tbl><w:tc>{table_cell}</w:tc></w:sdtContent></w:sdt><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
        <w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:sdt><w:sdtContent>{row_paragraph}<w:tbl><w:tr><w:tc>{row_table}</w:tc></w:tr></w:tbl><w:tr><w:tc>{row_row}</w:tc></w:tr></w:sdtContent></w:sdt><w:tc><w:p/></w:tc></w:tr></w:tbl>
        <w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:sdt><w:sdtContent><w:tr><w:tc>{cell_row}</w:tc></w:tr><w:tc>{cell_cell}</w:tc><w:p/></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>
        "#
    )
}

fn assert_invalid_block_control_sources_are_opaque(document: &mut Document) {
    let xml = document_xml(document);
    for label in [
        "body_row",
        "body_cell",
        "table_paragraph",
        "table_table",
        "table_cell",
        "row_paragraph",
        "row_table",
        "row_row",
        "cell_row",
        "cell_cell",
    ] {
        assert!(
            xml.contains(&format!(r#"<w:p data-case="{label}">"#)),
            "invalid source {label} was not preserved exactly: {xml}"
        );
        assert!(
            document
                .bookmarks()
                .iter()
                .all(|bookmark| bookmark.name() != Some(&format!("invalid_{label}"))),
            "invalid source {label} exposed a public bookmark"
        );
    }
}

#[test]
fn toc_rebuild_ignores_every_block_control_owner_mismatch_without_page_fields() {
    let body = format!(
        r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \f C \n &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        {}
        "#,
        invalid_block_control_owner_matrix()
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    assert_eq!(
        document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap()
            .len(),
        0,
        "invalid SEQ and TC fields must remain opaque"
    );
    let report = document
        .rebuild_toc()
        .expect("invalid children stay opaque");
    assert_eq!(report.entry_count, 0);
    assert!(toc_entry_signatures(&document_xml(&mut document)).is_empty());
    assert_invalid_block_control_sources_are_opaque(&mut document);
}

#[test]
fn toc_rebuild_does_not_treat_owner_mismatches_as_unresolved_pageref_sources() {
    let body = format!(
        r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \f C \s Chapter</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        {}
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Valid heading</w:t></w:r></w:p>
        "#,
        invalid_block_control_owner_matrix()
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    assert_eq!(
        document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap()
            .len(),
        0,
        "invalid SEQ and TC fields must remain opaque"
    );
    let report = document
        .rebuild_toc()
        .expect("invalid children are not unresolved valid sources");
    assert_eq!(report.entry_count, 1);
    assert_eq!(
        toc_entry_signatures(&document_xml(&mut document)),
        [(
            "Valid heading\t1".to_owned(),
            "TOC1".to_owned(),
            None,
            Some("_Toc1".to_owned()),
        )]
    );
    assert_invalid_block_control_sources_are_opaque(&mut document);
}

#[test]
fn invalid_block_control_paragraphs_before_toc_do_not_shift_raw_coordinates() {
    let body = format!(
        r#"
        {}
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \h \n &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p data-case="valid-target"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Valid target</w:t></w:r></w:p>
        "#,
        invalid_block_control_owner_matrix()
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    let report = document
        .rebuild_toc()
        .expect("raw and typed coordinates agree");
    assert_eq!(report.entry_count, 1);
    assert_eq!(report.bookmark_count, 1);
    let xml = document_xml(&mut document);
    assert_eq!(
        toc_entry_signatures(&xml),
        [(
            "Valid target".to_owned(),
            "TOC1".to_owned(),
            Some("_Toc1".to_owned()),
            None,
        )]
    );
    let target = document
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("_Toc1"))
        .expect("generated target is public");
    assert_eq!(target.text(), "Valid target");
    let valid_text = xml.rfind(">Valid target<").unwrap();
    let valid_start = xml[..valid_text].rfind("<w:p").unwrap();
    let valid_end = valid_start + xml[valid_start..].find("</w:p>").unwrap();
    let marker = xml.find(r#"w:name="_Toc1""#).unwrap();
    assert!((valid_start..valid_end).contains(&marker), "{xml}");
    assert_invalid_block_control_sources_are_opaque(&mut document);
}

#[test]
fn invalidly_owned_complex_and_simple_tocs_remain_opaque() {
    let body = r#"
        <w:tbl><w:tblGrid/><w:sdt><w:sdtContent>
          <w:p data-case="invalid-complex"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>opaque cache</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
          <w:p data-case="invalid-simple"><w:fldSimple w:instr="TOC \o &quot;1-1&quot;"><w:r><w:t>opaque simple</w:t></w:r></w:fldSimple></w:p>
        </w:sdtContent></w:sdt><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let before = document_xml(&mut document);

    assert_eq!(document.rebuild_toc().unwrap(), TocRebuildReport::default());
    assert_eq!(document_xml(&mut document), before);
}

#[test]
fn malformed_block_control_shell_keeps_complex_and_simple_tocs_opaque() {
    let body = r#"
        <w:sdt data-case="malformed-block"><w:sdtPr><w:id w:val="not-an-integer"/></w:sdtPr><w:sdtContent>
          <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
          <w:p><w:r><w:t>opaque cache</w:t></w:r></w:p>
          <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r><w:fldSimple w:instr="TOC \o &quot;1-1&quot;"><w:r><w:t>opaque simple</w:t></w:r></w:fldSimple></w:p>
        </w:sdtContent></w:sdt>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let before = document_xml(&mut document);

    assert_eq!(document.rebuild_toc().unwrap(), TocRebuildReport::default());
    assert_eq!(document_xml(&mut document), before);
}

#[test]
fn malformed_inline_control_shell_keeps_complex_and_simple_tocs_opaque() {
    let body = r#"
        <w:p><w:sdt data-case="malformed-inline"><w:sdtPr><w:tag/></w:sdtPr><w:sdtContent>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:fldSimple w:instr="TOC \o &quot;1-1&quot;"><w:r><w:t>opaque simple</w:t></w:r></w:fldSimple><w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:sdtContent></w:sdt></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let before = document_xml(&mut document);

    assert_eq!(document.rebuild_toc().unwrap(), TocRebuildReport::default());
    assert_eq!(document_xml(&mut document), before);
}

#[test]
fn malformed_control_fields_and_bookmarks_do_not_shift_later_toc_coordinates() {
    let body = r#"
        <w:sdt data-case="malformed-block-prefix"><w:sdtPr><w:id w:val="invalid"/></w:sdtPr><w:sdtContent>
          <w:p><w:bookmarkStart w:id="700" w:name="opaqueBlock"/><w:fldSimple w:instr="TOC \o &quot;1-1&quot;"><w:r><w:t>opaque TOC</w:t></w:r></w:fldSimple><w:fldSimple w:instr="SEQ Chapter"><w:r><w:t>99</w:t></w:r></w:fldSimple><w:r><w:t>opaque block</w:t></w:r><w:bookmarkEnd w:id="700"/></w:p>
        </w:sdtContent></w:sdt>
        <w:p><w:sdt data-case="malformed-inline-prefix"><w:sdtPr><w:tag/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id="701" w:name="opaqueInline"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>SEQ Chapter</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>99</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="701"/></w:sdtContent></w:sdt></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \h \n &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Valid target</w:t></w:r></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert!(document.bookmarks().is_empty());
    assert!(
        document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap()
            .is_empty()
    );
    let report = document
        .rebuild_toc()
        .expect("opaque prefixes do not shift coordinates");
    assert_eq!(report.entry_count, 1);
    assert_eq!(report.bookmark_count, 1);
    assert_eq!(report.diagnostic_count, 0);
    let target = document
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("_Toc1"))
        .expect("generated target is public");
    assert_eq!(target.id(), Some(1));
    assert_eq!(target.text(), "Valid target");
    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
}

#[test]
fn toc_rebuild_does_not_reuse_a_whole_bookmark_for_the_begin_boundary_fragment() {
    let body = r#"
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="41" w:name="wholeBegin"/><w:bookmarkStart w:id="42" w:name="wholeBeginInner"/><w:r><w:t>Begin fragment</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:bookmarkEnd w:id="42"/><w:bookmarkEnd w:id="41"/></w:p>
        <w:p><w:r><w:t>old cache</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    let report = document.rebuild_toc().unwrap();
    assert_eq!(report.entry_count, 1);
    assert_eq!(report.bookmark_count, 1);
    let xml = document_xml(&mut document);
    assert_eq!(
        toc_entry_signatures(&xml),
        [(
            "Begin fragment\t1".to_owned(),
            "TOC1".to_owned(),
            Some("_Toc1".to_owned()),
            Some("_Toc1".to_owned())
        )]
    );
    assert!(xml.contains(r#"w:id="43" w:name="_Toc1""#), "{xml}");
    let target = document
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("_Toc1"))
        .expect("generated hyperlink and page field have a surviving target");
    assert_eq!(target.text(), "Begin fragment");
    let original = document
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("wholeBegin"))
        .expect("the original whole-paragraph bookmark remains matched");
    assert_eq!(original.text(), "Begin fragment");
    let inner = document
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("wholeBeginInner"))
        .expect("every nested whole-paragraph bookmark remains matched");
    assert_eq!(inner.text(), "Begin fragment");
}

#[test]
fn toc_rebuild_repairs_a_partial_begin_bookmark_with_link_and_page_targets() {
    let body = r#"
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Outside </w:t></w:r><w:bookmarkStart w:id="71" w:name="partialBegin"/><w:r><w:t>selected</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:bookmarkEnd w:id="71"/></w:p>
        <w:p><w:r><w:t>old cache</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:hyperlink w:anchor="partialBegin"><w:r><w:t>existing reference</w:t></w:r></w:hyperlink></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    let xml = document_xml(&mut document);
    assert_eq!(
        toc_entry_signatures(&xml)[0],
        (
            "Outside selected\t1".to_owned(),
            "TOC1".to_owned(),
            Some("_Toc1".to_owned()),
            Some("_Toc1".to_owned())
        )
    );
    assert!(xml.contains(r#"w:anchor="partialBegin""#), "{xml}");
    let original = document
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("partialBegin"))
        .expect("partial begin bookmark remains matched");
    assert_eq!(original.text(), "selected");
}

#[test]
fn toc_rebuild_repairs_a_cross_paragraph_bookmark_with_a_hyperlink_only_entry() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \h \n &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:bookmarkStart w:id="72" w:name="crossEnd"/><w:r><w:t>old cache</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>End selected</w:t></w:r><w:bookmarkEnd w:id="72"/></w:p>
        <w:p><w:hyperlink w:anchor="crossEnd"><w:r><w:t>existing reference</w:t></w:r></w:hyperlink></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
    let xml = document_xml(&mut document);
    assert_eq!(
        toc_entry_signatures(&xml)[0],
        (
            "End selected".to_owned(),
            "TOC1".to_owned(),
            Some("_Toc1".to_owned()),
            None
        )
    );
    assert!(xml.contains(r#"w:anchor="crossEnd""#), "{xml}");
    let original = document
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("crossEnd"))
        .expect("cross-paragraph bookmark remains matched");
    assert_eq!(original.text(), "End selected");
}

#[test]
fn toc_rebuild_repairs_a_partial_end_bookmark_without_a_generated_target() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \n &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old cache</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>outside old</w:t></w:r><w:bookmarkStart w:id="73" w:name="partialEnd"/><w:r><w:t>inside old</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>End selected</w:t></w:r><w:bookmarkEnd w:id="73"/></w:p>
        <w:p><w:hyperlink w:anchor="partialEnd"><w:r><w:t>existing reference</w:t></w:r></w:hyperlink></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    let report = document.rebuild_toc().unwrap();
    assert_eq!(report.entry_count, 1);
    assert_eq!(report.bookmark_count, 0);
    let xml = document_xml(&mut document);
    assert_eq!(
        toc_entry_signatures(&xml)[0],
        ("End selected".to_owned(), "TOC1".to_owned(), None, None)
    );
    assert!(xml.contains(r#"w:anchor="partialEnd""#), "{xml}");
    let original = document
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("partialEnd"))
        .expect("partial end bookmark remains matched");
    assert_eq!(original.text(), "End selected");
}

#[test]
fn toc_rebuild_does_not_reuse_a_whole_bookmark_for_the_end_boundary_fragment() {
    let body = r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \h \n &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old cache</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="51" w:name="wholeEnd"/><w:r><w:t>old boundary</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>End fragment</w:t></w:r><w:bookmarkEnd w:id="51"/></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    let report = document.rebuild_toc().unwrap();
    assert_eq!(report.entry_count, 1);
    assert_eq!(report.bookmark_count, 1);
    let xml = document_xml(&mut document);
    assert_eq!(
        toc_entry_signatures(&xml),
        [(
            "End fragment".to_owned(),
            "TOC1".to_owned(),
            Some("_Toc1".to_owned()),
            None
        )]
    );
    assert!(xml.contains(r#"w:id="52" w:name="_Toc1""#), "{xml}");
    let target = document
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("_Toc1"))
        .expect("generated hyperlink has a surviving target");
    assert_eq!(target.text(), "End fragment");
    let original = document
        .bookmarks()
        .into_iter()
        .find(|bookmark| bookmark.name() == Some("wholeEnd"))
        .expect("the original whole-paragraph bookmark remains matched");
    assert_eq!(original.text(), "End fragment");
}

#[test]
fn toc_rebuild_repairs_boundary_bookmarks_when_links_and_page_fields_are_disabled() {
    let body = r#"
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="61" w:name="wholeBegin"/><w:r><w:t>Begin fragment</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o &quot;1-1&quot; \n &quot;1-1&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:bookmarkEnd w:id="61"/></w:p>
        <w:p><w:r><w:t>old cache</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="62" w:name="wholeEnd"/><w:r><w:t>old boundary</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>End fragment</w:t></w:r><w:bookmarkEnd w:id="62"/></w:p>
        <w:p><w:hyperlink w:anchor="wholeBegin"><w:r><w:t>existing link</w:t></w:r></w:hyperlink><w:fldSimple w:instr="PAGEREF wholeEnd"><w:r><w:t>stored page</w:t></w:r></w:fldSimple></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    let report = document.rebuild_toc().unwrap();
    assert_eq!(report.entry_count, 2);
    assert_eq!(report.bookmark_count, 0);
    let xml = document_xml(&mut document);
    assert_eq!(
        toc_entry_signatures(&xml),
        [
            ("Begin fragment".to_owned(), "TOC1".to_owned(), None, None),
            ("End fragment".to_owned(), "TOC1".to_owned(), None, None)
        ]
    );
    assert!(xml.contains(r#"w:anchor="wholeBegin""#), "{xml}");
    assert!(xml.contains("PAGEREF wholeEnd"), "{xml}");
    let bookmarks = document.bookmarks();
    assert_eq!(
        bookmarks
            .iter()
            .find(|bookmark| bookmark.name() == Some("wholeBegin"))
            .expect("begin bookmark remains matched")
            .text(),
        "Begin fragment"
    );
    assert_eq!(
        bookmarks
            .iter()
            .find(|bookmark| bookmark.name() == Some("wholeEnd"))
            .expect("end bookmark remains matched")
            .text(),
        "End fragment"
    );
}

#[test]
fn toc_rebuild_accepts_a_content_control_nested_inside_an_accepted_revision() {
    for wrapper in ["ins", "moveTo"] {
        let body = format!(
            r#"
            <w:p><w:{wrapper} w:id="94" w:author="Ada"><w:sdt><w:sdtContent>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
            </w:sdtContent></w:sdt></w:{wrapper}></w:p>
            <w:p><w:r><w:t>old</w:t></w:r></w:p>
            <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Nested owner</w:t></w:r></w:p>
            "#
        );
        let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

        assert_eq!(document.rebuild_toc().unwrap().entry_count, 1);
        assert_eq!(
            toc_entry_signatures(&document_xml(&mut document))[0].0,
            "Nested owner\t1"
        );
    }
}

#[test]
fn toc_rebuild_layouts_controls_nested_inside_insertions() {
    let long_title = "nested ".repeat(900);
    let body = format!(
        r#"
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>old</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:ins w:id="95" w:author="Ada"><w:sdt><w:sdtContent><w:r><w:t>{long_title}</w:t></w:r></w:sdtContent></w:sdt></w:ins></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Later nested heading</w:t></w:r></w:p>
        "#
    );
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    assert_eq!(document.rebuild_toc().unwrap().entry_count, 2);
    let entries = toc_entry_signatures(&document_xml(&mut document));
    assert_eq!(entries[0].0, format!("{long_title}\t3"));
    assert_eq!(entries[1].0, "Later nested heading\t5");
}

#[test]
fn document_properties_variables_and_author_use_package_values() {
    let body = r#"
        <w:p><w:fldSimple w:instr="DOCPROPERTY ClientTier"><w:r><w:t>stored tier</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="DOCVARIABLE Region"><w:r><w:t>stored region</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="AUTHOR"><w:r><w:t>stored author</w:t></w:r></w:fldSimple></w:p>
    "#;
    let settings = r#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:docVars><q:docVar q:name="Region" q:val="North"/></q:docVars></q:settings>"#;
    let custom = r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ClientTier"><vt:lpwstr>Gold</vt:lpwstr></property></Properties>"#;
    let mut document =
        document_with_field_parts(&wrap_word_body(body), Some(settings), Some(custom));
    document.set_author("Package Author");
    let before = document.to_bytes().unwrap();
    let outcomes = document
        .evaluate_fields(&FieldEvaluationContext::default())
        .unwrap()
        .into_iter()
        .map(|evaluation| evaluation.outcome)
        .collect::<Vec<_>>();
    assert_eq!(
        outcomes,
        [
            FieldOutcome::Resolved("Gold".to_owned()),
            FieldOutcome::Resolved("North".to_owned()),
            FieldOutcome::Resolved("Package Author".to_owned()),
        ]
    );
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn resolved_values_drive_date_pictures_and_empty_merge_omits_affixes() {
    let body = r#"
        <w:p><w:fldSimple w:instr="DOCPROPERTY EventDate \@ &quot;MMMM d, yyyy&quot;"><w:r><w:t>stored property</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD MergeDate \@ &quot;yyyy-MM-dd&quot;"><w:r><w:t>stored merge date</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD Empty \b &quot;Dear &quot; \f &quot;!&quot;"><w:r><w:t>stored empty</w:t></w:r></w:fldSimple></w:p>
    "#;
    let custom = r#"<p:Properties xmlns:p="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:v="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><p:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="EventDate"><v:filetime>2025-12-14T21:07:05Z</v:filetime></p:property></p:Properties>"#;
    let document = document_with_field_parts(&wrap_word_body(body), None, Some(custom));
    let context = FieldEvaluationContext {
        now: Some(FieldDateTime {
            year: 1999,
            month: 1,
            day: 1,
            hour: 0,
            minute: 0,
            second: 0,
        }),
        merge_fields: BTreeMap::from([
            ("MergeDate".to_owned(), "2026-01-02".to_owned()),
            ("Empty".to_owned(), String::new()),
        ]),
        ..FieldEvaluationContext::default()
    };
    assert_eq!(
        document
            .evaluate_fields(&context)
            .unwrap()
            .into_iter()
            .map(|result| result.outcome)
            .collect::<Vec<_>>(),
        [
            FieldOutcome::Resolved("December 14, 2025".to_owned()),
            FieldOutcome::Resolved("2026-01-02".to_owned()),
            FieldOutcome::Resolved(String::new()),
        ]
    );
}

#[test]
fn foreign_custom_property_lookalikes_are_preserved_but_not_evaluated() {
    let body = r#"<w:p><w:fldSimple w:instr="DOCPROPERTY ClientTier"><w:r><w:t>stored tier</w:t></w:r></w:fldSimple></w:p>"#;
    let foreign = r#"<x:Properties xmlns:x="urn:not-custom-properties" xmlns:v="urn:not-variant-types"><x:property x:fmtid="foreign" x:pid="2" x:name="ClientTier"><v:lpwstr>Injected</v:lpwstr></x:property></x:Properties>"#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, Some(foreign));
    assert!(matches!(
        document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap()[0]
            .outcome,
        FieldOutcome::KeepStored { .. }
    ));
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    assert_eq!(
        package.get_part("/metadata/producer-properties.xml"),
        Some(foreign.as_bytes())
    );
}

#[test]
fn nested_switch_and_positional_fields_keep_source_order_indices() {
    let body = concat!(
        r#"<w:p>"#,
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
        r#"<w:r><w:instrText xml:space="preserve">MERGEFIELD \b </w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
        r#"<w:r><w:instrText>AUTHOR</w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
        r#"<w:r><w:t>stored author</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        r#"<w:r><w:instrText xml:space="preserve"> </w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
        r#"<w:r><w:instrText>MERGEFIELD Name</w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
        r#"<w:r><w:t>stored name</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
        r#"<w:r><w:t>stored outer</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        r#"</w:p>"#,
    );
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    document.set_author("Package Author");
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([("Name".to_owned(), "Ada".to_owned())]),
        ..FieldEvaluationContext::default()
    };
    let evaluations = document.evaluate_fields(&context).unwrap();
    assert_eq!(
        evaluations
            .iter()
            .map(|evaluation| (evaluation.field_index, evaluation.instruction.as_str()))
            .collect::<Vec<_>>(),
        [(0, r"MERGEFIELD \b"), (1, "AUTHOR"), (2, "MERGEFIELD Name"),]
    );
    assert_eq!(
        evaluations[1].outcome,
        FieldOutcome::Resolved("Package Author".to_owned())
    );
    assert_eq!(
        evaluations[2].outcome,
        FieldOutcome::Resolved("Ada".to_owned())
    );

    assert_eq!(document.update_fields(&context).unwrap(), 3);
    let updated = document.evaluate_fields(&context).unwrap();
    assert_eq!(updated[0].cached_result, "stored outer");
    assert_eq!(updated[1].cached_result, "Package Author");
    assert_eq!(updated[2].cached_result, "Ada");
    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let reopened = reopened.evaluate_fields(&context).unwrap();
    assert_eq!(reopened[1].cached_result, "Package Author");
    assert_eq!(reopened[2].cached_result, "Ada");
}

#[test]
fn nested_if_operand_keeps_its_position_after_cache_updates() {
    let body = concat!(
        r#"<w:p><w:bookmarkStart w:id="7" w:name="destination"/><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>"#,
        r#"<w:p>"#,
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
        r#"<w:r><w:instrText xml:space="preserve">IF </w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>"#,
        r#"<w:r><w:instrText>REF destination</w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
        r#"<w:r><w:t>old reference</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        r#"<w:r><w:instrText xml:space="preserve"> = &quot;x&quot; &quot;yes&quot; &quot;no&quot;</w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
        r#"<w:r><w:t>old outcome</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        r#"</w:p>"#,
    );
    let context = FieldEvaluationContext::default();
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let initial = document.evaluate_fields(&context).unwrap();
    assert_eq!(initial[0].outcome, FieldOutcome::Resolved("yes".to_owned()));
    assert_eq!(initial[1].outcome, FieldOutcome::Resolved("x".to_owned()));

    assert_eq!(document.update_fields(&context).unwrap(), 2);
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(
        xml.find("REF destination").unwrap() < xml.find(" = &quot;x&quot;").unwrap(),
        "{xml}"
    );

    let reopened = Document::from_bytes(&saved).unwrap();
    let reopened = reopened.evaluate_fields(&context).unwrap();
    assert_eq!(reopened[0].cached_result, "yes");
    assert_eq!(
        reopened[0].outcome,
        FieldOutcome::Resolved("yes".to_owned())
    );
    assert_eq!(reopened[1].cached_result, "x");
}

#[test]
fn missing_context_and_unsupported_fields_keep_their_cached_display() {
    let body = r#"
        <w:p><w:fldSimple w:instr="DATE"><w:r><w:t>stored date</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="UNKNOWN"><w:r><w:t>stored unknown</w:t></w:r></w:fldSimple></w:p>
    "#;
    let document = document_with_field_parts(&wrap_word_body(body), None, None);
    let results = document
        .evaluate_fields(&FieldEvaluationContext::default())
        .unwrap();
    assert_eq!(results[0].cached_result, "stored date");
    assert_eq!(
        results[0].outcome,
        FieldOutcome::KeepStored {
            diagnostic: "DATE requires an explicit date and time, stored display retained"
                .to_owned(),
        }
    );
    assert_eq!(results[1].cached_result, "stored unknown");
    assert_eq!(
        results[1].outcome,
        FieldOutcome::KeepStored {
            diagnostic: "field UNKNOWN is unsupported, stored display retained".to_owned(),
        }
    );
}

#[test]
fn field_update_policies_produce_the_expected_result_cache_and_dirty_flag() {
    let body = concat!(
        r#"<w:p><w:fldSimple w:instr="MERGEFIELD Name" w:dirty="1"><w:r><w:t>stored name</w:t></w:r></w:fldSimple></w:p>"#,
        r#"<w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>7</w:t></w:r></w:fldSimple></w:p>"#,
    );
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([("Name".to_owned(), "Ada".to_owned())]),
        ..FieldEvaluationContext::default()
    };

    let mut on_demand = document_with_field_parts(&wrap_word_body(body), None, None);
    assert_eq!(on_demand.update_fields(&context).unwrap(), 2);
    let updated = document_xml(&mut on_demand);
    assert!(updated.contains(r#"w:instr="MERGEFIELD Name" w:dirty="0""#));
    assert!(updated.contains("<w:t>Ada</w:t>"));
    assert!(updated.contains(r#"w:instr="PAGE" w:dirty="1""#));
    assert!(updated.contains("<w:t>7</w:t>"));

    let mut on_save = document_with_field_parts(&wrap_word_body(body), None, None);
    let bytes = on_save.to_bytes_with_field_updates(&context).unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    let outcomes = reopened.evaluate_fields(&context).unwrap();
    assert_eq!(outcomes[0].cached_result, "Ada");
    assert_eq!(outcomes[1].cached_result, "7");

    let mut file_save = document_with_field_parts(&wrap_word_body(body), None, None);
    let path = std::env::temp_dir().join(format!(
        "rdocx-f162-field-update-{}.docx",
        std::process::id()
    ));
    file_save.save_with_field_updates(&path, &context).unwrap();
    let reopened = Document::open(&path).unwrap();
    std::fs::remove_file(&path).unwrap();
    let outcomes = reopened.evaluate_fields(&context).unwrap();
    assert_eq!(outcomes[0].cached_result, "Ada");
    assert_eq!(outcomes[1].cached_result, "7");
}

#[test]
fn unsupported_fields_keep_their_cached_result_when_updates_run() {
    let body = concat!(
        r#"<w:p><w:fldSimple w:instr="UNKNOWN"><w:r><w:t>producer display</w:t></w:r></w:fldSimple></w:p>"#,
        r#"<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>stored date</w:t></w:r></w:fldSimple></w:p>"#,
    );
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    assert_eq!(
        document
            .update_fields(&FieldEvaluationContext::default())
            .unwrap(),
        2
    );
    let xml = document_xml(&mut document);
    assert!(xml.contains("<w:t>producer display</w:t>"));
    assert!(xml.contains("<w:t>stored date</w:t>"));
    assert_eq!(xml.matches(r#"w:dirty="1""#).count(), 2);
}

#[test]
fn ordinary_save_leaves_cached_field_results_and_dirty_flags_alone() {
    let body = concat!(
        r#"<w:p xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:fldSimple q:instr="MERGEFIELD Name" q:dirty="on" data="producer"><q:r><q:t>stored</q:t><q:producer/></q:r></q:fldSimple></w:p>"#,
        r#"<w:p><w:fldSimple w:instr="PAGE" w:dirty="off"><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>"#,
    );
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    let xml = document_xml(&mut document);
    assert!(xml.contains(r#"q:dirty="on" data="producer""#));
    assert!(xml.contains("<q:t>stored</q:t><q:producer/>"));
    assert!(xml.contains(r#"w:dirty="off""#));
    assert!(xml.contains("<w:t>9</w:t>"));
}

#[test]
fn field_update_failure_leaves_document_bytes_unchanged() {
    let body = r#"<w:p><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>"#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let before = document.to_bytes().unwrap();
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([("Name".to_owned(), "invalid\0xml".to_owned())]),
        ..FieldEvaluationContext::default()
    };

    assert!(document.update_fields(&context).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn typed_story_fields_have_stable_order_and_physical_parts_are_evaluated_once() {
    let body = r#"
        <w:p><w:fldSimple w:instr="MERGEFIELD Main"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:tbl><w:tblPr/><w:tblGrid/><w:sdt><w:sdtPr><w:richText/></w:sdtPr><w:sdtContent><w:tr><w:tc><w:p><w:fldSimple w:instr="MERGEFIELD TableControl"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr></w:sdtContent></w:sdt><w:tr><w:tc><w:tcPr/><w:sdt><w:sdtPr><w:richText/></w:sdtPr><w:sdtContent><w:p><w:fldSimple w:instr="MERGEFIELD CellControl"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:sdtContent></w:sdt><w:p><w:fldSimple w:instr="MERGEFIELD Table"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr></w:tbl>
        <w:sdt><w:sdtPr><w:richText/></w:sdtPr><w:sdtContent><w:p><w:fldSimple w:instr="MERGEFIELD Control"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:sdtContent></w:sdt>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
            .unwrap();
    let header = r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:fldSimple w:instr="MERGEFIELD Header"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:hdr>"#;
    let earlier_header = r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:fldSimple w:instr="MERGEFIELD EarlierHeader"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:hdr>"#;
    let orphan_header = r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:fldSimple w:instr="MERGEFIELD Orphan"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:hdr>"#;
    let footer = r#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:fldSimple w:instr="MERGEFIELD Footer"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:ftr>"#;
    let footnotes = r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="2"><w:p><w:fldSimple w:instr="MERGEFIELD Footnote"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:footnote></w:footnotes>"#;
    let endnotes = r#"<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="2"><w:p><w:fldSimple w:instr="MERGEFIELD Endnote"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:endnote></w:endnotes>"#;
    package.set_part("/word/header1.xml", header.as_bytes().to_vec());
    package.set_part("/word/header2.xml", earlier_header.as_bytes().to_vec());
    package.set_part("/word/orphan-header.xml", orphan_header.as_bytes().to_vec());
    package.set_part("/word/footer1.xml", footer.as_bytes().to_vec());
    package.set_part("/word/footnotes.xml", footnotes.as_bytes().to_vec());
    package.set_part("/word/notes/end-stream.xml", endnotes.as_bytes().to_vec());
    let (header_id, duplicate_header_id, earlier_header_id, footer_id) = {
        let relationships = package.get_or_create_part_rels("/word/document.xml");
        let header_id = relationships.add(oxml_opc::relationship::rel_types::HEADER, "header1.xml");
        let duplicate_header_id =
            relationships.add(oxml_opc::relationship::rel_types::HEADER, "header1.xml");
        let earlier_header_id =
            relationships.add(oxml_opc::relationship::rel_types::HEADER, "header2.xml");
        relationships.add(
            oxml_opc::relationship::rel_types::HEADER,
            "orphan-header.xml",
        );
        let footer_id = relationships.add(oxml_opc::relationship::rel_types::FOOTER, "footer1.xml");
        relationships.add(
            oxml_opc::relationship::rel_types::FOOTNOTES,
            "footnotes.xml",
        );
        relationships.add(
            oxml_opc::relationship::rel_types::ENDNOTES,
            "notes/end-stream.xml",
        );
        (header_id, duplicate_header_id, earlier_header_id, footer_id)
    };
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="{earlier_header_id}"/></w:sectPr></w:pPr></w:p><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/><w:headerReference w:type="even" r:id="{duplicate_header_id}"/><w:footerReference w:type="default" r:id="{footer_id}"/></w:sectPr></w:body></w:document>"#
        )
        .into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(bytes.get_ref()).unwrap();
    let names = [
        "Main",
        "TableControl",
        "CellControl",
        "Table",
        "Control",
        "EarlierHeader",
        "Header",
        "Footer",
        "Footnote",
        "Endnote",
    ];
    let context = FieldEvaluationContext {
        merge_fields: names
            .iter()
            .map(|name| ((*name).to_owned(), format!("{name} value")))
            .collect(),
        ..FieldEvaluationContext::default()
    };
    let results = document.evaluate_fields(&context).unwrap();
    assert_eq!(
        results
            .iter()
            .map(|result| result.field_index)
            .collect::<Vec<_>>(),
        (0..10).collect::<Vec<_>>()
    );
    assert_eq!(
        results
            .into_iter()
            .map(|result| result.outcome)
            .collect::<Vec<_>>(),
        names
            .iter()
            .map(|name| FieldOutcome::Resolved(format!("{name} value")))
            .collect::<Vec<_>>()
    );

    assert_eq!(document.update_fields(&context).unwrap(), 10);
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    for (part_name, expected_values) in [
        ("/word/document.xml", &names[..5]),
        ("/word/header2.xml", &names[5..6]),
        ("/word/header1.xml", &names[6..7]),
        ("/word/footer1.xml", &names[7..8]),
        ("/word/footnotes.xml", &names[8..9]),
        ("/word/notes/end-stream.xml", &names[9..10]),
    ] {
        let xml = String::from_utf8(package.get_part(part_name).unwrap().to_vec()).unwrap();
        for name in expected_values {
            assert!(xml.contains(&format!("<w:t>{name} value</w:t>")), "{xml}");
        }
    }
    let orphan = String::from_utf8(
        package
            .get_part("/word/orphan-header.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(orphan.contains("<w:t>stored</w:t>"));
}

#[test]
fn package_story_updates_preserve_aliased_and_unmodelled_part_boundaries() {
    let mut document = document_with_field_parts(&wrap_word_body(""), None, None);
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
            .unwrap();
    let word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let header = format!(
        r#"<q:hdr xmlns:q="{word_namespace}" xmlns:x="urn:producer"><x:headerBefore data="A"/><q:p data="header-paragraph"><q:fldSimple q:instr="MERGEFIELD Header" q:dirty="on" x:token="header"><q:r><q:rPr><q:b/></q:rPr><q:t>stored header</q:t><x:insideHeader/></q:r></q:fldSimple></q:p><x:headerAfter data="B"/></q:hdr>"#,
    );
    let footer = format!(
        r#"<q:ftr xmlns:q="{word_namespace}" xmlns:x="urn:producer"><x:footerBefore/><q:p><q:fldSimple q:instr="MERGEFIELD Footer"><q:r><q:t>stored footer</q:t></q:r></q:fldSimple></q:p><x:footerAfter/></q:ftr>"#,
    );
    let endnotes = format!(
        r#"<q:endnotes xmlns:q="{word_namespace}" xmlns:x="urn:producer"><x:rootBefore/><q:endnote q:id="2" x:token="note"><x:noteBefore/><q:p data="endnote-paragraph"><q:fldSimple q:instr="MERGEFIELD Endnote"><q:r><q:rPr><q:i/></q:rPr><q:t>stored endnote</q:t><x:insideEndnote/></q:r></q:fldSimple></q:p><x:noteAfter/></q:endnote><x:rootAfter/></q:endnotes>"#,
    );
    package.set_part("/word/header-preserve.xml", header.as_bytes().to_vec());
    package.set_part("/word/footer-preserve.xml", footer.as_bytes().to_vec());
    package.set_part("/word/endnotes-preserve.xml", endnotes.as_bytes().to_vec());
    let (header_id, footer_id) = {
        let relationships = package.get_or_create_part_rels("/word/document.xml");
        let header_id = relationships.add(
            oxml_opc::relationship::rel_types::HEADER,
            "header-preserve.xml",
        );
        let footer_id = relationships.add(
            oxml_opc::relationship::rel_types::FOOTER,
            "footer-preserve.xml",
        );
        relationships.add(
            oxml_opc::relationship::rel_types::ENDNOTES,
            "endnotes-preserve.xml",
        );
        (header_id, footer_id)
    };
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="{word_namespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/><w:footerReference w:type="default" r:id="{footer_id}"/></w:sectPr></w:body></w:document>"#,
        )
        .into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(bytes.get_ref()).unwrap();
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([
            ("Header".to_owned(), "updated header".to_owned()),
            ("Footer".to_owned(), "updated footer".to_owned()),
            ("Endnote".to_owned(), "updated endnote".to_owned()),
        ]),
        ..FieldEvaluationContext::default()
    };

    assert_eq!(document.update_fields(&context).unwrap(), 3);
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let header =
        std::str::from_utf8(package.get_part("/word/header-preserve.xml").unwrap()).unwrap();
    let footer =
        std::str::from_utf8(package.get_part("/word/footer-preserve.xml").unwrap()).unwrap();
    let endnotes =
        std::str::from_utf8(package.get_part("/word/endnotes-preserve.xml").unwrap()).unwrap();

    assert!(header.contains(r#"<x:headerBefore data="A"/>"#), "{header}");
    assert!(header.contains(r#"x:token="header""#), "{header}");
    assert!(header.contains(r#"<q:rPr><q:b/></q:rPr>"#), "{header}");
    assert!(header.contains("<x:insideHeader/>"), "{header}");
    assert!(header.contains("updated header"), "{header}");
    assert!(
        header.find("<x:headerBefore").unwrap()
            < header.find(r#"<q:p data="header-paragraph""#).unwrap()
            && header.find(r#"<q:p data="header-paragraph""#).unwrap()
                < header.find("<x:headerAfter").unwrap(),
        "{header}"
    );
    assert!(footer.contains("updated footer"), "{footer}");
    assert!(
        footer.find("<x:footerBefore").unwrap() < footer.find("<q:p>").unwrap()
            && footer.find("<q:p>").unwrap() < footer.find("<x:footerAfter").unwrap(),
        "{footer}"
    );
    for preserved in [
        "<x:rootBefore/>",
        r#"x:token="note""#,
        "<x:noteBefore/>",
        r#"<q:rPr><q:i/></q:rPr>"#,
        "<x:insideEndnote/>",
        "<x:noteAfter/>",
        "<x:rootAfter/>",
    ] {
        assert!(
            endnotes.contains(preserved),
            "missing {preserved}: {endnotes}"
        );
    }
    assert!(endnotes.contains("updated endnote"), "{endnotes}");
    assert!(
        endnotes.find("<x:rootBefore").unwrap() < endnotes.find(r#"<q:endnote q:id="2""#).unwrap()
            && endnotes.find("<x:noteBefore").unwrap()
                < endnotes.find(r#"<q:p data="endnote-paragraph""#).unwrap()
            && endnotes.find(r#"<q:p data="endnote-paragraph""#).unwrap()
                < endnotes.find("<x:noteAfter").unwrap()
            && endnotes.find(r#"<q:endnote q:id="2""#).unwrap()
                < endnotes.find("<x:rootAfter").unwrap(),
        "{endnotes}"
    );
}

#[test]
fn pretty_printed_complex_package_field_updates_in_place() {
    let mut document = document_with_field_parts(&wrap_word_body(""), None, None);
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
            .unwrap();
    let word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let header = format!(
        r#"<q:hdr xmlns:q="{word_namespace}" xmlns:x="urn:producer">
  <x:before/>
  <q:p>
    <q:r><q:fldChar q:fldCharType="begin" q:dirty="on"/></q:r>
    <q:r><q:instrText>MERGEFIELD Pretty</q:instrText></q:r>
    <q:r><q:fldChar q:fldCharType="separate"/></q:r>
    <q:r><q:rPr><q:b/></q:rPr><q:t>stored pretty</q:t><x:inside/></q:r>
    <q:r><q:fldChar q:fldCharType="end"/></q:r>
  </q:p>
  <x:after/>
</q:hdr>"#,
    );
    package.set_part("/word/header-pretty.xml", header.as_bytes().to_vec());
    let header_id = package.get_or_create_part_rels("/word/document.xml").add(
        oxml_opc::relationship::rel_types::HEADER,
        "header-pretty.xml",
    );
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="{word_namespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/></w:sectPr></w:body></w:document>"#,
        )
        .into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(bytes.get_ref()).unwrap();
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([("Pretty".to_owned(), "updated pretty".to_owned())]),
        ..FieldEvaluationContext::default()
    };

    assert_eq!(document.update_fields(&context).unwrap(), 1);
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let header = std::str::from_utf8(package.get_part("/word/header-pretty.xml").unwrap()).unwrap();
    assert!(header.contains("updated pretty"), "{header}");
    assert!(header.contains("\n    <q:r>"), "{header}");
    assert!(header.contains("<q:rPr><q:b/></q:rPr>"), "{header}");
    assert!(header.contains("<x:inside/>"), "{header}");
    assert!(
        header.find("<x:before").unwrap() < header.find("updated pretty").unwrap()
            && header.find("updated pretty").unwrap() < header.find("<x:after").unwrap(),
        "{header}"
    );
}

#[test]
fn complex_field_updates_preserve_inter_run_comments_and_instructions() {
    let mut seed = document_with_field_parts(&wrap_word_body(""), None, None);
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
    let word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let header_field = concat!(
        r#"<q:r><q:fldChar q:fldCharType="begin" q:dirty="on"/></q:r>"#,
        "<!-- header-before-instruction -->",
        r#"<q:r><q:instrText>MERGEFIELD Producer</q:instrText></q:r>"#,
        "<?producer header-before-separator?>",
        r#"<q:r><q:fldChar q:fldCharType="separate"/></q:r>"#,
        "<!-- header-before-result -->",
        r#"<q:r><q:t>stored header</q:t></q:r>"#,
        "<?producer header-before-end?>",
        r#"<q:r><q:fldChar q:fldCharType="end"/></q:r>"#,
    );
    let header = format!(r#"<q:hdr xmlns:q="{word_namespace}"><q:p>{header_field}</q:p></q:hdr>"#,);
    package.set_part("/word/header-events.xml", header.into_bytes());
    let header_id = package.get_or_create_part_rels("/word/document.xml").add(
        oxml_opc::relationship::rel_types::HEADER,
        "header-events.xml",
    );
    let main_field = concat!(
        r#"<w:r><w:fldChar w:fldCharType="begin" w:dirty="on"/></w:r>"#,
        "<!-- main-before-instruction -->",
        r#"<w:r><w:instrText>MERGEFIELD Producer</w:instrText></w:r>"#,
        "<?producer main-before-separator?>",
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
        "<!-- main-before-result -->",
        r#"<w:r><w:t>stored main</w:t></w:r>"#,
        "<?producer main-before-end?>",
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
    );
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="{word_namespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p>{main_field}</w:p><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/></w:sectPr></w:body></w:document>"#,
        )
        .into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(bytes.get_ref()).unwrap();
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([("Producer".to_owned(), "updated producer".to_owned())]),
        ..FieldEvaluationContext::default()
    };

    assert_eq!(document.update_fields(&context).unwrap(), 2);
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let main = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    let header = std::str::from_utf8(package.get_part("/word/header-events.xml").unwrap()).unwrap();
    for preserved in [
        "<!-- main-before-instruction -->",
        "<?producer main-before-separator?>",
        "<!-- main-before-result -->",
        "<?producer main-before-end?>",
    ] {
        assert!(main.contains(preserved), "missing {preserved}: {main}");
    }
    for preserved in [
        "<!-- header-before-instruction -->",
        "<?producer header-before-separator?>",
        "<!-- header-before-result -->",
        "<?producer header-before-end?>",
    ] {
        assert!(header.contains(preserved), "missing {preserved}: {header}");
    }
    assert_eq!(main.matches("updated producer").count(), 1, "{main}");
    assert_eq!(header.matches("updated producer").count(), 1, "{header}");
}

#[test]
fn hyperlink_field_trivia_survives_package_update() {
    let mut seed = document_with_field_parts(&wrap_word_body(""), None, None);
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
    let word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let header = format!(
        concat!(
            r#"<q:hdr xmlns:q="{0}"><q:p><q:hyperlink q:anchor="destination">"#,
            "\n  ",
            r#"<q:r><q:fldChar q:fldCharType="begin" q:dirty="on"/></q:r>"#,
            "\n  <!-- hyperlink-package-before-instruction -->",
            r#"<q:r><q:instrText>MERGEFIELD Link</q:instrText></q:r>"#,
            "<?producer hyperlink-package-before-separator?>",
            r#"<q:r><q:fldChar q:fldCharType="separate"/></q:r>"#,
            "\n  <!-- hyperlink-package-before-result -->",
            r#"<q:r><q:t>stored link</q:t></q:r>"#,
            "<?producer hyperlink-package-before-end?>",
            r#"<q:r><q:fldChar q:fldCharType="end"/></q:r>"#,
            "\n",
            r#"</q:hyperlink></q:p></q:hdr>"#,
        ),
        word_namespace,
    );
    package.set_part("/word/header-hyperlink-events.xml", header.into_bytes());
    let header_id = package.get_or_create_part_rels("/word/document.xml").add(
        oxml_opc::relationship::rel_types::HEADER,
        "header-hyperlink-events.xml",
    );
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="{word_namespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/></w:sectPr></w:body></w:document>"#,
        )
        .into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(bytes.get_ref()).unwrap();
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([("Link".to_owned(), "updated link".to_owned())]),
        ..FieldEvaluationContext::default()
    };

    assert_eq!(document.update_fields(&context).unwrap(), 1);
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let header = std::str::from_utf8(
        package
            .get_part("/word/header-hyperlink-events.xml")
            .unwrap(),
    )
    .unwrap();
    for preserved in [
        "<!-- hyperlink-package-before-instruction -->",
        "<?producer hyperlink-package-before-separator?>",
        "<!-- hyperlink-package-before-result -->",
        "<?producer hyperlink-package-before-end?>",
    ] {
        assert!(header.contains(preserved), "missing {preserved}: {header}");
    }
    assert_eq!(header.matches("updated link").count(), 1, "{header}");
    assert!(header.contains("<q:hyperlink"), "{header}");
}

#[test]
fn identical_aliased_same_run_nested_fields_update_and_reopen() {
    let word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let nested = concat!(
        r#"<q:fldChar q:fldCharType="begin" q:dirty="on"/>"#,
        r#"<q:instrText>MERGEFIELD Same</q:instrText>"#,
        r#"<q:fldChar q:fldCharType="separate"/>"#,
        r#"<q:t>stored same run</q:t>"#,
        r#"<q:fldChar q:fldCharType="end"/>"#,
    );
    let body = format!(
        concat!(
            r#"<w:p><q:r xmlns:q="{0}" xmlns:x="urn:producer"><x:before/>"#,
            r#"<q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText xml:space="preserve">IF </q:instrText>"#,
            "{1}",
            r#"<q:instrText xml:space="preserve"> = </q:instrText>"#,
            "{1}",
            r#"<q:instrText xml:space="preserve"> &quot;yes&quot; &quot;no&quot;</q:instrText>"#,
            r#"<q:fldChar q:fldCharType="separate"/><q:t>stored outer</q:t>"#,
            r#"<q:fldChar q:fldCharType="end"/><x:after/></q:r></w:p>"#,
        ),
        word_namespace, nested,
    );
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([("Same".to_owned(), "updated same run".to_owned())]),
        ..FieldEvaluationContext::default()
    };
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    assert_eq!(document.update_fields(&context).unwrap(), 3);
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert_eq!(xml.matches("updated same run").count(), 2, "{xml}");
    assert!(!xml.contains("stored same run"), "{xml}");
    assert!(xml.contains("<x:before/>"), "{xml}");
    assert!(xml.contains("<x:after/>"), "{xml}");

    let reopened = Document::from_bytes(&saved).unwrap();
    let evaluations = reopened.evaluate_fields(&context).unwrap();
    assert_eq!(evaluations.len(), 3);
    assert_eq!(evaluations[0].cached_result, "yes");
    assert_eq!(evaluations[1].cached_result, "updated same run");
    assert_eq!(evaluations[2].cached_result, "updated same run");
}

#[test]
fn shared_boundary_run_nested_fields_update_in_order_and_reopen() {
    let word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let body = format!(
        concat!(
            r#"<w:p><q:r xmlns:q="{0}"><q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText xml:space="preserve">IF </q:instrText></q:r>"#,
            r#"<q:r xmlns:q="{0}"><q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText>MERGEFIELD First</q:instrText></q:r>"#,
            r#"<q:r xmlns:q="{0}" xmlns:x="urn:producer"><q:fldChar q:fldCharType="separate"/><q:t>stored first</q:t><q:fldChar q:fldCharType="end"/><x:between/><q:instrText xml:space="preserve"> = </q:instrText><q:fldChar q:fldCharType="begin" q:dirty="on"/></q:r>"#,
            r#"<q:r xmlns:q="{0}"><q:instrText>MERGEFIELD Second</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored second</q:t><q:fldChar q:fldCharType="end"/><q:instrText xml:space="preserve"> &quot;yes&quot; &quot;no&quot;</q:instrText></q:r>"#,
            r#"<q:r xmlns:q="{0}"><q:fldChar q:fldCharType="separate"/><q:t>stored outer</q:t><q:fldChar q:fldCharType="end"/></q:r></w:p>"#,
        ),
        word_namespace,
    );
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([
            ("First".to_owned(), "updated first boundary".to_owned()),
            ("Second".to_owned(), "updated second boundary".to_owned()),
        ]),
        ..FieldEvaluationContext::default()
    };
    let mut document = document_with_field_parts(&wrap_word_body(&body), None, None);

    assert_eq!(document.update_fields(&context).unwrap(), 3);
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert_eq!(xml.matches("updated first boundary").count(), 1, "{xml}");
    assert_eq!(xml.matches("updated second boundary").count(), 1, "{xml}");
    assert!(xml.contains("<x:between/>"), "{xml}");
    assert!(
        xml.find("updated first boundary").unwrap() < xml.find("<x:between/>").unwrap()
            && xml.find("<x:between/>").unwrap() < xml.find("updated second boundary").unwrap(),
        "{xml}"
    );

    let reopened = Document::from_bytes(&saved).unwrap();
    let evaluations = reopened.evaluate_fields(&context).unwrap();
    assert_eq!(evaluations.len(), 3);
    assert_eq!(evaluations[0].cached_result, "no");
    assert_eq!(evaluations[1].cached_result, "updated first boundary");
    assert_eq!(evaluations[2].cached_result, "updated second boundary");
}

#[test]
fn package_field_patching_skips_opaque_lookalikes_and_maps_identical_aliases() {
    let mut document = document_with_field_parts(&wrap_word_body(""), None, None);
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
            .unwrap();
    let word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let field =
        r#"<q:fldSimple q:instr="MERGEFIELD Same"><q:r><q:t>stored same</q:t></q:r></q:fldSimple>"#;
    let header = format!(
        r#"<q:hdr xmlns:q="{word_namespace}" xmlns:x="urn:producer"><x:opaque>{field}</x:opaque><q:p data="first"><x:innerOpaque>{field}</x:innerOpaque>{field}</q:p><q:p data="second">{field}</q:p></q:hdr>"#,
    );
    package.set_part("/word/header-lookalike.xml", header.as_bytes().to_vec());
    let header_id = package.get_or_create_part_rels("/word/document.xml").add(
        oxml_opc::relationship::rel_types::HEADER,
        "header-lookalike.xml",
    );
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="{word_namespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/></w:sectPr></w:body></w:document>"#,
        )
        .into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(bytes.get_ref()).unwrap();
    let context = FieldEvaluationContext {
        merge_fields: BTreeMap::from([("Same".to_owned(), "updated same".to_owned())]),
        ..FieldEvaluationContext::default()
    };

    assert_eq!(document.update_fields(&context).unwrap(), 2);
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let header =
        std::str::from_utf8(package.get_part("/word/header-lookalike.xml").unwrap()).unwrap();
    assert!(
        header.contains(&format!(r#"<x:opaque>{field}</x:opaque>"#)),
        "{header}"
    );
    assert!(
        header.contains(&format!(r#"<x:innerOpaque>{field}</x:innerOpaque>"#)),
        "{header}"
    );
    assert_eq!(header.matches("updated same").count(), 2, "{header}");
    assert_eq!(header.matches("stored same").count(), 2, "{header}");
    assert!(
        header.find(r#"<q:p data="first""#).unwrap()
            < header.find(r#"<q:p data="second""#).unwrap(),
        "{header}"
    );
}

#[test]
fn styleref_searches_the_approved_direction_and_scope() {
    let body = r#"
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>First heading</w:t></w:r></w:p>
        <w:p><w:fldSimple w:instr="STYLEREF &quot;Heading 1&quot;"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="STYLEREF &quot;Heading 1&quot; \p"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Last heading</w:t></w:r></w:p>
        <w:p><w:fldSimple w:instr="STYLEREF &quot;Heading 1&quot; \l"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
    "#;
    let document = document_with_field_parts(&wrap_word_body(body), None, None);
    let outcomes = document
        .evaluate_fields(&FieldEvaluationContext::default())
        .unwrap()
        .into_iter()
        .map(|result| result.outcome)
        .collect::<Vec<_>>();
    assert_eq!(
        outcomes,
        [
            FieldOutcome::Resolved("First heading".to_owned()),
            FieldOutcome::Resolved("First heading above".to_owned()),
            FieldOutcome::Resolved("Last heading".to_owned()),
        ]
    );
}

#[test]
fn date_time_filename_mergefield_and_includetext_use_only_explicit_context() {
    let body = r#"
        <w:p><w:fldSimple w:instr="DATE \@ &quot;yyyy-MM-dd&quot;"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="TIME \@ &quot;HH:mm:ss&quot;"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="FILENAME \p"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD Name \b &quot;Dear &quot; \f &quot;!&quot; \m \v"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="INCLUDETEXT &quot;chapter.docx&quot;"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="INCLUDETEXT &quot;chapter.docx&quot; &quot;summary&quot; \!"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
    "#;
    let document = document_with_field_parts(&wrap_word_body(body), None, None);
    let context = FieldEvaluationContext {
        now: Some(FieldDateTime {
            year: 2025,
            month: 12,
            day: 14,
            hour: 21,
            minute: 7,
            second: 5,
        }),
        file_name: Some("report.docx".to_owned()),
        file_path: Some("/safe/input/report.docx".to_owned()),
        merge_fields: BTreeMap::from([("Name".to_owned(), "Ada".to_owned())]),
        included_text: BTreeMap::from([
            ("chapter.docx".to_owned(), "Whole chapter".to_owned()),
            (
                "chapter.docx#summary".to_owned(),
                "Chapter summary".to_owned(),
            ),
        ]),
        merge_record_number: None,
        merge_sequence_number: None,
    };
    assert_eq!(
        document
            .evaluate_fields(&context)
            .unwrap()
            .into_iter()
            .map(|result| result.outcome)
            .collect::<Vec<_>>(),
        [
            FieldOutcome::Resolved("2025-12-14".to_owned()),
            FieldOutcome::Resolved("21:07:05".to_owned()),
            FieldOutcome::Resolved("/safe/input/report.docx".to_owned()),
            FieldOutcome::Resolved("Dear Ada!".to_owned()),
            FieldOutcome::Resolved("Whole chapter".to_owned()),
            FieldOutcome::Resolved("Chapter summary".to_owned()),
        ]
    );
}

#[test]
fn both_revision_views_render_and_accepted_matches_resolved_document() {
    let xml = wrap_word_body(
        r#"<w:p><w:r><w:t>ordinary </w:t></w:r><w:ins w:id="1" w:author="Ada"><w:r><w:t>inserted </w:t></w:r></w:ins><w:del w:id="2" w:author="Ben"><w:r><w:delText>deleted </w:delText></w:r></w:del><w:moveFrom w:id="3" w:author="Cy"><w:r><w:t>old </w:t></w:r></w:moveFrom><w:moveTo w:id="4" w:author="Dee"><w:r><w:t>moved</w:t></w:r></w:moveTo></w:p>"#,
    );
    let document = document_with_content_controls(&xml);
    let accepted = document
        .render_page_to_png_deterministic_with_options(
            0,
            150.0,
            RenderOptions {
                revision_view: RevisionView::Accepted,
            },
        )
        .unwrap()
        .expect("accepted page");
    let tracked = document
        .render_page_to_png_deterministic_with_options(
            0,
            150.0,
            RenderOptions {
                revision_view: RevisionView::Tracked,
            },
        )
        .unwrap()
        .expect("tracked page");
    assert_ne!(accepted, tracked);

    let mut resolved = document_with_content_controls(&xml);
    assert_eq!(resolved.accept_all().unwrap(), 4);
    let resolved = resolved
        .render_page_to_png_deterministic(0, 150.0)
        .unwrap()
        .expect("resolved page");
    assert_eq!(accepted, resolved);
}

#[test]
fn default_render_methods_keep_the_accepted_view() {
    let xml = wrap_word_body(
        r#"<w:p><w:r><w:t>ordinary </w:t></w:r><w:ins w:id="1" w:author="Ada"><w:r><w:t>accepted</w:t></w:r></w:ins><w:del w:id="2" w:author="Ben"><w:r><w:delText>omitted</w:delText></w:r></w:del></w:p>"#,
    );
    let document = document_with_content_controls(&xml);
    let options = RenderOptions::default();

    assert_eq!(
        document.to_pdf().unwrap(),
        document.to_pdf_with_options(options).unwrap()
    );
    assert_eq!(
        document.to_pdf_deterministic().unwrap(),
        document.to_pdf_deterministic_with_options(options).unwrap()
    );
    assert_eq!(
        document.render_page_to_png(0, 96.0).unwrap(),
        document
            .render_page_to_png_with_options(0, 96.0, options)
            .unwrap()
    );
    assert_eq!(
        document.render_page_to_png_deterministic(0, 96.0).unwrap(),
        document
            .render_page_to_png_deterministic_with_options(0, 96.0, options)
            .unwrap()
    );
    assert_eq!(
        document.render_all_pages(96.0).unwrap(),
        document
            .render_all_pages_with_options(96.0, options)
            .unwrap()
    );
    assert_eq!(
        format!("{:?}", document.layout_page(0).unwrap()),
        format!(
            "{:?}",
            document.layout_page_with_options(0, options).unwrap()
        )
    );
}

#[test]
fn native_image_export_ranges_are_zero_based_and_keep_selected_order() {
    let mut document = Document::new();
    document.add_paragraph("first");
    document.add_paragraph("second").page_break_before(true);
    document.add_paragraph("third").page_break_before(true);

    let output = document
        .render_pages_deterministic(
            &[2, 0],
            RasterOptions {
                dpi: 72.0,
                format: RasterFormat::Jpeg { quality: 80 },
            },
        )
        .expect("selected pages render");
    let RasterOutput::SeparatePages(pages) = output else {
        panic!("JPEG output should be separate pages");
    };
    assert_eq!(pages.len(), 2);
    assert!(pages.iter().all(|page| page.starts_with(&[0xff, 0xd8])));
    assert_ne!(pages[0], pages[1]);

    let duplicate = document.render_pages_deterministic(
        &[0, 0],
        RasterOptions {
            dpi: 72.0,
            format: RasterFormat::Png {
                transparent_background: false,
            },
        },
    );
    assert!(duplicate.is_err());
}

#[test]
fn downstream_renderers_can_traverse_the_complete_public_layout_result() {
    let mut document = Document::new();
    document.add_paragraph("public layout integration");

    let result = document.layout().expect("public layout should succeed");
    let mut saw_resolvable_run = false;
    for page in &result.layout.pages {
        oxml_layout::walk(&page.elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Text(run) = element {
                let font = result
                    .layout
                    .fonts
                    .iter()
                    .find(|font| font.id == run.font_id)
                    .expect("glyph run font should resolve");
                assert!(!font.data.is_empty());
                if let Some(source) = run.source {
                    assert!(result.source_node(source.node).is_some());
                }
                saw_resolvable_run = true;
            }
        });
    }
    assert!(saw_resolvable_run);
}

#[test]
fn caller_font_layout_options_select_the_tracked_revision_projection() {
    let xml = wrap_word_body(
        r#"<w:p><w:r><w:t>ordinary </w:t></w:r><w:ins w:id="1" w:author="Ada"><w:r><w:t>accepted</w:t></w:r></w:ins><w:del w:id="2" w:author="Ben"><w:r><w:delText> omitted</w:delText></w:r></w:del></w:p>"#,
    );
    let document = document_with_content_controls(&xml);
    let (family, bytes) = oxml_layout::bundled_fonts::bundled_font_data()[0];
    let accepted = document
        .layout_with_fonts(&[(family, bytes)])
        .expect("accepted caller-font layout should succeed");
    let tracked = document
        .layout_with_fonts_and_options(
            &[(family, bytes)],
            RenderOptions {
                revision_view: RevisionView::Tracked,
            },
        )
        .expect("tracked caller-font layout should succeed");

    let visible_text = |result: &rdocx_layout::WordLayoutResult| {
        let mut text = String::new();
        for page in &result.layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| {
                if let oxml_layout::PositionedElement::Text(run) = element {
                    text.push_str(&run.text);
                }
            });
        }
        text
    };

    let accepted_text = visible_text(&accepted);
    let tracked_text = visible_text(&tracked);
    assert_eq!(accepted.revision_view, RevisionView::Accepted);
    assert_eq!(tracked.revision_view, RevisionView::Tracked);
    assert!(accepted_text.contains("ordinary accepted"));
    assert!(!accepted_text.contains("omitted"));
    assert!(tracked_text.contains("ordinary accepted omitted"));
}

fn document_with_content_controls(document_xml: &str) -> Document {
    document_with_bound_content_controls(document_xml, None)
}

#[test]
fn legacy_horizontal_rule_package_reopens_with_exact_raw_xml() {
    let raw = br#"<w:pict><v:rect o:hr="true"/></w:pict>"#;
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p xmlns:v="urn:schemas-microsoft-com:vml"><w:r xmlns:o="urn:schemas-microsoft-com:office:office"><w:t>before</w:t>{}<w:t>after</w:t></w:r></w:p></w:body></w:document>"#,
        std::str::from_utf8(raw).unwrap()
    );
    let mut document = document_with_content_controls(&xml);

    let snapshot = |document: &Document| {
        document
            .paragraph(0)
            .unwrap()
            .run(0)
            .unwrap()
            .items()
            .map(|item| match item {
                RunItemRef::Text(text) => format!("text:{text}"),
                RunItemRef::LegacyHorizontalRule(rule) => {
                    format!("rule:{}", std::str::from_utf8(rule.raw_xml()).unwrap())
                }
                RunItemRef::UnsupportedXml(bytes) => {
                    format!("unsupported:{}", std::str::from_utf8(bytes).unwrap())
                }
                _ => panic!("unexpected run item"),
            })
            .collect::<Vec<_>>()
    };
    let expected = [
        "text:before".to_owned(),
        format!("rule:{}", std::str::from_utf8(raw).unwrap()),
        "text:after".to_owned(),
    ];
    assert_eq!(snapshot(&document), expected);

    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    assert_eq!(snapshot(&reopened), expected);
}

#[test]
fn legacy_horizontal_rule_classification_participates_in_run_equality() {
    let raw = r#"<w:pict><v:rect o:hr="true"/></w:pict>"#;
    let parsed_run = |vml_namespace: &str| {
        let body = body_from_xml(&format!(
            r#"<w:p xmlns:v="{vml_namespace}" xmlns:o="urn:schemas-microsoft-com:office:office"><w:r>{raw}</w:r></w:p>"#
        ));
        let BodyContent::Paragraph(paragraph) = &body.content[0] else {
            panic!("expected paragraph");
        };
        paragraph.runs[0].clone()
    };

    let legacy = parsed_run("urn:schemas-microsoft-com:vml");
    let foreign = parsed_run("urn:foreign");
    assert!(CT_R::raw_child_is_legacy_horizontal_rule(
        legacy.extra_xml_positions[0]
    ));
    assert!(!CT_R::raw_child_is_legacy_horizontal_rule(
        foreign.extra_xml_positions[0]
    ));
    assert_ne!(legacy, foreign);
}

#[test]
fn namespace_classification_metadata_exists_only_for_raw_children() {
    let body = body_from_xml(
        r#"<w:p xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:r><w:t>ordinary</w:t></w:r><w:r><w:pict><v:rect o:hr="true"/></w:pict></w:r></w:p>"#,
    );
    let BodyContent::Paragraph(paragraph) = &body.content[0] else {
        panic!("expected paragraph");
    };

    assert!(paragraph.runs[0].extra_xml.is_empty());
    assert!(paragraph.runs[0].extra_xml_positions.is_empty());
    assert_eq!(paragraph.runs[1].extra_xml_positions.len(), 1);
    assert!(CT_R::raw_child_is_legacy_horizontal_rule(
        paragraph.runs[1].extra_xml_positions[0]
    ));
}

fn ordered_reader_fixture() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <q:body xmlns:x="urn:foreign">
            <q:p>
              <q:r><q:t>first</q:t><x:run-raw/></q:r>
              <x:paragraph-raw/>
              <q:commentRangeStart q:id="2"/>
              <q:sdt><q:sdtContent><q:r><q:t>control</q:t></q:r></q:sdtContent></q:sdt>
              <q:hyperlink r:id="rId3">
                <q:r><q:t>link before</q:t></q:r><x:inside-link-before/>
                <q:ins q:id="6" q:author="Bea"><q:r><q:t>linked revision</q:t></q:r></q:ins>
                <x:inside-link-after/><q:r><q:t>link after</q:t></q:r>
              </q:hyperlink>
              <q:ins q:id="4" q:author="Ada"><q:r><q:t>inserted</q:t></q:r></q:ins>
              <q:bookmarkStart q:id="5" q:name="target"/>
              <q:r><q:t>last</q:t></q:r>
              <q:bookmarkEnd q:id="5"/>
              <q:commentRangeEnd q:id="2"/>
            </q:p>
            <x:body-raw><x:child/></x:body-raw>
            <q:tbl><q:tr><q:tc>
              <q:p><q:r><q:t>cell</q:t></q:r></q:p>
              <x:cell-raw/>
              <q:tbl><q:tr><q:tc><q:p><q:r><q:t>nested</q:t></q:r></q:p></q:tc></q:tr></q:tbl>
              <q:sdt><q:sdtContent><q:p><q:r><q:t>cell control</q:t></q:r></q:p></q:sdtContent></q:sdt>
            </q:tc></q:tr></q:tbl>
            <q:sdt><q:sdtContent><q:p><q:r><q:t>body control</q:t></q:r></q:p></q:sdtContent></q:sdt>
          </q:body>
        </q:document>"#
}

#[test]
fn ordered_reader_items_keep_every_direct_child_and_preserved_boundary() {
    let xml = ordered_reader_fixture();
    let document = document_with_content_controls(xml);

    let body = document
        .body_items()
        .map(|item| match item {
            BodyItemRef::Paragraph(paragraph) => format!("paragraph:{}", paragraph.text()),
            BodyItemRef::Table(_) => "table".to_owned(),
            BodyItemRef::ContentControl(control) => format!("control:{}", control.text()),
            BodyItemRef::UnsupportedXml(raw) => {
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
        })
        .collect::<Vec<_>>();
    assert_eq!(
        body,
        [
            "paragraph:firstcontrollink beforelink afterlast",
            "raw:<x:body-raw xmlns:x=\"urn:foreign\"><x:child/></x:body-raw>",
            "table",
            "control:body control",
        ]
    );

    let paragraph = document.paragraph(0).unwrap();
    let paragraph_items = paragraph
        .items()
        .map(|item| match item {
            ParagraphItemRef::Run(run) => format!("run:{}", run.text()),
            ParagraphItemRef::Hyperlink(link) => {
                let children = link
                    .items()
                    .map(|item| match item {
                        HyperlinkItemRef::Run(run) => format!("run:{}", run.text()),
                        HyperlinkItemRef::Revision(revision) => {
                            format!("revision:{}", revision.id())
                        }
                        HyperlinkItemRef::UnsupportedXml(raw) => {
                            format!("raw:{}", std::str::from_utf8(raw).unwrap())
                        }
                        _ => panic!("unexpected hyperlink item"),
                    })
                    .collect::<Vec<_>>();
                format!(
                    "hyperlink:{}:{}",
                    link.relationship_id().unwrap(),
                    children.join(",")
                )
            }
            ParagraphItemRef::ContentControl(control) => format!("control:{}", control.text()),
            ParagraphItemRef::Revision(revision) => format!("revision:{}", revision.id()),
            ParagraphItemRef::CommentRangeStart(id) => format!("comment-start:{id}"),
            ParagraphItemRef::CommentRangeEnd(id) => format!("comment-end:{id}"),
            ParagraphItemRef::BookmarkStart { id, name } => {
                format!("bookmark-start:{}:{}", id.unwrap(), name.unwrap())
            }
            ParagraphItemRef::BookmarkEnd { id } => {
                format!("bookmark-end:{}", id.unwrap())
            }
            ParagraphItemRef::UnsupportedXml(raw) => {
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
            _ => panic!("unexpected paragraph item"),
        })
        .collect::<Vec<_>>();
    assert_eq!(
        paragraph_items,
        [
            "run:first",
            "raw:<x:paragraph-raw/>",
            "comment-start:2",
            "control:control",
            "hyperlink:rId3:run:link before,raw:<x:inside-link-before/>,revision:6,raw:<x:inside-link-after/>,run:link after",
            "revision:4",
            "bookmark-start:5:target",
            "run:last",
            "bookmark-end:5",
            "comment-end:2",
        ]
    );

    let first_run = paragraph.run(0).unwrap();
    let run_items = first_run
        .items()
        .map(|item| match item {
            RunItemRef::Text(text) => format!("text:{text}"),
            RunItemRef::DeletedText(text) => format!("deleted:{text}"),
            RunItemRef::Tab => "tab".to_owned(),
            RunItemRef::Break(BreakKind::Line) => "break:line".to_owned(),
            RunItemRef::Break(BreakKind::Page) => "break:page".to_owned(),
            RunItemRef::Break(BreakKind::Column) => "break:column".to_owned(),
            RunItemRef::UnsupportedXml(raw) => {
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
            _ => panic!("unexpected run item"),
        })
        .collect::<Vec<_>>();
    assert_eq!(run_items, ["text:first", "raw:<x:run-raw/>"]);

    let table = document.table(0).unwrap();
    let cell = table.cell(0, 0).unwrap();
    let cell_items = cell
        .items()
        .map(|item| match item {
            CellItemRef::Paragraph(paragraph) => format!("paragraph:{}", paragraph.text()),
            CellItemRef::Table(table) => format!("table:{}", table.row_count()),
            CellItemRef::ContentControl(control) => format!("control:{}", control.text()),
            CellItemRef::UnsupportedXml(raw) => {
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
            _ => panic!("unexpected cell item"),
        })
        .collect::<Vec<_>>();
    assert_eq!(
        cell_items,
        [
            "paragraph:cell",
            "raw:<x:cell-raw xmlns:x=\"urn:foreign\"/>",
            "table:1",
            "control:cell control",
        ]
    );
}

#[test]
fn ordered_reader_items_resolve_aliases_without_flattening_containers() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:w="urn:not-word" xmlns:x="urn:foreign">
        <q:body><q:p><w:r><w:t>foreign</w:t></w:r><q:sdt><q:sdtContent>
          <q:r><q:t>nested</q:t></q:r>
        </q:sdtContent></q:sdt><q:r><q:t>direct</q:t></q:r></q:p></q:body>
      </q:document>"#;
    let document = document_with_content_controls(xml);
    let items = document
        .paragraph(0)
        .unwrap()
        .items()
        .map(|item| match item {
            ParagraphItemRef::Run(run) => format!("run:{}", run.text()),
            ParagraphItemRef::ContentControl(control) => format!("control:{}", control.text()),
            ParagraphItemRef::UnsupportedXml(raw) => {
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
            _ => "other".to_owned(),
        })
        .collect::<Vec<_>>();
    assert_eq!(
        items,
        [
            "raw:<w:r><w:t>foreign</w:t></w:r>",
            "control:nested",
            "run:direct",
        ]
    );
    assert_eq!(document.paragraph(0).unwrap().run_count(), 1);
}

#[test]
fn ordered_run_items_expose_every_retained_typed_fact() {
    let xml = wrap_word_body(
        r#"<w:p><w:r><x:r0 xmlns:x="urn:foreign"/><w:t>first</w:t><x:r1 xmlns:x="urn:foreign"/><w:delText>removed</w:delText><x:r2 xmlns:x="urn:foreign"/><w:tab/><x:r3 xmlns:x="urn:foreign"/><w:br/><x:r4 xmlns:x="urn:foreign"/><w:br w:type="page"/><x:r5 xmlns:x="urn:foreign"/><w:br w:type="column"/><x:r6 xmlns:x="urn:foreign"/><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:extent cx="1" cy="2"/><wp:docPr id="11" name="Picture" descr="Alternative"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing><x:r7 xmlns:x="urn:foreign"/><w:footnoteReference w:id="7"/><x:r8 xmlns:x="urn:foreign"/><w:endnoteReference w:id="8"/><x:r9 xmlns:x="urn:foreign"/><w:commentReference w:id="9"/><x:r10 xmlns:x="urn:foreign"/></w:r><w:fldSimple w:instr=" PAGE " w:dirty="true"><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p>"#,
    );
    let document = document_with_content_controls(&xml);
    let paragraph = document.paragraph(0).unwrap();
    let items = paragraph
        .run(0)
        .unwrap()
        .items()
        .map(|item| match item {
            RunItemRef::Text(text) => format!("text:{text}"),
            RunItemRef::DeletedText(text) => format!("deleted:{text}"),
            RunItemRef::Tab => "tab".to_owned(),
            RunItemRef::Break(BreakKind::Line) => "break:line".to_owned(),
            RunItemRef::Break(BreakKind::Page) => "break:page".to_owned(),
            RunItemRef::Break(BreakKind::Column) => "break:column".to_owned(),
            RunItemRef::FootnoteReference(id) => format!("footnote:{id}"),
            RunItemRef::EndnoteReference(id) => format!("endnote:{id}"),
            RunItemRef::CommentReference(id) => format!("comment:{id}"),
            RunItemRef::Drawing(drawing) => format!(
                "drawing:{}:{}:{}:{}:{}:{}:{}",
                drawing.is_inline(),
                drawing.is_anchor(),
                drawing.relationship_id().unwrap(),
                drawing.name().unwrap(),
                drawing.description().unwrap(),
                drawing.width().unwrap().to_emu(),
                drawing.height().unwrap().to_emu(),
            ),
            RunItemRef::Field(_) => panic!("field belongs to the next run"),
            RunItemRef::UnsupportedXml(raw) => {
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
            _ => panic!("unexpected run item"),
        })
        .collect::<Vec<_>>();
    assert_eq!(
        items,
        [
            "raw:<x:r0 xmlns:x=\"urn:foreign\"/>",
            "text:first",
            "raw:<x:r1 xmlns:x=\"urn:foreign\"/>",
            "deleted:removed",
            "raw:<x:r2 xmlns:x=\"urn:foreign\"/>",
            "tab",
            "raw:<x:r3 xmlns:x=\"urn:foreign\"/>",
            "break:line",
            "raw:<x:r4 xmlns:x=\"urn:foreign\"/>",
            "break:page",
            "raw:<x:r5 xmlns:x=\"urn:foreign\"/>",
            "break:column",
            "raw:<x:r6 xmlns:x=\"urn:foreign\"/>",
            "drawing:true:false:rId9:Picture:Alternative:1:2",
            "raw:<x:r7 xmlns:x=\"urn:foreign\"/>",
            "footnote:7",
            "raw:<x:r8 xmlns:x=\"urn:foreign\"/>",
            "endnote:8",
            "raw:<x:r9 xmlns:x=\"urn:foreign\"/>",
            "comment:9",
            "raw:<x:r10 xmlns:x=\"urn:foreign\"/>",
        ]
    );

    let field_run = paragraph.run(1).unwrap();
    let mut field_items = field_run.items();
    let field = field_items.next().unwrap();
    let RunItemRef::Field(field) = field else {
        panic!("expected field fact");
    };
    assert_eq!(field.instruction(), "PAGE");
    assert_eq!(field.name(), "PAGE");
    assert_eq!(field.cached_result(), "3");
    assert_eq!(field.dirty(), Some(true));
    assert!(field_items.next().is_none());
}

#[test]
fn ordered_run_items_keep_raw_children_before_properties() {
    let xml = wrap_word_body(
        r#"<w:p><w:r><x:before xmlns:x="urn:foreign"/><w:rPr><w:b/></w:rPr><w:t>typed</w:t></w:r></w:p>"#,
    );
    let document = document_with_content_controls(&xml);
    let items = document
        .paragraph(0)
        .unwrap()
        .run(0)
        .unwrap()
        .items()
        .map(|item| match item {
            RunItemRef::Text(text) => format!("text:{text}"),
            RunItemRef::UnsupportedXml(raw) => {
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
            _ => "other".to_owned(),
        })
        .collect::<Vec<_>>();
    assert_eq!(
        items,
        ["raw:<x:before xmlns:x=\"urn:foreign\"/>", "text:typed"]
    );
}

#[test]
fn modeled_unsupported_body_facts_do_not_invent_raw_xml() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:x="urn:inherited"><q:body><x:custom><x:child/></x:custom>
        <q:sdt><q:sdtContent><q:p><q:r><q:t>inside</q:t></q:r></q:p></q:sdtContent></q:sdt>
      </q:body></q:document>"#;
    let document = document_with_content_controls(xml);
    let facts = document.body_content().collect::<Vec<_>>();
    let BodyContentRef::UnsupportedXml(raw) = &facts[0] else {
        panic!("expected raw unsupported fact");
    };
    assert_eq!(raw.qualified_name(), Some("x:custom"));
    assert_eq!(raw.local_name(), "custom");
    assert_eq!(raw.namespace_uri(), Some("urn:inherited"));
    assert_eq!(
        raw.raw_xml(),
        Some(b"<x:custom><x:child/></x:custom>".as_slice())
    );
    assert!(raw.has_child_content());

    let BodyContentRef::UnsupportedXml(modeled) = &facts[1] else {
        panic!("expected modeled unsupported fact");
    };
    assert_eq!(modeled.qualified_name(), Some("w:sdt"));
    assert_eq!(modeled.local_name(), "sdt");
    assert_eq!(
        modeled.namespace_uri(),
        Some("http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    );
    assert_eq!(modeled.raw_xml(), None);
}

#[test]
fn unsupported_body_facts_respect_shadowed_conventional_prefixes() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:w="urn:foreign"><q:body><w:producer/></q:body></q:document>"#;
    let mut document = document_with_content_controls(xml);
    let source = ordered_reader_snapshot(&document);
    let fact = document.body_content().next().unwrap();
    let BodyContentRef::UnsupportedXml(fact) = fact else {
        panic!("expected unsupported fact");
    };
    assert_eq!(fact.qualified_name(), Some("w:producer"));
    assert_eq!(fact.namespace_uri(), Some("urn:foreign"));

    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let reopened_fact = reopened.body_content().next().unwrap();
    let BodyContentRef::UnsupportedXml(reopened_fact) = reopened_fact else {
        panic!("expected unsupported fact");
    };
    assert_eq!(reopened_fact.namespace_uri(), Some("urn:foreign"));
    assert_eq!(ordered_reader_snapshot(&reopened), source);
}

#[test]
fn body_local_canonical_prefix_shadows_survive_save_and_reopen() {
    for prefix in ["w", "r", "mc"] {
        let namespace = format!("urn:foreign-{prefix}");
        let raw = format!("<{prefix}:producer/>");
        let xml = format!(
            r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                 <q:body xmlns:{prefix}="{namespace}">{raw}<q:p><q:r><q:t>typed</q:t></q:r></q:p></q:body>
               </q:document>"#,
        );
        let mut document = document_with_content_controls(&xml);
        let source = ordered_reader_snapshot(&document);
        let source_fact = document.body_content().next().unwrap();
        let BodyContentRef::UnsupportedXml(source_fact) = source_fact else {
            panic!("expected unsupported fact");
        };
        assert_eq!(source_fact.namespace_uri(), Some(namespace.as_str()));

        let saved = document.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(saved_xml.contains(&format!(r#"xmlns:{prefix}="{namespace}""#)));
        assert!(saved_xml.contains(&raw));

        let reopened = Document::from_bytes(&saved).unwrap();
        assert_eq!(reopened.paragraph(0).unwrap().text(), "typed");
        let reopened_fact = reopened.body_content().next().unwrap();
        let BodyContentRef::UnsupportedXml(reopened_fact) = reopened_fact else {
            panic!("expected unsupported fact");
        };
        assert_eq!(reopened_fact.namespace_uri(), Some(namespace.as_str()));
        assert_eq!(ordered_reader_snapshot(&reopened), source);
    }
}

#[test]
fn root_and_body_prefix_collisions_keep_distinct_scopes() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:x="urn:root"><x:background/><q:body xmlns:x="urn:body">
                    <x:producer/><q:p><q:r><q:t>typed</q:t></q:r></q:p>
                  </q:body></q:document>"#;
    let mut document = document_with_content_controls(xml);
    let source = ordered_reader_snapshot(&document);
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let saved_xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    let document_start = saved_xml.find(":document").unwrap();
    let root_end = saved_xml[document_start..].find('>').unwrap() + document_start;
    let body_start = saved_xml.find(":body").unwrap();
    let body_end = saved_xml[body_start..].find('>').unwrap() + body_start;
    assert!(saved_xml[..root_end].contains(r#"xmlns:x="urn:root""#));
    assert!(saved_xml[body_start..body_end].contains(r#"xmlns:x="urn:body""#));
    assert!(saved_xml.contains("<x:background/>"));
    assert!(saved_xml.contains("<x:producer/>"));

    let reopened = Document::from_bytes(&saved).unwrap();
    let reopened_fact = reopened.body_content().next().unwrap();
    let BodyContentRef::UnsupportedXml(reopened_fact) = reopened_fact else {
        panic!("expected unsupported fact");
    };
    assert_eq!(reopened_fact.namespace_uri(), Some("urn:body"));
    assert_eq!(ordered_reader_snapshot(&reopened), source);
}

#[test]
fn unsupported_body_facts_resolve_body_local_namespaces() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <q:body xmlns:x="urn:body"><x:producer/></q:body></q:document>"#;
    let mut document = document_with_content_controls(xml);
    let fact = document.body_content().next().unwrap();
    let BodyContentRef::UnsupportedXml(fact) = fact else {
        panic!("expected unsupported fact");
    };
    assert_eq!(fact.qualified_name(), Some("x:producer"));
    assert_eq!(fact.namespace_uri(), Some("urn:body"));

    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let saved_xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(saved_xml.contains(r#"xmlns:x="urn:body""#));
    let reopened = Document::from_bytes(&saved).unwrap();
    let reopened_fact = reopened.body_content().next().unwrap();
    let BodyContentRef::UnsupportedXml(reopened_fact) = reopened_fact else {
        panic!("expected unsupported fact");
    };
    assert_eq!(reopened_fact.namespace_uri(), Some("urn:body"));
}

#[test]
fn unsupported_body_facts_decode_local_namespace_uris() {
    let xml = wrap_word_body(r#"<x:producer xmlns:x="urn:a&amp;b"/>"#);
    let mut document = document_with_content_controls(&xml);
    let fact = document.body_content().next().unwrap();
    let BodyContentRef::UnsupportedXml(fact) = fact else {
        panic!("expected unsupported fact");
    };
    assert_eq!(fact.namespace_uri(), Some("urn:a&b"));

    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let reopened_fact = reopened.body_content().next().unwrap();
    let BodyContentRef::UnsupportedXml(reopened_fact) = reopened_fact else {
        panic!("expected unsupported fact");
    };
    assert_eq!(reopened_fact.namespace_uri(), Some("urn:a&b"));
}

#[test]
fn unsupported_body_facts_resolve_the_implicit_xml_prefix() {
    let xml = wrap_word_body(r#"<xml:producer/>"#);
    let mut document = document_with_content_controls(&xml);
    let fact = document.body_content().next().unwrap();
    let BodyContentRef::UnsupportedXml(fact) = fact else {
        panic!("expected unsupported fact");
    };
    assert_eq!(fact.qualified_name(), Some("xml:producer"));
    assert_eq!(
        fact.namespace_uri(),
        Some("http://www.w3.org/XML/1998/namespace")
    );

    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let BodyContentRef::UnsupportedXml(fact) = reopened.body_content().next().unwrap() else {
        panic!("expected reopened unsupported fact");
    };
    assert_eq!(
        fact.namespace_uri(),
        Some("http://www.w3.org/XML/1998/namespace")
    );
}

#[test]
fn unsupported_body_facts_accept_empty_default_namespace_undeclarations() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <q:body><producer xmlns=""/></q:body></q:document>"#;
    let document = document_with_content_controls(xml);
    let fact = document.body_content().next().unwrap();
    let BodyContentRef::UnsupportedXml(fact) = fact else {
        panic!("expected unsupported fact");
    };
    assert_eq!(fact.qualified_name(), Some("producer"));
    assert_eq!(fact.namespace_uri(), Some(""));
}

#[test]
fn unsupported_body_facts_detect_cdata_and_entity_content() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:x="urn:foreign"><q:body>
                    <x:cdata><![CDATA[visible]]></x:cdata><x:entity>&amp;</x:entity>
                  </q:body></q:document>"#;
    let document = document_with_content_controls(xml);
    let facts = document.body_content().collect::<Vec<_>>();
    assert_eq!(facts.len(), 2);
    for fact in facts {
        let BodyContentRef::UnsupportedXml(fact) = fact else {
            panic!("expected unsupported fact");
        };
        assert!(fact.has_child_content());
    }
}

#[test]
fn nested_foreign_descendants_keep_their_scope_and_bytes_after_reopen() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:w="urn:foreign"><q:body>
        <q:p><w:r><w:t>foreign paragraph</w:t></w:r><q:r><q:t>typed paragraph</q:t></q:r></q:p>
        <q:tbl><q:tr><q:tc><w:p><w:r>foreign cell</w:r></w:p><q:p><q:r><q:t>typed cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl>
        <q:sdt><q:sdtContent><q:p><w:r><w:t>foreign control</w:t></w:r><q:r><q:t>typed control</q:t></q:r></q:p></q:sdtContent></q:sdt>
      </q:body></q:document>"#;
    let mut document = document_with_content_controls(xml);
    let source = ordered_reader_snapshot(&document);
    assert!(matches!(
        document.paragraph(0).unwrap().items().next().unwrap(),
        ParagraphItemRef::UnsupportedXml(raw)
            if raw == b"<w:r><w:t>foreign paragraph</w:t></w:r>"
    ));
    assert!(matches!(
        document.table(0).unwrap().cell(0, 0).unwrap().items().next().unwrap(),
        CellItemRef::UnsupportedXml(raw)
            if raw == b"<w:p><w:r>foreign cell</w:r></w:p>"
    ));
    assert_eq!(document.content_controls().len(), 1);

    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    assert_eq!(
        package.get_part("/word/document.xml").unwrap(),
        xml.as_bytes()
    );
    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(ordered_reader_snapshot(&reopened), source);
}

#[test]
fn hyperlink_and_run_shadows_keep_exposed_raw_bytes_after_reopen() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <q:body><q:p><q:hyperlink xmlns:w="urn:foreign">
          <q:r><q:t>typed</q:t><w:run-producer/></q:r><w:link-producer/>
        </q:hyperlink></q:p></q:body></q:document>"#;
    let mut document = document_with_content_controls(xml);
    let source = ordered_reader_snapshot(&document);
    let paragraph = document.paragraph(0).unwrap();
    let hyperlink = paragraph
        .items()
        .find_map(|item| match item {
            ParagraphItemRef::Hyperlink(hyperlink) => Some(hyperlink),
            _ => None,
        })
        .unwrap();
    let raw = hyperlink
        .items()
        .flat_map(|item| match item {
            HyperlinkItemRef::Run(run) => run
                .items()
                .filter_map(|item| match item {
                    RunItemRef::UnsupportedXml(raw) => Some(raw.to_vec()),
                    _ => None,
                })
                .collect::<Vec<_>>(),
            HyperlinkItemRef::UnsupportedXml(raw) => vec![raw.to_vec()],
            _ => Vec::new(),
        })
        .collect::<Vec<_>>();
    assert_eq!(
        raw,
        [
            b"<w:run-producer/>".to_vec(),
            b"<w:link-producer/>".to_vec(),
        ]
    );

    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    assert_eq!(
        package.get_part("/word/document.xml").unwrap(),
        xml.as_bytes()
    );
    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(ordered_reader_snapshot(&reopened), source);
}

#[test]
fn serializer_prefix_collisions_keep_modeled_drawings_bound_after_reopen() {
    let scopes = [
        r#"xmlns:wp="urn:root-wp" xmlns:a="urn:root-a" xmlns:pic="urn:root-pic" xmlns:c="urn:root-c""#,
        "",
    ];
    for (index, root_scope) in scopes.into_iter().enumerate() {
        let body_scope = if index == 0 {
            ""
        } else {
            r#"xmlns:wp="urn:body-wp" xmlns:a="urn:body-a" xmlns:pic="urn:body-pic" xmlns:c="urn:body-c""#
        };
        let xml = format!(
            r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" {root_scope}>
              <q:body {body_scope}><q:p><q:r><q:drawing><dwp:inline
                xmlns:dwp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                xmlns:da="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:dpic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <dwp:extent cx="1" cy="2"/><dwp:docPr id="11" name="Picture" descr="Alternative"/>
                <da:graphic><da:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <dpic:pic><dpic:blipFill><da:blip rel:embed="rId9"/></dpic:blipFill></dpic:pic>
                </da:graphicData></da:graphic>
              </dwp:inline></q:drawing></q:r></q:p></q:body></q:document>"#,
        );
        let mut document = document_with_content_controls(&xml);
        let source = ordered_reader_snapshot(&document);
        let paragraph = document.paragraph(0).unwrap();
        let drawing = paragraph.run(0).unwrap();
        assert!(matches!(
            drawing.items().next().unwrap(),
            RunItemRef::Drawing(drawing) if drawing.is_inline()
        ));

        let saved = document.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        assert_eq!(
            package.get_part("/word/document.xml").unwrap(),
            xml.as_bytes()
        );
        let reopened = Document::from_bytes(&saved).unwrap();
        assert_eq!(ordered_reader_snapshot(&reopened), source);
    }
}

#[test]
fn every_modeled_container_replays_nested_namespaces_after_modification() {
    let bodies = [
        (
            "paragraph",
            r#"<q:p xmlns:x="urn:paragraph"><x:producer/><q:r><q:t>typed</q:t></q:r></q:p>"#,
        ),
        (
            "table",
            r#"<q:tbl xmlns:x="urn:table"><x:producer/><q:tr><q:tc><q:p/></q:tc></q:tr></q:tbl>"#,
        ),
        (
            "cell",
            r#"<q:tbl><q:tr><q:tc xmlns:x="urn:cell"><x:producer/><q:p/></q:tc></q:tr></q:tbl>"#,
        ),
        (
            "control",
            r#"<q:sdt xmlns:x="urn:control"><x:producer/><q:sdtContent><q:p/></q:sdtContent></q:sdt>"#,
        ),
        (
            "hyperlink",
            r#"<q:p><q:hyperlink xmlns:x="urn:hyperlink"><x:producer/><q:r><q:t>typed</q:t></q:r></q:hyperlink></q:p>"#,
        ),
        (
            "run",
            r#"<q:p><q:r xmlns:x="urn:run"><x:producer/><q:t>typed</q:t></q:r></q:p>"#,
        ),
    ];
    for (owner, body) in bodies {
        let xml = format!(
            r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body>{body}</q:body></q:document>"#,
        );
        let mut unchanged = document_with_content_controls(&xml);
        let source = ordered_reader_snapshot(&unchanged);
        let saved = unchanged.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(saved_xml.contains(&format!(r#"xmlns:x="urn:{owner}""#)));
        assert!(saved_xml.contains("<x:producer"));
        let reopened = Document::from_bytes(&saved).unwrap();
        assert_eq!(
            ordered_reader_snapshot(&reopened),
            source,
            "{owner} facts changed after reopen",
        );

        let mut modified = document_with_content_controls(&xml);
        modified.add_paragraph("changed");
        let saved = modified.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(saved_xml.contains(&format!(r#"xmlns:x="urn:{owner}""#)));
        assert!(saved_xml.contains("<x:producer"));
        let reopened = Document::from_bytes(&saved).unwrap();
        assert!(
            reopened
                .paragraphs()
                .iter()
                .any(|paragraph| paragraph.text() == "changed")
        );
    }
}

#[test]
fn word_namespace_alias_used_by_raw_marker_replays_after_save_and_reopen() {
    let word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    for raw in ["<x:producer/>", "<x:t>raw same-local-name child</x:t>"] {
        let xml = format!(
            r#"<q:document xmlns:q="{word_namespace}"><q:body><q:p xmlns:x="{word_namespace}">{raw}<q:r><q:t>typed</q:t></q:r></q:p></q:body></q:document>"#,
        );

        for modified in [false, true] {
            let mut document = document_with_content_controls(&xml);
            if modified {
                document.add_paragraph("changed");
            }
            let saved = document.to_bytes().unwrap();
            let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
            let saved_xml =
                std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
            assert_namespace_on_raw_owner(
                saved_xml,
                "p",
                &format!(r#"xmlns:x="{word_namespace}""#),
                raw,
            );

            let mut reopened = Document::from_bytes(&saved).unwrap();
            assert!(reopened.paragraph(0).unwrap().items().any(
                |item| matches!(item, ParagraphItemRef::UnsupportedXml(bytes) if bytes == raw.as_bytes())
            ));
            let resaved = reopened.to_bytes().unwrap();
            let package =
                oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&resaved)).unwrap();
            let resaved_xml =
                std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
            assert_namespace_on_raw_owner(
                resaved_xml,
                "p",
                &format!(r#"xmlns:x="{word_namespace}""#),
                raw,
            );
        }
    }

    let typed_only = format!(
        r#"<q:document xmlns:q="{word_namespace}"><q:body><q:p xmlns:x="{word_namespace}"><x:r><x:t>typed alias</x:t></x:r></q:p></q:body></q:document>"#,
    );
    let mut document = document_with_content_controls(&typed_only);
    document.add_paragraph("changed");
    let saved = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(reopened.paragraph(0).unwrap().text(), "typed alias");
    assert_eq!(reopened.paragraph(1).unwrap().text(), "changed");

    let shared_alias = format!(
        r#"<q:document xmlns:q="{word_namespace}"><q:body><q:p xmlns:x="{word_namespace}"><x:producer/><x:r><x:t>shared alias</x:t></x:r></q:p></q:body></q:document>"#,
    );
    let mut document = document_with_content_controls(&shared_alias);
    document.add_paragraph("changed");
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let saved_xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert_namespace_on_raw_owner(
        saved_xml,
        "p",
        &format!(r#"xmlns:x="{word_namespace}""#),
        "<x:producer/>",
    );
    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(reopened.paragraph(0).unwrap().text(), "shared alias");
    assert_eq!(reopened.paragraph(1).unwrap().text(), "changed");
}

#[test]
fn intermediate_raw_shadow_is_safe_but_direct_fixed_prefix_use_fails_closed() {
    for (owner_uri, child_uri) in [("urn:owner", "urn:child"), ("urn:foreign", "urn:foreign")] {
        let raw = format!(
            r#"<x:wrapper xmlns:x="urn:x" xmlns:wp="{child_uri}"><wp:producer/></x:wrapper>"#,
        );
        let safe_xml = format!(
            r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body>
              <q:p xmlns:wp="{owner_uri}">{raw}<q:r><q:t>typed</q:t></q:r></q:p>
            </q:body></q:document>"#,
        );
        let mut safe = document_with_content_controls(&safe_xml);
        safe.add_paragraph("changed");
        let saved = safe.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(saved_xml.contains(&raw), "{saved_xml}");
        if owner_uri != child_uri {
            assert!(!saved_xml.contains(owner_uri), "{saved_xml}");
        }
        let reopened = Document::from_bytes(&saved).unwrap();
        assert!(reopened.paragraph(0).unwrap().items().any(
            |item| matches!(item, ParagraphItemRef::UnsupportedXml(bytes) if bytes == raw.as_bytes())
        ));
        assert_eq!(reopened.paragraph(1).unwrap().text(), "changed");
    }

    let direct_xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body>
      <q:p xmlns:wp="urn:owner"><wp:producer/><q:r><q:t>typed</q:t></q:r></q:p>
    </q:body></q:document>"#;
    let mut direct = document_with_content_controls(direct_xml);
    direct.add_paragraph("changed");
    let error = direct.to_bytes().unwrap_err();
    assert!(error.to_string().contains("shadowed `wp` namespace"));
}

fn assert_namespace_on_raw_owner(xml: &str, owner: &str, declaration: &str, raw: &str) {
    let raw_prefix = raw.strip_suffix("/>").unwrap_or(raw);
    let raw_start = xml.find(raw_prefix).expect("raw producer is present");
    let owner_start = xml[..raw_start]
        .rfind(&format!("<w:{owner}"))
        .expect("modeled owner starts before its raw producer");
    let owner_end = owner_start
        + xml[owner_start..]
            .find('>')
            .expect("modeled owner start tag is complete");
    let raw_end = raw_start
        + xml[raw_start..]
            .find('>')
            .expect("raw producer start tag is complete");
    assert!(
        xml[owner_start..=owner_end].contains(declaration)
            || xml[raw_start..=raw_end].contains(declaration),
        "{declaration} was not retained on the {owner} or its raw child {raw}: {xml}",
    );
}

fn assert_namespace_on_named_owner(xml: &str, owner: &str, declaration: &str, owner_text: &str) {
    let text_start = xml.find(owner_text).expect("named owner text is present");
    let owner_start = xml[..text_start]
        .rfind(&format!("<w:{owner}"))
        .expect("named modeled owner starts before its text");
    let owner_end = owner_start
        + xml[owner_start..]
            .find('>')
            .expect("named modeled owner start tag is complete");
    assert!(
        xml[owner_start..=owner_end].contains(declaration),
        "{declaration} was not replayed on the {owner} containing {owner_text}",
    );
}

#[test]
fn nested_namespace_replay_tracks_owners_across_insert_remove_and_reorder() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body>
      <q:p><q:hyperlink><q:r><q:t>paragraph decoy</q:t></q:r></q:hyperlink></q:p>
      <q:tbl><q:tr><q:tc><q:p><q:r><q:t>table decoy</q:t></q:r></q:p></q:tc></q:tr></q:tbl>
      <q:sdt><q:sdtContent><q:p><q:r><q:t>control decoy</q:t></q:r></q:p></q:sdtContent></q:sdt>
      <q:p xmlns:px="urn:target-paragraph"><px:producer/><q:hyperlink xmlns:hx="urn:target-hyperlink"><hx:producer/><q:r xmlns:rx="urn:target-run"><rx:producer/><q:t>target paragraph</q:t></q:r></q:hyperlink></q:p>
      <q:tbl xmlns:tx="urn:target-table"><tx:producer/><q:tr><q:tc xmlns:cx="urn:target-cell"><cx:producer/><q:sdt xmlns:sx="urn:target-control"><sx:producer/><q:sdtContent><q:p><q:r><q:t>target table</q:t></q:r></q:p></q:sdtContent></q:sdt></q:tc></q:tr></q:tbl>
    </q:body></q:document>"#;
    let owners = [
        ("p", r#"xmlns:px="urn:target-paragraph""#, "<px:producer/>"),
        (
            "hyperlink",
            r#"xmlns:hx="urn:target-hyperlink""#,
            "<hx:producer/>",
        ),
        ("r", r#"xmlns:rx="urn:target-run""#, "<rx:producer/>"),
        ("tbl", r#"xmlns:tx="urn:target-table""#, "<tx:producer/>"),
        ("tc", r#"xmlns:cx="urn:target-cell""#, "<cx:producer/>"),
        ("sdt", r#"xmlns:sx="urn:target-control""#, "<sx:producer/>"),
    ];

    for mutation in ["insert", "remove", "reorder"] {
        let mut document = document_with_content_controls(xml);
        match mutation {
            "insert" => {
                let relationship =
                    document.add_hyperlink_relationship("https://inserted.example.invalid");
                document
                    .insert_paragraph(0, "inserted")
                    .add_hyperlink("inserted link", &relationship);
                document.insert_table(1, 1, 1);
            }
            "remove" => {
                assert!(document.remove_content(0));
                assert!(document.remove_content(0));
                assert!(document.remove_content(0));
            }
            "reorder" => {
                assert!(document.remove_content(0));
                assert!(document.remove_content(0));
                assert!(document.remove_content(0));
                let relationship =
                    document.add_hyperlink_relationship("https://moved.example.invalid");
                document
                    .add_paragraph("moved paragraph decoy")
                    .add_hyperlink("moved link", &relationship);
                document.add_table(1, 1);
            }
            _ => unreachable!(),
        }

        let saved = document
            .to_bytes()
            .unwrap_or_else(|error| panic!("{mutation} namespace replay failed: {error}"));
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        for (owner, declaration, raw) in owners {
            assert_namespace_on_raw_owner(saved_xml, owner, declaration, raw);
        }

        let mut reopened = Document::from_bytes(&saved).unwrap();
        let resaved = reopened.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&resaved)).unwrap();
        let resaved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        for (owner, declaration, raw) in owners {
            assert_namespace_on_raw_owner(resaved_xml, owner, declaration, raw);
        }
    }
}

#[test]
fn duplicate_raw_markers_cannot_replace_removed_paragraph_or_table_owners() {
    let cases = [
        (
            "p",
            r#"<q:p><x:producer/><q:r><q:t>survivor</q:t></q:r></q:p>
               <q:p xmlns:x="urn:target"><x:producer/><q:r><q:t>removed</q:t></q:r></q:p>"#,
        ),
        (
            "tbl",
            r#"<q:tbl><x:producer/><q:tr><q:tc><q:p><q:r><q:t>survivor</q:t></q:r></q:p></q:tc></q:tr></q:tbl>
               <q:tbl xmlns:x="urn:target"><x:producer/><q:tr><q:tc><q:p><q:r><q:t>removed</q:t></q:r></q:p></q:tc></q:tr></q:tbl>"#,
        ),
    ];
    for (owner, body) in cases {
        let xml = format!(
            r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:root"><q:body>{body}</q:body></q:document>"#,
        );
        let mut document = document_with_content_controls(&xml);
        assert!(document.remove_content(1));
        let error = document.to_bytes().unwrap_err();
        assert!(
            error
                .to_string()
                .contains(&format!("retained `{owner}` nested namespace owner")),
            "removed {owner} must not transfer its declaration: {error}",
        );
    }
}

#[test]
fn foreign_namespace_decoy_cannot_replace_a_removed_namespace_owner() {
    for decoy_namespace in ["urn:decoy", "urn:target"] {
        let xml = format!(
            r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body>
      <q:p><x:producer xmlns:x="{decoy_namespace}"/><q:r><q:t>same</q:t></q:r></q:p>
      <q:p xmlns:x="urn:target"><x:producer/><q:r><q:t>same</q:t></q:r></q:p>
    </q:body></q:document>"#
        );
        let mut document = document_with_content_controls(&xml);
        assert!(document.remove_content(1));
        let error = document.to_bytes().unwrap_err();
        assert!(
            error
                .to_string()
                .contains("retained `p` nested namespace owner"),
            "foreign expanded names must keep the removed owner unidentifiable: {error}",
        );
    }
}

#[test]
fn duplicate_scope_markers_keep_the_target_owner_through_reorder_and_modification() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:root"><q:body>
      <q:p><x:producer/><q:r><q:t>scope decoy</q:t></q:r></q:p>
      <q:p xmlns:x="urn:target"><x:producer/><q:r><q:t>scope target</q:t></q:r></q:p>
    </q:body></q:document>"#;
    for mutation in ["reorder", "modify"] {
        let mut document = document_with_content_controls(xml);
        match mutation {
            "reorder" => {
                assert!(document.remove_content(0));
                let decoy = document_with_content_controls(
                    r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:root"><q:body><q:p><x:producer/><q:r><q:t>scope decoy</q:t></q:r></q:p></q:body></q:document>"#,
                );
                document.insert_document(document.content_count(), &decoy);
            }
            "modify" => {
                document.add_paragraph("unrelated modification");
            }
            _ => unreachable!(),
        }

        let saved = document.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert_namespace_on_named_owner(saved_xml, "p", r#"xmlns:x="urn:target""#, "scope target");
        assert_eq!(saved_xml.matches("<x:producer/>").count(), 2);

        let mut reopened = Document::from_bytes(&saved).unwrap();
        let resaved = reopened.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&resaved)).unwrap();
        let resaved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert_namespace_on_named_owner(
            resaved_xml,
            "p",
            r#"xmlns:x="urn:target""#,
            "scope target",
        );
    }
}

#[test]
fn inherited_and_self_bound_namespace_uses_keep_exact_owner_marker_cardinality() {
    let cases = [
        ("sibling elements", r#"<x:a/><x:b xmlns:x="urn:target"/>"#),
        (
            "nested elements",
            r#"<u:outer xmlns:u="urn:opaque"><x:a/></u:outer><u:independent xmlns:u="urn:opaque" xmlns:x="urn:target"><x:b/></u:independent>"#,
        ),
        (
            "namespaced attributes",
            r#"<u:dependent xmlns:u="urn:opaque" x:a="one"/><u:independent xmlns:u="urn:opaque" xmlns:x="urn:target" x:b="two"/>"#,
        ),
    ];
    for (case, raw) in cases {
        let xml = format!(
            r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body>
      <q:p xmlns:x="urn:target">{raw}<q:r><q:t>{case}</q:t></q:r></q:p>
    </q:body></q:document>"#,
        );
        let mut document = document_with_content_controls(&xml);
        document.add_paragraph("unrelated mutation");

        let saved = document
            .to_bytes()
            .unwrap_or_else(|error| panic!("{case} rejected a valid owner: {error}"));
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert_eq!(saved_xml.matches(r#"xmlns:x="urn:target""#).count(), 2);
        assert!(
            saved_xml.contains(raw),
            "{case} raw XML changed: {saved_xml}"
        );

        let mut reopened = Document::from_bytes(&saved).unwrap();
        let resaved = reopened.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&resaved)).unwrap();
        let resaved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert_eq!(resaved_xml.matches(r#"xmlns:x="urn:target""#).count(), 2);
        assert!(
            resaved_xml.contains(raw),
            "{case} raw XML changed: {resaved_xml}"
        );
    }
}

#[test]
fn nested_table_cell_owner_ignores_independent_same_uri_local_binding() {
    let raw = r#"<x:a/><x:b xmlns:x="urn:target"/>"#;
    let stabilized_raw = r#"<x:a xmlns:x="urn:target"/><x:b xmlns:x="urn:target"/>"#;
    let xml = format!(
        r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body>
      <q:tbl><q:tr><q:tc xmlns:x="urn:target">{raw}<q:p><q:r><q:t>nested owner</q:t></q:r></q:p></q:tc></q:tr></q:tbl>
    </q:body></q:document>"#,
    );
    let mut document = document_with_content_controls(&xml);
    document.add_paragraph("unrelated mutation");

    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let saved_xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert_eq!(
        saved_xml.matches(r#"xmlns:x="urn:target""#).count(),
        3,
        "{saved_xml}"
    );
    assert!(
        saved_xml.contains(stabilized_raw),
        "nested owner bindings were not stabilized: {saved_xml}"
    );

    let mut reopened = Document::from_bytes(&saved).unwrap();
    let resaved = reopened.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&resaved)).unwrap();
    let resaved_xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert_eq!(
        resaved_xml.matches(r#"xmlns:x="urn:target""#).count(),
        2,
        "{resaved_xml}"
    );
    assert!(
        resaved_xml.contains(stabilized_raw),
        "nested owner bindings changed after reopen: {resaved_xml}"
    );
}

#[test]
fn materially_changed_duplicate_run_owner_fails_closed() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:root"><q:body><q:p>
      <q:r><x:producer/><q:t>run decoy</q:t></q:r>
      <q:r xmlns:x="urn:target"><x:producer/><q:t>run target</q:t></q:r>
    </q:p></q:body></q:document>"#;
    let mut document = document_with_content_controls(xml);
    document
        .paragraph_mut(0)
        .unwrap()
        .run_mut(1)
        .unwrap()
        .set_text("materially changed");
    let error = document.to_bytes().unwrap_err();
    assert!(
        error
            .to_string()
            .contains("retained `r` nested namespace owner"),
        "changed run identity must fail closed: {error}",
    );
}

#[test]
fn unused_fixed_prefix_declarations_do_not_reject_safe_raw_replay() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body>
      <q:p xmlns:x="urn:producer" xmlns:wp="urn:unused"><x:producer/><q:r><q:t>target</q:t></q:r></q:p>
    </q:body></q:document>"#;
    let mut document = document_with_content_controls(xml);
    document.add_paragraph("changed");
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let saved_xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert_namespace_on_raw_owner(saved_xml, "p", r#"xmlns:x="urn:producer""#, "<x:producer/>");
    assert!(!saved_xml.contains("urn:unused"));
    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(reopened.paragraph(1).unwrap().text(), "changed");
}

#[test]
fn expanded_raw_markers_disambiguate_a_valid_owner_edit() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body>
      <q:p xmlns:x="urn:same"><x:a/><x:b mark="same"></x:b><q:r><q:t>first</q:t></q:r></q:p>
      <q:p xmlns:x="urn:same"><x:a/><x:b mark='same'></x:b><q:r><q:t>second</q:t></q:r></q:p>
    </q:body></q:document>"#;
    let mut document = document_with_content_controls(xml);
    document
        .paragraph_mut(0)
        .unwrap()
        .run_mut(0)
        .unwrap()
        .set_text("first edited");
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let saved_xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(saved_xml.contains(r#"<x:b mark="same"></x:b>"#));
    assert!(saved_xml.contains("first edited"));
    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(reopened.paragraph(0).unwrap().text(), "first edited");
}

#[test]
fn exact_marker_cardinality_disambiguates_a_valid_owner_edit() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body>
      <q:p xmlns:x="urn:same"><x:a/><x:a/><q:r><q:t>first</q:t></q:r></q:p>
      <q:p xmlns:x="urn:same"><x:a/><x:a /><q:r><q:t>second</q:t></q:r></q:p>
    </q:body></q:document>"#;
    let mut document = document_with_content_controls(xml);
    document
        .paragraph_mut(0)
        .unwrap()
        .run_mut(0)
        .unwrap()
        .set_text("first edited");
    let saved = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let saved_xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert_eq!(saved_xml.matches("<x:a/>").count(), 3);
    assert_eq!(saved_xml.matches("<x:a />").count(), 1);
    assert!(saved_xml.contains("first edited"));
    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(reopened.paragraph(0).unwrap().text(), "first edited");
}

#[test]
fn nested_wp_collisions_preserve_unchanged_bytes_and_fail_closed_when_modified() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <q:body><q:p xmlns:wp="urn:foreign"><wp:producer/><q:r><q:t>typed</q:t></q:r></q:p></q:body>
      </q:document>"#;
    let mut unchanged = document_with_content_controls(xml);
    let source = ordered_reader_snapshot(&unchanged);
    let saved = unchanged.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    assert_eq!(
        package.get_part("/word/document.xml").unwrap(),
        xml.as_bytes()
    );
    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(ordered_reader_snapshot(&reopened), source);

    let mut modified = document_with_content_controls(xml);
    modified.add_paragraph("changed");
    let error = modified.to_bytes().unwrap_err();
    assert!(error.to_string().contains("shadowed `wp` namespace"));
}

#[test]
fn fixed_prefix_collisions_fail_closed_after_owner_insert_remove_and_reorder() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body>
      <q:p><q:r><q:t>decoy</q:t></q:r></q:p>
      <q:p xmlns:wp="urn:foreign"><wp:producer/><q:r><q:t>target</q:t></q:r></q:p>
    </q:body></q:document>"#;

    let mut unchanged = document_with_content_controls(xml);
    let saved = unchanged.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    assert_eq!(
        package.get_part("/word/document.xml").unwrap(),
        xml.as_bytes()
    );
    Document::from_bytes(&saved).unwrap();

    for mutation in ["insert", "remove", "reorder"] {
        let mut document = document_with_content_controls(xml);
        match mutation {
            "insert" => {
                document.insert_paragraph(0, "inserted");
            }
            "remove" => {
                assert!(document.remove_content(0));
            }
            "reorder" => {
                assert!(document.remove_content(0));
                document.add_paragraph("moved decoy");
            }
            _ => unreachable!(),
        }
        let error = document.to_bytes().unwrap_err();
        assert!(
            error.to_string().contains("shadowed `wp` namespace"),
            "{mutation} must fail closed: {error}",
        );
    }
}

#[test]
fn escaped_root_namespace_uris_decode_once_across_unchanged_and_modified_saves() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:x="urn:a&amp;b"><q:body><x:producer/></q:body></q:document>"#;
    for modified in [false, true] {
        let mut document = document_with_content_controls(xml);
        let BodyContentRef::UnsupportedXml(fact) = document.body_content().next().unwrap() else {
            panic!("expected unsupported fact");
        };
        assert_eq!(fact.namespace_uri(), Some("urn:a&b"));
        if modified {
            document.add_paragraph("changed");
        }

        let saved = document.to_bytes().unwrap();
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(saved_xml.contains(r#"xmlns:x="urn:a&amp;b""#));
        assert!(!saved_xml.contains("urn:a&amp;amp;b"));

        let reopened = Document::from_bytes(&saved).unwrap();
        let BodyContentRef::UnsupportedXml(fact) = reopened.body_content().next().unwrap() else {
            panic!("expected reopened unsupported fact");
        };
        assert_eq!(fact.namespace_uri(), Some("urn:a&b"));
        if modified {
            assert_eq!(reopened.paragraph(0).unwrap().text(), "changed");
        }
    }
}

#[test]
fn empty_modeled_controls_and_numeric_references_report_visible_content_accurately() {
    let xml = r#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:x="urn:foreign"><q:body><q:sdt></q:sdt>
        <x:space>&#32;</x:space><x:tab>&#x9;</x:tab><x:newline>&#xA;</x:newline>
        <x:return>&#13;</x:return><x:visible>&#65;</x:visible><x:named>&amp;</x:named>
      </q:body></q:document>"#;
    let document = document_with_content_controls(xml);
    let facts = document.body_content().collect::<Vec<_>>();
    assert_eq!(facts.len(), 7);
    let BodyContentRef::UnsupportedXml(control) = &facts[0] else {
        panic!("expected modeled control fact");
    };
    assert_eq!(control.raw_xml(), None);
    assert!(!control.has_child_content());
    for fact in &facts[1..5] {
        let BodyContentRef::UnsupportedXml(fact) = fact else {
            panic!("expected raw unsupported fact");
        };
        assert!(!fact.has_child_content());
    }
    for fact in &facts[5..] {
        let BodyContentRef::UnsupportedXml(fact) = fact else {
            panic!("expected raw unsupported fact");
        };
        assert!(fact.has_child_content());
    }
}

#[test]
fn producer_defined_number_formats_survive_save_and_reopen() {
    let document_xml = wrap_word_body(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>"#,
    );
    let numbering_xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="chicago"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>"#;
    let mut seed = Document::new();
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
    package.set_part("/word/document.xml", document_xml.into_bytes());
    package.set_part("/word/numbering.xml", numbering_xml.to_vec());
    package.content_types.add_override(
        "/word/numbering.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    );
    package.get_or_create_part_rels("/word/document.xml").add(
        oxml_opc::relationship::rel_types::NUMBERING,
        "numbering.xml",
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();

    let mut document = Document::from_bytes(bytes.get_ref()).unwrap();
    assert_eq!(document.numbering_is_bullet(1), None);
    assert!(!document.to_html_fragment().contains("<ol>"));
    assert!(!document.to_markdown().contains("1. item"));
    assert!(
        document
            .to_rtf_bytes()
            .unwrap()
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message.contains("numbering format"))
    );

    let saved = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(reopened.numbering_is_bullet(1), None);
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let numbering = std::str::from_utf8(package.get_part("/word/numbering.xml").unwrap()).unwrap();
    assert!(numbering.contains(r#"<w:numFmt w:val="chicago"/>"#));
}

#[test]
fn legacy_flattened_accessors_keep_their_recursive_results() {
    let xml = wrap_word_body(
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>direct</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sdt><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>",
    );
    let document = document_with_content_controls(&xml);
    let table = document.table(0).unwrap();
    let cell = table.cell(0, 0).unwrap();
    assert_eq!(
        cell.paragraphs()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>(),
        ["direct", "control"]
    );
    assert_eq!(
        document
            .paragraphs()
            .into_iter()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>(),
        Vec::<String>::new()
    );
}

fn unsupported_fact_snapshot(
    fact: UnsupportedXmlRef<'_>,
    raw_subtrees: &mut Vec<Vec<u8>>,
) -> String {
    let raw = fact.raw_xml().map(|raw| {
        raw_subtrees.push(raw.to_vec());
        std::str::from_utf8(raw).unwrap().to_owned()
    });
    format!(
        "unsupported:{:?}:{:?}:{}:{}:{raw:?}",
        fact.qualified_name(),
        fact.namespace_uri(),
        fact.local_name(),
        fact.has_child_content(),
    )
}

fn run_snapshot(run: RunRef<'_>, raw_subtrees: &mut Vec<Vec<u8>>) -> String {
    run.items()
        .map(|item| match item {
            RunItemRef::Text(text) => format!("text:{text}"),
            RunItemRef::DeletedText(text) => format!("deleted:{text}"),
            RunItemRef::Tab => "tab".to_owned(),
            RunItemRef::Break(BreakKind::Line) => "break:line".to_owned(),
            RunItemRef::Break(BreakKind::Page) => "break:page".to_owned(),
            RunItemRef::Break(BreakKind::Column) => "break:column".to_owned(),
            RunItemRef::Drawing(drawing) => format!(
                "drawing:{}:{}:{:?}:{:?}:{:?}:{:?}:{:?}",
                drawing.is_inline(),
                drawing.is_anchor(),
                drawing.relationship_id(),
                drawing.name(),
                drawing.description(),
                drawing.width().map(Length::to_emu),
                drawing.height().map(Length::to_emu),
            ),
            RunItemRef::Field(field) => format!(
                "field:{}:{}:{}:{:?}",
                field.instruction(),
                field.name(),
                field.cached_result(),
                field.dirty(),
            ),
            RunItemRef::FootnoteReference(id) => format!("footnote:{id}"),
            RunItemRef::EndnoteReference(id) => format!("endnote:{id}"),
            RunItemRef::CommentReference(id) => format!("comment:{id}"),
            RunItemRef::UnsupportedXml(raw) => {
                raw_subtrees.push(raw.to_vec());
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
            _ => panic!("unexpected run item"),
        })
        .collect::<Vec<_>>()
        .join(",")
}

fn hyperlink_snapshot(hyperlink: HyperlinkRef<'_>, raw_subtrees: &mut Vec<Vec<u8>>) -> String {
    let items = hyperlink
        .items()
        .map(|item| match item {
            HyperlinkItemRef::Run(run) => format!("run:[{}]", run_snapshot(run, raw_subtrees)),
            HyperlinkItemRef::Revision(revision) => format!(
                "revision:{}:{}:{:?}",
                revision.id(),
                revision.author(),
                revision.kind(),
            ),
            HyperlinkItemRef::UnsupportedXml(raw) => {
                raw_subtrees.push(raw.to_vec());
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
            _ => panic!("unexpected hyperlink item"),
        })
        .collect::<Vec<_>>()
        .join(",");
    format!(
        "hyperlink:{:?}:{:?}:[{items}]",
        hyperlink.relationship_id(),
        hyperlink.anchor(),
    )
}

fn paragraph_snapshot(paragraph: ParagraphRef<'_>, raw_subtrees: &mut Vec<Vec<u8>>) -> String {
    paragraph
        .items()
        .map(|item| match item {
            ParagraphItemRef::Run(run) => format!("run:[{}]", run_snapshot(run, raw_subtrees)),
            ParagraphItemRef::Hyperlink(hyperlink) => hyperlink_snapshot(hyperlink, raw_subtrees),
            ParagraphItemRef::ContentControl(control) => {
                format!("control:{}", control.text())
            }
            ParagraphItemRef::Revision(revision) => format!(
                "revision:{}:{}:{:?}",
                revision.id(),
                revision.author(),
                revision.kind(),
            ),
            ParagraphItemRef::CommentRangeStart(id) => format!("comment-start:{id}"),
            ParagraphItemRef::CommentRangeEnd(id) => format!("comment-end:{id}"),
            ParagraphItemRef::BookmarkStart { id, name } => {
                format!("bookmark-start:{id:?}:{name:?}")
            }
            ParagraphItemRef::BookmarkEnd { id } => format!("bookmark-end:{id:?}"),
            ParagraphItemRef::UnsupportedXml(raw) => {
                raw_subtrees.push(raw.to_vec());
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
            _ => panic!("unexpected paragraph item"),
        })
        .collect::<Vec<_>>()
        .join(",")
}

fn table_snapshot(table: TableRef<'_>, raw_subtrees: &mut Vec<Vec<u8>>) -> String {
    let rows = (0..table.row_count())
        .map(|row_index| {
            let row = table.row(row_index).unwrap();
            let cells = (0..row.cell_count())
                .map(|cell_index| cell_snapshot(row.cell(cell_index).unwrap(), raw_subtrees))
                .collect::<Vec<_>>()
                .join("|");
            format!("row:[{cells}]")
        })
        .collect::<Vec<_>>()
        .join("|");
    format!("table:[{rows}]")
}

fn cell_snapshot(cell: CellRef<'_>, raw_subtrees: &mut Vec<Vec<u8>>) -> String {
    cell.items()
        .map(|item| match item {
            CellItemRef::Paragraph(paragraph) => {
                format!(
                    "paragraph:[{}]",
                    paragraph_snapshot(paragraph, raw_subtrees)
                )
            }
            CellItemRef::Table(table) => table_snapshot(table, raw_subtrees),
            CellItemRef::ContentControl(control) => format!("control:{}", control.text()),
            CellItemRef::UnsupportedXml(raw) => {
                raw_subtrees.push(raw.to_vec());
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
            _ => panic!("unexpected cell item"),
        })
        .collect::<Vec<_>>()
        .join(",")
}

fn ordered_reader_snapshot(document: &Document) -> (Vec<String>, Vec<Vec<u8>>) {
    let mut raw_subtrees = Vec::new();
    let facts = document
        .body_content()
        .map(|item| match item {
            BodyContentRef::Paragraph(paragraph) => {
                format!(
                    "paragraph:[{}]",
                    paragraph_snapshot(paragraph, &mut raw_subtrees)
                )
            }
            BodyContentRef::Table(table) => table_snapshot(table, &mut raw_subtrees),
            BodyContentRef::UnsupportedXml(fact) => {
                unsupported_fact_snapshot(fact, &mut raw_subtrees)
            }
            _ => panic!("unexpected body item"),
        })
        .collect::<Vec<_>>();
    (facts, raw_subtrees)
}

#[test]
fn ordered_reader_source_survives_save_and_reopen() {
    let xml = ordered_reader_fixture();
    let mut document = document_with_content_controls(xml);
    let source = ordered_reader_snapshot(&document);
    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let reopened = ordered_reader_snapshot(&reopened);
    assert_eq!(reopened.0, source.0, "ordered public facts changed");
    assert_eq!(reopened.1, source.1, "raw subtrees changed");
}

fn document_with_bound_content_controls(document_xml: &str, custom_xml: Option<&str>) -> Document {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    package.set_part("/word/document.xml", document_xml.as_bytes().to_vec());

    if let Some(custom_xml) = custom_xml {
        package.set_part("/customXml/item1.xml", custom_xml.as_bytes().to_vec());
        package.set_part(
            "/customXml/itemProps1.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="{11111111-1111-1111-1111-111111111111}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>"#
                .to_vec(),
        );
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(CUSTOM_XML_REL_TYPE, "../customXml/item1.xml");
        package
            .get_or_create_part_rels("/customXml/item1.xml")
            .add(CUSTOM_XML_PROPS_REL_TYPE, "itemProps1.xml");
        package.content_types.add_override(
            "/customXml/itemProps1.xml",
            "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
        );
    }

    let mut output = std::io::Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();
    Document::from_bytes(output.get_ref()).unwrap()
}

fn wrap_word_body(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
    )
}

fn body_from_document(document: &mut Document) -> CT_Body {
    let package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
            .unwrap();
    CT_Document::from_xml(package.get_part("/word/document.xml").unwrap())
        .unwrap()
        .body
}

fn document_xml(document: &mut Document) -> String {
    let package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
            .unwrap();
    String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap()
}

fn document_with_comment_paragraphs(paragraphs: &str) -> (Document, i32) {
    let mut document = Document::new();
    document.add_paragraph("seed");
    let comment_id = document
        .add_comment(
            RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 0,
                    run_index: 1,
                },
            },
            "Ada",
            None,
            "remove",
        )
        .unwrap();
    let bytes = document.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let mut xml =
        String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
    let start = xml.find("<w:p>").unwrap();
    let end = start + xml[start..].find("</w:p>").unwrap() + "</w:p>".len();
    xml.replace_range(start..end, paragraphs);
    package.set_part("/word/document.xml", xml.into_bytes());
    let mut output = std::io::Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();
    (Document::from_bytes(output.get_ref()).unwrap(), comment_id)
}

fn body_from_xml(body: &str) -> CT_Body {
    CT_Document::from_xml(wrap_word_body(body).as_bytes())
        .unwrap()
        .body
}

#[test]
fn accepting_every_revision_matches_word_normalized_body_xml() {
    let xml = wrap_word_body(
        r#"<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="5" w:author="Ada"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><w:rPrChange w:id="6" w:author="Ada"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>kept</w:t></w:r><w:ins w:id="1" w:author="Ada"><w:r><w:t> inserted</w:t></w:r></w:ins><w:del w:id="2" w:author="Ada"><w:r><w:delText> deleted</w:delText></w:r></w:del><w:moveFrom w:id="3" w:author="Ada"><w:r><w:t> old-place</w:t></w:r></w:moveFrom><w:moveTo w:id="4" w:author="Ada"><w:r><w:t> moved</w:t></w:r></w:moveTo></w:p><w:tbl><w:tblPr><w:jc w:val="center"/><w:tblPrChange w:id="7" w:author="Ada"><w:tblPr><w:jc w:val="right"/></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sectPr><w:titlePg/><w:sectPrChange w:id="8" w:author="Ada"><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:sectPrChange></w:sectPr>"#,
    );
    let mut document = document_with_content_controls(&xml);

    assert_eq!(document.accept_all().unwrap(), 8);
    assert!(document.revisions().is_empty());
    let expected = body_from_xml(
        r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>kept</w:t></w:r><w:r><w:t> inserted</w:t></w:r><w:r><w:t> moved</w:t></w:r></w:p><w:tbl><w:tblPr><w:jc w:val="center"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sectPr><w:titlePg/></w:sectPr>"#,
    );
    assert_eq!(
        body_from_document(&mut document),
        expected,
        "normalized body must match the pinned {WORD_REVISION_ORACLE} oracle"
    );
}

#[test]
fn rejecting_insertions_and_deletions_restores_the_recorded_content() {
    let xml = wrap_word_body(
        r#"<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="5" w:author="Ada"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><w:rPrChange w:id="6" w:author="Ada"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>kept</w:t></w:r><w:ins w:id="1" w:author="Ada"><w:r><w:t> inserted</w:t></w:r></w:ins><w:del w:id="2" w:author="Ada"><w:r><w:delText> deleted</w:delText></w:r></w:del><w:moveFrom w:id="3" w:author="Ada"><w:r><w:t> old-place</w:t></w:r></w:moveFrom><w:moveTo w:id="4" w:author="Ada"><w:r><w:t> moved</w:t></w:r></w:moveTo></w:p><w:tbl><w:tblPr><w:jc w:val="center"/><w:tblPrChange w:id="7" w:author="Ada"><w:tblPr><w:jc w:val="right"/></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sectPr><w:titlePg/><w:sectPrChange w:id="8" w:author="Ada"><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:sectPrChange></w:sectPr>"#,
    );
    let mut document = document_with_content_controls(&xml);

    assert_eq!(document.reject_all().unwrap(), 8);
    assert!(document.revisions().is_empty());
    let expected = body_from_xml(
        r#"<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>kept</w:t></w:r><w:r><w:t> deleted</w:t></w:r><w:r><w:t> old-place</w:t></w:r></w:p><w:tbl><w:tblPr><w:jc w:val="right"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#,
    );
    assert_eq!(body_from_document(&mut document), expected);
}

#[test]
fn scoped_revision_actions_change_only_matching_revisions() {
    let xml = wrap_word_body(
        r#"<w:p><w:ins w:id="7" w:author="Ada" w:date="2026-08-17T09:00:00+01:00"><w:r><w:t>A</w:t></w:r></w:ins><w:ins w:id="7" w:author="Ben" w:date="2026-08-17T09:00:00Z"><w:r><w:t>B</w:t></w:r></w:ins><w:del w:id="8" w:author="Ada"><w:r><w:delText>C</w:delText></w:r></w:del></w:p>"#,
    );

    let mut by_author = document_with_content_controls(&xml);
    assert_eq!(by_author.accept_revisions_by_author("Ada").unwrap(), 2);
    assert!(document_xml(&mut by_author).contains(
        r#"<w:ins w:id="7" w:author="Ben" w:date="2026-08-17T09:00:00Z"><w:r><w:t>B</w:t></w:r></w:ins>"#
    ));
    assert_eq!(
        by_author
            .revisions()
            .iter()
            .map(|revision| revision.author())
            .collect::<Vec<_>>(),
        vec!["Ben"]
    );

    let mut by_id = document_with_content_controls(&xml);
    assert_eq!(by_id.reject_revision_id(7).unwrap(), 2);
    assert_eq!(
        by_id
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>(),
        vec![8]
    );

    let mut by_date = document_with_content_controls(&xml);
    assert_eq!(
        by_date
            .accept_revisions_in_date_range("2026-08-17T08:00:00Z", "2026-08-17T08:00:00Z")
            .unwrap(),
        1
    );
    assert_eq!(
        by_date
            .revisions()
            .iter()
            .map(|revision| revision.author())
            .collect::<Vec<_>>(),
        vec!["Ben", "Ada"]
    );

    let mut by_lowercase_date = document_with_content_controls(&xml);
    assert_eq!(
        by_lowercase_date
            .reject_revisions_in_date_range("2026-08-17t08:00:00z", "2026-08-17t08:00:00z",)
            .unwrap(),
        1
    );
}

#[test]
fn unmodelled_revision_lookalikes_remain_byte_identical() {
    let xml = wrap_word_body(
        r#"<w:customXml><w:ins w:id="99" w:author="raw"><w:r><w:t>opaque</w:t></w:r></w:ins></w:customXml><w:p><w:r><w:t>typed</w:t></w:r></w:p>"#,
    );
    let mut document = document_with_content_controls(&xml);
    let before = document.to_bytes().unwrap();

    assert_eq!(document.accept_all().unwrap(), 0);
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn unwrapped_revisions_keep_wrapper_namespace_bindings() {
    let xml = wrap_word_body(
        r#"<w:p><w:ins xmlns:vendor="urn:vendor" w:id="1" w:author="Ada"><w:r><w:t>kept</w:t><vendor:opaque vendor:value="yes"/></w:r></w:ins></w:p>"#,
    );
    let mut document = document_with_content_controls(&xml);

    assert_eq!(document.accept_all().unwrap(), 1);
    let saved = document_xml(&mut document);
    assert!(saved.contains("xmlns:vendor=\"urn:vendor\""));
    assert!(saved.contains("<vendor:opaque"));
    assert!(saved.contains("vendor:value=\"yes\""));
}

#[test]
fn rejected_property_changes_keep_owner_namespace_bindings() {
    let xml = wrap_word_body(
        r#"<w:p><w:r><w:rPr xmlns:vendor="urn:vendor"><w:b/><w:rPrChange w:id="1" w:author="Ada"><w:rPr><vendor:rPrChange vendor:value="prior"/></w:rPr></w:rPrChange></w:rPr><w:t>text</w:t></w:r></w:p>"#,
    );
    let mut document = document_with_content_controls(&xml);

    assert_eq!(document.reject_all().unwrap(), 1);
    let saved = document_xml(&mut document);
    assert!(saved.contains("xmlns:vendor=\"urn:vendor\""));
    assert!(saved.contains("<vendor:rPrChange"));
    assert!(saved.contains("vendor:value=\"prior\""));
}

#[test]
fn invalid_date_ranges_leave_the_document_unchanged() {
    let xml = wrap_word_body(
        r#"<w:p><w:ins w:id="1" w:author="Ada" w:date="2026-08-17T09:00:00Z"><w:r><w:t>A</w:t></w:r></w:ins></w:p>"#,
    );
    let mut document = document_with_content_controls(&xml);
    let before = document.to_bytes().unwrap();

    assert!(
        document
            .accept_revisions_in_date_range("not-a-date", "2026-08-17T09:00:00Z")
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), before);
    assert!(
        document
            .reject_revisions_in_date_range("2026-08-18T09:00:00Z", "2026-08-17T09:00:00Z")
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), before);

    for malformed in [
        "2026-08-17T-1:00:00Z",
        "2026-8-17T09:00:00Z",
        "2026-08-17T9:00:00Z",
        "9223372036854775807-08-17T09:00:00Z",
        "2026-08-17T12:00:60Z",
        "2016-12-31T23:59:60Z",
    ] {
        assert!(
            document
                .accept_revisions_in_date_range(malformed, "2026-08-17T09:00:00Z")
                .is_err(),
            "{malformed} must not parse as RFC 3339"
        );
        assert_eq!(document.to_bytes().unwrap(), before);
    }
}

#[test]
fn contextual_row_markers_keep_or_remove_the_owning_row() {
    let xml = wrap_word_body(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:trPr><w:ins w:id="1" w:author="Ada"/></w:trPr><w:tc><w:p><w:r><w:t>inserted</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:trPr><w:del w:id="2" w:author="Ada"/></w:trPr><w:tc><w:p><w:r><w:t>deleted</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );

    let mut accepted = document_with_content_controls(&xml);
    assert_eq!(accepted.accept_all().unwrap(), 2);
    assert!(accepted.text().contains("inserted"));
    assert!(!accepted.text().contains("deleted"));

    let mut rejected = document_with_content_controls(&xml);
    assert_eq!(rejected.reject_all().unwrap(), 2);
    assert!(!rejected.text().contains("inserted"));
    assert!(rejected.text().contains("deleted"));

    let conflicting_xml = wrap_word_body(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:trPr><w:ins w:id="3" w:author="Ada"/><w:del w:id="4" w:author="Ada"/></w:trPr><w:tc><w:p><w:r><w:t>ambiguous</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let mut conflicting = document_with_content_controls(&conflicting_xml);
    assert_eq!(conflicting.accept_all().unwrap(), 2);
    assert!(!conflicting.text().contains("ambiguous"));
}

#[test]
fn contextual_cleanup_preserves_foreign_property_shells() {
    let xml = wrap_word_body(
        r#"<w:p><w:r><x:rPr xmlns:x="urn:foreign-property"><w:del w:id="91" w:author="lookalike"/></x:rPr><w:t>kept</w:t></w:r><w:ins w:id="93" w:author="Ada"><w:r><w:t> accepted</w:t></w:r></w:ins></w:p>"#,
    );
    let mut document = document_with_content_controls(&xml);

    assert_eq!(document.accept_all().unwrap(), 1);
    let saved = document_xml(&mut document);
    assert!(saved.contains(r#"<x:rPr xmlns:x="urn:foreign-property">"#));
    assert!(saved.contains(r#"w:id="91" w:author="lookalike""#));
    assert!(saved.contains("<w:t>kept</w:t>"));
    assert!(saved.contains("<w:t> accepted</w:t>"));
}

#[test]
fn table_cleanup_retains_rows_owned_by_content_controls() {
    let xml = wrap_word_body(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:sdt><w:sdtPr><w:tag w:val="retained-row"/></w:sdtPr><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>retained</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt><w:tr><w:trPr><w:del w:id="92" w:author="Ada"/></w:trPr><w:tc><w:p><w:r><w:t>deleted</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let mut document = document_with_content_controls(&xml);

    assert_eq!(document.accept_all().unwrap(), 1);
    let saved = document_xml(&mut document);
    assert!(saved.contains(r#"w:val="retained-row""#), "{saved}");
    assert!(saved.contains("<w:t>retained</w:t>"), "{saved}");
    assert!(!saved.contains("<w:t>deleted</w:t>"), "{saved}");
}

#[test]
fn contextual_paragraph_markers_merge_the_adjacent_paragraphs() {
    let xml = wrap_word_body(
        r#"<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="Ada"/></w:rPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t> second</w:t></w:r></w:p>"#,
    );
    let mut accepted = document_with_content_controls(&xml);
    assert_eq!(accepted.accept_all().unwrap(), 1);
    assert_eq!(accepted.paragraph_count(), 1);
    assert_eq!(accepted.text(), "first second\n");
    assert_eq!(
        body_from_document(&mut accepted),
        body_from_xml(
            r#"<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>first</w:t></w:r><w:r><w:t> second</w:t></w:r></w:p>"#,
        )
    );

    let insertion_xml = wrap_word_body(
        r#"<w:p><w:pPr><w:rPr><w:ins w:id="2" w:author="Ada"/></w:rPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t> second</w:t></w:r></w:p>"#,
    );
    let mut rejected = document_with_content_controls(&insertion_xml);
    assert_eq!(rejected.reject_all().unwrap(), 1);
    assert_eq!(rejected.paragraph_count(), 1);
    assert_eq!(rejected.text(), "first second\n");

    let chained_xml = wrap_word_body(
        r#"<w:p><w:pPr><w:rPr><w:del w:id="3" w:author="Ada"/></w:rPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:del w:id="4" w:author="Ada"/></w:rPr></w:pPr><w:r><w:t> two</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t> three</w:t></w:r></w:p>"#,
    );
    let mut chained = document_with_content_controls(&chained_xml);
    assert_eq!(chained.accept_all().unwrap(), 2);
    assert_eq!(chained.paragraph_count(), 1);
    assert_eq!(chained.text(), "one two three\n");
    assert_eq!(
        body_from_document(&mut chained),
        body_from_xml(
            r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>one</w:t></w:r><w:r><w:t> two</w:t></w:r><w:r><w:t> three</w:t></w:r></w:p>"#,
        )
    );
}

#[test]
fn malformed_selected_property_changes_fail_atomically() {
    let xml = wrap_word_body(
        r#"<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="1" w:author="Ada"/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>"#,
    );
    let mut document = document_with_content_controls(&xml);
    let before = document.to_bytes().unwrap();
    let normal_layout = document.layout_page(0).unwrap();
    let deterministic_layout = document.to_pdf_deterministic().unwrap();

    assert!(document.reject_all().is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
    assert_eq!(document.revisions().len(), 1);
    assert_eq!(
        format!("{:?}", document.layout_page(0).unwrap()),
        format!("{normal_layout:?}")
    );
    assert_eq!(
        document.to_pdf_deterministic().unwrap(),
        deterministic_layout
    );

    let wrong_prior_xml = wrap_word_body(
        r#"<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="2" w:author="Ada"><w:pPr/></w:rPrChange></w:rPr><w:t>text</w:t></w:r></w:p>"#,
    );
    let mut wrong_prior = document_with_content_controls(&wrong_prior_xml);
    let wrong_prior_before = wrong_prior.to_bytes().unwrap();
    assert!(wrong_prior.reject_all().is_err());
    assert_eq!(wrong_prior.to_bytes().unwrap(), wrong_prior_before);

    let extra_prior_xml = wrap_word_body(
        r#"<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="5" w:author="Ada"><w:rPr><w:i/></w:rPr><w:rPr/></w:rPrChange></w:rPr><w:t>text</w:t></w:r></w:p>"#,
    );
    let mut extra_prior = document_with_content_controls(&extra_prior_xml);
    let extra_prior_before = extra_prior.to_bytes().unwrap();
    assert!(extra_prior.accept_all().is_err());
    assert_eq!(extra_prior.to_bytes().unwrap(), extra_prior_before);

    let hidden_xml = wrap_word_body(
        r#"<w:p><w:ins w:id="3" w:author="Ada"><w:r><w:rPr><w:rPrChange w:id="4" w:author="Ada"/></w:rPr><w:t>hidden</w:t></w:r></w:ins></w:p>"#,
    );
    let mut hidden = document_with_content_controls(&hidden_xml);
    let hidden_before = hidden.to_bytes().unwrap();
    assert!(hidden.reject_all().is_err());
    assert_eq!(hidden.to_bytes().unwrap(), hidden_before);
}

#[test]
fn nested_selected_revisions_resolve_inside_out_and_count_once() {
    let xml = wrap_word_body(
        r#"<w:p><w:ins w:id="1" w:author="Ada"><w:r><w:rPr><w:b/><w:rPrChange w:id="2" w:author="Ada"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>nested</w:t></w:r></w:ins></w:p>"#,
    );
    let mut accepted = document_with_content_controls(&xml);

    assert_eq!(accepted.accept_all().unwrap(), 2);
    assert!(accepted.revisions().is_empty());
    assert_eq!(accepted.text(), "nested\n");
}

#[test]
fn duplicate_property_revisions_round_trip_and_resolve_one_identity_at_a_time() {
    let xml = wrap_word_body(
        r#"<w:p><w:pPr><w:numPr><w:ins w:id="501" w:author="Ada"/><x:ins xmlns:x="urn:producer" x:mark="num"/><w:ins w:id="502" w:author="Ada"/></w:numPr><w:pPrChange w:id="101" w:author="Ada"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange><x:pPrChange xmlns:x="urn:producer" x:mark="paragraph"/><w:pPrChange w:id="102" w:author="Ada"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:rPrChange w:id="201" w:author="Ada"><w:rPr><w:b/></w:rPr></w:rPrChange><x:rPrChange xmlns:x="urn:producer" x:mark="run"/><w:rPrChange w:id="202" w:author="Ada"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblPrChange w:id="301" w:author="Ada"><w:tblPr><w:jc w:val="left"/></w:tblPr></w:tblPrChange><x:tblPrChange xmlns:x="urn:producer" x:mark="table"/><w:tblPrChange w:id="302" w:author="Ada"><w:tblPr><w:jc w:val="right"/></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sectPr><w:sectPrChange w:id="401" w:author="Ada"><w:sectPr><w:titlePg/></w:sectPr></w:sectPrChange><x:sectPrChange xmlns:x="urn:producer" x:mark="section"/><w:sectPrChange w:id="402" w:author="Ada"><w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:sectPrChange></w:sectPr>"#,
    );
    let mut document = document_with_content_controls(&xml);
    let sorted_ids = |document: &Document| {
        let mut ids = document
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>();
        ids.sort_unstable();
        ids
    };
    assert_eq!(sorted_ids(&document), [102, 202, 302, 402, 502]);

    let initial = document_xml(&mut document);
    for id in [101, 102, 201, 202, 301, 302, 401, 402, 501, 502] {
        assert_eq!(initial.matches(&format!(r#"w:id="{id}""#)).count(), 1);
    }
    for mark in ["num", "paragraph", "run", "table", "section"] {
        assert_eq!(initial.matches(&format!(r#"x:mark="{mark}""#)).count(), 1);
    }

    for id in [102, 202, 302, 402, 502] {
        assert_eq!(document.accept_revision_id(id).unwrap(), 1);
    }
    assert_eq!(sorted_ids(&document), [101, 201, 301, 401, 501]);
    let staged = document_xml(&mut document);
    for id in [101, 201, 301, 401, 501] {
        assert_eq!(staged.matches(&format!(r#"w:id="{id}""#)).count(), 1);
    }
    for id in [102, 202, 302, 402, 502] {
        assert!(!staged.contains(&format!(r#"w:id="{id}""#)));
    }
    for mark in ["num", "paragraph", "run", "table", "section"] {
        assert_eq!(staged.matches(&format!(r#"x:mark="{mark}""#)).count(), 1);
    }

    let mut reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    assert_eq!(sorted_ids(&reopened), [101, 201, 301, 401, 501]);
    for id in [101, 201, 301, 401, 501] {
        assert_eq!(reopened.accept_revision_id(id).unwrap(), 1);
    }
    assert!(reopened.revisions().is_empty());

    let mut rejected = document_with_content_controls(&xml);
    for id in [102, 202, 302, 402, 502] {
        assert_eq!(rejected.reject_revision_id(id).unwrap(), 1, "id {id}");
    }
    assert_eq!(sorted_ids(&rejected), [101, 201, 301, 401, 501]);
    let rejected_xml = document_xml(&mut rejected);
    for id in [101, 201, 301, 401, 501] {
        assert_eq!(rejected_xml.matches(&format!(r#"w:id="{id}""#)).count(), 1);
    }
    for mark in ["num", "paragraph", "run", "table", "section"] {
        assert_eq!(
            rejected_xml.matches(&format!(r#"x:mark="{mark}""#)).count(),
            1
        );
    }

    let mut rejected = Document::from_bytes(&rejected.to_bytes().unwrap()).unwrap();
    assert_eq!(sorted_ids(&rejected), [101, 201, 301, 401, 501]);
    for id in [101, 201, 301, 401, 501] {
        assert_eq!(rejected.reject_revision_id(id).unwrap(), 1);
    }
    assert!(rejected.revisions().is_empty());
    let rejected_xml = document_xml(&mut rejected);
    for mark in ["num", "paragraph", "run", "table", "section"] {
        assert_eq!(
            rejected_xml.matches(&format!(r#"x:mark="{mark}""#)).count(),
            1
        );
    }

    let mut rejected_all = document_with_content_controls(&xml);
    assert_eq!(rejected_all.reject_all().unwrap(), 10);
    assert!(rejected_all.revisions().is_empty());
    let rejected_all_xml = document_xml(&mut rejected_all);
    for mark in ["num", "paragraph", "run", "table", "section"] {
        assert_eq!(
            rejected_all_xml
                .matches(&format!(r#"x:mark="{mark}""#))
                .count(),
            1
        );
    }
}

#[test]
fn hyperlink_nested_revisions_resolve_inside_out_when_scoped() {
    let xml = wrap_word_body(
        r#"<w:p><w:hyperlink r:id="rId5"><w:r><w:t xml:space="preserve">before </w:t></w:r><w:ins w:id="11" w:author="Ada"><w:del w:id="12" w:author="Ben"><w:r><w:delText>nested</w:delText></w:r></w:del></w:ins><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:hyperlink></w:p>"#,
    );
    let mut document = document_with_content_controls(&xml);

    assert_eq!(document.reject_revision_id(12).unwrap(), 1);
    assert_eq!(
        document
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>(),
        [11]
    );
    let staged = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(staged)).unwrap();
    let document_xml =
        std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(document_xml.contains(r#"<w:ins w:id="11" w:author="Ada">"#));
    assert!(document_xml.contains("<w:t>nested</w:t>"));
    assert!(!document_xml.contains(r#"w:id="12""#));

    assert_eq!(document.accept_revision_id(11).unwrap(), 1);
    assert!(document.revisions().is_empty());
    assert_eq!(document.text(), "before nested after\n");
}

#[test]
fn targetless_revision_only_hyperlinks_keep_sibling_order_when_resolved() {
    let xml = wrap_word_body(
        r#"<w:p><w:hyperlink><w:ins w:id="21" w:author="Ada"><w:r><w:t>H</w:t></w:r></w:ins></w:hyperlink><w:del w:id="22" w:author="Ben"><w:r><w:delText>B</w:delText></w:r></w:del><w:ins w:id="23" w:author="Cy"><w:r><w:t>C</w:t></w:r></w:ins><w:hyperlink><w:del w:id="24" w:author="Dee"><w:r><w:delText>D</w:delText></w:r></w:del></w:hyperlink></w:p>"#,
    );
    let mut document = document_with_content_controls(&xml);

    assert_eq!(
        document
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>(),
        [21, 22, 23, 24]
    );
    assert_eq!(document.accept_revision_id(21).unwrap(), 1);
    assert_eq!(document.reject_revision_id(24).unwrap(), 1);
    assert_eq!(
        document
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>(),
        [22, 23]
    );

    let staged = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(staged)).unwrap();
    let document_xml =
        std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    let positions = [
        document_xml.find("<w:t>H</w:t>").unwrap(),
        document_xml.find(r#"w:id="22""#).unwrap(),
        document_xml.find(r#"w:id="23""#).unwrap(),
        document_xml.find("<w:t>D</w:t>").unwrap(),
    ];
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));

    assert_eq!(document.accept_all().unwrap(), 2);
    assert!(document.revisions().is_empty());
    assert_eq!(document.text(), "HCD\n");
}

#[test]
fn resolving_a_modeled_hyperlink_keeps_unreported_raw_children() {
    let malformed = r#"<w:ins w:id="bad"><w:r><w:t>raw revision</w:t></w:r></w:ins>"#;
    let foreign = r#"<x:opaque xmlns:x="urn:opaque" x:flag="1"/>"#;
    let xml = wrap_word_body(&format!(
        r#"<w:p><w:hyperlink r:id="rId5"><w:r><w:t>before</w:t></w:r>{malformed}<w:ins w:id="31" w:author="Ada"><w:r><w:t>reported</w:t></w:r></w:ins>{foreign}<w:r><w:t>after</w:t></w:r></w:hyperlink></w:p>"#
    ));
    let mut document = document_with_content_controls(&xml);

    assert_eq!(
        document
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>(),
        [31]
    );
    assert_eq!(document.accept_revision_id(31).unwrap(), 1);
    assert!(document.revisions().is_empty());
    assert_eq!(document.text(), "beforereportedafter\n");

    let staged = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(staged)).unwrap();
    let document_xml =
        std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(document_xml.contains(malformed));
    assert!(document_xml.contains(foreign));
    let positions = [
        document_xml.find("<w:t>before</w:t>").unwrap(),
        document_xml.find(malformed).unwrap(),
        document_xml.find("<w:t>reported</w:t>").unwrap(),
        document_xml.find(foreign).unwrap(),
        document_xml.find("<w:t>after</w:t>").unwrap(),
    ];
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
}

#[test]
fn malformed_revision_wrappers_are_opaque_to_every_resolution_scope() {
    let malformed = r#"<w:ins w:id="bad"><w:del w:id="51" w:author="Ada" w:date="2026-08-17T10:00:00Z"><w:r><w:delText>hidden</w:delText></w:r></w:del></w:ins>"#;
    let xml = wrap_word_body(&format!(
        r#"<w:p><w:hyperlink>{malformed}</w:hyperlink></w:p>"#
    ));
    let assert_opaque =
        |mut document: Document, action: fn(&mut Document) -> rdocx::Result<usize>| {
            assert!(document.revisions().is_empty());
            let before = document_xml(&mut document);
            assert_eq!(action(&mut document).unwrap(), 0);
            assert_eq!(document_xml(&mut document), before);
            assert!(before.contains(malformed));
        };

    assert_opaque(document_with_content_controls(&xml), Document::accept_all);
    assert_opaque(document_with_content_controls(&xml), Document::reject_all);
    assert_opaque(document_with_content_controls(&xml), |document| {
        document.accept_revision_id(51)
    });
    assert_opaque(document_with_content_controls(&xml), |document| {
        document.reject_revision_id(51)
    });
    assert_opaque(document_with_content_controls(&xml), |document| {
        document.accept_revisions_by_author("Ada")
    });
    assert_opaque(document_with_content_controls(&xml), |document| {
        document.reject_revisions_by_author("Ada")
    });
    assert_opaque(document_with_content_controls(&xml), |document| {
        document.accept_revisions_in_date_range("2026-08-17T09:00:00Z", "2026-08-17T11:00:00Z")
    });
    assert_opaque(document_with_content_controls(&xml), |document| {
        document.reject_revisions_in_date_range("2026-08-17T09:00:00Z", "2026-08-17T11:00:00Z")
    });
}

#[test]
fn comment_removal_remaps_and_retains_hyperlink_owned_revision_content() {
    let raw_solo = r#"<x:solo xmlns:x="urn:opaque"/>"#;
    let raw_middle = r#"<x:middle xmlns:x="urn:opaque"/>"#;
    let paragraphs = format!(
        r#"<w:p><w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/><w:hyperlink><w:r><w:commentReference w:id="0"/></w:r><w:ins w:id="41" w:author="Ada"><w:r><w:t>solo</w:t></w:r></w:ins>{raw_solo}</w:hyperlink></w:p><w:p><w:hyperlink><w:r><w:commentReference w:id="0"/></w:r><w:r><w:t>before</w:t></w:r><w:ins w:id="42" w:author="Ada"><w:r><w:t>middle</w:t></w:r></w:ins>{raw_middle}<w:r><w:t>after</w:t></w:r></w:hyperlink></w:p>"#
    );
    let (mut document, comment_id) = document_with_comment_paragraphs(&paragraphs);

    assert_eq!(
        document
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>(),
        [41, 42]
    );
    assert!(document.remove_comment(comment_id).unwrap());
    let removed = document_xml(&mut document);
    assert!(!removed.contains("commentReference"));
    for raw in [raw_solo, raw_middle] {
        assert!(removed.contains(raw), "missing {raw} in {removed}");
    }
    let positions = [
        removed.find("<w:t>before</w:t>").unwrap(),
        removed.find(r#"w:id="42""#).unwrap(),
        removed.find(raw_middle).unwrap(),
        removed.find("<w:t>after</w:t>").unwrap(),
    ];
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));

    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>(),
        [41, 42]
    );
    let mut reopened = reopened;
    assert_eq!(reopened.accept_revision_id(41).unwrap(), 1);
    assert_eq!(reopened.accept_revision_id(42).unwrap(), 1);
    assert!(reopened.revisions().is_empty());
    let resolved = document_xml(&mut reopened);
    assert!(resolved.contains(raw_solo));
    assert!(resolved.contains(raw_middle));
    assert_eq!(reopened.text(), "solo\nbeforemiddleafter\n");
}

#[test]
fn run_mutations_keep_raw_children_at_live_boundaries() {
    let body = r#"<w:p><w:r><w:before/><w:t>A</w:t><w:between/><w:t>B</w:t><w:after/></w:r></w:p>"#;

    let mut replaced = document_with_content_controls(&wrap_word_body(body));
    replaced
        .paragraph_mut(0)
        .unwrap()
        .run_mut(0)
        .unwrap()
        .set_text("replacement");
    let replaced_xml = document_xml(&mut replaced);
    let replaced_positions = [
        replaced_xml.find("<w:before/>").unwrap(),
        replaced_xml.find("<w:t>replacement</w:t>").unwrap(),
        replaced_xml.find("<w:between/>").unwrap(),
        replaced_xml.find("<w:after/>").unwrap(),
    ];
    assert!(replaced_positions.windows(2).all(|pair| pair[0] < pair[1]));

    let mut appended = document_with_content_controls(&wrap_word_body(body));
    appended
        .paragraph_mut(0)
        .unwrap()
        .run_mut(0)
        .unwrap()
        .add_text("C");
    let appended_xml = document_xml(&mut appended);
    let appended_positions = [
        appended_xml.find("<w:before/>").unwrap(),
        appended_xml.find("<w:t>A</w:t>").unwrap(),
        appended_xml.find("<w:between/>").unwrap(),
        appended_xml.find("<w:t>B</w:t>").unwrap(),
        appended_xml.find("<w:after/>").unwrap(),
        appended_xml.find("<w:t>C</w:t>").unwrap(),
    ];
    assert!(appended_positions.windows(2).all(|pair| pair[0] < pair[1]));

    let raw_property_body =
        r#"<w:p><w:r><w:rPr><w:rStyle><w:opaque/></w:rStyle></w:rPr><w:t>A</w:t></w:r></w:p>"#;
    let mut formatted = document_with_content_controls(&wrap_word_body(raw_property_body));
    formatted
        .paragraph_mut(0)
        .unwrap()
        .run_mut(0)
        .unwrap()
        .set_bold(true);
    let formatted_xml = document_xml(&mut formatted);
    let formatted_positions = [
        formatted_xml.find("<w:rPr>").unwrap(),
        formatted_xml.find("<w:rStyle>").unwrap(),
        formatted_xml.find("<w:b/>").unwrap(),
        formatted_xml.find("<w:t>A</w:t>").unwrap(),
    ];
    assert!(formatted_positions.windows(2).all(|pair| pair[0] < pair[1]));
}

#[test]
fn comment_removal_remaps_direct_and_control_run_raw_boundaries() {
    let paragraphs = r#"<w:p><w:r><w:directBefore/><w:t>A</w:t><w:directMiddle/><w:commentReference w:id="0"/><w:directAfter/><w:t>B</w:t><w:directLast/></w:r></w:p><w:p><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:controlBefore/><w:t>C</w:t><w:controlMiddle/><w:commentReference w:id="0"/><w:controlAfter/><w:t>D</w:t><w:controlLast/></w:r></w:sdtContent></w:sdt></w:p>"#;
    let (mut document, comment_id) = document_with_comment_paragraphs(paragraphs);

    assert!(document.remove_comment(comment_id).unwrap());
    let output = document_xml(&mut document);
    assert!(!output.contains("commentReference"));
    for ordered in [
        [
            "<w:directBefore/>",
            "<w:t>A</w:t>",
            "<w:directMiddle/>",
            "<w:directAfter/>",
            "<w:t>B</w:t>",
            "<w:directLast/>",
        ],
        [
            "<w:controlBefore/>",
            "<w:t>C</w:t>",
            "<w:controlMiddle/>",
            "<w:controlAfter/>",
            "<w:t>D</w:t>",
            "<w:controlLast/>",
        ],
    ] {
        let positions = ordered
            .iter()
            .map(|needle| output.find(needle).unwrap())
            .collect::<Vec<_>>();
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
    }

    let mut reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let reopened_xml = document_xml(&mut reopened);
    for raw in [
        "<w:directBefore/>",
        "<w:directMiddle/>",
        "<w:directAfter/>",
        "<w:directLast/>",
        "<w:controlBefore/>",
        "<w:controlMiddle/>",
        "<w:controlAfter/>",
        "<w:controlLast/>",
    ] {
        assert_eq!(reopened_xml.matches(raw).count(), 1, "raw child {raw}");
    }
}

#[test]
fn comment_insertion_keeps_content_at_the_hyperlink_end_boundary() {
    let raw = r#"<x:end xmlns:x="urn:opaque"/>"#;
    let xml = wrap_word_body(&format!(
        r#"<w:p><w:hyperlink><w:r><w:t>one</w:t></w:r><w:r><w:t>two</w:t></w:r><w:ins w:id="43" w:author="Ada"><w:r><w:t>end</w:t></w:r></w:ins>{raw}</w:hyperlink></w:p>"#
    ));
    let mut document = document_with_content_controls(&xml);
    let comment_id = document
        .add_comment(
            RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 0,
                    run_index: 2,
                },
            },
            "Ada",
            None,
            "review",
        )
        .unwrap();

    let inserted = document_xml(&mut document);
    let revision = inserted.find(r#"w:id="43""#).unwrap();
    let raw_position = inserted.find(raw).unwrap();
    let hyperlink_end = inserted.find("</w:hyperlink>").unwrap();
    let reference = inserted.find("<w:commentReference").unwrap();
    assert!(revision < raw_position && raw_position < hyperlink_end && hyperlink_end < reference);

    assert!(document.remove_comment(comment_id).unwrap());
    let removed = document_xml(&mut document);
    assert!(removed.contains(r#"w:id="43""#));
    assert!(removed.contains(raw));
    assert_eq!(document.accept_revision_id(43).unwrap(), 1);
    assert_eq!(document.text(), "onetwoend\n");
}

#[test]
fn hyperlink_relationship_ids_use_expanded_names_and_safe_output_prefixes() {
    let relationship_namespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let xml = wrap_word_body(&format!(
        r#"<w:p><w:hyperlink xmlns:r="urn:foreign" xmlns:rel="{relationship_namespace}" xmlns:x="urn:other" r:id="wrong" rel:id="right" x:id="other"><w:r><w:t>one</w:t></w:r></w:hyperlink><w:hyperlink xmlns:r="urn:foreign" r:id="still-wrong"><w:r><w:t>two</w:t></w:r></w:hyperlink></w:p>"#
    ));
    let mut document = document_with_content_controls(&xml);

    let paragraph = document.paragraph(0).unwrap();
    let spans = paragraph.hyperlink_spans();
    assert_eq!(spans[0].2, Some("right"));
    assert_eq!(spans[1].2, None);
    let output = document_xml(&mut document);
    assert!(output.contains(r#"xmlns:r="urn:foreign""#));
    assert!(output.contains(r#"r:id="wrong""#));
    assert!(output.contains(r#"r:id="still-wrong""#));
    assert!(output.contains(r#"x:id="other""#));
    assert!(output.contains(
        r#"xmlns:rdocxR="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
    ));
    assert!(output.contains(r#"rdocxR:id="right""#));

    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let paragraph = reopened.paragraph(0).unwrap();
    let spans = paragraph.hyperlink_spans();
    assert_eq!(spans[0].2, Some("right"));
    assert_eq!(spans[1].2, None);
}

#[test]
fn content_control_display_insertion_keeps_trailing_raw_after_the_value() {
    let xml = wrap_word_body(
        r#"<w:sdt><w:sdtPr><w:tag w:val="customer"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:before/><w:tab/><w:after/></w:r></w:p></w:sdtContent></w:sdt>"#,
    );
    let mut document = document_with_content_controls(&xml);

    assert_eq!(
        document
            .set_content_control_value_by_tag("customer", "Ada")
            .unwrap(),
        1
    );
    let output = document_xml(&mut document);
    let positions = [
        output.find("<w:before/>").unwrap(),
        output.find("<w:t>Ada</w:t>").unwrap(),
        output.find("<w:after/>").unwrap(),
    ];
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));

    let mut reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let reopened_xml = document_xml(&mut reopened);
    assert_eq!(reopened_xml.matches("<w:before/>").count(), 1);
    assert_eq!(reopened_xml.matches("<w:after/>").count(), 1);
}

#[test]
fn tag_precedes_alias_and_each_control_updates_once() {
    let xml = wrap_word_body(
        r#"<w:sdt><w:sdtPr><w:tag w:val="customer"/><w:alias w:val="shared"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>one</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sdt><w:sdtPr><w:alias w:val="shared"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>two</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sdt><w:sdtPr><w:tag w:val="missing"/><w:alias w:val="shared"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>three</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
    );
    let mut document = document_with_content_controls(&xml);
    let values = HashMap::from([
        ("customer".to_string(), "tag value".to_string()),
        ("shared".to_string(), "alias value".to_string()),
    ]);

    assert_eq!(document.bind_content_controls(&values).unwrap(), 3);
    let controls = document.content_controls();
    assert_eq!(
        controls
            .iter()
            .map(|control| control.text())
            .collect::<Vec<_>>(),
        ["tag value", "alias value", "alias value"]
    );
    assert_eq!(controls[0].tag(), Some("customer"));
    assert_eq!(controls[0].alias(), Some("shared"));
    assert_eq!(controls[0].id(), None);
    assert_eq!(
        controls[0].control_type(),
        Some(rdocx_oxml::SdtType::PlainText)
    );
    assert_eq!(document.content_controls_by_tag("customer").len(), 1);
    assert_eq!(document.content_controls_by_alias("shared").len(), 3);

    let mut alias_document = document_with_content_controls(&xml);
    assert_eq!(
        alias_document
            .set_content_control_value_by_alias("shared", "direct alias")
            .unwrap(),
        3
    );
    assert!(
        alias_document
            .content_controls()
            .iter()
            .all(|control| control.text() == "direct alias")
    );
}

#[test]
fn a_control_map_updates_every_matching_display_value() {
    let xml = wrap_word_body(
        r#"<w:sdt><w:sdtPr><w:tag w:val="outer"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>old outer</w:t><w:tab/><w:br/></w:r></w:p><w:sdt><w:sdtPr><w:tag w:val="inner"/><w:text/><w:temporary w:val="1"/></w:sdtPr><w:sdtContent><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>old inner</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sdt><w:sdtPr><w:tag w:val="unrelated-control"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>nested untouched</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt><w:p><w:r><w:t>unrelated</w:t></w:r></w:p>"#,
    );
    let mut document = document_with_content_controls(&xml);
    let values = HashMap::from([
        ("outer".to_string(), "new outer".to_string()),
        ("inner".to_string(), "new inner".to_string()),
    ]);

    assert_eq!(document.bind_content_controls(&values).unwrap(), 2);
    let bytes = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let saved = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(saved.contains("<w:b/>"), "{saved}");
    assert!(saved.contains("<w:t>new outer</w:t>"), "{saved}");
    assert!(saved.contains("<w:i/>"), "{saved}");
    assert!(saved.contains("<w:t>new inner</w:t>"), "{saved}");
    assert!(saved.contains("<w:temporary w:val=\"1\"/>"), "{saved}");
    assert!(saved.contains("<w:t>nested untouched</w:t>"), "{saved}");
    assert!(saved.contains("<w:t>unrelated</w:t>"), "{saved}");
    assert!(!saved.contains("<w:tab"), "{saved}");
    assert!(!saved.contains("<w:br"), "{saved}");
    let controls = document.content_controls();
    assert_eq!(controls[0].text(), "new outer");
    assert_eq!(controls[1].text(), "new inner");
    assert_eq!(controls[2].text(), "nested untouched");
}

#[test]
fn a_bound_custom_xml_value_updates_the_part_and_display_text_atomically() {
    let xml = wrap_word_body(
        r#"<w:sdt><w:sdtPr><w:tag w:val="customer"/><w:dataBinding w:storeItemID="{11111111-1111-1111-1111-111111111111}" w:xpath="/c:root/c:customer[2]/c:name" w:prefixMappings="xmlns:c='urn:customer'"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>old display</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
    );
    let custom_xml = r#"<?producer keep?><c:root xmlns:c="urn:customer" keep="same"><c:customer code="A"><c:name>First</c:name></c:customer><c:customer xmlns:c="urn:other"><c:name>shadow</c:name></c:customer><!--marker--><c:customer code="B"><c:name old="1">Old &amp; text</c:name><c:other>untouched</c:other></c:customer></c:root>"#;
    let mut document = document_with_bound_content_controls(&xml, Some(custom_xml));

    assert_eq!(
        document
            .set_content_control_value_by_tag("customer", "Ada & Co")
            .unwrap(),
        1
    );
    let bytes = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&bytes)).unwrap();
    assert_eq!(
        package.get_part("/customXml/item1.xml").unwrap(),
        br#"<?producer keep?><c:root xmlns:c="urn:customer" keep="same"><c:customer code="A"><c:name>First</c:name></c:customer><c:customer xmlns:c="urn:other"><c:name>shadow</c:name></c:customer><!--marker--><c:customer code="B"><c:name old="1">Ada &amp; Co</c:name><c:other>untouched</c:other></c:customer></c:root>"#
    );

    let reopened = Document::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.content_controls()[0].text(), "Ada & Co");

    let document_xml =
        std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    let tag = document_xml.find("<w:tag").unwrap();
    let binding = document_xml.find("<w:dataBinding").unwrap();
    let control_type = document_xml.find("<w:text/>").unwrap();
    assert!(tag < binding && binding < control_type);
}

#[test]
fn an_invalid_binding_changes_neither_document_nor_custom_xml() {
    let cases = [
        (
            "{22222222-2222-2222-2222-222222222222}",
            "/c:root/c:name",
            "xmlns:c='urn:customer'",
        ),
        (
            "{11111111-1111-1111-1111-111111111111}",
            "//c:name",
            "xmlns:c='urn:customer'",
        ),
        (
            "{11111111-1111-1111-1111-111111111111}",
            "/c:root/c:name",
            "xmlns:c='urn:customer'",
        ),
        (
            "{11111111-1111-1111-1111-111111111111}",
            "/c:root/c:name[1]",
            "xmlns:c=urn:customer",
        ),
    ];
    let custom_xml =
        r#"<c:root xmlns:c="urn:customer"><c:name>one</c:name><c:name>two</c:name></c:root>"#;

    for (store_item_id, xpath, prefix_mappings) in cases {
        let xml = wrap_word_body(&format!(
            r#"<w:sdt><w:sdtPr><w:tag w:val="customer"/><w:dataBinding w:storeItemID="{{11111111-1111-1111-1111-111111111111}}" w:xpath="/c:root/c:name[1]" w:prefixMappings="xmlns:c='urn:customer'"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>valid unchanged</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sdt><w:sdtPr><w:tag w:val="customer"/><w:dataBinding w:storeItemID="{store_item_id}" w:xpath="{xpath}" w:prefixMappings="{prefix_mappings}"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>invalid unchanged</w:t></w:r></w:p></w:sdtContent></w:sdt>"#
        ));
        let mut document = document_with_bound_content_controls(&xml, Some(custom_xml));
        let before = document.to_bytes().unwrap();

        assert!(
            document
                .set_content_control_value_by_tag("customer", "changed")
                .is_err(),
            "invalid binding unexpectedly succeeded for {store_item_id} {xpath}"
        );
        assert_eq!(document.to_bytes().unwrap(), before);
    }

    let xml = wrap_word_body(
        r#"<w:sdt><w:sdtPr><w:tag w:val="customer"/><w:dataBinding w:storeItemID="{11111111-1111-1111-1111-111111111111}" w:xpath="/c:root/c:name[1]" w:prefixMappings="xmlns:c='urn:customer'"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>unchanged</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
    );
    let invalid_properties = [
        br#"<ds:datastoreItem bad:itemID="{11111111-1111-1111-1111-111111111111}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" xmlns:bad="urn:producer"><ds:schemaRefs/></ds:datastoreItem>"#
            .as_slice(),
        br#"<bad:wrapper xmlns:bad="urn:producer" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:datastoreItem ds:itemID="{11111111-1111-1111-1111-111111111111}"/></bad:wrapper>"#
            .as_slice(),
    ];
    for properties_xml in invalid_properties {
        let mut document = document_with_bound_content_controls(&xml, Some(custom_xml));
        let bytes = document.to_bytes().unwrap();
        let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        package.set_part("/customXml/itemProps1.xml", properties_xml.to_vec());
        let mut output = std::io::Cursor::new(Vec::new());
        package.write_to(&mut output).unwrap();
        let mut document = Document::from_bytes(output.get_ref()).unwrap();
        let before = document.to_bytes().unwrap();
        assert!(
            document
                .set_content_control_value_by_tag("customer", "changed")
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), before);
    }
}

#[test]
fn run_ranges_reject_reverse_missing_and_nonparagraph_positions() {
    let mut document = Document::new();
    document.add_paragraph("first");
    document.add_table(1, 1);
    document.add_paragraph("second");
    let before = document.to_bytes().unwrap();

    let invalid = [
        RunRange {
            start: RunPosition {
                body_index: 2,
                run_index: 0,
            },
            end: RunPosition {
                body_index: 0,
                run_index: 1,
            },
        },
        RunRange {
            start: RunPosition {
                body_index: 99,
                run_index: 0,
            },
            end: RunPosition {
                body_index: 99,
                run_index: 0,
            },
        },
        RunRange {
            start: RunPosition {
                body_index: 1,
                run_index: 0,
            },
            end: RunPosition {
                body_index: 1,
                run_index: 0,
            },
        },
    ];

    for range in invalid {
        assert!(
            document
                .add_comment(range, "Ada", Some("AL"), "review")
                .is_err()
        );
    }
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn a_ranged_comment_reply_and_resolution_keep_one_intact_thread() {
    let mut document = Document::new();
    let mut paragraph = document.add_paragraph("");
    paragraph.add_run("first ");
    paragraph.add_run("second");

    let comment_id = document
        .add_comment(
            RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 0,
                    run_index: 2,
                },
            },
            "Ada",
            Some("AL"),
            "Please review",
        )
        .unwrap();
    let reply_id = document.reply_to(comment_id, "Ben", "Looks good").unwrap();
    assert!(document.resolve_comment(comment_id, true).unwrap());

    let bytes = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    let comments = reopened.comments();
    assert_eq!(comments.len(), 2);
    assert_eq!(comments[0].id(), comment_id);
    assert_eq!(comments[0].text(), "Please review");
    assert!(comments[0].resolved());
    assert_eq!(comments[1].id(), reply_id);
    assert_eq!(comments[1].text(), "Looks good");
    assert_eq!(comments[1].parent_id(), Some(comment_id));
}

#[test]
fn removing_a_comment_removes_only_its_anchors_and_thread_metadata() {
    let mut document = Document::new();
    let mut paragraph = document.add_paragraph("");
    paragraph.add_run("left ");
    paragraph.add_run("middle ");
    paragraph.add_run("right");
    let first = document
        .add_comment(
            RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 0,
                    run_index: 1,
                },
            },
            "Ada",
            None,
            "remove me",
        )
        .unwrap();
    let second = document
        .add_comment(
            RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 2,
                },
                end: RunPosition {
                    body_index: 0,
                    run_index: 3,
                },
            },
            "Ben",
            None,
            "keep me",
        )
        .unwrap();

    let bytes = document.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let extension = String::from_utf8(
        package
            .get_part("/word/commentsExtended.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap()
    .replace(
        "<w15:commentsEx ",
        "<w15:commentsEx xmlns:ext=\"urn:producer\" ",
    )
    .replace(
        "</w15:commentsEx>",
        "<ext:kept token=\"same\"/></w15:commentsEx>",
    );
    package.set_part("/word/commentsExtended.xml", extension.into_bytes());
    let mut seeded = std::io::Cursor::new(Vec::new());
    package.write_to(&mut seeded).unwrap();

    let mut reopened = Document::from_bytes(seeded.get_ref()).unwrap();
    assert!(reopened.remove_comment(first).unwrap());
    let saved = reopened.to_bytes().unwrap();
    let saved_package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let document_xml = String::from_utf8(
        saved_package
            .get_part("/word/document.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    let extension_xml = String::from_utf8(
        saved_package
            .get_part("/word/commentsExtended.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();

    assert_eq!(reopened.text(), "left middle right\n");
    assert_eq!(reopened.comments().len(), 1);
    assert_eq!(reopened.comments()[0].id(), second);
    assert!(!document_xml.contains(&format!("w:id=\"{first}\"")));
    assert!(document_xml.contains(&format!("w:id=\"{second}\"")));
    assert!(extension_xml.contains("<ext:kept token=\"same\"/>"));
}

#[test]
fn comment_mutations_keep_bookmarks_and_run_controls_at_their_boundaries() {
    let xml = wrap_word_body(
        r#"<w:p><w:bookmarkStart w:id="4" w:name="destination"/><w:r><w:t>left</w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="run-control"/></w:sdtPr><w:sdtContent><w:r><w:t>wrapped</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>right</w:t></w:r><w:bookmarkEnd w:id="4"/></w:p>"#,
    );
    let mut document = document_with_content_controls(&xml);
    assert_eq!(document.bookmarks()[0].range().unwrap().end.run_index, 3);
    assert_eq!(document.bookmarks()[0].text(), "leftwrappedright");

    let comment_id = document
        .add_comment(
            RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 0,
                    run_index: 1,
                },
            },
            "Ada",
            None,
            "Review",
        )
        .unwrap();
    assert_eq!(document.bookmarks()[0].range().unwrap().end.run_index, 4);
    assert_eq!(document.bookmarks()[0].text(), "leftwrappedright");
    assert_eq!(document.content_controls()[0].text(), "wrapped");

    let bytes = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let document_xml =
        String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
    let end = document_xml.find("<w:commentRangeEnd").unwrap();
    let reference = document_xml.find("<w:commentReference").unwrap();
    let control = document_xml.find("<w:sdt>").unwrap();
    let right = document_xml.find("<w:t>right</w:t>").unwrap();
    let bookmark_end = document_xml.find("<w:bookmarkEnd").unwrap();
    assert!(end < reference && reference < control && control < right && right < bookmark_end);

    let mut reopened = Document::from_bytes(
        &document
            .to_bytes()
            .expect("serialize document with comment boundaries"),
    )
    .unwrap();
    assert!(reopened.remove_comment(comment_id).unwrap());
    assert_eq!(reopened.bookmarks()[0].range().unwrap().end.run_index, 3);
    assert_eq!(reopened.bookmarks()[0].text(), "leftwrappedright");
    assert_eq!(reopened.content_controls()[0].text(), "wrapped");
    let saved = reopened.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let document_xml =
        String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
    assert!(!document_xml.contains("commentRange"));
    assert!(!document_xml.contains("commentReference"));
    let left = document_xml.find("<w:t>left</w:t>").unwrap();
    let control = document_xml.find("<w:sdt>").unwrap();
    let right = document_xml.find("<w:t>right</w:t>").unwrap();
    let bookmark_end = document_xml.find("<w:bookmarkEnd").unwrap();
    assert!(left < control && control < right && right < bookmark_end);
}

#[test]
fn new_bookmark_after_an_accepted_control_has_live_accepted_coordinates() {
    let body = r#"
        <w:p><w:fldSimple w:instr="REF destination"><w:r><w:t>stale</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:sdt><w:sdtContent><w:r><w:t>wrapped prefix</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>direct target</w:t></w:r></w:p>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);

    document
        .add_bookmark(
            "destination",
            RunRange {
                start: RunPosition {
                    body_index: 1,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 1,
                    run_index: 1,
                },
            },
        )
        .unwrap();
    let bookmark = &document.bookmarks()[0];
    assert_eq!(bookmark.text(), "direct target");
    assert_eq!(bookmark.range().unwrap().start.run_index, 1);
    assert_eq!(bookmark.range().unwrap().end.run_index, 2);

    let page = document.layout_page(0).unwrap().unwrap();
    let text = compatibility_page_elements(&page.elements)
        .into_iter()
        .filter_map(|element| match element {
            oxml_layout::PositionedElement::Text(run) => Some(run.text.as_str()),
            _ => None,
        })
        .collect::<String>();
    assert!(text.contains("direct target"), "{text}");
    assert!(!text.contains("stale"), "{text}");
}

#[test]
fn removing_a_comment_recurses_through_every_typed_control_placement() {
    let mut document = Document::new();
    document.add_paragraph("anchor");
    let comment_id = document
        .add_comment(
            RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 0,
                    run_index: 1,
                },
            },
            "Ada",
            None,
            "Nested",
        )
        .unwrap();
    let bytes = document.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let mut document_xml =
        String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();

    let reference = document_xml.find("<w:commentReference").unwrap();
    let reference_run_start = document_xml[..reference].rfind("<w:r").unwrap();
    let reference_run_end =
        reference + document_xml[reference..].find("</w:r>").unwrap() + "</w:r>".len();
    let reference_run = document_xml[reference_run_start..reference_run_end].to_owned();
    document_xml.replace_range(
        reference_run_start..reference_run_end,
        &format!("<w:sdt><w:sdtContent>{reference_run}</w:sdtContent></w:sdt>"),
    );

    let paragraph_start = document_xml.find("<w:p").unwrap();
    let paragraph_end =
        paragraph_start + document_xml[paragraph_start..].find("</w:p>").unwrap() + "</w:p>".len();
    let paragraph = document_xml[paragraph_start..paragraph_end].to_owned();
    let nested = format!(
        "<w:sdt><w:sdtContent><w:sdt><w:sdtContent><w:tbl><w:sdt><w:sdtContent><w:tr><w:sdt><w:sdtContent><w:tc><w:sdt><w:sdtContent>{paragraph}</w:sdtContent></w:sdt></w:tc></w:sdtContent></w:sdt></w:tr></w:sdtContent></w:sdt></w:tbl></w:sdtContent></w:sdt></w:sdtContent></w:sdt>"
    );
    document_xml.replace_range(paragraph_start..paragraph_end, &nested);
    package.set_part("/word/document.xml", document_xml.into_bytes());
    let mut output = std::io::Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();

    let mut reopened = Document::from_bytes(output.get_ref()).unwrap();
    assert_eq!(reopened.comments().len(), 1);
    assert!(reopened.remove_comment(comment_id).unwrap());
    assert!(reopened.comments().is_empty());
    assert_eq!(reopened.paragraphs()[0].text(), "anchor");
    let saved = reopened.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let document_xml =
        String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
    assert!(!document_xml.contains("commentRange"));
    assert!(!document_xml.contains("commentReference"));
}

#[test]
fn mutable_document_indexes_match_recursive_immutable_indexes() {
    let xml = wrap_word_body(
        r#"<w:p><w:r><w:t>one</w:t></w:r></w:p><w:sdt><w:sdtContent><w:p><w:r><w:t>wrapped</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>three</w:t></w:r></w:p><w:sdt><w:sdtContent><w:tbl><w:tr><w:tc><w:p><w:r><w:t>table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:sdtContent></w:sdt><w:tbl><w:tr><w:tc><w:p><w:r><w:t>direct table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let mut document = document_with_content_controls(&xml);
    assert_eq!(
        document
            .paragraphs()
            .iter()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>(),
        ["one", "wrapped", "three", "table cell"]
    );
    assert_eq!(document.tables().len(), 2);

    document
        .paragraph_mut(1)
        .expect("wrapped paragraph uses immutable index order")
        .add_run(" changed");
    document
        .paragraph_mut(3)
        .expect("deep wrapped paragraph uses immutable index order")
        .add_run(" deep");
    document
        .table_mut(0)
        .expect("wrapped table uses immutable index order")
        .set_style("WrappedTable");

    assert_eq!(document.paragraph(1).unwrap().text(), "wrapped changed");
    assert_eq!(document.paragraph(3).unwrap().text(), "table cell deep");
    assert_eq!(document.table(0).unwrap().style_id(), Some("WrappedTable"));
    assert_eq!(document.table(1).unwrap().style_id(), None);
}

#[test]
fn mislabelled_jpeg_uses_sniffed_package_metadata() {
    let jpeg = [0xff, 0xd8, 0xff, 0xd9];
    let mut seed = Document::new();
    let seed_bytes = seed.to_bytes().unwrap();
    let mut seed_package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(seed_bytes)).unwrap();
    seed_package
        .content_types
        .add_default("jpeg", "application/octet-stream");
    let mut reopened_bytes = std::io::Cursor::new(Vec::new());
    seed_package.write_to(&mut reopened_bytes).unwrap();
    let mut document = Document::from_bytes(&reopened_bytes.into_inner()).unwrap();

    document.add_picture(
        &jpeg,
        "mislabelled.png",
        Length::inches(1.0),
        Length::inches(1.0),
    );

    let bytes = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let part_name = "/word/media/image1.jpeg";

    assert_eq!(package.get_part(part_name), Some(jpeg.as_slice()));
    assert_eq!(
        package.content_types.content_type_for(part_name),
        Some("image/jpeg")
    );

    let image_relationship = package
        .get_part_rels("/word/document.xml")
        .and_then(|relationships| {
            relationships.get_by_type(oxml_opc::relationship::rel_types::IMAGE)
        })
        .expect("document should relate the sniffed image part");
    assert_eq!(image_relationship.target, "media/image1.jpeg");
}

#[test]
fn next_image_name_uses_the_highest_existing_index_not_the_part_count() {
    let cases: &[(&[&str], &str)] = &[
        (
            &["/word/media/image1.png", "/word/media/image5.png"],
            "/word/media/image6.png",
        ),
        (
            &[
                "/word/media/image1.png",
                "/word/media/image2.png",
                "/word/media/image4.png",
            ],
            "/word/media/image5.png",
        ),
    ];

    for (existing_names, expected_name) in cases {
        let mut seed = Document::new();
        let seed_bytes = seed.to_bytes().unwrap();
        let mut package =
            oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(seed_bytes)).unwrap();
        for name in *existing_names {
            package.set_part(name, vec![0xAA]);
        }

        let mut package_bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut package_bytes).unwrap();
        let mut reopened = Document::from_bytes(&package_bytes.into_inner()).unwrap();
        reopened.add_picture(
            &[0x11, 0x22, 0x33],
            "added.png",
            Length::inches(1.0),
            Length::inches(1.0),
        );

        let saved = reopened.to_bytes().unwrap();
        let saved_package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        assert_eq!(
            saved_package.get_part(expected_name),
            Some([0x11, 0x22, 0x33].as_slice()),
            "existing names: {existing_names:?}"
        );
    }
}

#[test]
fn malformed_media_names_do_not_change_the_highest_image_index() {
    let mut seed = Document::new();
    let seed_bytes = seed.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(seed_bytes)).unwrap();
    for name in [
        "/word/media/image4.png",
        "/word/media/image.png",
        "/word/media/imagezero.png",
        "/word/media/image-7.png",
        "/word/media/image0.png",
        "/word/media/images99.png",
        "/ppt/media/image99.png",
    ] {
        package.set_part(name, vec![0xAA]);
    }

    let mut package_bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut package_bytes).unwrap();
    let mut reopened = Document::from_bytes(&package_bytes.into_inner()).unwrap();
    reopened.add_picture(
        &[0x44, 0x55, 0x66],
        "added.png",
        Length::inches(1.0),
        Length::inches(1.0),
    );

    let saved = reopened.to_bytes().unwrap();
    let saved_package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    assert_eq!(
        saved_package.get_part("/word/media/image5.png"),
        Some([0x44, 0x55, 0x66].as_slice())
    );
}

#[test]
fn occupied_max_image_suffix_wraps_to_a_free_low_number() {
    let mut document = Document::new();
    let bytes = document.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let lower_name = format!("/word/media/image{}.png", usize::MAX - 1);
    let max_name = format!("/word/media/image{}.png", usize::MAX);

    package.set_part("/word/media/image1.png", vec![0xaa]);
    package.set_part(&lower_name, vec![0xbb]);
    package.set_part(&max_name, vec![0xbb]);

    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let mut reopened = Document::from_bytes(&input.into_inner()).unwrap();
    reopened.add_picture(
        &[0x44, 0x55, 0x66],
        "added.png",
        Length::inches(1.0),
        Length::inches(1.0),
    );

    let saved = reopened.to_bytes().unwrap();
    let saved_package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();

    assert_eq!(
        saved_package.get_part("/word/media/image2.png"),
        Some([0x44, 0x55, 0x66].as_slice())
    );
    assert!(saved_package.get_part("/word/media/image0.png").is_none());
    assert_eq!(
        saved_package.get_part("/word/media/image1.png"),
        Some([0xaa].as_slice())
    );
    assert_eq!(saved_package.get_part(&lower_name), Some([0xbb].as_slice()));
    assert_eq!(saved_package.get_part(&max_name), Some([0xbb].as_slice()));
}

#[test]
fn max_minus_one_allocates_max_then_wraps_safely() {
    let mut document = Document::new();
    let bytes = document.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let lower_name = format!("/word/media/image{}.png", usize::MAX - 1);
    let max_name = format!("/word/media/image{}.png", usize::MAX);

    package.set_part(&lower_name, vec![0xaa]);

    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let mut reopened = Document::from_bytes(&input.into_inner()).unwrap();
    reopened.add_picture(
        &[0x11, 0x22, 0x33],
        "first.png",
        Length::inches(1.0),
        Length::inches(1.0),
    );
    reopened.add_picture(
        &[0x44, 0x55, 0x66],
        "second.png",
        Length::inches(1.0),
        Length::inches(1.0),
    );

    let saved = reopened.to_bytes().unwrap();
    let saved_package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();

    assert_eq!(saved_package.get_part(&lower_name), Some([0xaa].as_slice()));
    assert_eq!(
        saved_package.get_part(&max_name),
        Some([0x11, 0x22, 0x33].as_slice())
    );
    assert_eq!(
        saved_package.get_part("/word/media/image1.png"),
        Some([0x44, 0x55, 0x66].as_slice())
    );
    assert!(saved_package.get_part("/word/media/image0.png").is_none());
}

/// A replacement that contains the placeholder used to restart the search from
/// offset 0 and match its own output forever.
#[test]
fn replacement_containing_the_placeholder_terminates() {
    let mut doc = Document::new();
    doc.add_paragraph("Hello NAME!");

    let count = doc.replace_text("NAME", "NAME Smith");

    assert_eq!(count, 1);
    assert_eq!(doc.paragraphs()[0].text(), "Hello NAME Smith!");
}

#[test]
fn repeated_placeholders_are_all_replaced() {
    let mut doc = Document::new();
    doc.add_paragraph("a X b X c X d");

    let count = doc.replace_text("X", "Y");

    assert_eq!(count, 3);
    assert_eq!(doc.paragraphs()[0].text(), "a Y b Y c Y d");
}

#[test]
fn overlapping_replacement_does_not_rescan_its_own_output() {
    let mut doc = Document::new();
    doc.add_paragraph("aaa");

    let count = doc.replace_text("a", "aa");

    assert_eq!(count, 3, "each source 'a' should be replaced exactly once");
    assert_eq!(doc.paragraphs()[0].text(), "aaaaaa");
}

/// The regex path had the same non-termination hazard, plus one for patterns
/// that can match the empty string.
#[test]
fn regex_replacement_containing_the_pattern_terminates() {
    let mut doc = Document::new();
    doc.add_paragraph("value: 42");

    let count = doc.replace_regex(r"\d+", "[$0]").unwrap();

    assert_eq!(count, 1);
    assert_eq!(doc.paragraphs()[0].text(), "value: [42]");
}

#[test]
fn zero_width_regex_match_terminates() {
    let mut doc = Document::new();
    doc.add_paragraph("abc");

    // `x*` matches the empty string at every position.
    let count = doc.replace_regex("x*", "-").unwrap();

    assert!(count <= 4, "should not loop indefinitely, got {count}");
}

/// `9360 / cols` panicked when a caller asked for a zero-column table.
#[test]
fn zero_column_tables_do_not_panic() {
    let mut doc = Document::new();

    let table = doc.add_table(2, 0);
    assert_eq!(table.row_count(), 2);

    doc.insert_table(0, 1, 0);

    let mut with_cell = doc.add_table(1, 1);
    let mut cell = with_cell.cell(0, 0).unwrap();
    cell.add_table(1, 0);
}

/// Styles were always written to `/word/styles.xml`, so a document without a
/// styles part gained an orphan that Word would ignore.
#[test]
fn styles_part_is_reachable_after_save() {
    let mut doc = Document::new();
    doc.add_paragraph("Body");
    let bytes = doc.to_bytes().unwrap();

    let pkg = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();

    let rels = pkg
        .get_part_rels("/word/document.xml")
        .expect("document should have relationships");
    let styles_rel = rels
        .get_by_type(oxml_opc::relationship::rel_types::STYLES)
        .expect("styles relationship must exist");
    let target = oxml_opc::OpcPackage::resolve_rel_target("/word/document.xml", &styles_rel.target);

    assert!(
        pkg.get_part(&target).is_some(),
        "styles relationship must point at a part that exists"
    );
    assert!(
        pkg.content_types.content_type_for(&target).is_some(),
        "styles part needs a content type"
    );
}

/// Adding a list twice must not produce two numbering relationships.
#[test]
fn numbering_relationship_is_added_once() {
    let mut doc = Document::new();
    doc.add_bullet_list_item("first", 0);
    doc.add_numbered_list_item("second", 0);
    let bytes = doc.to_bytes().unwrap();

    let pkg = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let rels = pkg.get_part_rels("/word/document.xml").unwrap();
    let numbering_rels = rels.get_all_by_type(oxml_opc::relationship::rel_types::NUMBERING);

    assert_eq!(
        numbering_rels.len(),
        1,
        "expected exactly one numbering rel"
    );
}

/// Saving the same document twice must produce identical bytes.
#[test]
fn saving_is_reproducible() {
    let mut doc = Document::new();
    doc.add_paragraph("Reproducible");
    for i in 0..25 {
        doc.add_picture(
            &[0u8, 1, 2, 3, i],
            &format!("img{i}.png"),
            Length::inches(1.0),
            Length::inches(1.0),
        );
    }

    let first = doc.to_bytes().unwrap();
    let second = doc.to_bytes().unwrap();

    assert_eq!(first, second);
}

/// Two documents built the same way must produce identical deterministic PDFs.
///
/// The PDF writer keyed its prepared fonts, its font references and its glyph
/// to Unicode table on hashed maps, and iterated all three to write the file.
/// Iteration order differs between map instances, so the same document written
/// twice differed in its `/Font` dictionary order, its font object order and
/// its ToUnicode CMap line order. Nothing about the rendered page changed and
/// nothing failed, which is why it survived until the hash harness went looking.
///
/// Two separately built documents are the point. One document written twice
/// reuses the same map instances and cannot see this.
#[test]
fn two_identical_documents_produce_identical_deterministic_pdfs() {
    let build = || {
        let mut doc = Document::new();
        doc.add_paragraph("A heading in the document");
        let mut styled = doc.add_paragraph("");
        styled.add_run("bold text").bold(true);
        styled.add_run(" and italic text").italic(true);
        styled.add_run(" and plain text");
        doc.add_paragraph("A closing paragraph with enough words to shape.");
        doc
    };

    let first = build()
        .to_pdf_deterministic()
        .expect("deterministic PDF rendering should succeed");
    let second = build()
        .to_pdf_deterministic()
        .expect("deterministic PDF rendering should succeed");

    assert!(first.starts_with(b"%PDF-"));
    assert_eq!(
        first, second,
        "the same document produced two different PDFs"
    );
}

/// Batching must produce the same result as replacing one field at a time.
#[test]
fn batch_replacement_matches_sequential() {
    let build = || {
        let mut doc = Document::new();
        doc.add_paragraph("Dear {{name}}, your order {{order}} ships {{date}}.");
        doc
    };

    let mut sequential = build();
    let mut n = sequential.replace_text("{{name}}", "Ada");
    n += sequential.replace_text("{{order}}", "A-1");
    n += sequential.replace_text("{{date}}", "Friday");

    let mut batched = build();
    let map = HashMap::from([
        ("{{name}}", "Ada"),
        ("{{order}}", "A-1"),
        ("{{date}}", "Friday"),
    ]);
    let batch_count = batched.replace_all(&map);

    assert_eq!(n, batch_count);
    assert_eq!(
        sequential.paragraphs()[0].text(),
        batched.paragraphs()[0].text()
    );
    assert_eq!(
        batched.paragraphs()[0].text(),
        "Dear Ada, your order A-1 ships Friday."
    );
}

/// Entity references must survive a parse/serialise round trip.
#[test]
fn xml_entities_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Ampersand & <angle> \"quote\" 'apos'");
    doc.set_title("Title & Co. <tagged>");

    let bytes = doc.to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();

    assert_eq!(
        reopened.paragraphs()[0].text(),
        "Ampersand & <angle> \"quote\" 'apos'"
    );
    assert_eq!(reopened.title(), Some("Title & Co. <tagged>"));
}

/// A crafted font name or colour must not be able to break out of the `style`
/// attribute it is written into.
#[test]
fn hostile_run_properties_cannot_inject_html() {
    let mut doc = Document::new();
    {
        let mut para = doc.add_paragraph("");
        para.add_run("text")
            .font("Arial\" onmouseover=\"alert(1)")
            .color("red;} body{display:none} .x{");
    }

    let html = doc.to_html_fragment();

    assert!(!html.contains("onmouseover"), "attribute injection: {html}");
    assert!(!html.contains("display:none"), "css injection: {html}");
    assert!(html.contains("text"));
}

#[test]
fn a_bookmark_inserted_over_a_range_is_listed_with_its_text() {
    let mut document = Document::new();
    document.add_paragraph("before ").add_run("inside");
    document.add_paragraph(" across ").add_run("after");

    let id = document
        .add_bookmark(
            "destination",
            RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 1,
                },
                end: RunPosition {
                    body_index: 1,
                    run_index: 1,
                },
            },
        )
        .expect("valid bookmark range");

    let bookmarks = document.bookmarks();
    assert_eq!(bookmarks.len(), 1);
    assert_eq!(bookmarks[0].id(), Some(id));
    assert_eq!(bookmarks[0].name(), Some("destination"));
    assert_eq!(bookmarks[0].text(), "inside\n across ");
    assert_eq!(bookmarks[0].issue(), None);

    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.bookmarks()[0].text(), "inside\n across ");
}

#[test]
fn ref_and_pageref_resolve_to_the_bookmark_text_and_final_page() {
    let mut document = Document::new();
    document.add_paragraph("field-placeholder");
    document
        .add_paragraph("bookmarked text")
        .page_break_before(true);
    document
        .add_bookmark(
            "destination",
            RunRange {
                start: RunPosition {
                    body_index: 1,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 1,
                    run_index: 1,
                },
            },
        )
        .unwrap();

    let bytes = document.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let original_xml =
        String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
    let xml = original_xml.replace(
            "<w:r>\n        <w:t>field-placeholder</w:t>\n      </w:r>",
            r#"<w:fldSimple w:instr=" REF destination "><w:r><w:t>cached</w:t></w:r></w:fldSimple><w:r><w:t> page </w:t></w:r><w:fldSimple w:instr=" PAGEREF destination "><w:r><w:t>cached-page</w:t></w:r></w:fldSimple>"#,
        );
    assert_ne!(xml, original_xml, "{original_xml}");
    package.set_part("/word/document.xml", xml.into_bytes());
    let mut rewritten = std::io::Cursor::new(Vec::new());
    package.write_to(&mut rewritten).unwrap();

    let reopened = Document::from_bytes(rewritten.get_ref()).unwrap();
    let page = reopened.layout_page(0).unwrap().unwrap();
    let text = compatibility_page_elements(&page.elements)
        .into_iter()
        .filter_map(|element| match element {
            oxml_layout::PositionedElement::Text(run) => Some(run.text.as_str()),
            _ => None,
        })
        .collect::<String>();
    assert!(text.contains("bookmarked text"), "{text}");
    assert!(text.contains("page 2"), "{text}");
}

#[test]
fn missing_and_duplicate_bookmark_targets_fail_without_mutation() {
    let mut document = Document::new();
    document.add_paragraph("one");
    let range = RunRange {
        start: RunPosition {
            body_index: 0,
            run_index: 0,
        },
        end: RunPosition {
            body_index: 0,
            run_index: 1,
        },
    };
    document.add_bookmark("destination", range).unwrap();
    let before = document.to_bytes().unwrap();

    assert!(document.add_bookmark("destination", range).is_err());
    assert!(document.add_bookmark("_TocReserved", range).is_err());
    assert!(
        document
            .add_bookmark(
                "invalid",
                RunRange {
                    start: RunPosition {
                        body_index: 0,
                        run_index: 2,
                    },
                    end: range.end,
                },
            )
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn malformed_and_unmatched_bookmark_markers_are_reported_without_loss() {
    let mut document = Document::new();
    document.add_paragraph("content");
    let bytes = document.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let original_xml =
        String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
    let xml = original_xml
        .replace(
            "<w:r>\n        <w:t>content</w:t>\n      </w:r>",
            r#"<w:bookmarkStart w:name="missing-id"/><w:r><w:t>content</w:t></w:r><w:bookmarkEnd w:id="9"/>"#,
        );
    assert_ne!(xml, original_xml);
    package.set_part("/word/document.xml", xml.into_bytes());
    let mut rewritten = std::io::Cursor::new(Vec::new());
    package.write_to(&mut rewritten).unwrap();

    let mut reopened = Document::from_bytes(rewritten.get_ref()).unwrap();
    let bookmarks = reopened.bookmarks();
    assert_eq!(bookmarks.len(), 2);
    assert!(bookmarks.iter().all(|bookmark| bookmark.issue().is_some()));
    assert!(bookmarks.iter().any(|bookmark| bookmark.id().is_none()));
    assert!(bookmarks.iter().any(|bookmark| bookmark.id() == Some(9)));

    let round_trip = reopened.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(round_trip)).unwrap();
    let round_trip_xml =
        String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
    assert!(round_trip_xml.contains(r#"<w:bookmarkStart w:name="missing-id"/>"#));
    assert!(round_trip_xml.contains(r#"<w:bookmarkEnd w:id="9"/>"#));
}

#[test]
fn a_nested_loop_and_conditional_generate_the_expected_document() {
    let body = r#"
        <w:p><w:r><w:t>{% for group in groups %}</w:t></w:r></w:p>
        <w:p><w:r><w:t>Group {{ group.name }}</w:t></w:r></w:p>
        <w:p><w:r><w:t>{% if group.visible %}</w:t></w:r></w:p>
        <w:p><w:r><w:t>Visible {{ group.name }}</w:t></w:r></w:p>
        <w:p><w:r><w:t>{% endif %}</w:t></w:r></w:p>
        <w:tbl>
          <w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
          <w:tr><w:tc><w:p><w:r><w:t>{% for item in group.items %}</w:t></w:r></w:p></w:tc></w:tr>
          <w:tr><w:tc><w:p><w:r><w:t>{{ group.name }}:{{ item.label }}</w:t></w:r></w:p></w:tc></w:tr>
          <w:tr><w:tc><w:p><w:r><w:t>{% endfor %}</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>
        <w:p><w:r><w:t>{% endfor %}</w:t></w:r></w:p>
        <w:p><w:r><w:t>Root {{ title }}</w:t></w:r></w:p>
        <w:sectPr/>
    "#;
    let mut document = document_with_content_controls(&wrap_word_body(body));
    let data = serde_json::json!({
        "title": "Summary",
        "groups": [
            {
                "name": "Alpha",
                "visible": true,
                "items": [{"label": "one"}, {"label": "two"}]
            },
            {
                "name": "Beta",
                "visible": false,
                "items": [{"label": "three"}]
            }
        ]
    });

    assert_eq!(document.render_template(&data).unwrap(), 10);
    let body = body_from_document(&mut document);
    let paragraphs = body
        .content
        .iter()
        .filter_map(|content| match content {
            BodyContent::Paragraph(paragraph) => Some(paragraph.text()),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(
        paragraphs,
        ["Group Alpha", "Visible Alpha", "Group Beta", "Root Summary"]
    );
    let row_text = body
        .tables()
        .flat_map(|table| &table.rows)
        .map(|row| {
            row.cells
                .iter()
                .flat_map(|cell| &cell.content)
                .filter_map(|content| match content {
                    rdocx_oxml::table::CellContent::Paragraph(paragraph) => Some(paragraph.text()),
                    _ => None,
                })
                .collect::<String>()
        })
        .collect::<Vec<_>>();
    assert_eq!(row_text, ["Alpha:one", "Alpha:two", "Beta:three"]);
}

#[test]
fn three_template_rows_over_ten_records_produce_thirty_preserved_rows() {
    let body = r#"
        <w:tbl>
          <w:tblPr>
            <w:tblStyle w:val="BandedRows"/>
            <w:tblLook w:val="04A0" w:firstRow="1" w:noHBand="0"/>
          </w:tblPr>
          <w:tblGrid>
            <w:gridCol w:w="1800"/><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/>
          </w:tblGrid>
          <w:tr><w:tc><w:p><w:r><w:t>{% for record in records %}</w:t></w:r></w:p></w:tc></w:tr>
          <w:tr>
            <w:trPr><w:cnfStyle w:val="100000000000"/></w:trPr>
            <w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>A {{ record.id }}</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>top</w:t></w:r></w:p></w:tc>
          </w:tr>
          <w:tr>
            <w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr><w:p><w:r><w:t>B {{ record.id }}</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>middle</w:t></w:r></w:p></w:tc>
          </w:tr>
          <w:tr>
            <w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr><w:p><w:r><w:t>C {{ record.id }}</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>bottom</w:t></w:r></w:p></w:tc>
          </w:tr>
          <w:tr><w:tc><w:p><w:r><w:t>{% endfor %}</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>
        <w:sectPr/>
    "#;
    let mut document = document_with_content_controls(&wrap_word_body(body));
    let records = (0..10)
        .map(|id| serde_json::json!({"id": id}))
        .collect::<Vec<_>>();

    assert_eq!(
        document
            .render_template(&serde_json::json!({"records": records}))
            .unwrap(),
        30
    );
    let body = body_from_document(&mut document);
    let table = body.tables().next().unwrap();
    assert_eq!(table.rows.len(), 30);
    assert_eq!(table.grid.as_ref().unwrap().columns.len(), 3);
    let properties = table.properties.as_ref().unwrap();
    assert_eq!(properties.style_id.as_deref(), Some("BandedRows"));
    let look = properties.look.as_ref().unwrap();
    assert_eq!(look.first_row, Some(true));
    assert_eq!(look.no_h_band, Some(false));

    for (index, row) in table.rows.iter().enumerate() {
        let cell_properties = row.cells[0].properties.as_ref().unwrap();
        assert_eq!(cell_properties.grid_span, Some(2));
        assert_eq!(
            cell_properties.v_merge,
            Some(if index % 3 == 0 {
                rdocx_oxml::table::VMerge::Restart
            } else {
                rdocx_oxml::table::VMerge::Continue
            })
        );
        let prefix = ["A", "B", "C"][index % 3];
        assert_eq!(row.cells[0].text(), format!("{prefix} {}", index / 3));
    }

    let invalid_list = r#"
        <w:p><w:r><w:t>{% for item in items %}</w:t></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="99"/></w:numPr></w:pPr><w:r><w:t>{{ item }}</w:t></w:r></w:p>
        <w:p><w:r><w:t>{% endfor %}</w:t></w:r></w:p>
        <w:sectPr/>
    "#;
    let mut invalid = document_with_content_controls(&wrap_word_body(invalid_list));
    let before = invalid.to_bytes().unwrap();
    assert!(
        invalid
            .render_template(&serde_json::json!({"items": ["value"]}))
            .is_err()
    );
    assert_eq!(invalid.to_bytes().unwrap(), before);
}

#[test]
fn repeated_numbered_items_keep_one_continuous_sequence() {
    let mut document = Document::new();
    document.add_paragraph("{% for item in items %}");
    document.add_numbered_list_item("Item {{ item.name }}", 2);
    document.add_paragraph("Note {{ item.name }}");
    document.add_paragraph("{% endfor %}");
    let before = document.to_bytes().unwrap();
    let before_package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(before)).unwrap();
    let numbering_before = before_package
        .get_part("/word/numbering.xml")
        .unwrap()
        .to_vec();

    assert_eq!(
        document
            .render_template(&serde_json::json!({
                "items": [{"name": "one"}, {"name": "two"}, {"name": "three"}]
            }))
            .unwrap(),
        6
    );
    let saved = document.to_bytes().unwrap();
    let saved_package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    assert_eq!(
        saved_package.get_part("/word/numbering.xml").unwrap(),
        numbering_before
    );
    let reopened = Document::from_bytes(&saved).unwrap();
    let paragraphs = reopened.paragraphs();
    assert_eq!(
        paragraphs
            .iter()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>(),
        [
            "Item one",
            "Note one",
            "Item two",
            "Note two",
            "Item three",
            "Note three"
        ]
    );
    let numbering = paragraphs
        .iter()
        .step_by(2)
        .map(|paragraph| paragraph.numbering().unwrap())
        .collect::<Vec<_>>();
    assert!(numbering.iter().all(|value| *value == numbering[0]));
    assert_eq!(numbering[0].1, 2);

    let invalid_body = r#"
        <w:p><w:r><w:t>{% for item in items %}</w:t></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="99"/></w:numPr></w:pPr><w:r><w:t>{{ item }}</w:t></w:r></w:p>
        <w:p><w:r><w:t>{% endfor %}</w:t></w:r></w:p>
        <w:sectPr/>
    "#;
    let mut invalid = document_with_content_controls(&wrap_word_body(invalid_body));
    let before = invalid.to_bytes().unwrap();
    assert!(
        invalid
            .render_template(&serde_json::json!({"items": ["value"]}))
            .is_err()
    );
    assert_eq!(invalid.to_bytes().unwrap(), before);
}

fn mail_merge_record(values: &[(&str, &str)]) -> BTreeMap<String, String> {
    values
        .iter()
        .map(|(name, value)| ((*name).to_owned(), (*value).to_owned()))
        .collect()
}

fn document_with_mail_merge_header() -> Document {
    let mut seed = Document::new();
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
    let header = r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tbl xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:tblPr/><q:tblGrid/><q:tr><q:tc><q:p><q:fldSimple q:instr="MERGEFIELD &quot;Full Name&quot;"><q:r><q:t>stored table header</q:t></q:r></q:fldSimple></q:p></q:tc></q:tr></w:tbl><w:sdt><w:sdtPr><w:tag w:val="merge"/></w:sdtPr><w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> MERGEFIELD Name </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>stored control header</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt></w:hdr>"#;
    package.set_part("/word/header1.xml", header.as_bytes().to_vec());
    package.content_types.add_override(
        "/word/header1.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    );
    let header_id = package
        .get_or_create_part_rels("/word/document.xml")
        .add(oxml_opc::relationship::rel_types::HEADER, "header1.xml");
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/></w:sectPr></w:body></w:document>"#
        )
        .into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    Document::from_bytes(bytes.get_ref()).unwrap()
}

#[test]
fn a_fixture_record_set_produces_separate_and_sectioned_documents() {
    let body = r#"
        <w:p><w:fldSimple w:instr="MERGEFIELD Name \* Upper"><w:r><w:t>stored name</w:t></w:r></w:fldSimple></w:p>
        <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p><w:fldSimple w:instr="MERGEFIELD City"><w:r><w:t>stored city</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr></w:tbl>
        <w:sdt><w:sdtPr><w:tag w:val="role"/></w:sdtPr><w:sdtContent><w:p><w:fldSimple w:instr="MERGEFIELD Role"><w:r><w:t>stored role</w:t></w:r></w:fldSimple></w:p></w:sdtContent></w:sdt>
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
    "#;
    let document = document_with_content_controls(&wrap_word_body(body));
    let records = vec![
        mail_merge_record(&[("Name", "Ada"), ("City", "London"), ("Role", "Engineer")]),
        mail_merge_record(&[("Name", "Grace"), ("Role", "Admiral")]),
    ];

    let mut separate = document.mail_merge(&records).unwrap();
    assert_eq!(separate.len(), 2);
    let first = document_xml(&mut separate[0]);
    let second = document_xml(&mut separate[1]);
    assert!(first.contains(">ADA<"), "{first}");
    assert!(first.contains(">London<"), "{first}");
    assert!(first.contains(">Engineer<"), "{first}");
    assert!(second.contains(">GRACE<"), "{second}");
    assert!(second.contains(">Admiral<"), "{second}");
    assert!(!second.contains("stored city"), "{second}");

    let mut sectioned = document.mail_merge_sections(&records).unwrap();
    let xml = document_xml(&mut sectioned);
    let ada = xml.find(">ADA<").unwrap();
    let grace = xml.find(">GRACE<").unwrap();
    assert!(ada < grace, "{xml}");
    assert!(!xml.contains("stored city"), "{xml}");
    assert_eq!(xml.matches(r#"<w:type w:val="nextPage"/>"#).count(), 1);
}

#[test]
fn mail_merge_preserves_switches_and_general_field_policy() {
    let body = r#"
        <w:p><w:fldSimple w:instr="MERGEFIELD Name \* Upper"><w:r><w:t>stored name</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD Missing"><w:r><w:t>stored missing</w:t></w:r></w:fldSimple></w:p>
        <w:sectPr/>
    "#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let mut merged = document
        .mail_merge(&[mail_merge_record(&[("Name", "Ada")])])
        .unwrap()
        .pop()
        .unwrap();
    let merged_xml = document_xml(&mut merged);
    assert!(merged_xml.contains(r#"MERGEFIELD Name \* Upper"#));
    assert!(merged_xml.contains(">ADA<"), "{merged_xml}");
    assert!(!merged_xml.contains("stored missing"), "{merged_xml}");

    let outcomes = document
        .evaluate_fields(&FieldEvaluationContext::default())
        .unwrap();
    assert!(matches!(
        &outcomes[0].outcome,
        FieldOutcome::KeepStored { .. }
    ));
    assert!(matches!(
        &outcomes[1].outcome,
        FieldOutcome::KeepStored { .. }
    ));
    assert_eq!(
        document
            .update_fields(&FieldEvaluationContext::default())
            .unwrap(),
        2
    );
    let ordinary_xml = document_xml(&mut document);
    assert!(ordinary_xml.contains("stored name"), "{ordinary_xml}");
    assert!(ordinary_xml.contains("stored missing"), "{ordinary_xml}");
}

#[test]
fn sectioned_mail_merge_preserves_section_properties_and_unmodelled_xml() {
    const PRODUCER_BODY: &str = r#"<x:producer xmlns:x="urn:producer" mark="kept"/>"#;
    const PRODUCER_SECTION: &str = r#"<x:section xmlns:x="urn:producer" mark="kept"/>"#;
    let body = r#"
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>fixed</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
        <x:producer xmlns:x="urn:producer" mark="kept"/>
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><x:section xmlns:x="urn:producer" mark="kept"/></w:sectPr>
    "#;
    let document = document_with_field_parts(&wrap_word_body(body), None, None);
    let records = vec![
        mail_merge_record(&[("Name", "one")]),
        mail_merge_record(&[("Name", "two")]),
    ];
    let mut merged = document.mail_merge_sections(&records).unwrap();
    let bytes = merged.to_bytes().unwrap();
    let mut reopened = Document::from_bytes(&bytes).unwrap();
    let body = body_from_document(&mut reopened);

    assert_eq!(body.tables().count(), 2);
    assert_eq!(
        body.content
            .iter()
            .filter(|content| matches!(content, BodyContent::Paragraph(paragraph) if paragraph.properties.as_ref().and_then(|properties| properties.sect_pr.as_ref()).is_some()))
            .count(),
        1
    );
    let section_break = body.content.iter().find_map(|content| match content {
        BodyContent::Paragraph(paragraph) => paragraph
            .properties
            .as_ref()
            .and_then(|properties| properties.sect_pr.as_ref()),
        _ => None,
    });
    assert_eq!(
        section_break.unwrap().section_type,
        Some(rdocx_oxml::shared::ST_SectionType::NextPage)
    );
    assert_eq!(body.sect_pr.as_ref().unwrap().page_width.unwrap().0, 11906);

    let xml = document_xml(&mut reopened);
    assert_eq!(xml.matches(PRODUCER_BODY).count(), 2, "{xml}");
    assert_eq!(xml.matches(PRODUCER_SECTION).count(), 2, "{xml}");
    assert!(xml.rfind("<w:sectPr>").unwrap() < xml.rfind("</w:body>").unwrap());
}

#[test]
fn sectioned_mail_merge_remaps_record_local_body_identities() {
    let body = r#"
        <w:p><w:bookmarkStart w:id="7" w:name="Target"/><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>stored target</w:t></w:r></w:fldSimple><w:bookmarkEnd w:id="7"/></w:p>
        <w:p><w:fldSimple w:instr="REF Target"><w:r><w:t>stored reference</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="REF MailMerge1"><w:r><w:t>intentionally unresolved</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:hyperlink w:anchor="MailMerge2"><w:r><w:t>unresolved anchor</w:t></w:r></w:hyperlink></w:p>
        <w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt>
        <w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><wp:extent cx="1" cy="1"/><wp:docPr id="11" name="Picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="12" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
        <w:sectPr/>
    "#;
    let document = document_with_field_parts(&wrap_word_body(body), None, None);
    let records = [
        mail_merge_record(&[("Name", "one")]),
        mail_merge_record(&[("Name", "two")]),
    ];
    let mut merged = document.mail_merge_sections(&records).unwrap();

    let bookmarks = merged.bookmarks();
    assert_eq!(bookmarks.len(), 2);
    assert!(bookmarks.iter().all(|bookmark| bookmark.issue().is_none()));
    assert_ne!(bookmarks[0].id(), bookmarks[1].id());
    assert_ne!(bookmarks[0].name(), bookmarks[1].name());

    let evaluations = merged
        .evaluate_fields(&FieldEvaluationContext::default())
        .unwrap();
    assert!(evaluations.iter().any(|evaluation| {
        evaluation.instruction == "REF MailMerge1"
            && matches!(evaluation.outcome, FieldOutcome::KeepStored { .. })
    }));
    assert!(
        evaluations
            .iter()
            .filter(|evaluation| {
                matches!(
                    evaluation.instruction.as_str(),
                    "REF Target" | "REF MailMerge3"
                )
            })
            .all(|evaluation| matches!(evaluation.outcome, FieldOutcome::Resolved(_)))
    );
    let xml = document_xml(&mut merged);
    assert!(xml.contains(r#"w:instr="REF Target""#), "{xml}");
    assert!(xml.contains(r#"w:instr="REF MailMerge3""#), "{xml}");
    assert_eq!(
        xml.matches(r#"w:instr="REF MailMerge1""#).count(),
        2,
        "{xml}"
    );
    assert_eq!(xml.matches(r#"w:anchor="MailMerge2""#).count(), 2, "{xml}");
    assert!(!xml.contains(r#"w:name="MailMerge1""#), "{xml}");
    assert!(!xml.contains(r#"w:name="MailMerge2""#), "{xml}");
    assert_eq!(xml.matches(r#"w:val="9""#).count(), 1, "{xml}");
    assert_eq!(xml.matches(r#"wp:docPr id="11""#).count(), 1, "{xml}");
    assert_eq!(xml.matches(r#"pic:cNvPr id="12""#).count(), 1, "{xml}");
}

#[test]
fn sectioned_mail_merge_remaps_references_inside_preserved_body_xml() {
    let body = r#"
        <w:customXml>
          <w:p><w:bookmarkStart w:id="21" w:name="RawTarget"/><w:r><w:t>raw target</w:t></w:r><w:bookmarkEnd w:id="21"/></w:p>
          <w:p><w:fldSimple w:instr="REF RawTarget"><w:r><w:t>stored ref</w:t></w:r></w:fldSimple></w:p>
          <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText><w:instrText>REF RawTarget</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
          <w:p><w:hyperlink w:anchor="RawTarget"><w:r><w:t>raw link</w:t></w:r></w:hyperlink></w:p>
          <w:p><w:fldSimple w:instr="REF MailMerge1"><w:r><w:t>unresolved</w:t></w:r></w:fldSimple></w:p>
        </w:customXml>
        <w:sectPr/>
    "#;
    let document = document_with_field_parts(&wrap_word_body(body), None, None);
    let records = [mail_merge_record(&[]), mail_merge_record(&[])];
    let mut merged = document.mail_merge_sections(&records).unwrap();
    let xml = document_xml(&mut merged);

    assert_eq!(xml.matches(r#"w:name="RawTarget""#).count(), 1, "{xml}");
    assert_eq!(xml.matches(r#"w:name="MailMerge2""#).count(), 1, "{xml}");
    assert!(xml.contains(r#"w:instr="REF MailMerge2""#), "{xml}");
    assert!(xml.contains("PAGEREF MailMerge2"), "{xml}");
    assert!(xml.contains(r#"w:anchor="MailMerge2""#), "{xml}");
    assert_eq!(
        xml.matches(r#"w:instr="REF MailMerge1""#).count(),
        2,
        "{xml}"
    );
    assert!(!xml.contains(r#"w:name="MailMerge1""#), "{xml}");
}

#[test]
fn sectioned_mail_merge_correlates_entity_escaped_bookmark_names() {
    let body = r#"
        <w:p><w:bookmarkStart w:id="7" w:name="A&amp;B"/><w:r><w:t>target</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>
        <w:p><w:fldSimple w:instr="REF A&amp;B"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:sectPr/>
    "#;
    let document = document_with_field_parts(&wrap_word_body(body), None, None);
    let mut merged = document
        .mail_merge_sections(&[mail_merge_record(&[]), mail_merge_record(&[])])
        .unwrap();
    let evaluations = merged
        .evaluate_fields(&FieldEvaluationContext::default())
        .unwrap();
    assert!(
        evaluations
            .iter()
            .filter(|evaluation| evaluation.instruction.starts_with("REF "))
            .all(|evaluation| matches!(evaluation.outcome, FieldOutcome::Resolved(_)))
    );
    let xml = document_xml(&mut merged);
    assert_eq!(xml.matches(r#"w:name="A&amp;B""#).count(), 1, "{xml}");
    assert_eq!(xml.matches(r#"w:name="MailMerge1""#).count(), 1, "{xml}");
    assert!(xml.contains(r#"w:instr="REF MailMerge1""#), "{xml}");
}

#[test]
fn sectioned_mail_merge_ignores_foreign_same_local_name_attributes() {
    let mut seed = Document::new();
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
    let header = r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:producer"><w:p><w:fldSimple x:instr="MERGEFIELD Stable" w:instr="MERGEFIELD Vary"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:hdr>"#;
    package.set_part("/word/header-foreign.xml", header.as_bytes().to_vec());
    package.content_types.add_override(
        "/word/header-foreign.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    );
    let header_id = package.get_or_create_part_rels("/word/document.xml").add(
        oxml_opc::relationship::rel_types::HEADER,
        "header-foreign.xml",
    );
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:customXml xmlns:x="urn:producer"><w:p><w:bookmarkStart x:id="701" w:id="7" x:name="Foreign" w:name="Target"/><w:r><w:t>target</w:t></w:r><w:bookmarkEnd x:id="702" w:id="7"/></w:p><w:p><w:fldSimple x:instr="REF Foreign" w:instr="REF Target"><w:r><w:t>stored ref</w:t></w:r></w:fldSimple></w:p><w:sdt><w:sdtPr><w:id x:val="900" w:val="9"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><wp:extent cx="1" cy="1"/><wp:docPr x:id="1100" id="11" name="Picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr x:id="1200" id="12" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:customXml><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/></w:sectPr></w:body></w:document>"#
        )
        .into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let document = Document::from_bytes(bytes.get_ref()).unwrap();

    assert!(
        document
            .mail_merge_sections(&[
                mail_merge_record(&[("Stable", "same"), ("Vary", "one")]),
                mail_merge_record(&[("Stable", "same"), ("Vary", "two")]),
            ])
            .is_err()
    );

    let mut merged = document
        .mail_merge_sections(&[
            mail_merge_record(&[("Stable", "same"), ("Vary", "same")]),
            mail_merge_record(&[("Stable", "same"), ("Vary", "same")]),
        ])
        .unwrap();
    let xml = document_xml(&mut merged);
    assert_eq!(xml.matches(r#"x:id="701""#).count(), 2, "{xml}");
    assert_eq!(xml.matches(r#"x:id="702""#).count(), 2, "{xml}");
    assert_eq!(xml.matches(r#"x:name="Foreign""#).count(), 2, "{xml}");
    assert_eq!(xml.matches(r#"x:val="900""#).count(), 2, "{xml}");
    assert_eq!(xml.matches(r#"x:id="1100""#).count(), 2, "{xml}");
    assert_eq!(xml.matches(r#"x:id="1200""#).count(), 2, "{xml}");
    assert_eq!(xml.matches(r#"w:name="Target""#).count(), 1, "{xml}");
    assert!(xml.contains(r#"w:name="MailMerge1""#), "{xml}");
    assert_eq!(xml.matches(r#"w:instr="REF Target""#).count(), 1, "{xml}");
    assert!(xml.contains(r#"w:instr="REF MailMerge1""#), "{xml}");
    assert_eq!(xml.matches(r#"w:val="9""#).count(), 1, "{xml}");
    assert_eq!(
        xml.matches(r#"wp:docPr x:id="1100" id="11""#).count(),
        1,
        "{xml}"
    );
    assert_eq!(
        xml.matches(r#"pic:cNvPr x:id="1200" id="12""#).count(),
        1,
        "{xml}"
    );
}

#[test]
fn sectioned_mail_merge_scans_header_references_in_block_content_controls() {
    let mut seed = Document::new();
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
    let header = r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:fldSimple w:instr="MERGEFIELD Vary"><w:r><w:t>stored nested header</w:t></w:r></w:fldSimple></w:p></w:hdr>"#;
    package.set_part(
        "/word/sections/nested-header.xml",
        header.as_bytes().to_vec(),
    );
    package.content_types.add_override(
        "/word/sections/nested-header.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    );
    let header_id = package.get_or_create_part_rels("/word/document.xml").add(
        oxml_opc::relationship::rel_types::HEADER,
        "sections/nested-header.xml",
    );
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:sdt><w:sdtPr><w:tag w:val="nested-section"/></w:sdtPr><w:sdtContent><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/></w:sectPr></w:pPr><w:r><w:t>nested section</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr/></w:body></w:document>"#
        )
        .into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(bytes.get_ref()).unwrap();

    assert!(
        document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap()
            .is_empty()
    );
    assert_eq!(
        document
            .update_fields(&FieldEvaluationContext::default())
            .unwrap(),
        0
    );
    assert!(
        document
            .mail_merge_sections(&[
                mail_merge_record(&[("Vary", "one")]),
                mail_merge_record(&[("Vary", "two")]),
            ])
            .is_err()
    );
    assert!(
        document
            .mail_merge_sections(&[
                mail_merge_record(&[("Vary", "same")]),
                mail_merge_record(&[("Vary", "same")]),
            ])
            .is_ok()
    );
}

#[test]
fn mail_merge_uses_the_relationship_resolved_footnotes_part() {
    const RAW_TABLE: &str = r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>producer table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let mut document = document_with_field_parts(
        &wrap_word_body(r#"<w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr/>"#),
        None,
        None,
    );
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
            .unwrap();
    let footnotes = format!(
        r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="2"><w:p><w:fldSimple w:instr="MERGEFIELD Note"><w:r><w:t>stored note</w:t></w:r></w:fldSimple></w:p>{RAW_TABLE}</w:footnote></w:footnotes>"#
    );
    package.set_part("/word/notes/producer-footnotes.xml", footnotes.into_bytes());
    package.content_types.add_override(
        "/word/notes/producer-footnotes.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    );
    package.get_or_create_part_rels("/word/document.xml").add(
        oxml_opc::relationship::rel_types::FOOTNOTES,
        "notes/producer-footnotes.xml",
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let document = Document::from_bytes(bytes.get_ref()).unwrap();
    let records = [
        mail_merge_record(&[("Note", "one")]),
        mail_merge_record(&[("Note", "two")]),
    ];

    for (mut output, expected) in document
        .mail_merge(&records)
        .unwrap()
        .into_iter()
        .zip(["one", "two"])
    {
        let package =
            oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(output.to_bytes().unwrap()))
                .unwrap();
        let producer = std::str::from_utf8(
            package
                .get_part("/word/notes/producer-footnotes.xml")
                .unwrap(),
        )
        .unwrap();
        assert!(producer.contains(&format!(">{expected}<")), "{producer}");
        assert!(producer.contains(RAW_TABLE), "{producer}");
        assert!(package.get_part("/word/footnotes.xml").is_none());
    }
    assert!(document.mail_merge_sections(&records).is_err());
    let mut sectioned = document
        .mail_merge_sections(&[
            mail_merge_record(&[("Note", "same")]),
            mail_merge_record(&[("Note", "same")]),
        ])
        .unwrap();
    let package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(sectioned.to_bytes().unwrap()))
            .unwrap();
    let producer = std::str::from_utf8(
        package
            .get_part("/word/notes/producer-footnotes.xml")
            .unwrap(),
    )
    .unwrap();
    assert!(producer.contains(RAW_TABLE), "{producer}");
}

#[test]
fn a_failed_record_leaves_the_source_and_outputs_uncommitted() {
    let body = r#"<w:p><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p><w:sectPr/>"#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let before = document.to_bytes().unwrap();
    let records = vec![
        mail_merge_record(&[("Name", "valid")]),
        mail_merge_record(&[("Name", "invalid\u{000b}value")]),
    ];

    assert!(document.mail_merge(&records).is_err());
    assert!(document.mail_merge_sections(&records).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn empty_and_single_record_merges_have_stable_boundaries() {
    let body = r#"<w:p><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#;
    let document = document_with_field_parts(&wrap_word_body(body), None, None);
    assert!(document.mail_merge(&[]).is_err());
    assert!(document.mail_merge_sections(&[]).is_err());

    let mut single = document
        .mail_merge_sections(&[mail_merge_record(&[("Name", "only")])])
        .unwrap();
    let xml = document_xml(&mut single);
    assert!(xml.contains(">only<"), "{xml}");
    assert!(!xml.contains(r#"<w:type w:val="nextPage"/>"#), "{xml}");
    let body = body_from_document(&mut single);
    assert!(body.sect_pr.is_some());
    assert!(body.content.iter().all(|content| {
        !matches!(content, BodyContent::Paragraph(paragraph) if paragraph.properties.as_ref().and_then(|properties| properties.sect_pr.as_ref()).is_some())
    }));

    let mut document = document_with_mail_merge_header();
    let package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
            .unwrap();
    let header_before = package.get_part("/word/header1.xml").unwrap().to_vec();
    assert!(
        document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap()
            .is_empty()
    );
    assert_eq!(
        document
            .update_fields(&FieldEvaluationContext::default())
            .unwrap(),
        0
    );
    let package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
            .unwrap();
    assert_eq!(
        package.get_part("/word/header1.xml").unwrap(),
        header_before
    );
    let varying = vec![
        mail_merge_record(&[("Name", "same"), ("Full Name", "one")]),
        mail_merge_record(&[("Name", "same"), ("Full Name", "two")]),
    ];
    let mut separate = document.mail_merge(&varying).unwrap();
    for output in &mut separate {
        let package =
            oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(output.to_bytes().unwrap()))
                .unwrap();
        let header = std::str::from_utf8(package.get_part("/word/header1.xml").unwrap()).unwrap();
        assert!(header.contains("stored table header"), "{header}");
        assert!(header.contains("stored control header"), "{header}");
        assert!(!header.contains(">one<"), "{header}");
        assert!(!header.contains(">two<"), "{header}");
    }
    assert!(document.mail_merge_sections(&varying).is_err());
    assert!(
        document
            .mail_merge_sections(&[
                mail_merge_record(&[("Name", "one"), ("Full Name", "same")]),
                mail_merge_record(&[("Name", "two"), ("Full Name", "same")]),
            ])
            .is_err()
    );
    assert!(
        document
            .mail_merge_sections(&[
                mail_merge_record(&[("Name", "same"), ("Full Name", "same")]),
                mail_merge_record(&[("Name", "same"), ("Full Name", "same")]),
            ])
            .is_ok()
    );
}

#[test]
fn repeated_content_produces_a_deterministic_comparison() {
    let original_xml = wrap_word_body(
        r#"<w:p><w:r><w:t>repeat</w:t></w:r><w:r><w:t>repeat</w:t></w:r></w:p><w:p><w:r><w:t>repeat</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>repeat row</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>repeat row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let edited_xml = wrap_word_body(
        r#"<w:p><w:r><w:t>repeat</w:t></w:r><w:r><w:t>changed</w:t></w:r></w:p><w:p><w:r><w:t>repeat</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>repeat row</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>changed row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let edited = document_with_content_controls(&edited_xml);
    let mut first = document_with_content_controls(&original_xml);
    let mut second = document_with_content_controls(&original_xml);

    assert_eq!(
        first
            .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
            .unwrap(),
        second
            .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
            .unwrap()
    );
    assert_eq!(first.to_bytes().unwrap(), second.to_bytes().unwrap());
}

#[test]
fn comparison_metadata_is_escaped_and_ids_do_not_collide() {
    let original_xml = wrap_word_body(
        r#"<w:p><w:bookmarkStart w:id="0" w:name="kept"/><w:r><w:t>old</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>"#,
    );
    let edited_xml = wrap_word_body(
        r#"<w:p><w:bookmarkStart w:id="0" w:name="kept"/><w:r><w:t>new</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>"#,
    );
    let edited = document_with_content_controls(&edited_xml);
    let mut document = document_with_content_controls(&original_xml);
    document
        .compare(&edited, "Ada & \"Bob\"", "2026-08-21T09:30:00+01:00")
        .unwrap();

    let xml = document_xml(&mut document);
    assert!(xml.contains(r#"w:author="Ada &amp; &quot;Bob&quot;""#));
    assert!(
        document
            .revisions()
            .iter()
            .all(|revision| revision.id() != 0)
    );
    let before = document.to_bytes().unwrap();
    assert!(document.compare(&edited, "Ada", "not-a-date").is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn accepting_a_comparison_reproduces_the_edited_body_exactly() {
    let original_xml = wrap_word_body(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>old body</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>old cell</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>old nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr></w:tbl><w:sdt><w:sdtPr><w:tag w:val="scope"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>old control</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
    );
    let edited_xml = wrap_word_body(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>new body</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>new cell</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>new nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr></w:tbl><w:sdt><w:sdtPr><w:tag w:val="scope"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>new control</w:t></w:r></w:p><w:p><w:r><w:t>added control child</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
    );
    let edited = document_with_content_controls(&edited_xml);
    let mut compared = document_with_content_controls(&original_xml);
    compared
        .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    assert!(!compared.revisions().is_empty());
    compared.accept_all().unwrap();
    assert!(
        compared
            .compare(&edited, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
    assert!(compared.revisions().is_empty());
}

#[test]
fn rejecting_a_comparison_reproduces_the_original_body_exactly() {
    let original_xml = wrap_word_body(
        r#"<w:p><w:r><w:t>one</w:t></w:r><w:r><w:t>two</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let edited_xml = wrap_word_body(
        r#"<w:p><w:r><w:t>one</w:t></w:r><w:r><w:t>changed</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>changed row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let original = document_with_content_controls(&original_xml);
    let edited = document_with_content_controls(&edited_xml);
    let mut compared = document_with_content_controls(&original_xml);
    compared
        .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    compared.reject_all().unwrap();
    assert!(
        compared
            .compare(&original, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
    assert!(compared.revisions().is_empty());
}

#[test]
fn modeled_formatting_changes_are_tracked_without_diagnostics() {
    let original_xml =
        wrap_word_body(r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>same</w:t></w:r></w:p>"#);
    let edited_xml =
        wrap_word_body(r#"<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>same</w:t></w:r></w:p>"#);
    let original = document_with_content_controls(&original_xml);
    let edited = document_with_content_controls(&edited_xml);
    let mut compared = document_with_content_controls(&original_xml);
    let diagnostics = compared
        .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();

    assert!(diagnostics.is_empty(), "{diagnostics:?}");
    assert!(!compared.revisions().is_empty());
    let tracked_bytes = compared.to_bytes().unwrap();
    assert!(document_xml(&mut compared).contains("<w:rPrChange"));

    let mut accepted = Document::from_bytes(&tracked_bytes).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&edited, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
    let mut rejected = Document::from_bytes(&tracked_bytes).unwrap();
    rejected.reject_all().unwrap();
    assert!(
        rejected
            .compare(&original, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
}

#[test]
fn a_failed_comparison_leaves_the_original_package_unchanged() {
    let existing_xml = wrap_word_body(
        r#"<w:p><w:ins w:id="4" w:author="prior"><w:r><w:t>tracked</w:t></w:r></w:ins></w:p>"#,
    );
    let edited = Document::new();
    let mut document = document_with_content_controls(&existing_xml);
    let before = document.to_bytes().unwrap();

    assert!(
        document
            .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn comparison_preserves_unmodelled_xml_byte_for_byte() {
    let body_raw = r#"<x:bodyOpaque xmlns:x="urn:comparison-body" x:value="keep"/>"#;
    let paragraph_raw =
        r#"<x:paragraphOpaque xmlns:x="urn:comparison-paragraph"><x:nested/></x:paragraphOpaque>"#;
    let table_raw = r#"<x:tableOpaque xmlns:x="urn:comparison-table" x:value="keep"/>"#;
    let cell_raw = r#"<x:cellOpaque xmlns:x="urn:comparison-cell" x:value="keep"/>"#;
    let control_raw =
        r#"<x:controlOpaque xmlns:x="urn:comparison-control"><x:nested/></x:controlOpaque>"#;
    let body = |value: &str| {
        wrap_word_body(&format!(
            r#"{body_raw}<w:p><w:r><w:t>{value} body</w:t>{paragraph_raw}</w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/>{table_raw}<w:tr><w:tc>{cell_raw}<w:p><w:r><w:t>{value} cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sdt><w:sdtPr><w:tag w:val="raw-scope"/></w:sdtPr><w:sdtContent>{control_raw}<w:p><w:r><w:t>{value} control</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
        ))
    };
    let original_xml = body("old");
    let edited_xml = body("new");
    let edited = document_with_content_controls(&edited_xml);
    let mut document = document_with_content_controls(&original_xml);
    document
        .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();

    let xml = document_xml(&mut document);
    for raw in [body_raw, table_raw, cell_raw, control_raw] {
        assert_eq!(xml.matches(raw).count(), 1, "{xml}");
    }
    assert_eq!(xml.matches(paragraph_raw).count(), 2, "{xml}");
    let mut reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let reopened_xml = document_xml(&mut reopened);
    for raw in [body_raw, table_raw, cell_raw, control_raw] {
        assert_eq!(reopened_xml.matches(raw).count(), 1, "{reopened_xml}");
    }
    assert_eq!(reopened_xml.matches(paragraph_raw).count(), 2);
}

#[test]
fn inserted_and_deleted_body_blocks_resolve_without_empty_containers() {
    let original_xml = wrap_word_body(
        r#"<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>last</w:t></w:r></w:p>"#,
    );
    let edited_xml = wrap_word_body(
        r#"<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>inserted</w:t></w:r></w:p><w:p><w:r><w:t>last</w:t></w:r></w:p>"#,
    );
    let original = document_with_content_controls(&original_xml);
    let edited = document_with_content_controls(&edited_xml);
    let mut inserted = document_with_content_controls(&original_xml);
    inserted
        .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let mut accepted = Document::from_bytes(&inserted.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    let diagnostics = accepted
        .compare(&edited, "postcondition", "2026-08-21T09:31:00Z")
        .unwrap();
    assert!(diagnostics.is_empty(), "{diagnostics:?}");
    let mut rejected = Document::from_bytes(&inserted.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    let diagnostics = rejected
        .compare(&original, "postcondition", "2026-08-21T09:31:00Z")
        .unwrap();
    assert!(diagnostics.is_empty(), "{diagnostics:?}");

    let empty_xml = wrap_word_body(r#"<w:p><w:r><w:t>anchor</w:t></w:r></w:p>"#);
    let table_xml = wrap_word_body(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>only row</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>anchor</w:t></w:r></w:p>"#,
    );
    let empty = document_with_content_controls(&empty_xml);
    let table = document_with_content_controls(&table_xml);
    let mut inserted_table = document_with_content_controls(&empty_xml);
    inserted_table
        .compare(&table, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    inserted_table.reject_all().unwrap();
    assert!(
        inserted_table
            .compare(&empty, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );

    let final_xml = wrap_word_body(
        r#"<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>final</w:t></w:r></w:p>"#,
    );
    let final_document = document_with_content_controls(&final_xml);
    let mut inserted_final = document_with_content_controls(&empty_xml);
    inserted_final
        .compare(&final_document, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    inserted_final.reject_all().unwrap();
    assert!(
        inserted_final
            .compare(&empty, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
}

#[test]
fn comparison_revises_nested_control_content_without_replacing_its_shell() {
    let body = |value: &str, paragraph_tag: &str| {
        wrap_word_body(&format!(
            r#"<w:p><w:sdt><w:sdtPr><w:tag w:val="{paragraph_tag}"/></w:sdtPr><w:sdtContent><w:r><w:t>{value} paragraph control</w:t></w:r></w:sdtContent></w:sdt></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:sdt><w:sdtPr><w:tag w:val="table-control"/></w:sdtPr><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>{value} table control</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt><w:tr><w:sdt><w:sdtPr><w:tag w:val="row-control"/></w:sdtPr><w:sdtContent><w:tc><w:p><w:r><w:t>{value} row control</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt><w:tc><w:sdt><w:sdtPr><w:tag w:val="cell-control"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>{value} cell control</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>"#,
        ))
    };
    let original_xml = body("old", "paragraph-control");
    let edited_xml = body("new", "paragraph-control");
    let original = document_with_content_controls(&original_xml);
    let edited = document_with_content_controls(&edited_xml);
    let mut compared = document_with_content_controls(&original_xml);
    compared
        .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let tracked_xml = document_xml(&mut compared);
    for tag in [
        "paragraph-control",
        "table-control",
        "row-control",
        "cell-control",
    ] {
        assert_eq!(tracked_xml.matches(&format!(r#"w:val="{tag}""#)).count(), 1);
    }
    let mut accepted = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    let diagnostics = accepted
        .compare(&edited, "postcondition", "2026-08-21T09:31:00Z")
        .unwrap();
    assert!(diagnostics.is_empty(), "{diagnostics:?}");
    let mut rejected = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    let diagnostics = rejected
        .compare(&original, "postcondition", "2026-08-21T09:31:00Z")
        .unwrap();
    assert!(diagnostics.is_empty(), "{diagnostics:?}");

    let changed_shell_xml = body("old", "changed-shell");
    let changed_shell = document_with_content_controls(&changed_shell_xml);
    let mut unchanged = document_with_content_controls(&original_xml);
    let before = unchanged.to_bytes().unwrap();
    assert!(
        unchanged
            .compare(&changed_shell, "Ada", "2026-08-21T09:30:00Z")
            .is_err()
    );
    assert_eq!(unchanged.to_bytes().unwrap(), before);
}

#[test]
fn comparison_preserves_content_control_whitespace_slots() {
    let first_slot = "\r\n\t \t\r\n";
    let second_slot = "\n \t \n";
    let body = |value: &str| {
        wrap_word_body(&format!(
            r#"<w:sdt><w:sdtPr><w:tag w:val="pretty"/></w:sdtPr><w:sdtContent>{first_slot}<w:p><w:r><w:t>{value}</w:t></w:r></w:p>{second_slot}</w:sdtContent></w:sdt>"#,
        ))
    };
    let original_xml = body("old");
    let edited_xml = body("new");
    let edited = document_with_content_controls(&edited_xml);
    let mut compared = document_with_content_controls(&original_xml);

    compared
        .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let tracked = document_xml(&mut compared);
    assert!(tracked.contains(first_slot), "{tracked:?}");
    assert!(tracked.contains(second_slot), "{tracked:?}");

    let mut accepted = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&edited, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
}

#[test]
fn comparison_deletes_text_without_corrupting_tabs() {
    let original_xml =
        wrap_word_body(r#"<w:p><w:r><w:t>old</w:t><w:tab/><w:t>tail</w:t></w:r></w:p>"#);
    let edited_xml =
        wrap_word_body(r#"<w:p><w:r><w:t>new</w:t><w:tab/><w:t>tail</w:t></w:r></w:p>"#);
    let original = document_with_content_controls(&original_xml);
    let edited = document_with_content_controls(&edited_xml);
    let mut compared = document_with_content_controls(&original_xml);

    compared
        .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let tracked = document_xml(&mut compared);
    assert!(tracked.contains("<w:tab/>"), "{tracked}");
    assert!(!tracked.contains("delTextab"), "{tracked}");

    let mut accepted = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&edited, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
    let mut rejected = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    assert!(
        rejected
            .compare(&original, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
}

#[test]
fn comparison_reports_formatting_inside_matched_table_rows() {
    let original_xml = wrap_word_body(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>same</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let edited_xml = wrap_word_body(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:tcPr><w:shd w:fill="0000FF"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>same</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let edited = document_with_content_controls(&edited_xml);
    let mut compared = document_with_content_controls(&original_xml);

    let diagnostics = compared
        .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    assert_eq!(
        diagnostics
            .iter()
            .map(|diagnostic| diagnostic.location.as_str())
            .collect::<Vec<_>>(),
        vec!["body/table[0]/row[0]/cell[0]"]
    );
    assert!(!compared.revisions().is_empty());
    let tracked_bytes = compared.to_bytes().unwrap();
    let tracked = document_xml(&mut compared);
    assert!(tracked.contains("<w:pPrChange"), "{tracked}");
    assert!(tracked.contains("<w:rPrChange"), "{tracked}");

    let mut accepted = Document::from_bytes(&tracked_bytes).unwrap();
    accepted.accept_all().unwrap();
    let accepted = document_xml(&mut accepted);
    assert!(accepted.contains(r#"<w:jc w:val="right"/>"#), "{accepted}");
    assert!(accepted.contains("<w:i/>"), "{accepted}");
    assert!(accepted.contains(r#"w:fill="FF0000""#), "{accepted}");

    let mut rejected = Document::from_bytes(&tracked_bytes).unwrap();
    rejected.reject_all().unwrap();
    let rejected = document_xml(&mut rejected);
    assert!(rejected.contains(r#"<w:jc w:val="left"/>"#), "{rejected}");
    assert!(rejected.contains("<w:b/>"), "{rejected}");
    assert!(rejected.contains(r#"w:fill="FF0000""#), "{rejected}");
}

#[test]
fn a_new_paragraph_property_owner_does_not_adopt_edited_unmodelled_xml() {
    let edited_raw = r#"<x:edited xmlns:x="urn:f234-paragraph-property" x:value="discard"/>"#;
    let original_xml = wrap_word_body(r#"<w:p><w:r><w:t>same</w:t></w:r></w:p>"#);
    let edited_xml = wrap_word_body(&format!(
        r#"<w:p><w:pPr><w:jc w:val="center"/>{edited_raw}</w:pPr><w:r><w:t>same</w:t></w:r></w:p>"#,
    ));
    let edited = document_with_content_controls(&edited_xml);
    let mut compared = document_with_content_controls(&original_xml);

    assert_eq!(
        compared
            .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
            .unwrap(),
        vec![rdocx::ComparisonDiagnostic {
            location: "body/paragraph[0]".to_owned(),
            message: "formatting differs and the original formatting was retained".to_owned(),
        }]
    );
    let tracked_bytes = compared.to_bytes().unwrap();
    let tracked = document_xml(&mut compared);
    assert!(tracked.contains("<w:pPrChange"), "{tracked}");
    assert!(tracked.contains(r#"<w:jc w:val="center"/>"#), "{tracked}");
    assert!(!tracked.contains(edited_raw), "{tracked}");

    let mut accepted = Document::from_bytes(&tracked_bytes).unwrap();
    accepted.accept_all().unwrap();
    let accepted = document_xml(&mut accepted);
    assert!(accepted.contains(r#"<w:jc w:val="center"/>"#), "{accepted}");
    assert!(!accepted.contains(edited_raw), "{accepted}");

    let mut rejected = Document::from_bytes(&tracked_bytes).unwrap();
    rejected.reject_all().unwrap();
    let rejected = document_xml(&mut rejected);
    assert!(!rejected.contains("<w:pPr"), "{rejected}");
    assert!(!rejected.contains(edited_raw), "{rejected}");
}

#[test]
fn a_new_table_property_owner_does_not_adopt_edited_unmodelled_xml() {
    let edited_raw = r#"<x:edited xmlns:x="urn:f234-table-property" x:value="discard"/>"#;
    let original_xml = wrap_word_body(
        r#"<w:tbl><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>same</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let edited_xml = wrap_word_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="200" w:type="dxa"/>{edited_raw}</w:tblPr><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>same</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    ));
    let edited = document_with_content_controls(&edited_xml);
    let mut compared = document_with_content_controls(&original_xml);

    assert_eq!(
        compared
            .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
            .unwrap(),
        vec![rdocx::ComparisonDiagnostic {
            location: "body/table[0]".to_owned(),
            message: "formatting differs and the original formatting was retained".to_owned(),
        }]
    );
    let tracked_bytes = compared.to_bytes().unwrap();
    let tracked = document_xml(&mut compared);
    assert!(tracked.contains("<w:tblPrChange"), "{tracked}");
    assert!(
        tracked.contains(r#"<w:tblW w:w="200" w:type="dxa"/>"#),
        "{tracked}"
    );
    assert!(!tracked.contains(edited_raw), "{tracked}");

    let mut accepted = Document::from_bytes(&tracked_bytes).unwrap();
    accepted.accept_all().unwrap();
    let accepted = document_xml(&mut accepted);
    assert!(
        accepted.contains(r#"<w:tblW w:w="200" w:type="dxa"/>"#),
        "{accepted}"
    );
    assert!(!accepted.contains(edited_raw), "{accepted}");

    let mut rejected = Document::from_bytes(&tracked_bytes).unwrap();
    rejected.reject_all().unwrap();
    let rejected = document_xml(&mut rejected);
    assert!(!rejected.contains("<w:tblPr"), "{rejected}");
    assert!(!rejected.contains(edited_raw), "{rejected}");
}

#[test]
fn comparison_replaces_paragraphs_and_tables_before_an_anchor() {
    let paragraph_xml = wrap_word_body(
        r#"<w:p><w:r><w:t>old paragraph</w:t></w:r></w:p><w:p><w:r><w:t>anchor</w:t></w:r></w:p>"#,
    );
    let table_xml = wrap_word_body(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>new table</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>anchor</w:t></w:r></w:p>"#,
    );
    let paragraph = document_with_content_controls(&paragraph_xml);
    let table = document_with_content_controls(&table_xml);

    let mut paragraph_to_table = document_with_content_controls(&paragraph_xml);
    paragraph_to_table
        .compare(&table, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let mut accepted = Document::from_bytes(&paragraph_to_table.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&table, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
    let mut rejected = Document::from_bytes(&paragraph_to_table.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    assert!(
        rejected
            .compare(&paragraph, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );

    let mut table_to_paragraph = document_with_content_controls(&table_xml);
    table_to_paragraph
        .compare(&paragraph, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let mut accepted = Document::from_bytes(&table_to_paragraph.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&paragraph, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
    let mut rejected = Document::from_bytes(&table_to_paragraph.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    assert!(
        rejected
            .compare(&table, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
}

#[test]
fn comparison_preserves_unrelated_modeled_fields() {
    let field =
        r#"<w:fldSimple w:instr="AUTHOR"><w:r><w:t>stored author</w:t></w:r></w:fldSimple>"#;
    let original_xml = wrap_word_body(&format!(
        r#"<w:p>{field}<w:r><w:t>old text</w:t></w:r></w:p>"#
    ));
    let edited_xml = wrap_word_body(&format!(
        r#"<w:p>{field}<w:r><w:t>new text</w:t></w:r></w:p>"#
    ));
    let original = document_with_content_controls(&original_xml);
    let edited = document_with_content_controls(&edited_xml);
    let mut compared = document_with_content_controls(&original_xml);

    compared
        .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    assert_eq!(
        document_xml(&mut compared).matches("<w:fldSimple").count(),
        1
    );

    let mut accepted = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&edited, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
    assert_eq!(
        document_xml(&mut accepted).matches("<w:fldSimple").count(),
        1
    );

    let mut rejected = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    assert!(
        rejected
            .compare(&original, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
    assert_eq!(
        document_xml(&mut rejected).matches("<w:fldSimple").count(),
        1
    );
}

#[test]
fn final_paragraph_markers_stay_outside_formatted_run_properties() {
    let anchor = r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>anchor</w:t></w:r></w:p>"#;
    let added = r#"<w:p><w:r><w:t>added</w:t></w:r></w:p>"#;
    let original_xml = wrap_word_body(anchor);
    let edited_xml = wrap_word_body(&format!("{anchor}{added}"));
    let original = document_with_content_controls(&original_xml);
    let edited = document_with_content_controls(&edited_xml);
    let mut compared = document_with_content_controls(&original_xml);

    compared
        .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let tracked = document_xml(&mut compared);
    let marker = tracked.find("<w:ins").unwrap();
    let properties_start = tracked.find("<w:pPr").unwrap();
    let properties_end = tracked.find("</w:pPr>").unwrap();
    let first_run = tracked.find("<w:r>").unwrap();
    assert!(
        properties_start < marker && marker < properties_end,
        "{tracked}"
    );
    assert!(properties_end < first_run, "{tracked}");
    assert_eq!(tracked.matches("<w:b/>").count(), 1, "{tracked}");
    let mut rejected = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    assert!(
        rejected
            .compare(&original, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );

    let mut deletion = document_with_content_controls(&edited_xml);
    deletion
        .compare(&original, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let tracked = document_xml(&mut deletion);
    let marker = tracked.find("<w:del").unwrap();
    let properties_start = tracked.find("<w:pPr").unwrap();
    let properties_end = tracked.find("</w:pPr>").unwrap();
    let first_run = tracked.find("<w:r>").unwrap();
    assert!(
        properties_start < marker && marker < properties_end,
        "{tracked}"
    );
    assert!(properties_end < first_run, "{tracked}");
    let mut accepted = Document::from_bytes(&deletion.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&original, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
}

#[test]
fn whole_row_markers_target_the_outer_row_properties() {
    let anchor_xml = wrap_word_body(r#"<w:p><w:r><w:t>anchor</w:t></w:r></w:p>"#);
    let table_xml = wrap_word_body(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr></w:tbl><w:p><w:r><w:t>anchor</w:t></w:r></w:p>"#,
    );
    let anchor = document_with_content_controls(&anchor_xml);
    let table = document_with_content_controls(&table_xml);
    let mut compared = document_with_content_controls(&anchor_xml);

    compared
        .compare(&table, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let tracked = document_xml(&mut compared);
    let marker = tracked.find("<w:ins").unwrap();
    let outer_cell = tracked.find("<w:tc>").unwrap();
    let nested_properties = tracked.find("<w:tblHeader/>").unwrap();
    assert!(marker < outer_cell, "{tracked}");
    assert!(marker < nested_properties, "{tracked}");
    assert_eq!(tracked.matches("<w:ins").count(), 1, "{tracked}");

    let mut rejected = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    assert!(
        rejected
            .compare(&anchor, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );

    let mut deletion = document_with_content_controls(&table_xml);
    deletion
        .compare(&anchor, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let tracked = document_xml(&mut deletion);
    assert!(tracked.find("<w:del").unwrap() < tracked.find("<w:tc>").unwrap());
    let mut accepted = Document::from_bytes(&deletion.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&anchor, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
}

#[test]
fn whole_table_marking_includes_control_owned_rows() {
    let anchor_xml = wrap_word_body(r#"<w:p><w:r><w:t>anchor</w:t></w:r></w:p>"#);
    let table_xml = wrap_word_body(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:sdt><w:sdtPr><w:tag w:val="owned-row"/></w:sdtPr><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>same row</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt><w:tr><w:tc><w:p><w:r><w:t>same row</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>anchor</w:t></w:r></w:p>"#,
    );
    let anchor = document_with_content_controls(&anchor_xml);
    let table = document_with_content_controls(&table_xml);

    let mut insertion = document_with_content_controls(&anchor_xml);
    insertion
        .compare(&table, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let tracked = document_xml(&mut insertion);
    assert_eq!(tracked.matches(r#"<w:trPr><w:ins"#).count(), 2, "{tracked}");
    let mut rejected = Document::from_bytes(&insertion.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    assert!(
        rejected
            .compare(&anchor, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );

    let mut deletion = document_with_content_controls(&table_xml);
    deletion
        .compare(&anchor, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let tracked = document_xml(&mut deletion);
    assert_eq!(tracked.matches(r#"<w:trPr><w:del"#).count(), 2, "{tracked}");
    let mut accepted = Document::from_bytes(&deletion.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&anchor, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
}

#[test]
fn table_row_insertions_stay_inside_table_schema_boundaries() {
    let original_xml = wrap_word_body(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>anchor</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let edited_xmls = [
        wrap_word_body(
            r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>inserted</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>anchor</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        ),
        wrap_word_body(
            r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>anchor</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>inserted</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        ),
    ];
    let original = document_with_content_controls(&original_xml);

    for edited_xml in edited_xmls {
        let edited = document_with_content_controls(&edited_xml);
        let mut compared = document_with_content_controls(&original_xml);
        compared
            .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
            .unwrap();
        let tracked = document_xml(&mut compared);
        let table_start = tracked.find("<w:tbl>").unwrap();
        let table_end = tracked.find("</w:tbl>").unwrap();
        let inserted_row = tracked
            .find("<w:ins")
            .unwrap_or_else(|| panic!("{tracked}"));
        assert!(
            table_start < inserted_row && inserted_row < table_end,
            "{tracked}"
        );

        let mut accepted = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
        accepted.accept_all().unwrap();
        assert!(
            accepted
                .compare(&edited, "postcondition", "2026-08-21T09:31:00Z")
                .unwrap()
                .is_empty()
        );
        let mut rejected = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
        rejected.reject_all().unwrap();
        assert!(
            rejected
                .compare(&original, "postcondition", "2026-08-21T09:31:00Z")
                .unwrap()
                .is_empty()
        );
    }
}

#[test]
fn comparison_adds_and_removes_numbering_without_a_property_owner() {
    let plain_xml = wrap_word_body(r#"<w:p><w:r><w:t>item</w:t></w:r></w:p>"#);
    let numbered_xml = wrap_word_body(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>"#,
    );
    let plain = document_with_content_controls(&plain_xml);
    let numbered = document_with_content_controls(&numbered_xml);

    let mut addition = document_with_content_controls(&plain_xml);
    addition
        .compare(&numbered, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let addition_xml = document_xml(&mut addition);
    assert!(addition_xml.contains("<w:pPrChange"), "{addition_xml}");
    let mut accepted = Document::from_bytes(&addition.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&numbered, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
    let mut rejected = Document::from_bytes(&addition.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    assert!(
        rejected
            .compare(&plain, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );

    let mut removal = document_with_content_controls(&numbered_xml);
    removal
        .compare(&plain, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let removal_xml = document_xml(&mut removal);
    assert!(removal_xml.contains("<w:pPrChange"), "{removal_xml}");
    let mut accepted = Document::from_bytes(&removal.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&plain, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
    let mut rejected = Document::from_bytes(&removal.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    assert!(
        rejected
            .compare(&numbered, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
}

#[test]
fn numbering_changes_preserve_unrelated_paragraph_properties() {
    let unnumbered_xml = wrap_word_body(
        r#"<w:p><w:pPr><w:keepNext/><w:jc w:val="center"/></w:pPr><w:r><w:t>item</w:t></w:r></w:p>"#,
    );
    let numbered_xml = wrap_word_body(
        r#"<w:p><w:pPr><w:keepNext/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="9"/></w:numPr><w:jc w:val="center"/></w:pPr><w:r><w:t>item</w:t></w:r></w:p>"#,
    );

    for (original_xml, edited_xml) in [
        (&unnumbered_xml, &numbered_xml),
        (&numbered_xml, &unnumbered_xml),
    ] {
        let original = document_with_content_controls(original_xml);
        let edited = document_with_content_controls(edited_xml);
        let mut compared = document_with_content_controls(original_xml);
        compared
            .compare(&edited, "Ada", "2026-08-21T09:30:00Z")
            .unwrap();

        let mut accepted = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
        accepted.accept_all().unwrap();
        let accepted_xml = document_xml(&mut accepted);
        assert!(accepted_xml.contains("<w:keepNext"), "{accepted_xml}");
        assert!(
            accepted_xml.contains(r#"<w:jc w:val="center"/>"#),
            "{accepted_xml}"
        );
        assert!(
            accepted
                .compare(&edited, "postcondition", "2026-08-21T09:31:00Z")
                .unwrap()
                .is_empty()
        );

        let mut rejected = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
        rejected.reject_all().unwrap();
        let rejected_xml = document_xml(&mut rejected);
        assert!(rejected_xml.contains("<w:keepNext"), "{rejected_xml}");
        assert!(
            rejected_xml.contains(r#"<w:jc w:val="center"/>"#),
            "{rejected_xml}"
        );
        assert!(
            rejected
                .compare(&original, "postcondition", "2026-08-21T09:31:00Z")
                .unwrap()
                .is_empty()
        );
    }
}

#[test]
fn resolving_empty_paragraph_property_changes_cleans_only_empty_owners() {
    let empty_xml = wrap_word_body(
        r#"<w:p><w:pPr><w:pPrChange w:id="1" w:author="Ada"><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>item</w:t></w:r></w:p>"#,
    );
    let retained_xml = wrap_word_body(
        r#"<w:p><w:pPr><w:keepNext/><w:pPrChange w:id="1" w:author="Ada"><w:pPr><w:keepNext/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>item</w:t></w:r></w:p>"#,
    );

    for accept in [true, false] {
        let mut empty = document_with_content_controls(&empty_xml);
        if accept {
            empty.accept_all().unwrap();
        } else {
            empty.reject_all().unwrap();
        }
        let empty = document_xml(&mut empty);
        assert!(!empty.contains("<w:pPr"), "{empty}");

        let mut retained = document_with_content_controls(&retained_xml);
        if accept {
            retained.accept_all().unwrap();
        } else {
            retained.reject_all().unwrap();
        }
        let retained = document_xml(&mut retained);
        assert_eq!(retained.matches("<w:pPr>").count(), 1, "{retained}");
        assert_eq!(retained.matches("<w:keepNext").count(), 1, "{retained}");
        assert!(!retained.contains("<w:pPrChange"), "{retained}");
    }
}

#[test]
fn rejecting_an_attributed_empty_prior_paragraph_owner_preserves_it_exactly() {
    let prior = format!(
        r#"<old:pPr xmlns:old="{}" xmlns:ext="urn:producer" ext:flag="keep" ext:mode="exact"/>"#,
        rdocx_oxml::namespace::W_NS
    );
    let tracked_xml = wrap_word_body(&format!(
        r#"<w:p><w:pPr><w:pPrChange w:id="1" w:author="Ada">{prior}</w:pPrChange></w:pPr><w:r><w:t>item</w:t></w:r></w:p>"#,
    ));
    let mut tracked = document_with_content_controls(&tracked_xml);

    let mut accepted = Document::from_bytes(&tracked.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    let accepted = document_xml(&mut accepted);
    assert!(!accepted.contains("<w:pPr"), "{accepted}");
    assert!(!accepted.contains("ext:flag"), "{accepted}");

    let mut rejected = Document::from_bytes(&tracked.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    let rejected = document_xml(&mut rejected);
    assert!(rejected.contains(&prior), "{rejected}");
    assert!(!rejected.contains("<w:pPrChange"), "{rejected}");

    let tracked_xml = wrap_word_body(
        r#"<w:p><w:pPr><w:pPrChange xmlns:ext="urn:inherited" w:id="2" w:author="Ada"><w:pPr ext:flag="keep"/></w:pPrChange></w:pPr><w:r><w:t>item</w:t></w:r></w:p>"#,
    );
    let mut tracked = document_with_content_controls(&tracked_xml);
    let mut rejected = Document::from_bytes(&tracked.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    let rejected = document_xml(&mut rejected);
    assert!(rejected.contains("<w:pPr"), "{rejected}");
    assert!(
        rejected.contains(r#"xmlns:ext="urn:inherited""#),
        "{rejected}"
    );
    assert!(rejected.contains(r#"ext:flag="keep""#), "{rejected}");
    assert!(!rejected.contains("<w:pPrChange"), "{rejected}");
}

#[test]
fn control_block_replacement_keeps_whitespace_before_the_replacement() {
    let before = "\r\n\t  \t\r\n";
    let between = "\n\t \n";
    let control = |first: &str| {
        wrap_word_body(&format!(
            r#"<w:sdt><w:sdtPr><w:tag w:val="block-replacement"/></w:sdtPr><w:sdtContent>{before}{first}{between}<w:p><w:r><w:t>anchor</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
        ))
    };
    let paragraph_xml = control(r#"<w:p><w:r><w:t>old paragraph</w:t></w:r></w:p>"#);
    let table_xml = control(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>new table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let paragraph = document_with_content_controls(&paragraph_xml);
    let table = document_with_content_controls(&table_xml);
    let mut compared = document_with_content_controls(&paragraph_xml);

    compared
        .compare(&table, "Ada", "2026-08-21T09:30:00Z")
        .unwrap();
    let mut accepted = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    let accepted_xml = document_xml(&mut accepted);
    assert!(accepted_xml.contains(before), "{accepted_xml:?}");
    assert!(accepted_xml.contains(between), "{accepted_xml:?}");
    assert!(accepted_xml.find(before).unwrap() < accepted_xml.find("<w:tbl>").unwrap());
    assert!(accepted_xml.find("</w:tbl>").unwrap() < accepted_xml.find(between).unwrap());
    assert!(
        accepted
            .compare(&table, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );

    let mut rejected = Document::from_bytes(&compared.to_bytes().unwrap()).unwrap();
    rejected.reject_all().unwrap();
    let rejected_xml = document_xml(&mut rejected);
    assert!(rejected_xml.find(before).unwrap() < rejected_xml.find("<w:p>").unwrap());
    assert!(
        rejected
            .compare(&paragraph, "postcondition", "2026-08-21T09:31:00Z")
            .unwrap()
            .is_empty()
    );
}

fn document_with_comparison_header(inner: &str) -> Document {
    let mut seed = Document::new();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(
        seed.to_bytes().expect("serialize comparison seed"),
    ))
    .expect("open comparison seed package");
    package.set_part(
        "/word/header1.xml",
        format!(
            r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:comparison-producer">{inner}</w:hdr>"#,
        )
        .into_bytes(),
    );
    package.content_types.add_override(
        "/word/header1.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    );
    let header_id = package
        .get_or_create_part_rels("/word/document.xml")
        .add(oxml_opc::relationship::rel_types::HEADER, "header1.xml");
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>same body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/></w:sectPr></w:body></w:document>"#,
        )
        .into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package
        .write_to(&mut bytes)
        .expect("serialize comparison fixture");
    Document::from_bytes(bytes.get_ref()).expect("open comparison fixture")
}

fn document_with_comparison_stories(value: &str) -> Document {
    document_with_comparison_story_values(value, value)
}

fn document_with_comparison_story_values(value: &str, comment_value: &str) -> Document {
    let mut seed = Document::new();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(
        seed.to_bytes().expect("serialize full-story seed"),
    ))
    .expect("open full-story seed package");
    let relationships = package.get_or_create_part_rels("/word/document.xml");
    let header_id = relationships.add(oxml_opc::relationship::rel_types::HEADER, "header1.xml");
    let footer_id = relationships.add(oxml_opc::relationship::rel_types::FOOTER, "footer1.xml");
    relationships.add(oxml_opc::relationship::rel_types::COMMENTS, "comments.xml");
    relationships.add(
        oxml_opc::relationship::rel_types::FOOTNOTES,
        "footnotes.xml",
    );
    relationships.add(oxml_opc::relationship::rel_types::ENDNOTES, "endnotes.xml");
    let word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    package.set_part(
        "/word/header1.xml",
        format!(
            r#"<w:hdr xmlns:w="{word}"><w:p><w:r><w:t>{value} header</w:t></w:r></w:p></w:hdr>"#
        )
        .into_bytes(),
    );
    package.set_part(
        "/word/footer1.xml",
        format!(
            r#"<w:ftr xmlns:w="{word}"><w:p><w:r><w:t>{value} footer</w:t></w:r></w:p></w:ftr>"#
        )
        .into_bytes(),
    );
    package.set_part(
        "/word/comments.xml",
        format!(r#"<w:comments xmlns:w="{word}"><w:comment w:id="0" w:author="Ada" w:initials="AL" w:date="2026-09-04T09:00:00Z"><w:p><w:r><w:t>{comment_value} comment</w:t></w:r></w:p></w:comment></w:comments>"#).into_bytes(),
    );
    package.set_part(
        "/word/footnotes.xml",
        format!(r#"<w:footnotes xmlns:w="{word}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t xml:space="preserve"> {value} footnote</w:t></w:r></w:p></w:footnote></w:footnotes>"#).into_bytes(),
    );
    package.set_part(
        "/word/endnotes.xml",
        format!(r#"<w:endnotes xmlns:w="{word}"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="1"><w:p><w:r><w:endnoteRef/><w:t xml:space="preserve"> {value} endnote</w:t></w:r></w:p></w:endnote></w:endnotes>"#).into_bytes(),
    );
    for (part, content_type) in [
        (
            "/word/header1.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        ),
        (
            "/word/footer1.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        ),
        (
            "/word/comments.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        ),
        (
            "/word/footnotes.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        ),
        (
            "/word/endnotes.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
        ),
    ] {
        package.content_types.add_override(part, content_type);
    }
    package.set_part(
        "/word/document.xml",
        format!(r#"<?xml version="1.0"?><w:document xmlns:w="{word}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>{value} body</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:p><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/><w:footerReference w:type="default" r:id="{footer_id}"/></w:sectPr></w:body></w:document>"#).into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package
        .write_to(&mut bytes)
        .expect("write full-story package");
    Document::from_bytes(bytes.get_ref()).expect("open full-story package")
}

fn comparison_part_xml(document: &mut Document, part_name: &str) -> String {
    let bytes = document.to_bytes().expect("serialize comparison document");
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes))
        .expect("open comparison package");
    String::from_utf8(
        package
            .get_part(part_name)
            .expect("comparison part exists")
            .to_vec(),
    )
    .expect("comparison part is UTF-8")
}

fn comparison_story_with_run_properties(
    mut document: Document,
    part_name: &str,
    run_start: &str,
    properties: &str,
) -> Document {
    let bytes = document.to_bytes().expect("serialize comparison story");
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes))
        .expect("open comparison story package");
    let source = String::from_utf8(
        package
            .get_part(part_name)
            .expect("comparison story exists")
            .to_vec(),
    )
    .expect("comparison story is UTF-8");
    let replacement = run_start.replacen("<w:r>", &format!("<w:r><w:rPr>{properties}</w:rPr>"), 1);
    let rewritten = source.replacen(run_start, &replacement, 1);
    assert_ne!(rewritten, source, "missing run source in {part_name}");
    package.set_part(part_name, rewritten.into_bytes());
    let mut output = std::io::Cursor::new(Vec::new());
    package
        .write_to(&mut output)
        .expect("serialize rewritten comparison story");
    Document::from_bytes(output.get_ref()).expect("open rewritten comparison story")
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct ComparisonOracleRecord {
    kind: &'static str,
    order: usize,
    story: &'static str,
    pair: Option<&'static str>,
    field_owner: Option<&'static str>,
    formatting: Option<&'static str>,
}

fn pinned_word_comparison_records() -> Vec<ComparisonOracleRecord> {
    let mut records = Vec::new();
    let mut push = |kind, story, pair, field_owner, formatting| {
        records.push(ComparisonOracleRecord {
            kind,
            order: records.len(),
            story,
            pair,
            field_owner,
            formatting,
        });
    };
    for story in [
        "body",
        "header",
        "footer",
        "comments",
        "footnote",
        "endnote",
        "header/textbox[0]",
    ] {
        push("del", story, None, None, None);
        push("ins", story, None, None, None);
    }
    push("moveFrom", "header", Some("move-0"), None, None);
    push("moveTo", "header", Some("move-0"), None, None);
    push("del", "header", None, Some("field-result"), None);
    push("ins", "header", None, Some("field-result"), None);
    push("del", "header", None, Some("field"), None);
    push("ins", "header", None, Some("field"), None);
    for formatting in ["pPrChange", "rPrChange", "tblPrChange", "sectPrChange"] {
        push(formatting, "header", None, None, Some(formatting));
    }
    records.last_mut().expect("section record").story = "body";
    records
}

fn push_oracle_replacement(
    records: &mut Vec<ComparisonOracleRecord>,
    story: &'static str,
    xml: &str,
    field_owner: Option<&'static str>,
) {
    assert!(xml.contains("<w:del"), "missing deletion in {story}: {xml}");
    assert!(
        xml.contains("<w:ins"),
        "missing insertion in {story}: {xml}"
    );
    for kind in ["del", "ins"] {
        records.push(ComparisonOracleRecord {
            kind,
            order: records.len(),
            story,
            pair: None,
            field_owner,
            formatting: None,
        });
    }
}

fn push_oracle_move(records: &mut Vec<ComparisonOracleRecord>, xml: &str) {
    assert!(xml.contains("<w:moveFrom"), "missing move source: {xml}");
    assert!(xml.contains("<w:moveTo"), "missing move target: {xml}");
    for kind in ["moveFrom", "moveTo"] {
        records.push(ComparisonOracleRecord {
            kind,
            order: records.len(),
            story: "header",
            pair: Some("move-0"),
            field_owner: None,
            formatting: None,
        });
    }
}

fn push_oracle_formatting(
    records: &mut Vec<ComparisonOracleRecord>,
    story: &'static str,
    kind: &'static str,
    xml: &str,
) {
    assert!(xml.contains(&format!("<w:{kind}")), "missing {kind}: {xml}");
    records.push(ComparisonOracleRecord {
        kind,
        order: records.len(),
        story,
        pair: None,
        field_owner: None,
        formatting: Some(kind),
    });
}

fn normalized_rdocx_comparison_records() -> Vec<ComparisonOracleRecord> {
    let mut records = Vec::new();
    let mut stories = document_with_comparison_stories("original");
    let edited_stories = document_with_comparison_stories("edited");
    stories
        .compare(&edited_stories, "Word", "2026-09-04T09:00:00Z")
        .expect("compare story oracle inputs");
    for (part, story) in [
        ("/word/document.xml", "body"),
        ("/word/header1.xml", "header"),
        ("/word/footer1.xml", "footer"),
        ("/word/comments.xml", "comments"),
        ("/word/footnotes.xml", "footnote"),
        ("/word/endnotes.xml", "endnote"),
    ] {
        push_oracle_replacement(
            &mut records,
            story,
            &comparison_part_xml(&mut stories, part),
            None,
        );
    }

    let text_box = |value: &str| {
        format!(
            r#"<w:p><w:r><w:pict><v:shape id="f234-box" style="width:100pt;height:50pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>{value} box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"#,
        )
    };
    let mut text_box_original = document_with_comparison_header(&text_box("old"));
    let text_box_edited = document_with_comparison_header(&text_box("new"));
    text_box_original
        .compare(&text_box_edited, "Word", "2026-09-04T09:00:00Z")
        .expect("compare text-box oracle inputs");
    push_oracle_replacement(
        &mut records,
        "header/textbox[0]",
        &comparison_part_xml(&mut text_box_original, "/word/header1.xml"),
        None,
    );

    let mut move_original = document_with_comparison_header(
        r#"<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>"#,
    );
    let move_edited = document_with_comparison_header(
        r#"<w:p><w:r><w:t>second</w:t></w:r></w:p><w:p><w:r><w:t>first</w:t></w:r></w:p>"#,
    );
    move_original
        .compare(&move_edited, "Word", "2026-09-04T09:00:00Z")
        .expect("compare move oracle inputs");
    push_oracle_move(
        &mut records,
        &comparison_part_xml(&mut move_original, "/word/header1.xml"),
    );

    let mut field_original = document_with_comparison_header(
        r#"<w:p><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>old cache</w:t></w:r></w:fldSimple></w:p><w:p><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>same cache</w:t></w:r></w:fldSimple></w:p>"#,
    );
    let field_edited = document_with_comparison_header(
        r#"<w:p><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>new cache</w:t></w:r></w:fldSimple></w:p><w:p><w:fldSimple w:instr="MERGEFIELD Other"><w:r><w:t>same cache</w:t></w:r></w:fldSimple></w:p>"#,
    );
    field_original
        .compare(&field_edited, "Word", "2026-09-04T09:00:00Z")
        .expect("compare field oracle inputs");
    let field_xml = comparison_part_xml(&mut field_original, "/word/header1.xml");
    let first_field_end = field_xml.find("</w:fldSimple>").expect("first field end");
    push_oracle_replacement(
        &mut records,
        "header",
        &field_xml[..first_field_end],
        Some("field-result"),
    );
    push_oracle_replacement(
        &mut records,
        "header",
        &field_xml[first_field_end..],
        Some("field"),
    );

    let property_pairs = [
        (
            r#"<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>same</w:t></w:r></w:p>"#,
            r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>same</w:t></w:r></w:p>"#,
            "pPrChange",
        ),
        (
            r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>same</w:t></w:r></w:p>"#,
            r#"<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>same</w:t></w:r></w:p>"#,
            "rPrChange",
        ),
        (
            r#"<w:tbl><w:tblPr><w:tblW w:w="100" w:type="dxa"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>same</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
            r#"<w:tbl><w:tblPr><w:tblW w:w="200" w:type="dxa"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>same</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
            "tblPrChange",
        ),
    ];
    for (original, edited, kind) in property_pairs {
        let mut original = document_with_comparison_header(original);
        let edited = document_with_comparison_header(edited);
        original
            .compare(&edited, "Word", "2026-09-04T09:00:00Z")
            .expect("compare property oracle inputs");
        push_oracle_formatting(
            &mut records,
            "header",
            kind,
            &comparison_part_xml(&mut original, "/word/header1.xml"),
        );
    }
    let mut section_original = document_with_content_controls(&wrap_word_body(
        r#"<w:p><w:r><w:t>same</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="1000" w:h="2000"/></w:sectPr>"#,
    ));
    let section_edited = document_with_content_controls(&wrap_word_body(
        r#"<w:p><w:r><w:t>same</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="1100" w:h="2000"/></w:sectPr>"#,
    ));
    section_original
        .compare(&section_edited, "Word", "2026-09-04T09:00:00Z")
        .expect("compare section oracle inputs");
    push_oracle_formatting(
        &mut records,
        "body",
        "sectPrChange",
        &document_xml(&mut section_original),
    );
    records
}

fn capture_word_comparison_case(
    directory: &std::path::Path,
    name: &str,
    mut original: Document,
    mut edited: Document,
) -> oxml_opc::OpcPackage {
    use std::process::Command;

    let original_path = directory.join(format!("{name}-original.docx"));
    let edited_path = directory.join(format!("{name}-edited.docx"));
    let compared_path = directory.join(format!("{name}-word-compared.docx"));
    std::fs::write(
        &original_path,
        original.to_bytes().expect("serialize Word original"),
    )
    .expect("write Word original");
    std::fs::write(
        &edited_path,
        edited.to_bytes().expect("serialize Word edited"),
    )
    .expect("write Word edited");
    let script = format!(
        r#"with timeout of 60 seconds
tell application "Microsoft Word"
set revised_path to (POSIX file "{}") as text
set output_path to (POSIX file "{}") as text
set original_document to open file name "{}" read only false add to recent files false
compare original_document path revised_path author name "Word" target compare target new detect format changes true ignore all comparison warnings true add to recent files false
set compared_document to active document
save as compared_document file name output_path file format format document default add to recent files false
end tell
end timeout"#,
        edited_path.display(),
        compared_path.display(),
        original_path.display(),
    );
    let output = Command::new("osascript")
        .args(["-e", &script])
        .output()
        .expect("run Word comparison automation");
    assert!(
        output.status.success(),
        "Word comparison {name} failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let cleanup_script = format!(
        r#"tell application "Microsoft Word"
set directory_path to (POSIX file "{}/") as text
repeat 2 times
set gate_document to active document
set opened_path to full name of gate_document as text
if opened_path starts with directory_path then
close gate_document saving no
else
error "unsafe non-F-234 Word document: " & opened_path
end if
end repeat
end tell"#,
        directory.display(),
    );
    let cleanup = Command::new("osascript")
        .args(["-e", &cleanup_script])
        .output()
        .expect("clean up Word comparison documents");
    assert!(
        cleanup.status.success(),
        "Word comparison {name} cleanup failed: {}",
        String::from_utf8_lossy(&cleanup.stderr)
    );
    println!("F-234 Word oracle {name}: {}", compared_path.display());
    oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(
        std::fs::read(&compared_path).expect("read Word comparison"),
    ))
    .expect("open Word comparison package")
}

fn word_part_xml(package: &oxml_opc::OpcPackage, part: &str) -> String {
    String::from_utf8(package.get_part(part).expect("Word part exists").to_vec())
        .expect("Word part is UTF-8")
}

#[test]
#[ignore = "requires installed Microsoft Word 16.104 GUI automation"]
fn capture_full_story_comparison_word_records() {
    use std::process::Command;
    use std::time::{SystemTime, UNIX_EPOCH};

    let version = Command::new("/usr/libexec/PlistBuddy")
        .args([
            "-c",
            "Print:CFBundleShortVersionString",
            "/Applications/Microsoft Word.app/Contents/Info.plist",
        ])
        .output()
        .expect("read Word version");
    let build = Command::new("/usr/libexec/PlistBuddy")
        .args([
            "-c",
            "Print:CFBundleVersion",
            "/Applications/Microsoft Word.app/Contents/Info.plist",
        ])
        .output()
        .expect("read Word build");
    assert_eq!(String::from_utf8_lossy(&version.stdout).trim(), "16.104");
    assert_eq!(
        String::from_utf8_lossy(&build.stdout).trim(),
        "16.104.25121423"
    );

    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock")
        .as_nanos();
    let capture_name = format!("{}-{nonce}", std::process::id());
    let directory = std::path::Path::new(
        "/Users/atulsharma/Library/Containers/com.microsoft.Word/Data/Documents/rdocx-f234-word-oracle",
    )
    .join(&capture_name);
    std::fs::create_dir_all(&directory).expect("create Word oracle directory");
    let mut records = Vec::new();

    let story_package = capture_word_comparison_case(
        &directory,
        "stories",
        document_with_comparison_story_values("original", "same"),
        document_with_comparison_story_values("edited", "same"),
    );
    for (part, story) in [
        ("/word/document.xml", "body"),
        ("/word/header1.xml", "header"),
        ("/word/footer1.xml", "footer"),
        ("/word/footnotes.xml", "footnote"),
        ("/word/endnotes.xml", "endnote"),
    ] {
        push_oracle_replacement(
            &mut records,
            story,
            &word_part_xml(&story_package, part),
            None,
        );
    }
    records.insert(
        6,
        ComparisonOracleRecord {
            kind: "del",
            order: 6,
            story: "comments",
            pair: None,
            field_owner: None,
            formatting: None,
        },
    );
    records.insert(
        7,
        ComparisonOracleRecord {
            kind: "ins",
            order: 7,
            story: "comments",
            pair: None,
            field_owner: None,
            formatting: None,
        },
    );
    for (index, record) in records.iter_mut().enumerate() {
        record.order = index;
    }
    let comment_package = capture_word_comparison_case(
        &directory,
        "comments",
        document_with_comparison_story_values("same", "original"),
        document_with_comparison_story_values("same", "edited"),
    );
    let comments = word_part_xml(&comment_package, "/word/comments.xml");
    assert!(comments.contains("original comment"), "{comments}");
    assert!(comments.contains("edited comment"), "{comments}");

    let text_box = |value: &str| {
        format!(
            r#"<w:p><w:r><w:pict><v:shape id="f234-box" style="width:100pt;height:50pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>{value} box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"#,
        )
    };
    let text_box_package = capture_word_comparison_case(
        &directory,
        "textbox",
        document_with_comparison_header(&text_box("old")),
        document_with_comparison_header(&text_box("new")),
    );
    push_oracle_replacement(
        &mut records,
        "header/textbox[0]",
        &word_part_xml(&text_box_package, "/word/header1.xml"),
        None,
    );

    let move_package = capture_word_comparison_case(
        &directory,
        "move",
        document_with_comparison_header(
            r#"<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>"#,
        ),
        document_with_comparison_header(
            r#"<w:p><w:r><w:t>second</w:t></w:r></w:p><w:p><w:r><w:t>first</w:t></w:r></w:p>"#,
        ),
    );
    let move_xml = word_part_xml(&move_package, "/word/header1.xml");
    let move_name = move_xml
        .split("<w:moveFromRangeStart")
        .nth(1)
        .and_then(|xml| xml.split("w:name=\"").nth(1))
        .and_then(|xml| xml.split('"').next())
        .expect("Word move range name");
    assert!(
        move_xml
            .split("<w:moveToRangeStart")
            .nth(1)
            .is_some_and(|xml| xml.contains(&format!("w:name=\"{move_name}\""))),
        "{move_xml}"
    );
    push_oracle_move(&mut records, &move_xml);

    let fields_package = capture_word_comparison_case(
        &directory,
        "fields",
        document_with_comparison_header(
            r#"<w:p><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>old cache</w:t></w:r></w:fldSimple></w:p><w:p><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>same cache</w:t></w:r></w:fldSimple></w:p>"#,
        ),
        document_with_comparison_header(
            r#"<w:p><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>new cache</w:t></w:r></w:fldSimple></w:p><w:p><w:fldSimple w:instr="MERGEFIELD Other"><w:r><w:t>same cache</w:t></w:r></w:fldSimple></w:p>"#,
        ),
    );
    let fields = word_part_xml(&fields_package, "/word/header1.xml");
    assert!(fields.contains("MERGEFIELD Other"), "{fields}");
    assert!(fields.contains("MERGEFIELD Name"), "{fields}");
    push_oracle_replacement(&mut records, "header", &fields, Some("field-result"));
    push_oracle_replacement(&mut records, "header", &fields, Some("field"));

    let properties_package = capture_word_comparison_case(
        &directory,
        "properties",
        document_with_comparison_header(
            r#"<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>paragraph</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>run</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="100" w:type="dxa"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        ),
        document_with_comparison_header(
            r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>paragraph</w:t></w:r></w:p><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>run</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="200" w:type="dxa"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        ),
    );
    let properties = word_part_xml(&properties_package, "/word/header1.xml");
    for kind in ["pPrChange", "rPrChange", "tblPrChange"] {
        push_oracle_formatting(&mut records, "header", kind, &properties);
    }
    let section_package = capture_word_comparison_case(
        &directory,
        "section",
        document_with_content_controls(&wrap_word_body(
            r#"<w:p><w:r><w:t>same</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="1000" w:h="2000"/></w:sectPr>"#,
        )),
        document_with_content_controls(&wrap_word_body(
            r#"<w:p><w:r><w:t>same</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="1100" w:h="2000"/></w:sectPr>"#,
        )),
    );
    push_oracle_formatting(
        &mut records,
        "body",
        "sectPrChange",
        &word_part_xml(&section_package, "/word/document.xml"),
    );
    assert_eq!(records, pinned_word_comparison_records());
    let inspection_directory =
        std::path::Path::new("/Users/atulsharma/Downloads/rdocx-f234-word-oracle")
            .join(capture_name);
    std::fs::create_dir_all(&inspection_directory)
        .expect("create Word oracle inspection directory");
    for entry in std::fs::read_dir(&directory).expect("read Word oracle capture directory") {
        let entry = entry.expect("read Word oracle capture entry");
        if entry
            .file_type()
            .expect("read capture entry type")
            .is_file()
        {
            std::fs::copy(entry.path(), inspection_directory.join(entry.file_name()))
                .expect("copy Word oracle evidence for inspection");
        }
    }
    println!(
        "F-234 preserved Word oracle evidence: {}",
        inspection_directory.display()
    );
}

#[test]
fn full_story_comparison_matches_pinned_word_records() {
    assert_eq!(
        WORD_COMPARISON_ORACLE,
        "Microsoft Word 16.104 build 16.104.25121423"
    );
    assert_eq!(
        WORD_COMPARISON_ENVIRONMENT,
        "locale=en-US; normalization=revision-records-v1"
    );
    assert_eq!(
        normalized_rdocx_comparison_records(),
        pinned_word_comparison_records(),
        "{WORD_COMPARISON_ORACLE}"
    );
    let original_run = r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>old header</w:t></w:r></w:p>"#;
    let edited_run = r#"<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>new header</w:t></w:r></w:p>"#;
    let mut original = document_with_comparison_header(original_run);
    let edited = document_with_comparison_header(edited_run);

    original
        .compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("full-story comparison");
    let tracked = comparison_part_xml(&mut original, "/word/header1.xml");
    assert!(tracked.contains("<w:del"), "{tracked}");
    assert!(tracked.contains("<w:ins"), "{tracked}");

    for (original_owner, edited_owner, revision_kind) in [
        (
            r#"<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>same</w:t></w:r></w:p>"#,
            r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>same</w:t></w:r></w:p>"#,
            "pPrChange",
        ),
        (
            r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>same</w:t></w:r></w:p>"#,
            r#"<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>same</w:t></w:r></w:p>"#,
            "rPrChange",
        ),
        (
            r#"<w:tbl><w:tblPr><w:tblW w:w="100" w:type="dxa"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>same</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
            r#"<w:tbl><w:tblPr><w:tblW w:w="200" w:type="dxa"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>same</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
            "tblPrChange",
        ),
    ] {
        let mut original = document_with_comparison_header(original_owner);
        let edited = document_with_comparison_header(edited_owner);
        original
            .compare(&edited, "Word", "2026-09-04T09:00:00Z")
            .expect("modeled property comparison");
        let xml = comparison_part_xml(&mut original, "/word/header1.xml");
        assert!(xml.contains(&format!("<w:{revision_kind}")), "{xml}");
        let bytes = original.to_bytes().expect("serialize property comparison");
        let mut accepted = Document::from_bytes(&bytes).expect("open accepted property copy");
        accepted.accept_all().expect("accept property change");
        assert!(
            accepted
                .compare(&edited, "postcondition", "2026-09-04T09:01:00Z")
                .expect("accepted property postcondition")
                .is_empty()
        );
    }

    let mut original = document_with_comparison_header(
        r#"<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>old cache</w:t></w:r></w:fldSimple></w:p>"#,
    );
    let edited = document_with_comparison_header(
        r#"<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>new cache</w:t></w:r></w:fldSimple></w:p>"#,
    );
    original
        .compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("field result comparison");
    let field = comparison_part_xml(&mut original, "/word/header1.xml");
    let field_start = field.find("<w:fldSimple").expect("field owner");
    let field_end = field.find("</w:fldSimple>").expect("field owner end");
    assert!(field[field_start..field_end].contains("<w:del"), "{field}");
    assert!(field[field_start..field_end].contains("<w:ins"), "{field}");

    let mut original = document_with_comparison_header(
        r#"<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>cache</w:t></w:r></w:fldSimple></w:p>"#,
    );
    let edited = document_with_comparison_header(
        r#"<w:p><w:fldSimple w:instr="TIME"><w:r><w:t>cache</w:t></w:r></w:fldSimple></w:p>"#,
    );
    original
        .compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("field instruction comparison");
    let instruction = comparison_part_xml(&mut original, "/word/header1.xml");
    assert_eq!(
        instruction.matches("<w:fldSimple").count(),
        2,
        "{instruction}"
    );
    let bytes = original
        .to_bytes()
        .expect("serialize instruction comparison");
    let mut accepted = Document::from_bytes(&bytes).expect("open accepted instruction copy");
    accepted.accept_all().expect("accept instruction change");
    assert!(comparison_part_xml(&mut accepted, "/word/header1.xml").contains("TIME"));
    let mut rejected = Document::from_bytes(&bytes).expect("open rejected instruction copy");
    rejected.reject_all().expect("reject instruction change");
    assert!(comparison_part_xml(&mut rejected, "/word/header1.xml").contains("DATE"));

    let original_main = wrap_word_body(
        r#"<w:p><w:r><w:t>same</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="1000" w:h="2000"/></w:sectPr>"#,
    );
    let edited_main = wrap_word_body(
        r#"<w:p><w:r><w:t>same</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="1100" w:h="2000"/></w:sectPr>"#,
    );
    let mut original = document_with_content_controls(&original_main);
    let edited = document_with_content_controls(&edited_main);
    original
        .compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("section property comparison");
    assert!(document_xml(&mut original).contains("<w:sectPrChange"));

    let text_box = |value: &str| {
        format!(
            r#"<w:p><w:r><w:drawing><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>{value} box</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></w:drawing></w:r></w:p>"#,
        )
    };
    let mut original = document_with_comparison_header(&text_box("old"));
    let edited = document_with_comparison_header(&text_box("new"));
    original
        .compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("nested text-box comparison");
    let tracked_box = comparison_part_xml(&mut original, "/word/header1.xml");
    assert!(tracked_box.contains("txbxContent"), "{tracked_box}");
    assert!(tracked_box.contains("<w:del"), "{tracked_box}");
    assert!(tracked_box.contains("<w:ins"), "{tracked_box}");

    let mut original = document_with_comparison_header(&format!(
        r#"<w:p><w:r><w:t>old host</w:t></w:r></w:p>{}"#,
        text_box("old")
    ));
    let edited = document_with_comparison_header(&format!(
        r#"<w:p><w:r><w:t>new host</w:t></w:r></w:p>{}"#,
        text_box("new")
    ));
    original
        .compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("simultaneous host and text-box comparison");
    let tracked_box = comparison_part_xml(&mut original, "/word/header1.xml");
    assert!(tracked_box.contains("old host"), "{tracked_box}");
    assert!(tracked_box.contains("new host"), "{tracked_box}");
    assert!(tracked_box.contains("old box"), "{tracked_box}");
    assert!(tracked_box.contains("new box"), "{tracked_box}");
    let mut accepted = Document::from_bytes(&original.to_bytes().unwrap()).unwrap();
    accepted.accept_all().unwrap();
    assert!(
        accepted
            .compare(&edited, "postcondition", "2026-09-04T09:01:00Z")
            .unwrap()
            .is_empty()
    );

    let aliased_box = |value: &str| {
        text_box(value)
            .replace(
                "<w:txbxContent>",
                r#"<q:txbxContent xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
            )
            .replace("</w:txbxContent>", "</q:txbxContent>")
    };
    let mut aliased = document_with_comparison_header(&aliased_box("old"));
    let aliased_edited = document_with_comparison_header(&aliased_box("new"));
    aliased
        .compare(&aliased_edited, "Word", "2026-09-04T09:00:00Z")
        .expect("locally aliased nested text box");
    let aliased_xml = comparison_part_xml(&mut aliased, "/word/header1.xml");
    assert!(aliased_xml.contains("<q:txbxContent"), "{aliased_xml}");
    assert!(aliased_xml.contains("old box"), "{aliased_xml}");
    assert!(aliased_xml.contains("new box"), "{aliased_xml}");
}

#[test]
fn run_property_changes_preserve_original_unmodelled_children_and_report_the_difference() {
    let original_raw =
        r#"<x:original xmlns:x="urn:f234-run-property" x:value="keep"><x:nested/></x:original>"#;
    let edited_raw = r#"<x:edited xmlns:x="urn:f234-run-property" x:value="discard"/>"#;
    let mut original = document_with_comparison_header(&format!(
        r#"<w:p><w:r><w:rPr><w:b/>{original_raw}</w:rPr><w:t>same</w:t></w:r></w:p>"#,
    ));
    let edited = document_with_comparison_header(&format!(
        r#"<w:p><w:r><w:rPr><w:i/>{edited_raw}</w:rPr><w:t>same</w:t></w:r></w:p>"#,
    ));

    let diagnostics = original
        .compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("compare modeled and unmodelled run properties");
    assert_eq!(
        diagnostics,
        vec![rdocx::ComparisonDiagnostic {
            location: "header:/word/header1.xml/paragraph[0]/run[0]".to_owned(),
            message: "formatting differs and the original formatting was retained".to_owned(),
        }]
    );

    let tracked_bytes = original.to_bytes().expect("serialize tracked comparison");
    let tracked = comparison_part_xml(&mut original, "/word/header1.xml");
    assert!(tracked.contains("<w:i/>"), "{tracked}");
    assert!(tracked.contains("<w:rPrChange"), "{tracked}");
    assert_eq!(tracked.matches(original_raw).count(), 1, "{tracked}");
    assert!(!tracked.contains(edited_raw), "{tracked}");

    let mut accepted = Document::from_bytes(&tracked_bytes).expect("open accepted copy");
    accepted.accept_all().expect("accept run property change");
    let accepted = comparison_part_xml(&mut accepted, "/word/header1.xml");
    assert!(accepted.contains("<w:i/>"), "{accepted}");
    assert!(!accepted.contains("<w:b/>"), "{accepted}");
    assert_eq!(accepted.matches(original_raw).count(), 1, "{accepted}");
    assert!(!accepted.contains(edited_raw), "{accepted}");

    let mut rejected = Document::from_bytes(&tracked_bytes).expect("open rejected copy");
    rejected.reject_all().expect("reject run property change");
    let rejected = comparison_part_xml(&mut rejected, "/word/header1.xml");
    assert!(rejected.contains("<w:b/>"), "{rejected}");
    assert!(!rejected.contains("<w:i/>"), "{rejected}");
    assert_eq!(rejected.matches(original_raw).count(), 1, "{rejected}");
    assert!(!rejected.contains(edited_raw), "{rejected}");
}

#[test]
fn related_story_formatting_diagnostics_retain_deterministic_story_locations() {
    let original_properties = r#"<w:b/><x:original xmlns:x="urn:f234-diagnostic" x:value="keep"/>"#;
    let edited_properties = r#"<w:b/><x:edited xmlns:x="urn:f234-diagnostic" x:value="discard"/>"#;
    for (part_name, run_start, expected_location) in [
        (
            "/word/header1.xml",
            "<w:r><w:t>same header",
            "header:/word/header1.xml/paragraph[0]/run[0]",
        ),
        (
            "/word/footer1.xml",
            "<w:r><w:t>same footer",
            "footer:/word/footer1.xml/paragraph[0]/run[0]",
        ),
        (
            "/word/comments.xml",
            "<w:r><w:t>same comment",
            "comments:/word/comments.xml/comment[0]/paragraph[0]/run[0]",
        ),
        (
            "/word/footnotes.xml",
            "<w:r><w:footnoteRef/>",
            "footnotes:/word/footnotes.xml/footnote[2]/paragraph[0]/run[0]",
        ),
        (
            "/word/endnotes.xml",
            "<w:r><w:endnoteRef/>",
            "endnotes:/word/endnotes.xml/endnote[2]/paragraph[0]/run[0]",
        ),
    ] {
        let mut original = comparison_story_with_run_properties(
            document_with_comparison_stories("same"),
            part_name,
            run_start,
            original_properties,
        );
        let edited = comparison_story_with_run_properties(
            document_with_comparison_stories("same"),
            part_name,
            run_start,
            edited_properties,
        );

        assert_eq!(
            original
                .compare(&edited, "Word", "2026-09-04T09:00:00Z")
                .expect("compare related-story formatting"),
            vec![rdocx::ComparisonDiagnostic {
                location: expected_location.to_owned(),
                message: "formatting differs and the original formatting was retained".to_owned(),
            }],
            "{part_name}"
        );
    }
}

#[test]
fn nested_text_box_formatting_diagnostic_retains_its_host_story_location() {
    let text_box = |raw_property: &str| {
        format!(
            r#"<w:p><w:r><w:drawing><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:b/>{raw_property}</w:rPr><w:t>same box</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></w:drawing></w:r></w:p>"#,
        )
    };
    let mut original = document_with_comparison_header(&text_box(
        r#"<x:original x:value="keep"><x:nested/></x:original>"#,
    ));
    let edited = document_with_comparison_header(&text_box(r#"<x:edited x:value="discard"/>"#));

    assert_eq!(
        original
            .compare(&edited, "Word", "2026-09-04T09:00:00Z")
            .expect("compare nested text-box formatting"),
        vec![rdocx::ComparisonDiagnostic {
            location: "header:/word/header1.xml/text-box[0:0]/paragraph[0]/run[0]".to_owned(),
            message: "formatting differs and the original formatting was retained".to_owned(),
        }]
    );
}

#[test]
fn full_story_comparison_differential_rejects_kind_order_and_story_perturbations() {
    let actual = normalized_rdocx_comparison_records();
    let expected = pinned_word_comparison_records();
    assert_eq!(actual, expected, "{WORD_COMPARISON_ORACLE}");
    for (label, index, replacement) in [
        (
            "kind",
            14usize,
            ComparisonOracleRecord {
                kind: "del",
                ..expected[14].clone()
            },
        ),
        (
            "order",
            14,
            ComparisonOracleRecord {
                order: 15,
                ..expected[14].clone()
            },
        ),
        (
            "story",
            14,
            ComparisonOracleRecord {
                story: "footer",
                ..expected[14].clone()
            },
        ),
        (
            "move-pair",
            15,
            ComparisonOracleRecord {
                pair: Some("wrong-pair"),
                ..expected[15].clone()
            },
        ),
        (
            "field-owner",
            16,
            ComparisonOracleRecord {
                field_owner: None,
                ..expected[16].clone()
            },
        ),
        (
            "formatting",
            20,
            ComparisonOracleRecord {
                formatting: Some("rPrChange"),
                ..expected[20].clone()
            },
        ),
    ] {
        let mut mutated = expected.clone();
        mutated[index] = replacement;
        assert_ne!(
            mutated, expected,
            "{label} mutation must fail the predicate"
        );
    }
}

#[test]
fn accepting_and_rejecting_full_story_comparison_reproduces_each_supported_story() {
    let mut tracked = document_with_comparison_stories("original");
    let edited = document_with_comparison_stories("edited");
    tracked
        .compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("full-story comparison");

    let tracked_bytes = tracked.to_bytes().expect("serialize tracked document");
    let mut accepted = Document::from_bytes(&tracked_bytes).expect("open accepted copy");
    accepted.accept_all().expect("accept comparison");
    for part in [
        "/word/document.xml",
        "/word/header1.xml",
        "/word/footer1.xml",
        "/word/comments.xml",
        "/word/footnotes.xml",
        "/word/endnotes.xml",
    ] {
        assert!(
            comparison_part_xml(&mut accepted, part).contains("edited"),
            "{part}"
        );
    }
    let mut rejected = Document::from_bytes(&tracked_bytes).expect("open rejected copy");
    rejected.reject_all().expect("reject comparison");
    for part in [
        "/word/document.xml",
        "/word/header1.xml",
        "/word/footer1.xml",
        "/word/comments.xml",
        "/word/footnotes.xml",
        "/word/endnotes.xml",
    ] {
        assert!(
            comparison_part_xml(&mut rejected, part).contains("original"),
            "{part}"
        );
    }
}

#[test]
fn scoped_revision_resolution_visits_every_compared_story_once() {
    let mut document = document_with_comparison_header(
        r#"<w:p><w:ins w:id="71" w:author="Word"><w:r><w:t>selected header</w:t></w:r></w:ins><w:ins w:id="72" w:author="Other"><w:r><w:t>retained header</w:t></w:r></w:ins></w:p>"#,
    );

    assert_eq!(
        document
            .accept_revision_id(71)
            .expect("accept header revision"),
        1
    );
    let resolved = comparison_part_xml(&mut document, "/word/header1.xml");
    assert!(!resolved.contains(r#"w:id="71""#), "{resolved}");
    assert!(resolved.contains(r#"w:id="72""#), "{resolved}");
}

#[test]
fn comparison_preserves_word_source_paths_in_every_revision_view() {
    let mut original =
        document_with_comparison_header(r#"<w:p><w:r><w:t>source original</w:t></w:r></w:p>"#);
    let edited =
        document_with_comparison_header(r#"<w:p><w:r><w:t>source edited</w:t></w:r></w:p>"#);
    original
        .compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("full-story comparison");

    let tracked_bytes = original
        .to_bytes()
        .expect("serialize tracked source fixture");
    let tracked = Document::from_bytes(&tracked_bytes).expect("open tracked source view");
    let mut accepted = Document::from_bytes(&tracked_bytes).expect("open accepted source view");
    accepted.accept_all().expect("accept source view");
    let mut rejected = Document::from_bytes(&tracked_bytes).expect("open rejected source view");
    rejected.reject_all().expect("reject source view");
    for (view, document) in [
        (rdocx_layout::RevisionView::Tracked, tracked),
        (rdocx_layout::RevisionView::Accepted, accepted),
        (rdocx_layout::RevisionView::Accepted, rejected),
    ] {
        let result = document
            .layout_with_options(RenderOptions {
                revision_view: view,
            })
            .expect("layout revision source view");
        let mut header_sources = 0usize;
        let mut actual_paths = Vec::new();
        for page in &result.layout.pages {
            for element in compatibility_page_elements(&page.elements) {
                let oxml_layout::PositionedElement::Text(run) = element else {
                    continue;
                };
                let Some(source) = run.source else {
                    continue;
                };
                assert_eq!(
                    source.char_end - source.char_start,
                    run.text.chars().count() as u32,
                    "scalar source range for {view:?} and {:?}",
                    run.text
                );
                if matches!(
                    result.source_node(source.node),
                    Some(rdocx_layout::WordSourcePath {
                        story: rdocx_layout::WordStory::Header { relationship_id },
                        children,
                    }) if relationship_id == "rId1" && children == &[0]
                ) {
                    header_sources += 1;
                }
                actual_paths.push((run.text.clone(), result.source_node(source.node).cloned()));
            }
        }
        assert!(
            header_sources > 0,
            "missing exact header source path in {view:?}: {actual_paths:?}"
        );
    }
}

#[test]
fn comparison_preserves_unmodelled_story_xml_and_relationships_byte_for_byte() {
    let leading = "\n  <!--leading exact-->\n<?producer leading?>\t";
    let raw = r#"<x:owner xmlns:q="urn:f234-local-binding" x:value="keep"><q:nested/><?producer inside?></x:owner>"#;
    let before_paragraph = "\r\n <!--before paragraph--><?producer paragraph?> \n";
    let before_table = "\n\t<!--before table--> <?producer table?>\r\n";
    let before_text_box = "\r\n<?producer text-box?><!--before text box-->  ";
    let trailing = "\n <?producer trailing?><!--trailing exact-->\t\n";
    let text_box = |value: &str| {
        format!(
            r#"<w:p><w:r><w:drawing><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>{value} box</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></w:drawing></w:r></w:p>"#,
        )
    };
    let story = |value: &str| {
        format!(
            r#"{leading}{raw}{before_paragraph}<w:p><w:r><w:t>{value} paragraph</w:t></w:r></w:p>{before_table}<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>{value} table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>{before_text_box}{}{trailing}"#,
            text_box(value)
        )
    };
    let mut original = document_with_comparison_header(&story("original"));
    let edited = document_with_comparison_header(&story("edited"));

    original
        .compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("full-story comparison");
    let tracked_bytes = original.to_bytes().expect("save tracked comparison");
    let header_xml = |bytes: &[u8]| {
        let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes))
            .expect("open comparison view");
        String::from_utf8(
            package
                .get_part("/word/header1.xml")
                .expect("header story exists")
                .to_vec(),
        )
        .expect("header story is UTF-8")
    };
    let unowned = [
        leading,
        raw,
        before_paragraph,
        before_table,
        before_text_box,
        trailing,
    ];
    let assert_unowned = |view: &str, xml: &str| {
        let mut previous = 0usize;
        for segment in unowned {
            let matches = xml.match_indices(segment).collect::<Vec<_>>();
            assert_eq!(matches.len(), 1, "{view} changed {segment:?}: {xml}");
            let at = matches[0].0;
            assert!(at >= previous, "{view} reordered {segment:?}: {xml}");
            assert_eq!(
                &xml.as_bytes()[at..at + segment.len()],
                segment.as_bytes(),
                "{view} changed bytes in {segment:?}"
            );
            previous = at + segment.len();
        }
    };

    let tracked = header_xml(&tracked_bytes);
    assert_unowned("tracked save", &tracked);
    assert!(tracked.contains("<w:ins"), "{tracked}");
    assert!(tracked.contains("original paragraph"), "{tracked}");
    assert!(tracked.contains("edited paragraph"), "{tracked}");
    assert!(tracked.contains("original table"), "{tracked}");
    assert!(tracked.contains("edited table"), "{tracked}");
    assert!(tracked.contains("original box"), "{tracked}");
    assert!(tracked.contains("edited box"), "{tracked}");

    let mut tracked_reopened =
        Document::from_bytes(&tracked_bytes).expect("reopen tracked comparison");
    let tracked_reopened = tracked_reopened
        .to_bytes()
        .expect("save reopened tracked view");
    assert_unowned("tracked reopen", &header_xml(&tracked_reopened));

    let mut accepted = Document::from_bytes(&tracked_bytes).expect("open accepted view");
    accepted.accept_all().expect("accept comparison");
    let accepted_bytes = accepted.to_bytes().expect("save accepted view");
    let accepted_xml = header_xml(&accepted_bytes);
    assert_unowned("accepted save", &accepted_xml);
    assert!(accepted_xml.contains("edited paragraph"), "{accepted_xml}");
    assert!(accepted_xml.contains("edited table"), "{accepted_xml}");
    assert!(accepted_xml.contains("edited box"), "{accepted_xml}");
    assert!(
        !accepted_xml.contains("original paragraph"),
        "{accepted_xml}"
    );
    let mut accepted_reopened =
        Document::from_bytes(&accepted_bytes).expect("reopen accepted view");
    let accepted_reopened = accepted_reopened
        .to_bytes()
        .expect("save reopened accepted view");
    assert_unowned("accepted reopen", &header_xml(&accepted_reopened));

    let mut rejected = Document::from_bytes(&tracked_bytes).expect("open rejected view");
    rejected.reject_all().expect("reject comparison");
    let rejected_bytes = rejected.to_bytes().expect("save rejected view");
    let rejected_xml = header_xml(&rejected_bytes);
    assert_unowned("rejected save", &rejected_xml);
    assert!(
        rejected_xml.contains("original paragraph"),
        "{rejected_xml}"
    );
    assert!(rejected_xml.contains("original table"), "{rejected_xml}");
    assert!(rejected_xml.contains("original box"), "{rejected_xml}");
    assert!(!rejected_xml.contains("edited paragraph"), "{rejected_xml}");
    let mut rejected_reopened =
        Document::from_bytes(&rejected_bytes).expect("reopen rejected view");
    let rejected_reopened = rejected_reopened
        .to_bytes()
        .expect("save reopened rejected view");
    assert_unowned("rejected reopen", &header_xml(&rejected_reopened));
}

#[test]
fn a_failed_full_story_comparison_leaves_package_typed_state_and_caches_unchanged() {
    let mut original =
        document_with_comparison_header(r#"<w:p><w:r><w:t>same header</w:t></w:r></w:p>"#);
    let edited = document_with_comparison_header(
        r#"<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"#,
    );
    let before = original.to_bytes().expect("serialize original");

    assert!(
        original
            .compare(&edited, "Word", "2026-09-04T09:00:00Z")
            .is_err()
    );
    assert_eq!(
        original.to_bytes().expect("serialize after failure"),
        before
    );
}

#[test]
fn repeated_story_content_and_moves_use_deterministic_matches_and_ids() {
    let first = r#"<w:p><w:r><w:t>repeat</w:t></w:r></w:p><w:p><w:r><w:t>anchor</w:t></w:r></w:p><w:p><w:r><w:t>repeat</w:t></w:r></w:p>"#;
    let second = r#"<w:p><w:r><w:t>repeat</w:t></w:r></w:p><w:p><w:r><w:t>repeat</w:t></w:r></w:p><w:p><w:r><w:t>anchor</w:t></w:r></w:p>"#;
    let edited = document_with_comparison_header(second);
    let mut left = document_with_comparison_header(first);
    let mut right = document_with_comparison_header(first);

    left.compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("first comparison");
    right
        .compare(&edited, "Word", "2026-09-04T09:00:00Z")
        .expect("second comparison");
    let left_xml = comparison_part_xml(&mut left, "/word/header1.xml");
    let right_xml = comparison_part_xml(&mut right, "/word/header1.xml");
    assert_eq!(left_xml, right_xml);
    assert!(left_xml.contains("<w:moveFrom"), "{left_xml}");
}

#[test]
fn fixed_break_runs_match_pdf_and_raster_backends() {
    let mut document = Document::new();
    document.add_paragraph(&"financial planning ttf-parser double  spaces allocated ".repeat(12));
    let (family, bytes) = oxml_layout::bundled_fonts::bundled_font_data()[0];
    let layout = document
        .layout_with_fonts(&[(family, bytes)])
        .expect("caller-font deterministic layout");
    let mut fonts =
        oxml_layout::FontManager::new_with_fonts(vec![(family.to_owned(), bytes.to_vec())]);
    for run in layout.layout.pages.iter().flat_map(|page| {
        page.elements.iter().filter_map(|element| match element {
            oxml_layout::PositionedElement::Text(run) if run.source.is_some() => Some(run),
            _ => None,
        })
    }) {
        let font_family = layout
            .layout
            .fonts
            .iter()
            .find(|font| font.id == run.font_id)
            .expect("run font is in result")
            .family
            .clone();
        let font_id = fonts
            .resolve_font(Some(&font_family), run.bold, run.italic)
            .expect("caller run font resolves");
        let independently_shaped = fonts
            .shape_text(font_id, &run.text, run.font_size)
            .expect("emitted chunk reshapes");
        assert_eq!(run.glyph_ids, independently_shaped.glyph_ids);
        assert_eq!(run.advances, independently_shaped.advances);
    }

    let direct_pdf = oxml_pdf::render_to_pdf(&layout.layout);
    let direct_png = oxml_pdf::render_page_to_png(&layout.layout, 0, 96.0)
        .expect("raster backend renders first page");
    assert_eq!(
        document
            .to_pdf_with_fonts(&[(family, bytes)])
            .expect("PDF facade"),
        direct_pdf
    );
    assert!(!direct_png.is_empty());
}

#[test]
fn complex_shaping_preserves_clusters_offsets_and_logical_source_spans_across_word_backends() {
    let samples = [
        ("Noto Sans Arabic", "ar-SA", "العربية مرحبا بالعالم"),
        ("Noto Sans Devanagari", "hi-IN", "देवनागरी नमस्ते दुनिया"),
        ("Noto Sans Thai", "th-TH", "ภาษาไทยยินดีต้อนรับ"),
        ("Noto Sans SC", "zh-CN", "〈中〉、你好世界"),
    ];
    let mut document = Document::new();
    let mut paragraph = document.add_paragraph("");
    for (index, (family, language, text)) in samples.iter().enumerate() {
        if index > 0 {
            paragraph.add_run("  ");
        }
        paragraph.add_run(text).font(family).language(language);
    }

    let result = document
        .layout_deterministic()
        .expect("deterministic multilingual Word layout");
    let mut rich_runs = Vec::new();
    for page in &result.layout.pages {
        oxml_layout::walk(&page.elements, &mut |element, _| {
            let oxml_layout::PositionedElement::MultilingualText(run) = element else {
                return;
            };
            assert!(run.is_valid(), "Word layout emits a complete rich run");
            assert!(
                run.source
                    .is_some_and(|span| result.source_node(span.node).is_some()),
                "rich run retains resolvable Word source provenance"
            );
            assert_eq!(run.glyph_ids.len(), run.x_advances.len());
            assert_eq!(run.glyph_ids.len(), run.y_advances.len());
            assert_eq!(run.glyph_ids.len(), run.x_offsets.len());
            assert_eq!(run.glyph_ids.len(), run.y_offsets.len());
            assert!(!run.clusters.is_empty());
            rich_runs.push(run.clone());
        });
    }
    assert!(rich_runs.len() >= samples.len(), "one rich span per script");
    for (_, language, text) in samples {
        let mut language_runs = rich_runs
            .iter()
            .filter(|run| run.language.as_deref() == Some(language))
            .collect::<Vec<_>>();
        language_runs.sort_by_key(|run| run.logical_index);
        assert_eq!(
            language_runs
                .iter()
                .map(|run| run.logical_text.as_str())
                .collect::<String>(),
            text,
            "rich output retains {language} logical text"
        );
    }

    let pdf = document
        .to_pdf_deterministic()
        .expect("multilingual PDF renders");
    assert!(pdf.starts_with(b"%PDF-"));
    let png = document
        .render_page_to_png_deterministic(0, 96.0)
        .expect("multilingual raster renders")
        .expect("multilingual first page exists");
    assert!(png.starts_with(b"\x89PNG\r\n\x1a\n"));
    let svg = document
        .render_page_to_svg_deterministic(0)
        .expect("multilingual SVG renders")
        .expect("multilingual first SVG page exists");
    for run in &rich_runs {
        assert!(
            svg.svg.contains(&run.logical_text),
            "SVG preserves searchable logical text for {}",
            run.logical_text
        );
    }
}

fn empty_story_layout_input() -> rdocx_layout::LayoutInput {
    use rdocx_oxml::footnotes::{CT_Footnote, CT_Footnotes, NoteType};
    use rdocx_oxml::header_footer::CT_HdrFtr;

    let document = CT_Document::from_xml(
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p/><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:pPr/></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:footnoteReference w:id="4"/><w:endnoteReference w:id="9"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>"#,
    )
    .expect("empty story document parses");
    let mut header = CT_HdrFtr::new();
    header.paragraphs.push(rdocx_oxml::text::CT_P::new());
    let mut footer = CT_HdrFtr::new();
    footer.paragraphs.push(rdocx_oxml::text::CT_P::new());
    let empty_note = |id| CT_Footnote {
        id,
        note_type: NoteType::Normal,
        paragraphs: vec![rdocx_oxml::text::CT_P::new()],
    };

    rdocx_layout::LayoutInput {
        automatic_hyphenation: false,
        math_properties: None,
        document,
        styles: rdocx_oxml::styles::CT_Styles::new_default(),
        numbering: None,
        headers: HashMap::from([("rIdHeader".to_owned(), header)]),
        footers: HashMap::from([("rIdFooter".to_owned(), footer)]),
        images: HashMap::new(),
        charts: HashMap::new(),
        chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
        chart_color_map: oxml_drawing::color::ColorMap::default(),
        core_properties: None,
        hyperlink_urls: HashMap::new(),
        footnotes: Some(CT_Footnotes {
            footnotes: vec![empty_note(4)],
        }),
        endnotes: Some(CT_Footnotes {
            footnotes: vec![empty_note(9)],
        }),
        theme: None,
        fonts: Vec::new(),
        revision_view: rdocx_layout::RevisionView::Accepted,
    }
}

#[test]
fn empty_word_stories_emit_one_attributed_zero_width_segment() {
    use rdocx_layout::{WordSourcePath, WordStory};

    let result =
        rdocx_layout::layout_document_deterministic_with_provenance(&empty_story_layout_input())
            .expect("empty Word stories lay out");
    let attributed = result
        .layout
        .pages
        .iter()
        .flat_map(|page| compatibility_page_elements(&page.elements))
        .filter_map(|element| match element {
            oxml_layout::PositionedElement::Text(run) if run.text.is_empty() => Some(run),
            _ => None,
        })
        .collect::<Vec<_>>();
    let actual = attributed
        .iter()
        .filter_map(|run| run.source)
        .filter_map(|source| result.source_node(source.node))
        .collect::<Vec<_>>();
    assert_eq!(attributed.len(), 6, "attributed empty stories: {actual:?}");
    let expected = [
        WordSourcePath {
            story: WordStory::Document,
            children: vec![0],
        },
        WordSourcePath {
            story: WordStory::Document,
            children: vec![1, 0, 0, 0],
        },
        WordSourcePath {
            story: WordStory::Header {
                relationship_id: "rIdHeader".to_owned(),
            },
            children: vec![0],
        },
        WordSourcePath {
            story: WordStory::Footer {
                relationship_id: "rIdFooter".to_owned(),
            },
            children: vec![0],
        },
        WordSourcePath {
            story: WordStory::Footnote { id: 4 },
            children: vec![0],
        },
        WordSourcePath {
            story: WordStory::Endnote { id: 9 },
            children: vec![0],
        },
    ];
    for path in expected {
        let matching = attributed
            .iter()
            .filter(|run| {
                run.source.is_some_and(|source| {
                    source.char_start == 0
                        && source.char_end == 0
                        && result.source_node(source.node) == Some(&path)
                })
            })
            .collect::<Vec<_>>();
        assert_eq!(matching.len(), 1, "missing or duplicate caret for {path:?}");
        assert_eq!(matching[0].advances.iter().sum::<f64>(), 0.0);
        assert_eq!(matching[0].advances, Vec::<f64>::new());
        assert_eq!(matching[0].glyph_ids, Vec::<u16>::new());
    }
}

#[test]
fn empty_paragraph_uses_resolved_default_metrics() {
    use rdocx_oxml::properties::{CT_PPr, CT_RPr};
    use rdocx_oxml::units::HalfPoint;

    let mut input = empty_story_layout_input();
    input.document = CT_Document::from_xml(
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:p/></w:body></w:document>"#,
    )
    .expect("metric document parses");
    input.styles.styles[0].rpr = Some(CT_RPr {
        font_ascii: Some("Caladea".to_owned()),
        font_hansi: Some("Caladea".to_owned()),
        sz: Some(HalfPoint(28)),
        ..Default::default()
    });
    let BodyContent::Paragraph(direct) = &mut input.document.body.content[0] else {
        panic!("direct paragraph");
    };
    direct.properties = Some(CT_PPr {
        rpr: Some(CT_RPr {
            font_ascii: Some("Carlito".to_owned()),
            font_hansi: Some("Carlito".to_owned()),
            sz: Some(HalfPoint(36)),
            ..Default::default()
        }),
        ..Default::default()
    });

    let media = rdocx_layout::MediaRegistry::new(&input.images);
    let mut font_manager =
        oxml_layout::FontManager::new_deterministic().expect("deterministic metric fonts load");
    let mut numbering = rdocx_layout::style_resolver::NumberingState::new();
    let mut diagnostics = Vec::new();
    for index in 0..2 {
        let BodyContent::Paragraph(paragraph) = &input.document.body.content[index] else {
            panic!("metric paragraph");
        };
        let block = rdocx_layout::engine::layout_paragraph(
            paragraph,
            400.0,
            &input.styles,
            &input,
            &media,
            &mut font_manager,
            &mut numbering,
            &mut diagnostics,
        )
        .expect("metric paragraph lays out");
        let oxml_layout::LineItem::Text(segment) = &block.lines[0].items[0] else {
            panic!("empty paragraph carrier");
        };
        assert_eq!(segment.width, 0.0);
        let resolved = font_manager
            .metrics(segment.font_id, segment.font_size)
            .expect("carrier metrics resolve");
        assert_eq!(segment.ascent, resolved.ascent);
        assert_eq!(segment.descent, resolved.descent);
        if index == 0 {
            assert_eq!(block.lines[0].ascent, resolved.ascent);
            assert_eq!(block.lines[0].descent, resolved.descent);
        }
    }

    let result = rdocx_layout::layout_document_deterministic_with_provenance(&input)
        .expect("empty metric paragraphs lay out");
    let empty_runs = result
        .layout
        .pages
        .iter()
        .flat_map(|page| compatibility_page_elements(&page.elements))
        .filter_map(|element| match element {
            oxml_layout::PositionedElement::Text(run) if run.text.is_empty() => Some(run),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(empty_runs.len(), 2);
    let metric = |children: &[usize]| {
        let run = empty_runs
            .iter()
            .find(|run| {
                run.source.is_some_and(|source| {
                    matches!(
                        result.source_node(source.node),
                        Some(rdocx_layout::WordSourcePath {
                            story: rdocx_layout::WordStory::Document,
                            children: actual,
                        }) if actual == children
                    )
                })
            })
            .expect("empty paragraph run");
        let family = result
            .layout
            .fonts
            .iter()
            .find(|font| font.id == run.font_id)
            .expect("caret font retained")
            .family
            .as_str();
        (family, run.font_size)
    };
    assert_eq!(metric(&[0]), ("Carlito", 18.0));
    assert_eq!(metric(&[1]), ("Caladea", 14.0));
}

#[test]
fn empty_segment_is_backend_invisible_and_layout_compatible() {
    let mut input = empty_story_layout_input();
    input.document = CT_Document::from_xml(
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:rPr><w:u w:val="double"/><w:highlight w:val="yellow"/><w:strike/></w:rPr></w:pPr></w:p><w:p><w:r><w:t>visible</w:t></w:r></w:p></w:body></w:document>"#,
    )
    .expect("compatibility document parses");
    let ordinary = rdocx_layout::layout_document_deterministic(&input).expect("ordinary layout");
    let mut attributed = rdocx_layout::layout_document_deterministic_with_provenance(&input)
        .expect("attributed layout")
        .into_layout_result();
    assert!(
        attributed
            .pages
            .iter()
            .flat_map(|page| compatibility_page_elements(&page.elements))
            .any(
                |element| matches!(element, oxml_layout::PositionedElement::Text(run)
            if run.text == "visible" && !run.glyph_ids.is_empty())
            )
    );
    for page in &mut attributed.pages {
        clear_compatibility_sources(&mut std::sync::Arc::make_mut(page).elements);
    }
    assert_eq!(format!("{ordinary:?}"), format!("{attributed:?}"));

    let mut without_empty = attributed.clone();
    for page in &mut without_empty.pages {
        remove_compatibility_empty_text(&mut std::sync::Arc::make_mut(page).elements);
    }
    assert_eq!(
        oxml_pdf::render_to_pdf(&attributed),
        oxml_pdf::render_to_pdf(&without_empty)
    );
    assert_eq!(
        oxml_pdf::render_page_to_png(&attributed, 0, 96.0),
        oxml_pdf::render_page_to_png(&without_empty, 0, 96.0)
    );
}

#[test]
fn empty_form_paragraphs_use_mark_metrics_and_new_runs_inherit_them() {
    let body = r#"<w:p><w:pPr><w:rPr><w:b/><w:sz w:val="14"/></w:rPr></w:pPr></w:p>"#;
    let mut document = document_with_field_parts(&wrap_word_body(body), None, None);
    let layout = document.layout().expect("empty mark layout");
    let carrier = layout
        .layout
        .pages
        .iter()
        .flat_map(|page| compatibility_page_elements(&page.elements))
        .find_map(|element| match element {
            oxml_layout::PositionedElement::Text(run) if run.text.is_empty() => Some(run),
            _ => None,
        })
        .expect("zero-width paragraph-mark carrier");
    assert_eq!(carrier.font_size, 7.0);
    assert!(carrier.glyph_ids.is_empty());
    assert!(carrier.advances.is_empty());

    {
        let mut paragraph = document.paragraph_mut(0).unwrap();
        paragraph.add_run_inheriting_mark("typed");
        let run = paragraph.run(0).unwrap();
        assert_eq!(run.size(), Some(7.0));
        assert!(run.is_bold());
    }

    let xml = document_xml(&mut document);
    assert!(xml.contains("<w:b"));
    assert!(xml.contains(r#"<w:sz w:val="14""#));
    assert!(xml.contains("typed"));
}

fn dense_form_document() -> Document {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
 <w:body>
  <w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t>Dense form receipt</w:t></w:r></w:p>
  <w:tbl>
   <w:tblPr>
    <w:tblStyle w:val="DenseForm"/><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/>
    <w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>
   </w:tblPr>
   <w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
   <w:tr>
    <w:trPr><w:trHeight w:val="420" w:hRule="exact"/></w:trPr>
    <w:tc>
     <w:tcPr><w:tcW w:w="4680" w:type="dxa"/><w:vMerge w:val="restart"/><w:tcBorders><w:top w:val="nil"/><w:right w:val="nil"/></w:tcBorders></w:tcPr>
     <w:p><w:r><w:t>Patient details</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
     <w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>
     <w:p><w:r><w:drawing><wp:anchor behindDoc="1">
      <wp:positionH relativeFrom="column"><wp:posOffset>274320</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="548640" cy="274320"/><wp:wrapNone/><wp:docPr id="41" name="Behind stamp" descr="behind cell stamp"/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:solidFill><a:srgbClr val="DDEBFF"/></a:solidFill><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp></a:graphicData></a:graphic>
     </wp:anchor></w:drawing></w:r><w:r><w:t>Account 0042</w:t></w:r></w:p>
    </w:tc>
   </w:tr>
   <w:tr>
    <w:trPr><w:trHeight w:val="360" w:hRule="exact"/></w:trPr>
    <w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>
    <w:tc>
     <w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>
     <w:tbl>
      <w:tblPr><w:tblW w:w="4400" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="6" w:color="336699"/><w:left w:val="single" w:sz="6" w:color="336699"/><w:bottom w:val="single" w:sz="6" w:color="336699"/><w:right w:val="single" w:sz="6" w:color="336699"/><w:insideV w:val="single" w:sz="4" w:color="7799BB"/></w:tblBorders></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2200"/><w:gridCol w:w="2200"/></w:tblGrid>
      <w:tr><w:trPr><w:trHeight w:val="240" w:hRule="atLeast"/></w:trPr>
       <w:tc><w:p><w:r><w:t>Code</w:t></w:r></w:p></w:tc>
       <w:tc><w:p><w:r><w:t>A17</w:t></w:r></w:p></w:tc>
      </w:tr>
     </w:tbl>
    </w:tc>
   </w:tr>
   <w:tr>
    <w:trPr><w:trHeight w:val="360" w:hRule="atLeast"/></w:trPr>
    <w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/><w:tcBorders><w:top w:val="nil"/></w:tcBorders></w:tcPr><w:p><w:pPr><w:rPr><w:b/><w:sz w:val="14"/></w:rPr></w:pPr></w:p></w:tc>
    <w:tc>
     <w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>
     <w:p><w:r><w:drawing><wp:anchor behindDoc="0">
      <wp:positionH relativeFrom="column"><wp:posOffset>1371600</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="548640" cy="274320"/><wp:wrapNone/><wp:docPr id="42" name="Front stamp" descr="foreground cell stamp"/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:solidFill><a:srgbClr val="FFD7D7"/></a:solidFill><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp></a:graphicData></a:graphic>
     </wp:anchor></w:drawing></w:r><w:r><w:t>Total due 18.40</w:t></w:r></w:p>
    </w:tc>
   </w:tr>
  </w:tbl>
  <w:p><w:pPr><w:spacing w:before="120"/></w:pPr><w:r><w:t>Reviewed form footer</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1440" w:bottom="1080" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
 </w:body>
</w:document>"#;
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Carlito" w:hAnsi="Carlito"/><w:sz w:val="18"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>
 <w:style w:type="table" w:styleId="DenseBase"><w:name w:val="Dense Base"/><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8" w:color="A22B2B"/><w:left w:val="single" w:sz="8" w:color="A22B2B"/><w:bottom w:val="single" w:sz="8" w:color="A22B2B"/><w:right w:val="single" w:sz="8" w:color="A22B2B"/><w:insideH w:val="single" w:sz="4" w:color="A22B2B"/><w:insideV w:val="single" w:sz="4" w:color="A22B2B"/></w:tblBorders></w:tblPr></w:style>
 <w:style w:type="table" w:styleId="DenseForm"><w:name w:val="Dense Form"/><w:basedOn w:val="DenseBase"/><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:tblStylePr w:type="firstRow"><w:pPr><w:spacing w:after="40"/></w:pPr><w:tcPr><w:shd w:val="clear" w:fill="E8F1F8"/></w:tcPr></w:tblStylePr></w:style>
</w:styles>"#;

    let mut seed = Document::new();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(
        seed.to_bytes().expect("seed package"),
    ))
    .expect("seed opens");
    package.set_part("/word/document.xml", document_xml.as_bytes().to_vec());
    package.set_part("/word/styles.xml", styles_xml.as_bytes().to_vec());
    let mut output = std::io::Cursor::new(Vec::new());
    package.write_to(&mut output).expect("dense form package");
    Document::from_bytes(output.get_ref()).expect("dense form reopens")
}

fn decode_rgba_png(png: &[u8]) -> (u32, u32, Vec<u8>) {
    assert!(png.starts_with(b"\x89PNG\r\n\x1a\n"));
    let mut cursor = 8;
    let mut width = 0;
    let mut height = 0;
    let mut compressed = Vec::new();
    while cursor + 12 <= png.len() {
        let length = u32::from_be_bytes(png[cursor..cursor + 4].try_into().unwrap()) as usize;
        let kind = &png[cursor + 4..cursor + 8];
        let data = &png[cursor + 8..cursor + 8 + length];
        match kind {
            b"IHDR" => {
                width = u32::from_be_bytes(data[0..4].try_into().unwrap());
                height = u32::from_be_bytes(data[4..8].try_into().unwrap());
                assert_eq!(&data[8..13], &[8, 6, 0, 0, 0]);
            }
            b"IDAT" => compressed.extend_from_slice(data),
            b"IEND" => break,
            _ => {}
        }
        cursor += 12 + length;
    }
    let filtered = miniz_oxide::inflate::decompress_to_vec_zlib(&compressed).unwrap();
    let stride = width as usize * 4;
    assert_eq!(filtered.len(), (stride + 1) * height as usize);
    let mut rgba = vec![0; stride * height as usize];
    for row in 0..height as usize {
        let source = &filtered[row * (stride + 1)..(row + 1) * (stride + 1)];
        let (filter, source) = (source[0], &source[1..]);
        for index in 0..stride {
            let left = if index >= 4 {
                rgba[row * stride + index - 4]
            } else {
                0
            };
            let up = if row > 0 {
                rgba[(row - 1) * stride + index]
            } else {
                0
            };
            let upper_left = if row > 0 && index >= 4 {
                rgba[(row - 1) * stride + index - 4]
            } else {
                0
            };
            let predictor = match filter {
                0 => 0,
                1 => left,
                2 => up,
                3 => ((left as u16 + up as u16) / 2) as u8,
                4 => {
                    let estimate = left as i32 + up as i32 - upper_left as i32;
                    let distances = [
                        (estimate - left as i32).unsigned_abs(),
                        (estimate - up as i32).unsigned_abs(),
                        (estimate - upper_left as i32).unsigned_abs(),
                    ];
                    if distances[0] <= distances[1] && distances[0] <= distances[2] {
                        left
                    } else if distances[1] <= distances[2] {
                        up
                    } else {
                        upper_left
                    }
                }
                other => panic!("unsupported PNG filter {other}"),
            };
            rgba[row * stride + index] = source[index].wrapping_add(predictor);
        }
    }
    (width, height, rgba)
}

#[test]
fn dense_form_matches_reviewed_one_page_geometry() {
    let document = dense_form_document();
    let layout = document
        .layout_with_fonts_and_bundled_fallback(&[])
        .expect("deterministic dense-form layout");
    assert_eq!(layout.layout.pages.len(), 1, "{WORD_DENSE_FORM_ORACLE}");
    let page = &layout.layout.pages[0];
    assert_eq!((page.width, page.height), (612.0, 792.0));

    let positioned = compatibility_page_elements(&page.elements);
    let visible_text = positioned
        .iter()
        .filter_map(|element| match element {
            oxml_layout::PositionedElement::Text(run) if !run.text.is_empty() => {
                Some(run.text.as_str())
            }
            _ => None,
        })
        .collect::<String>();
    for expected in [
        "Dense form receipt",
        "Patient details",
        "Account 0042",
        "Code",
        "A17",
        "Total due 18.40",
        "Reviewed form footer",
    ] {
        assert!(
            visible_text.contains(expected),
            "{expected}: {visible_text}"
        );
    }
    let empty_mark = positioned
        .iter()
        .find_map(|element| match element {
            oxml_layout::PositionedElement::Text(run) if run.text.is_empty() => Some(run),
            _ => None,
        })
        .expect("empty 7pt form carrier");
    assert_eq!(empty_mark.font_size, 7.0);
    assert!(empty_mark.glyph_ids.is_empty());

    let table_lines = positioned
        .iter()
        .filter_map(|element| match element {
            oxml_layout::PositionedElement::Line { start, end, .. }
                if start.x >= 71.0 && end.x <= 541.0 && start.y >= 60.0 && end.y <= 180.0 =>
            {
                Some((*start, *end))
            }
            _ => None,
        })
        .collect::<Vec<_>>();
    let has_line = |x1, y1, x2, y2| {
        table_lines
            .iter()
            .any(|(start, end)| (start.x, start.y, end.x, end.y) == (x1, y1, x2, y2))
    };
    assert_eq!(table_lines.len(), 26, "table geometry: {table_lines:?}");
    assert!(has_line(72.0, 70.0, 306.0, 70.0));
    assert!(has_line(72.0, 70.0, 72.0, 109.0));
    assert!(!has_line(72.0, 91.0, 306.0, 91.0));
    assert!(has_line(72.0, 109.0, 306.0, 109.0));
    assert!(has_line(311.4, 91.0, 421.4, 91.0));
    assert!(has_line(421.4, 103.0, 531.4, 103.0));
    assert!(has_line(306.0, 127.0, 540.0, 127.0));

    let pdf = document.to_pdf_deterministic().expect("dense form PDF");
    assert!(pdf.starts_with(b"%PDF-"));
    assert_eq!(pdf, document.to_pdf_deterministic().unwrap());
    let png = document
        .render_page_to_png_deterministic(0, 96.0)
        .expect("dense form raster")
        .expect("page zero");
    assert_eq!(
        png,
        document
            .render_page_to_png_deterministic(0, 96.0)
            .unwrap()
            .unwrap()
    );
    let (width, height, rgba) = decode_rgba_png(&png);
    assert_eq!((width, height), (816, 1056));
    let checksum = rgba.iter().fold(0xcbf29ce484222325u64, |hash, byte| {
        (hash ^ u64::from(*byte)).wrapping_mul(0x100000001b3)
    });
    let non_white_pixels = rgba
        .chunks_exact(4)
        .filter(|pixel| *pixel != [255, 255, 255, 255])
        .count();
    let behind_pixels = rgba
        .chunks_exact(4)
        .filter(|pixel| *pixel == [221, 235, 255, 255])
        .count();
    let foreground_pixels = rgba
        .chunks_exact(4)
        .filter(|pixel| *pixel == [255, 215, 215, 255])
        .count();
    assert_eq!(checksum, 0x2319_bcbe_502e_4fe8);
    assert_eq!(non_white_pixels, 32_221);
    assert_eq!(
        behind_pixels, 0,
        "page-behind stamp is covered by cell shading"
    );
    assert_eq!(foreground_pixels, 1_682);
}

#[test]
fn document_facing_aliases_share_one_caller_font() {
    let bundled_bytes = include_bytes!("../../oxml-layout/fonts/Caladea-Regular.ttf").as_slice();
    let mut caller_bytes = bundled_bytes.to_vec();
    caller_bytes.push(0);
    let mut labelled_document = Document::new();
    labelled_document
        .add_paragraph("")
        .add_run("label-derived alias")
        .font("Document Serif");
    let labelled = labelled_document
        .layout_with_fonts(&[("Document Serif", caller_bytes.as_slice())])
        .expect("caller font label resolves through the strict facade");
    assert!(
        labelled.layout.fonts.iter().any(|font| {
            font.family == "Caladea" && font.data.as_ref() == caller_bytes.as_slice()
        })
    );

    let mut document = Document::new();
    for (family, text) in [
        ("Document Serif", "first alias"),
        ("Legacy Serif", "second alias"),
    ] {
        document.add_paragraph("").add_run(text).font(family);
    }

    let result = document
        .layout_with_fonts_aliases_and_bundled_fallback(
            &[("Caladea", caller_bytes.as_slice())],
            &[("Document Serif", "Caladea"), ("Legacy Serif", "Caladea")],
        )
        .expect("document-facing aliases resolve");

    let mut alias_fonts = Vec::new();
    for page in &result.layout.pages {
        oxml_layout::walk(&page.elements, &mut |element, _| {
            let oxml_layout::PositionedElement::Text(run) = element else {
                return;
            };
            let font = result
                .layout
                .fonts
                .iter()
                .find(|font| font.id == run.font_id)
                .expect("alias run font exists");
            assert_eq!(font.family, "Caladea");
            assert_eq!(font.data.as_ref(), caller_bytes.as_slice());
            assert_ne!(font.data.as_ref(), bundled_bytes);
            assert!(
                run.source
                    .is_some_and(|span| result.source_node(span.node).is_some()),
                "alias run retains provenance"
            );
            alias_fonts.push(Arc::clone(&font.data));
        });
    }
    assert!(alias_fonts.len() >= 2);
    assert!(
        alias_fonts
            .iter()
            .skip(1)
            .all(|font| Arc::ptr_eq(&alias_fonts[0], font))
    );
    assert!(result.layout.diagnostics.is_empty());
}

#[test]
fn five_large_caller_fonts_and_forty_aliases_keep_warm_and_fresh_layouts_equal() {
    const TOTAL_FONT_BYTES: usize = 22 * 1024 * 1024;
    let bundled = oxml_layout::bundled_fonts::bundled_font_data();
    let generated_fonts = [0, 4, 8, 12, 16]
        .into_iter()
        .enumerate()
        .map(|(index, bundled_index)| {
            let (family, source) = bundled[bundled_index];
            let mut data = source.to_vec();
            let target = TOTAL_FONT_BYTES / 5 + usize::from(index < TOTAL_FONT_BYTES % 5);
            data.resize(
                target,
                u8::try_from(index).expect("five font indices fit in u8"),
            );
            (family.to_owned(), data)
        })
        .collect::<Vec<_>>();
    assert_eq!(
        generated_fonts
            .iter()
            .map(|(_, data)| data.len())
            .sum::<usize>(),
        TOTAL_FONT_BYTES
    );
    let font_files = generated_fonts
        .iter()
        .map(|(family, data)| (family.as_str(), data.as_slice()))
        .collect::<Vec<_>>();
    let owned_aliases = (0..40)
        .map(|index| {
            (
                format!("Editor Family {index}"),
                generated_fonts[index % generated_fonts.len()].0.clone(),
            )
        })
        .collect::<Vec<_>>();
    let aliases = owned_aliases
        .iter()
        .map(|(requested, target)| (requested.as_str(), target.as_str()))
        .collect::<Vec<_>>();

    let make_document = |changed: bool| {
        let mut document = Document::new();
        for index in 0..40 {
            let text = if changed && index == 20 {
                format!("paragraph {index:03} stable line changed")
            } else {
                format!("paragraph {index:03} stable line")
            };
            document.add_paragraph("").page_break_before(index > 0);
            document
                .paragraph_mut(index)
                .expect("generated paragraph")
                .add_run(&text)
                .font(&owned_aliases[index % owned_aliases.len()].0);
        }
        document
    };

    let mut warm_document = make_document(false);
    let initial = warm_document
        .layout_with_fonts_aliases_and_bundled_fallback(&font_files, &aliases)
        .expect("prime large caller-font layout");
    warm_document
        .paragraph_mut(20)
        .expect("middle paragraph")
        .run_mut(0)
        .expect("middle paragraph text run")
        .set_text("paragraph 020 stable line changed");
    let fresh_document = make_document(true);
    let warm = warm_document
        .layout_with_fonts_aliases_and_bundled_fallback(&font_files, &aliases)
        .expect("warm large caller-font layout");
    let fresh = fresh_document
        .layout_with_fonts_aliases_and_bundled_fallback(&font_files, &aliases)
        .expect("fresh large caller-font layout");

    let retained_pages = warm
        .layout
        .pages
        .iter()
        .zip(&initial.layout.pages)
        .filter(|(current, previous)| Arc::ptr_eq(current, previous))
        .count();
    assert!(
        retained_pages >= warm.layout.pages.len().saturating_sub(2),
        "the bounded edit retained {retained_pages} of {} pages",
        warm.layout.pages.len()
    );
    assert_eq!(warm.revision_view, fresh.revision_view);
    assert_eq!(warm.layout.pages.len(), fresh.layout.pages.len());
    for (warm_page, fresh_page) in warm.layout.pages.iter().zip(&fresh.layout.pages) {
        assert_eq!(warm_page.page_number, fresh_page.page_number);
        assert_eq!(warm_page.width, fresh_page.width);
        assert_eq!(warm_page.height, fresh_page.height);
        assert_eq!(warm_page.elements, fresh_page.elements);
        assert_eq!(warm_page.background, fresh_page.background);
    }
    assert_eq!(warm.layout.fonts.len(), fresh.layout.fonts.len());
    for (warm_font, fresh_font) in warm.layout.fonts.iter().zip(&fresh.layout.fonts) {
        assert_eq!(warm_font.id, fresh_font.id);
        assert_eq!(warm_font.family, fresh_font.family);
        assert_eq!(warm_font.data, fresh_font.data);
        assert_eq!(warm_font.face_index, fresh_font.face_index);
        assert_eq!(warm_font.bold, fresh_font.bold);
        assert_eq!(warm_font.italic, fresh_font.italic);
    }
    assert_eq!(warm.layout.diagnostics, fresh.layout.diagnostics);
    assert_eq!(
        format!("{:?}", warm.layout.outlines),
        format!("{:?}", fresh.layout.outlines)
    );
    for (warm_page, fresh_page) in warm.layout.pages.iter().zip(&fresh.layout.pages) {
        let mut warm_sources = Vec::new();
        oxml_layout::walk(&warm_page.elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Text(run) = element
                && let Some(source) = run.source
            {
                warm_sources.push(warm.source_node(source.node).cloned());
            }
        });
        let mut fresh_sources = Vec::new();
        oxml_layout::walk(&fresh_page.elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Text(run) = element
                && let Some(source) = run.source
            {
                fresh_sources.push(fresh.source_node(source.node).cloned());
            }
        });
        assert_eq!(warm_sources, fresh_sources);
    }
    assert_eq!(
        oxml_pdf::render_to_pdf(&warm.layout),
        oxml_pdf::render_to_pdf(&fresh.layout)
    );
}

fn redaction_fixture() -> Document {
    let mut seed = Document::new();
    seed.set_title("secret core title");
    let bytes = seed.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let relationship_namespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let relationships = package.get_or_create_part_rels("/word/document.xml");
    let header_id = relationships.add(
        oxml_opc::relationship::rel_types::HEADER,
        "stories/header1.xml",
    );
    let footer_id = relationships.add(
        oxml_opc::relationship::rel_types::FOOTER,
        "stories/footer1.xml",
    );
    relationships.add(
        oxml_opc::relationship::rel_types::FOOTNOTES,
        "stories/footnotes1.xml",
    );
    relationships.add(
        oxml_opc::relationship::rel_types::ENDNOTES,
        "stories/endnotes1.xml",
    );
    relationships.add(
        oxml_opc::relationship::rel_types::COMMENTS,
        "stories/comments1.xml",
    );

    package.set_part(
        "/word/document.xml",
        format!(
            r#"<w:document xmlns:w="{word_namespace}" xmlns:r="{relationship_namespace}" xmlns:p="urn:producer"><w:body><p:keep>producer bytes</p:keep><w:p><w:r><w:t>body secret</w:t></w:r><w:ins w:id="1" w:author="secret author"><w:r><w:t>inserted se</w:t></w:r><w:r><w:t>cret</w:t></w:r></w:ins><w:del w:id="2" w:author="secret author"><w:r><w:delText>deleted secret</w:delText></w:r></w:del></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>table secret</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>control secret</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/><w:footerReference w:type="default" r:id="{footer_id}"/></w:sectPr></w:body></w:document>"#
        )
        .into_bytes(),
    );
    for (part, root, content_type) in [
        (
            "/word/stories/header1.xml",
            "hdr",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        ),
        (
            "/word/stories/footer1.xml",
            "ftr",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        ),
    ] {
        package.set_part(
            part,
            format!(
                r#"<w:{root} xmlns:w="{word_namespace}"><w:p><w:r><w:t>{root} secret</w:t></w:r></w:p></w:{root}>"#
            )
            .into_bytes(),
        );
        package.content_types.add_override(part, content_type);
    }
    for (part, root, item, content_type) in [
        (
            "/word/stories/footnotes1.xml",
            "footnotes",
            "footnote",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        ),
        (
            "/word/stories/endnotes1.xml",
            "endnotes",
            "endnote",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
        ),
    ] {
        package.set_part(
            part,
            format!(
                r#"<w:{root} xmlns:w="{word_namespace}"><w:{item} w:id="2"><w:p><w:r><w:t>{item} secret</w:t></w:r></w:p></w:{item}></w:{root}>"#
            )
            .into_bytes(),
        );
        package.content_types.add_override(part, content_type);
    }
    package.set_part(
        "/word/stories/comments1.xml",
        format!(
            r#"<w:comments xmlns:w="{word_namespace}"><w:comment w:id="0" w:author="secret author" w:initials="secret"><w:p><w:r><w:t>comment secret</w:t></w:r></w:p></w:comment></w:comments>"#
        )
        .into_bytes(),
    );
    package.content_types.add_override(
        "/word/stories/comments1.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    );
    package.set_part(
        "/metadata/custom.xml",
        br#"<p:Properties xmlns:p="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:v="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><p:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Client"><v:lpwstr>custom secret</v:lpwstr></p:property></p:Properties>"#.to_vec(),
    );
    package.content_types.add_override(
        "/metadata/custom.xml",
        oxml_opc::content_types::CUSTOM_PROPERTIES,
    );
    package.package_rels.add(
        oxml_opc::relationship::rel_types::CUSTOM_PROPERTIES,
        "metadata/custom.xml",
    );
    package.set_part("/custom/unrelated.bin", b"producer bytes".to_vec());
    package
        .content_types
        .add_override("/custom/unrelated.bin", "application/octet-stream");
    let mut output = std::io::Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();
    Document::from_bytes(output.get_ref()).unwrap()
}

fn assert_raw_package_has_no_selector(bytes: &[u8], selector: &str) {
    fn absent(bytes: &[u8], selector: &str) {
        let utf16le = selector
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>();
        assert!(
            !bytes
                .windows(selector.len())
                .any(|part| part == selector.as_bytes())
        );
        assert!(!bytes.windows(utf16le.len()).any(|part| part == utf16le));
    }

    let package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).expect("outer package");
    absent(
        &package.content_types.to_xml().expect("content types XML"),
        selector,
    );
    absent(
        &package
            .package_rels
            .to_xml()
            .expect("package relationships XML"),
        selector,
    );
    for relationships in package.part_rels.values() {
        absent(
            &relationships.to_xml().expect("part relationships XML"),
            selector,
        );
    }
    for part in package.parts.values() {
        absent(part, selector);
        if part.starts_with(b"PK\x03\x04") {
            let nested = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(part))
                .expect("nested package");
            absent(
                &nested
                    .content_types
                    .to_xml()
                    .expect("nested content types XML"),
                selector,
            );
            absent(
                &nested
                    .package_rels
                    .to_xml()
                    .expect("nested package relationships XML"),
                selector,
            );
            for relationships in nested.part_rels.values() {
                absent(
                    &relationships
                        .to_xml()
                        .expect("nested part relationships XML"),
                    selector,
                );
            }
            for nested_part in nested.parts.values() {
                absent(nested_part, selector);
            }
        }
    }
}

#[test]
fn redaction_removes_body_comments_revisions_and_metadata_traces() {
    let mut document = redaction_fixture();
    let report = document
        .redact_text("secret")
        .expect("redact all Word stories");
    assert!(report.word_text >= 13, "{report:?}");
    assert_eq!(report.metadata, 2);
    let bytes = document.to_bytes().unwrap();
    assert_raw_package_has_no_selector(&bytes, "secret");
    let reopened = Document::from_bytes(&bytes).unwrap();
    assert!(!reopened.text().contains("secret"));
    assert_eq!(reopened.title(), Some(" core title"));
}

#[test]
fn redaction_removes_chart_cache_and_embedded_workbook_traces() {
    let mut document = Document::new();
    document
        .add_chart(
            ChartKind::Bar,
            Length::inches(5.0),
            Length::inches(3.0),
            &ChartData {
                categories: vec!["secret north".to_owned(), "public".to_owned()],
                series: vec![("secret revenue".to_owned(), vec![12.5, 19.0])],
                number_format: None,
            },
        )
        .unwrap();
    let report = document.redact_text("secret").expect("redact chart source");
    assert!(report.chart_caches >= 2, "{report:?}");
    assert!(report.embedded_workbooks >= 2, "{report:?}");
    assert_raw_package_has_no_selector(&document.to_bytes().unwrap(), "secret");

    let mut numeric_document = Document::new();
    numeric_document
        .add_chart(
            ChartKind::Bar,
            Length::inches(5.0),
            Length::inches(3.0),
            &ChartData {
                categories: vec!["north".to_owned(), "south".to_owned()],
                series: vec![("revenue".to_owned(), vec![12.5, 19.0])],
                number_format: None,
            },
        )
        .unwrap();
    let numeric_report = numeric_document
        .redact_text("12.5")
        .expect("redact numeric chart source");
    assert!(numeric_report.chart_caches >= 1, "{numeric_report:?}");
    assert!(numeric_report.embedded_workbooks >= 1, "{numeric_report:?}");
    assert_raw_package_has_no_selector(&numeric_document.to_bytes().unwrap(), "12.5");
}

fn assert_redaction_failure_preserves_document(document: &mut Document, selector: &str) {
    let before_bytes = document.to_bytes().expect("atomic snapshot serializes");
    let before_text = document.text();
    let before_title = document.title().map(str::to_owned);
    let before_paragraphs = document
        .paragraphs()
        .iter()
        .map(|paragraph| paragraph.text())
        .collect::<Vec<_>>();
    let before_layout = document.layout().expect("atomic layout cache primes");

    assert!(document.redact_text(selector).is_err());
    assert_eq!(document.to_bytes().unwrap(), before_bytes);
    assert_eq!(document.text(), before_text);
    assert_eq!(document.title(), before_title.as_deref());
    assert_eq!(
        document
            .paragraphs()
            .iter()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>(),
        before_paragraphs
    );
    assert!(std::sync::Arc::ptr_eq(
        &before_layout,
        &document.layout().expect("atomic layout cache remains")
    ));
}

#[test]
fn redaction_failure_is_atomic() {
    let mut empty_selector = redaction_fixture();
    assert_redaction_failure_preserves_document(&mut empty_selector, "");

    let mut document = redaction_fixture();
    let before = document.to_bytes().unwrap();
    let mut package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&before)).unwrap();
    package.set_part("/word/stories/header1.xml", b"<secret".to_vec());
    let mut malformed = std::io::Cursor::new(Vec::new());
    package.write_to(&mut malformed).unwrap();
    let mut document = Document::from_bytes(malformed.get_ref()).unwrap();
    assert_redaction_failure_preserves_document(&mut document, "secret");

    let mut utf16_document = redaction_fixture();
    let mut utf16_package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(utf16_document.to_bytes().unwrap()))
            .unwrap();
    let utf16_secret = "secret"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    utf16_package.set_part("/custom/unrelated.bin", utf16_secret);
    let mut utf16_bytes = std::io::Cursor::new(Vec::new());
    utf16_package.write_to(&mut utf16_bytes).unwrap();
    let mut utf16_document = Document::from_bytes(utf16_bytes.get_ref()).unwrap();
    assert_redaction_failure_preserves_document(&mut utf16_document, "secret");

    let mut chart_document = Document::new();
    chart_document
        .add_chart(
            ChartKind::Bar,
            Length::inches(5.0),
            Length::inches(3.0),
            &ChartData {
                categories: vec!["secret".to_owned()],
                series: vec![("public".to_owned(), vec![1.0])],
                number_format: None,
            },
        )
        .unwrap();
    let mut chart_package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(chart_document.to_bytes().unwrap()))
            .unwrap();
    let chart_part = chart_package
        .get_part_rels("/word/document.xml")
        .and_then(|relationships| {
            relationships.get_by_type(oxml_opc::relationship::rel_types::CHART)
        })
        .map(|relationship| {
            oxml_opc::OpcPackage::resolve_rel_target("/word/document.xml", &relationship.target)
        })
        .unwrap();
    let workbook_part = chart_package
        .get_part_rels(&chart_part)
        .and_then(|relationships| {
            relationships.get_by_type(oxml_opc::relationship::rel_types::PACKAGE)
        })
        .map(|relationship| {
            oxml_opc::OpcPackage::resolve_rel_target(&chart_part, &relationship.target)
        })
        .unwrap();

    let mut bounded_package = chart_package.clone();
    let mut oversized_workbook = oxml_opc::OpcPackage::new();
    oversized_workbook
        .content_types
        .add_default("bin", "application/octet-stream");
    for index in 0..1_024 {
        oversized_workbook.set_part(&format!("/payload/{index}.bin"), Vec::new());
    }
    let mut oversized_bytes = std::io::Cursor::new(Vec::new());
    oversized_workbook.write_to(&mut oversized_bytes).unwrap();
    bounded_package.set_part(&workbook_part, oversized_bytes.into_inner());
    let mut bounded_bytes = std::io::Cursor::new(Vec::new());
    bounded_package.write_to(&mut bounded_bytes).unwrap();
    let mut bounded_document = Document::from_bytes(bounded_bytes.get_ref()).unwrap();
    assert_redaction_failure_preserves_document(&mut bounded_document, "secret");

    let workbook_relationship = chart_package
        .get_or_create_part_rels(&chart_part)
        .items
        .iter_mut()
        .find(|relationship| relationship.rel_type == oxml_opc::relationship::rel_types::PACKAGE)
        .unwrap();
    workbook_relationship.target_mode = Some("External".to_owned());
    workbook_relationship.target = "https://example.invalid/secret.xlsx".to_owned();
    let mut external_bytes = std::io::Cursor::new(Vec::new());
    chart_package.write_to(&mut external_bytes).unwrap();
    let mut external_document = Document::from_bytes(external_bytes.get_ref()).unwrap();
    assert_redaction_failure_preserves_document(&mut external_document, "secret");
}

#[test]
fn redacted_package_preserves_unrelated_parts_and_relationships() {
    let mut document = redaction_fixture();
    let before = document.to_bytes().unwrap();
    let before = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(before)).unwrap();
    let before_content_types = before.content_types.to_xml().unwrap();
    let before_package_relationships = before.package_rels.to_xml().unwrap();
    let before_relationships = before
        .part_rels
        .iter()
        .map(|(source, relationships)| (source.clone(), relationships.to_xml().unwrap()))
        .collect::<std::collections::BTreeMap<_, _>>();
    let before_endnotes = before
        .get_part("/word/stories/endnotes1.xml")
        .unwrap()
        .to_vec();
    let before_custom = before.get_part("/metadata/custom.xml").unwrap().to_vec();
    document.redact_text("secret").unwrap();
    let after =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
            .unwrap();
    assert_eq!(
        after.get_part("/custom/unrelated.bin"),
        before.get_part("/custom/unrelated.bin")
    );
    assert_eq!(after.content_types.to_xml().unwrap(), before_content_types);
    assert_eq!(
        after.package_rels.to_xml().unwrap(),
        before_package_relationships
    );
    assert_eq!(
        after
            .part_rels
            .iter()
            .map(|(source, relationships)| (source.clone(), relationships.to_xml().unwrap()))
            .collect::<std::collections::BTreeMap<_, _>>(),
        before_relationships
    );
    assert_eq!(
        after.get_part("/word/stories/endnotes1.xml").unwrap(),
        String::from_utf8(before_endnotes)
            .unwrap()
            .replace("secret", "")
            .as_bytes()
    );
    assert_eq!(
        after.get_part("/metadata/custom.xml").unwrap(),
        String::from_utf8(before_custom)
            .unwrap()
            .replace("secret", "")
            .as_bytes()
    );
    let edited_parts = std::collections::HashSet::from([
        "/word/document.xml",
        "/word/stories/header1.xml",
        "/word/stories/footer1.xml",
        "/word/stories/footnotes1.xml",
        "/word/stories/endnotes1.xml",
        "/word/stories/comments1.xml",
        "/docProps/core.xml",
        "/metadata/custom.xml",
    ]);
    assert_eq!(
        after.parts.keys().collect::<std::collections::HashSet<_>>(),
        before.parts.keys().collect()
    );
    for (part_name, bytes) in &before.parts {
        if !edited_parts.contains(part_name.as_str()) {
            assert_eq!(
                after.get_part(part_name).unwrap(),
                bytes,
                "untouched part changed: {part_name}"
            );
        }
    }
    let document_xml = std::str::from_utf8(after.get_part("/word/document.xml").unwrap()).unwrap();
    let producer = document_xml
        .find("<p:keep>producer bytes</p:keep>")
        .unwrap();
    let paragraph = document_xml.find("<w:p>").unwrap();
    let table = document_xml.find("<w:tbl>").unwrap();
    let control = document_xml.find("<w:sdt>").unwrap();
    let section = document_xml.find("<w:sectPr>").unwrap();
    assert!(producer < paragraph && paragraph < table && table < control && control < section);
}

#[test]
fn raw_zip_scan_finds_no_redacted_value() {
    let mut document = Document::new();
    document.add_paragraph("secret secret");
    let report = document.redact_text("secret").unwrap();
    assert_eq!(report.total(), 2);
    assert_raw_package_has_no_selector(&document.to_bytes().unwrap(), "secret");
}

#[test]
fn svg_facade_options_share_the_existing_layout_paths_and_bounds_contract() {
    let mut document = Document::new();
    document.add_paragraph("SVG facade");
    let options = RenderOptions::default();

    assert_eq!(
        document.render_page_to_svg(0).unwrap(),
        document
            .render_page_to_svg_with_options(0, options)
            .unwrap()
    );
    assert_eq!(
        document.render_page_to_svg_deterministic(0).unwrap(),
        document
            .render_page_to_svg_deterministic_with_options(0, options)
            .unwrap()
    );
    assert!(document.render_page_to_svg(usize::MAX).unwrap().is_none());
    assert!(
        document
            .render_page_to_svg_deterministic(usize::MAX)
            .unwrap()
            .is_none()
    );
}

#[test]
fn automatic_hyphenation_authoring_round_trips_with_run_language() {
    let mut document = Document::new();
    document.set_auto_hyphenation(true).unwrap();
    document
        .add_paragraph("")
        .add_run("representation")
        .language("en-US");

    let bytes = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(&bytes)).unwrap();
    let settings = std::str::from_utf8(package.get_part("/word/settings.xml").unwrap()).unwrap();
    assert!(settings.contains("<w:autoHyphenation/>"));

    let reopened = Document::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened.paragraphs()[0].run(0).unwrap().language(),
        Some("en-US")
    );
}

#[test]
fn legacy_equation_editor_objects_remain_unmodelled_raw_xml() {
    let legacy =
        r#"<w:object><v:shape id="legacy"><o:OLEObject ProgID="Equation.3"/></v:shape></w:object>"#;
    let xml = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body><w:p>{legacy}</w:p><w:sectPr/></w:body></w:document>"#
    );
    let mut document = document_with_content_controls(&xml);
    let paragraph = document.paragraph(0).unwrap();
    assert_eq!(paragraph.equations().count(), 0);
    assert!(matches!(
        paragraph.items().next().unwrap(),
        rdocx::ParagraphItemRef::UnsupportedXml(raw) if raw == legacy.as_bytes()
    ));
    let bytes = document.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let output = package.get_part("/word/document.xml").unwrap();
    assert!(
        output
            .windows(legacy.len())
            .any(|window| window == legacy.as_bytes())
    );
}

#[test]
fn nested_source_built_records_generate_ordered_rich_content_without_stale_fields() {
    let mut seed = Document::new();
    let num_id = seed.add_list_definition(&[ListLevel::decimal()]);
    let body = format!(
        r#"
        <w:p><w:r><w:t>before</w:t></w:r></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD TableStart:Lines"><w:r><w:t>start</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr><w:fldSimple w:instr="MERGEFIELD ListItem"><w:r><w:t>stored list</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD Formatted \b &quot;[&quot; \f &quot;]&quot; \* Upper"><w:r><w:t>stored format</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD Logo"><w:r><w:t>stored logo</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD Insert"><w:r><w:t>stored fragment</w:t></w:r></w:fldSimple></w:p>
        <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>
          <w:tr><w:tc><w:p><w:fldSimple w:instr="MERGEFIELD TableStart:Rows"><w:r><w:t>row start</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr>
          <w:tr><w:tc><w:p><w:fldSimple w:instr="MERGEFIELD RowValue"><w:r><w:t>stored row</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr>
          <w:tr><w:tc><w:p><w:fldSimple w:instr="MERGEFIELD TableEnd:Rows"><w:r><w:t>row end</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr>
        </w:tbl>
        <w:p><w:fldSimple w:instr="MERGEFIELD TableEnd:Lines"><w:r><w:t>end</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:r><w:t>after</w:t></w:r></w:p>
        <w:sectPr/>
    "#
    );
    let mut package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
    package.set_part("/word/document.xml", wrap_word_body(&body).into_bytes());
    let mut template_bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut template_bytes).unwrap();
    let document = Document::from_bytes(template_bytes.get_ref()).unwrap();
    let mut fragment = Document::new();
    fragment.add_paragraph("fragment");
    let fragment = fragment.to_bytes().unwrap();
    let lines = [
        (
            "alpha",
            "list-a",
            "format-a",
            ["row-a1", "row-a2"].as_slice(),
        ),
        ("beta", "list-b", "format-b", ["row-b1"].as_slice()),
    ]
    .into_iter()
    .map(|(name, list, formatted, rows)| MailMergeRecord {
        values: BTreeMap::from([
            ("Name".to_owned(), MailMergeValue::Text(name.to_owned())),
            ("ListItem".to_owned(), MailMergeValue::Text(list.to_owned())),
            (
                "Formatted".to_owned(),
                MailMergeValue::Text(formatted.to_owned()),
            ),
        ]),
        regions: BTreeMap::from([(
            "Rows".to_owned(),
            rows.iter()
                .map(|value| MailMergeRecord {
                    values: BTreeMap::from([(
                        "RowValue".to_owned(),
                        MailMergeValue::Text((*value).to_owned()),
                    )]),
                    ..Default::default()
                })
                .collect(),
        )]),
    })
    .collect();
    let data = MailMergeData {
        records: vec![MailMergeRecord {
            values: BTreeMap::from([
                (
                    "Logo".to_owned(),
                    MailMergeValue::Image(MailMergeImage {
                        data: b"\x89PNG\r\n\x1a\nlogo".to_vec(),
                        filename: "logo.png".to_owned(),
                        width: Length::emu(11),
                        height: Length::emu(19),
                    }),
                ),
                ("Insert".to_owned(), MailMergeValue::Fragment(fragment)),
            ]),
            ..Default::default()
        }],
        sources: BTreeMap::from([("Lines".to_owned(), lines)]),
    };
    let mut formatter = |context: &rdocx::MailMergeFormatContext<'_>,
                         value: &mut MailMergeFormattedText| {
        if context.field_name == "Formatted" {
            value.text.push('!');
        }
        Ok(())
    };
    let mut output = document
        .mail_merge_rich(&data, Some(&mut formatter))
        .unwrap()
        .remove(0);
    let paragraph_text = output
        .paragraphs()
        .into_iter()
        .map(|paragraph| paragraph.text())
        .filter(|text| !text.is_empty())
        .collect::<Vec<_>>();
    assert_eq!(
        paragraph_text,
        [
            "before",
            "alpha",
            "list-a",
            "[FORMAT-A]!",
            "fragment",
            "beta",
            "list-b",
            "[FORMAT-B]!",
            "fragment",
            "after",
        ]
    );
    assert_eq!(output.table_count(), 2);
    let first_table = output.table(0).unwrap();
    assert_eq!(first_table.row_count(), 2);
    assert_eq!(first_table.cell(0, 0).unwrap().text(), "row-a1");
    assert_eq!(first_table.cell(1, 0).unwrap().text(), "row-a2");
    let second_table = output.table(1).unwrap();
    assert_eq!(second_table.row_count(), 1);
    assert_eq!(second_table.cell(0, 0).unwrap().text(), "row-b1");
    let xml = document_xml(&mut output);
    assert_eq!(
        xml.matches(r#"<wp:extent cx="11" cy="19"/>"#).count(),
        2,
        "{xml}"
    );
    assert_eq!(xml.matches(r#"<w:numId w:val=""#).count(), 2, "{xml}");
    assert!(!xml.contains("MERGEFIELD"), "{xml}");
    assert!(!xml.contains("TableStart:"), "{xml}");
    assert!(!xml.contains("TableEnd:"), "{xml}");
    assert!(!xml.contains("stored"), "{xml}");
}

#[test]
fn rich_merge_imports_images_and_fragments_without_relationship_or_identity_collisions() {
    let body = r#"
        <w:p><w:fldSimple w:instr="MERGEFIELD Logo"><w:r><w:t>logo</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD InsertA"><w:r><w:t>insert</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD InsertB"><w:r><w:t>insert</w:t></w:r></w:fldSimple></w:p>
        <w:sectPr/>
    "#;
    let mut document = document_with_content_controls(&wrap_word_body(body));
    document.add_style(StyleBuilder::paragraph("Collision", "Destination style"));
    let mut fragment = Document::new();
    fragment.add_style(StyleBuilder::paragraph("Collision", "Fragment style"));
    fragment.add_paragraph("fragment text").style("Collision");
    let fragment_num_id = fragment.add_list_definition(&[ListLevel::decimal()]);
    fragment
        .add_paragraph("fragment list")
        .set_numbering(fragment_num_id, 0);
    fragment.add_picture(
        b"\x89PNG\r\n\x1a\n",
        "fragment.png",
        Length::emu(31),
        Length::emu(47),
    );
    let mut fragment_package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(fragment.to_bytes().unwrap()))
            .unwrap();
    let fragment_image = fragment_package
        .get_part_rels("/word/document.xml")
        .unwrap()
        .get_by_type(oxml_opc::relationship::rel_types::IMAGE)
        .unwrap()
        .clone();
    let fragment_image_part =
        oxml_opc::OpcPackage::resolve_rel_target("/word/document.xml", &fragment_image.target);
    fragment_package.set_part(
        "/word/media/fragment-payload.bin",
        b"fragment payload".to_vec(),
    );
    fragment_package.content_types.add_override(
        "/word/media/fragment-payload.bin",
        "application/octet-stream",
    );
    fragment_package
        .get_or_create_part_rels(&fragment_image_part)
        .add("urn:rdocx:test:fragment-payload", "fragment-payload.bin");
    let fragment_xml = String::from_utf8(
        fragment_package
            .get_part("/word/document.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap()
    .replace(
        "<w:body>",
        r#"<w:body><w:p><w:bookmarkStart w:id="7" w:name="fragmentMark"/><w:hyperlink w:anchor="fragmentMark"><w:r><w:t>linked</w:t></w:r></w:hyperlink><w:bookmarkEnd w:id="7"/></w:p><w:sdt><w:sdtPr><w:id w:val="11"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>controlled</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
    );
    fragment_package.set_part("/word/document.xml", fragment_xml.into_bytes());
    let mut fragment_output = std::io::Cursor::new(Vec::new());
    fragment_package.write_to(&mut fragment_output).unwrap();
    let fragment_bytes = fragment_output.into_inner();
    let record = MailMergeRecord {
        values: BTreeMap::from([
            (
                "Logo".to_owned(),
                MailMergeValue::Image(MailMergeImage {
                    data: b"\x89PNG\r\n\x1a\nlogo".to_vec(),
                    filename: "logo.png".to_owned(),
                    width: Length::emu(11),
                    height: Length::emu(19),
                }),
            ),
            (
                "InsertA".to_owned(),
                MailMergeValue::Fragment(fragment_bytes.clone()),
            ),
            (
                "InsertB".to_owned(),
                MailMergeValue::Fragment(fragment_bytes),
            ),
        ]),
        ..Default::default()
    };
    let mut output = document
        .mail_merge_rich(
            &MailMergeData {
                records: vec![record],
                ..Default::default()
            },
            None,
        )
        .unwrap()
        .remove(0);
    let bytes = output.to_bytes().unwrap();
    let package = oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let image_relationships = package
        .get_part_rels("/word/document.xml")
        .unwrap()
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == oxml_opc::relationship::rel_types::IMAGE)
        .count();
    assert_eq!(image_relationships, 3);
    let payload_parts = package
        .parts
        .iter()
        .filter(|(_, bytes)| bytes.as_slice() == b"fragment payload")
        .map(|(name, _)| name)
        .collect::<Vec<_>>();
    assert_eq!(payload_parts.len(), 2);
    let nested_payload_relationships = package
        .parts
        .keys()
        .filter(|name| name.ends_with(".png"))
        .filter_map(|name| package.get_part_rels(name))
        .flat_map(|relationships| &relationships.items)
        .filter(|relationship| relationship.rel_type == "urn:rdocx:test:fragment-payload")
        .count();
    assert_eq!(nested_payload_relationships, 2);
    assert!(
        output
            .paragraphs()
            .iter()
            .any(|p| p.text() == "fragment text")
    );
    let xml = String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
    let attribute_values = |prefix: &str| {
        xml.match_indices(prefix)
            .filter_map(|(start, _)| {
                let value = &xml[start + prefix.len()..];
                value.find('"').map(|end| value[..end].to_owned())
            })
            .collect::<Vec<_>>()
    };
    let bookmark_ids = attribute_values(r#"w:bookmarkStart w:id=""#);
    let bookmark_names = attribute_values(r#"w:name=""#);
    let hyperlink_anchors = attribute_values(r#"w:anchor=""#);
    let content_control_ids = attribute_values(r#"<w:id w:val=""#);
    let drawing_ids = attribute_values(r#"wp:docPr id=""#);
    assert_eq!(bookmark_ids.len(), 2, "{xml}");
    assert_ne!(bookmark_ids[0], bookmark_ids[1], "{xml}");
    assert_eq!(bookmark_names.len(), 2, "{xml}");
    assert_eq!(hyperlink_anchors, bookmark_names, "{xml}");
    assert_eq!(content_control_ids.len(), 2, "{xml}");
    assert_ne!(content_control_ids[0], content_control_ids[1], "{xml}");
    assert_eq!(drawing_ids.len(), 3, "{xml}");
    assert_eq!(drawing_ids.iter().collect::<HashSet<_>>().len(), 3, "{xml}");
    assert_eq!(xml.matches(">controlled<").count(), 2, "{xml}");
    let fragment_numbering = output
        .paragraphs()
        .into_iter()
        .filter(|paragraph| paragraph.text() == "fragment list")
        .map(|paragraph| paragraph.numbering().unwrap().0)
        .collect::<Vec<_>>();
    assert_eq!(fragment_numbering.len(), 2);
    assert_ne!(fragment_numbering[0], fragment_numbering[1]);
    assert!(output.style("Collision").is_some());
    assert!(output.style("CollisionMerge1").is_some());
    assert!(output.style("CollisionMerge2").is_some());
    assert_eq!(
        xml.matches(r#"<w:pStyle w:val="CollisionMerge1"/>"#)
            .count(),
        1,
        "{xml}"
    );
    assert_eq!(
        xml.matches(r#"<w:pStyle w:val="CollisionMerge2"/>"#)
            .count(),
        1,
        "{xml}"
    );
}

#[test]
fn formatting_hooks_change_only_the_selected_merge_field_runs() {
    let body = r#"
        <w:p><w:r><w:t>fixed</w:t></w:r></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD Name \* Upper"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD TableStart:Items"><w:r><w:t>start</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD Value"><w:r><w:t>item</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD TableEnd:Items"><w:r><w:t>end</w:t></w:r></w:fldSimple></w:p>
        <w:sectPr/>
    "#;
    let document = document_with_content_controls(&wrap_word_body(body));
    let data = MailMergeData {
        records: vec![MailMergeRecord {
            values: BTreeMap::from([("Name".to_owned(), MailMergeValue::Text("ada".to_owned()))]),
            regions: BTreeMap::from([(
                "Items".to_owned(),
                ["one", "two"]
                    .into_iter()
                    .map(|text| MailMergeRecord {
                        values: BTreeMap::from([(
                            "Value".to_owned(),
                            MailMergeValue::Text(text.to_owned()),
                        )]),
                        ..Default::default()
                    })
                    .collect(),
            )]),
        }],
        ..Default::default()
    };
    let mut observed = Vec::new();
    let mut formatter = |context: &rdocx::MailMergeFormatContext<'_>,
                         value: &mut MailMergeFormattedText| {
        observed.push((
            context.source_name.to_owned(),
            context.region_path.to_vec(),
            context.field_name.to_owned(),
            context.record_number,
            context.output_sequence_number,
        ));
        value.text.push('!');
        value
            .run_properties
            .get_or_insert_with(Default::default)
            .bold = Some(true);
        Ok(())
    };
    let mut output = document
        .mail_merge_rich(&data, Some(&mut formatter))
        .unwrap()
        .remove(0);
    assert_eq!(
        observed,
        [
            ("records".to_owned(), vec![], "Name".to_owned(), 1, 1),
            (
                "Items".to_owned(),
                vec!["Items".to_owned()],
                "Value".to_owned(),
                1,
                2,
            ),
            (
                "Items".to_owned(),
                vec!["Items".to_owned()],
                "Value".to_owned(),
                2,
                3,
            ),
        ]
    );
    let xml = document_xml(&mut output);
    assert!(xml.contains(">fixed<"), "{xml}");
    assert!(xml.contains(">ADA!<"), "{xml}");
    assert!(xml.contains(">one!<"), "{xml}");
    assert!(xml.contains(">two!<"), "{xml}");
    assert_eq!(xml.matches("<w:b/>").count(), 3, "{xml}");
}

#[test]
fn rich_merge_preserves_schema_order_and_unmodelled_xml_after_reopen() {
    let body = r#"
        <w:p><w:pPr><w:keepNext/></w:pPr><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>stored</w:t></w:r></w:fldSimple><producer:opaque producer:v="1"/></w:p>
        <w:sectPr/>
    "#;
    let document = document_with_content_controls(&wrap_word_body(body));
    let data = MailMergeData {
        records: vec![MailMergeRecord {
            values: BTreeMap::from([("Name".to_owned(), MailMergeValue::Text("value".to_owned()))]),
            ..Default::default()
        }],
        ..Default::default()
    };
    let mut output = document.mail_merge_rich(&data, None).unwrap().remove(0);
    let xml = document_xml(&mut output);
    let opaque = r#"<producer:opaque producer:v="1"/>"#;
    assert_eq!(xml.matches(opaque).count(), 1, "{xml}");
    assert!(xml.find("<w:pPr").unwrap() < xml.find(">value<").unwrap());
    CT_Document::from_xml(xml.as_bytes()).unwrap();
}

#[test]
fn invalid_rich_merge_input_leaves_the_template_and_outputs_uncommitted() {
    let assert_atomic_error = |mut document: Document, data: MailMergeData| {
        let before = document.to_bytes().unwrap();
        assert!(document.mail_merge_rich(&data, None).is_err());
        assert_eq!(document.to_bytes().unwrap(), before);
    };
    let one_value = |name: &str, value: MailMergeValue| MailMergeData {
        records: vec![MailMergeRecord {
            values: BTreeMap::from([(name.to_owned(), value)]),
            ..Default::default()
        }],
        ..Default::default()
    };
    let linked_fragment = |external: bool| {
        let mut fragment = Document::new();
        fragment.add_paragraph("fragment link");
        let mut package =
            oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(fragment.to_bytes().unwrap()))
                .unwrap();
        let relationship_id = if external {
            package
                .get_or_create_part_rels("/word/document.xml")
                .add_external(
                    oxml_opc::relationship::rel_types::HYPERLINK,
                    "https://example.invalid",
                )
        } else {
            "rIdMissing".to_owned()
        };
        let xml = String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec())
            .unwrap()
            .replace(
                "<w:body>",
                &format!(
                    r#"<w:body><w:p><w:hyperlink r:id="{relationship_id}"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:p>"#
                ),
            );
        package.set_part("/word/document.xml", xml.into_bytes());
        let mut output = std::io::Cursor::new(Vec::new());
        package.write_to(&mut output).unwrap();
        output.into_inner()
    };

    let body = r#"
        <w:p><w:fldSimple w:instr="MERGEFIELD TableStart:Lines"><w:r><w:t>start</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="MERGEFIELD Value"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p>
        <w:sectPr/>
    "#;
    assert_atomic_error(
        document_with_content_controls(&wrap_word_body(body)),
        MailMergeData {
            records: vec![MailMergeRecord::default()],
            ..Default::default()
        },
    );

    let image_body = r#"<w:p><w:fldSimple w:instr="MERGEFIELD Logo"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p><w:sectPr/>"#;
    assert_atomic_error(
        document_with_content_controls(&wrap_word_body(image_body)),
        one_value(
            "Logo",
            MailMergeValue::Image(MailMergeImage {
                data: vec![1],
                filename: "x.png".to_owned(),
                width: Length::emu(0),
                height: Length::emu(1),
            }),
        ),
    );

    let scalar_body = r#"<w:p><w:fldSimple w:instr="MERGEFIELD Value"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p><w:sectPr/>"#;
    assert_atomic_error(
        document_with_content_controls(&wrap_word_body(scalar_body)),
        MailMergeData::default(),
    );
    assert_atomic_error(
        document_with_content_controls(&wrap_word_body(scalar_body)),
        one_value("Value", MailMergeValue::Text("\0".to_owned())),
    );
    assert_atomic_error(
        document_with_content_controls(&wrap_word_body(scalar_body)),
        one_value("Value", MailMergeValue::Fragment(vec![1, 2, 3])),
    );
    assert_atomic_error(
        document_with_content_controls(&wrap_word_body(scalar_body)),
        one_value("Value", MailMergeValue::Fragment(linked_fragment(false))),
    );
    assert_atomic_error(
        document_with_content_controls(&wrap_word_body(scalar_body)),
        one_value("Value", MailMergeValue::Fragment(linked_fragment(true))),
    );

    let inline_fragment_body = r#"<w:p><w:r><w:t>prefix</w:t></w:r><w:fldSimple w:instr="MERGEFIELD Value"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p><w:sectPr/>"#;
    let mut valid_fragment = Document::new();
    valid_fragment.add_paragraph("fragment");
    assert_atomic_error(
        document_with_content_controls(&wrap_word_body(inline_fragment_body)),
        one_value(
            "Value",
            MailMergeValue::Fragment(valid_fragment.to_bytes().unwrap()),
        ),
    );

    let mut callback_document = document_with_content_controls(&wrap_word_body(scalar_body));
    let callback_before = callback_document.to_bytes().unwrap();
    let callback_data = one_value("Value", MailMergeValue::Text("valid".to_owned()));
    let mut formatter = |_context: &rdocx::MailMergeFormatContext<'_>,
                         _value: &mut MailMergeFormattedText| {
        Err(rdocx::Error::Other("formatter rejected value".to_owned()))
    };
    assert!(
        callback_document
            .mail_merge_rich(&callback_data, Some(&mut formatter))
            .is_err()
    );
    assert_eq!(callback_document.to_bytes().unwrap(), callback_before);
}

#[test]
fn flat_mail_merge_output_remains_unchanged_after_rich_merge_is_added() {
    let body = r#"<w:p><w:fldSimple w:instr="MERGEFIELD Name \* Upper"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p><w:p><w:fldSimple w:instr="MERGEFIELD Amount \# 0.00"><w:r><w:t>stored amount</w:t></w:r></w:fldSimple></w:p><w:sectPr/>"#;
    let mut document = document_with_content_controls(&wrap_word_body(body));
    let records = [mail_merge_record(&[("Name", "Ada"), ("Amount", "12.5")])];
    let original = document.to_bytes().unwrap();
    let mut expected = Document::from_bytes(&original).unwrap();
    expected
        .update_fields(&FieldEvaluationContext {
            merge_fields: records[0].clone(),
            merge_record_number: Some(1),
            merge_sequence_number: Some(1),
            ..Default::default()
        })
        .unwrap();
    let expected = Document::from_bytes(&expected.to_bytes().unwrap()).unwrap();
    let mut expected_bytes = expected;
    let expected_bytes = expected_bytes.to_bytes().unwrap();

    let mut separate = document.mail_merge(&records).unwrap().remove(0);
    assert_eq!(separate.to_bytes().unwrap(), expected_bytes);
    let mut sectioned = document.mail_merge_sections(&records).unwrap();
    assert_eq!(sectioned.to_bytes().unwrap(), expected_bytes);
}
