# 00, Vision

## What this repository is

`rdocx` is a Rust port of `python-docx`: open a `.docx`, walk and edit it, save
it, and render it to PDF, PNG, HTML or Markdown. It ships on crates.io and is
used as a library, a CLI and a WASM module.

This repository is becoming three things instead of one:

- **`oxml-*`**, the format-neutral Office Open XML infrastructure. The ZIP
  container, relationships, content types, units, DrawingML, image handling,
  the layout primitives and the PDF backend.
- **`rdocx-*`**, WordprocessingML on top of that infrastructure.
- **`rpptx-*`**, PresentationML on top of the same infrastructure. A Rust port
  of `python-pptx`.

## Why rpptx

The same reason rdocx exists. The incumbent tools for programmatic PowerPoint
are Python libraries that are slow, memory-hungry and cannot render. A Rust
implementation that opens a real corporate deck, fills its placeholders and
emits a PDF in one process, with no Office install and no headless LibreOffice,
is a capability nothing in the ecosystem currently has.

## Why one workspace

Three arguments, in order of weight.

**`oxml-drawing` will churn for months.** It is the largest body of new type
work in the plan, and `rpptx-oxml` is its only consumer for most of that time.
Every change wants to be atomic with its consumer. A published-crate boundary
between them would mean a release cycle per iteration.

**One of everything.** One `deny.toml`, one `ci.yml`, one lockfile, one MSRV
story, one `cargo test --workspace` that covers the whole surface. The
alternative is three copies drifting apart.

**No new tooling risk.** `publish.yml` already publishes seven crates from one
workspace in dependency order. Adding members is incremental.

The repository keeps the name `tensorbee/rdocx`. Renaming it would buy a more
accurate name and cost a redirect that breaks the moment anything is created at
the old path. The name under-describing its contents is a cosmetic problem.

## What is deliberately reused

The audit that started this work found that roughly 22 percent of the existing
lines transfer directly. That number understates the value, because the 22
percent is the infrastructure that is painful to get right:

- **`rdocx-opc` is already about 97 percent format-agnostic.** `OpcPackage` is a
  string-keyed map with generic relationship path algebra. Its
  `main_document_part()` keys off the `officeDocument` relationship type, which
  is the same one PowerPoint uses, so it reads a `.pptx` today without
  modification.
- **`rdocx-pdf` is format-independent.** It depends on the layout crate and
  nothing else in the workspace, and it consumes only `LayoutResult`. A slide is
  a page with a fixed size. Font subsetting, ToUnicode CMaps, JPEG passthrough
  and the tiny-skia rasteriser all carry over untouched.
- **About 1,018 generic lines sit inside `rdocx-oxml`**: units, raw-XML capture,
  entity decoding, core properties.

The honest gap is that PowerPoint is roughly 90 percent DrawingML and rdocx has
almost none. Of `drawing.rs`, only `write_graphic_element` transfers. That work
is new, but it was always going to be new.

## What done looks like

A single release in which:

- `rdocx` continues to pass every test it passes today, on a shared
  infrastructure it did not have before.
- `rpptx` opens a corporate template, adds slides from its layouts, fills
  placeholders and tables, adds pictures and charts, and saves a file that
  PowerPoint opens without a repair prompt.
- `rpptx` renders that deck to PDF and to PNG at a quality that is
  indistinguishable at a glance from PowerPoint's own export.
- Both ship as Rust crates, CLIs, WASM modules and Python wheels.

The post-v1 roadmap now adds two Word authoring boundaries before the
conditional spreadsheet programme. M23 proves that five representative
business documents can be generated from `Document::new()` through public
`rdocx` APIs alone. M24 closes the audited modern DOCX authoring matrix across
stories, fields, forms, collaboration, drawings, package extensions,
accessibility, conformance, determinism, and bindings. Complete authoring does
not mean executing proprietary payloads or reproducing every Word rendering
effect.

## What this costs

Charts mean ChartML plus a minimal SpreadsheetML writer, because every chart
embeds its own workbook part. v1 therefore spans three OOXML formats.

Sized at story level in `14-development-backlog.md`, the backlog is 150 stories
and roughly 390 developer-days, which is **17 to 18 months solo**. An earlier
phase-level estimate said nine to twelve. The story-level number is the
trustworthy one.

Two ways to compress, neither requiring any rework: a second developer takes it
to roughly 9 to 11 months, since M7 and M8 parallelise once M6 lands and M12 is
self-contained throughout. Or cut a read-plus-render release at the end of M10,
which is about 12 months solo and is the point where the library becomes
genuinely useful.

This is recorded here so the trade is made deliberately rather than discovered
in month nine.

The active post-v1 delivery record is larger than this historical v1 estimate.
M23 adds 24 stories and roughly 112 developer-days. M24 adds 47 stories and
roughly 219 developer-days. Their planned feature sprints occupy S70 through
S75, S86, S89 through S91, S95.2 through S95.4, and S99 through S101. The S76 through S88 contribution and issue repair
programme comes before the remaining feature work. S102 measures Word PDF
conversion against an external Word-oracle benchmark, and conditional
spreadsheet work begins at S103. The F-240 audit closed 85 stable capability
rows against the public facade, modeled Word properties, bindings, package
ownership, and five anonymous private references. Its evidence confirmed the
F-243 through F-310 story boundaries, sizes, and dependencies without adding
another backlog or document model. The current plan places them across S71
through S75, S86, S89 through S91, S95.2 through S95.4, and S99 through S101.
