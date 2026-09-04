# Sprint Tracker

Velocity log. One row per completed F-ID, appended by `/complete-feature`, plus
a per-sprint summary appended by `/close-sprint`.

Estimates come from `docs/hld/14-development-backlog.md`. Actuals are recorded
so the velocity assumption can be corrected against reality rather than
defended.

`S = 1d`, `M = 2-3d`, `L = 4-5d`, `XL = split me`.

## Per-sprint summary

| Sprint | Milestone | Planned | Done | Carried | Est. days | Actual days | Notes |
|--------|-----------|---------|------|---------|-----------|-------------|-------|
| S01 | M1 | 6 | 6 | 0 | 10 | 6 | Completed with no carries |
| S02 | M1 | 6 | 6 | 0 | 8 | 6 | Completed M1 and published rdocx 0.4.1 |
| S03 | M2 | 5 | 3 | 2 | 8 | 3 | F-015 and F-016 carried to S04 to keep rdocx 0.5.0 independent of unpublished oxml-core |
| S04 | M2 | 7 | 4 | 3 | 9 | 4 | F-015, F-016, and F-022 carried to S32.2 so development crates remain unpublished until PowerPoint is complete |
| S05 | M3 | 4 | 4 | 0 | 8 | 4 | Completed isolated unpublished oxml-media staging, with F-027 and F-028 remaining planned for S32.2 |
| S06 | M4 | 3 | 3 | 0 | 8 | 3 | Completed unpublished oxml-layout staging, with M4 continuing in S07 |
| S07 | M4 | 5 | 5 | 0 | 8 | 5 | Completed M4 in unpublished oxml-layout with all 28 hashes unchanged |
| S08 | M5 | 3 | 3 | 0 | 7 | 3 | Staged unpublished oxml-pdf, installed the exact golden gate, and completed the global CTM rewrite |
| S09 | M5 | 4 | 4 | 0 | 7 | 4 | Completed nested groups, solid paths, transform-aware collection, and reusable alpha in unpublished oxml-pdf |
| S10 | M5 | 2 | 2 | 0 | 8 | 2 | Completed M5 with PDF gradients and recursive raster groups, paths, clips, gradients, dashes, and backgrounds |
| S11 | M6 | 0 | 0 | 0 | 0 | 1 | Confirmed the staged extraction boundary with no publication, consumer cutover, or implementation F-IDs |
| S12 | M7 | 5 | 5 | 0 | 11 | 5 | Completed the first M7 DrawingML slice with exact PowerPoint colour evidence and no publication |
| S13 | M7 | 4 | 4 | 0 | 12 | 4 | Completed transforms, custom geometry, and fills in unpublished oxml-drawing with all 28 hashes unchanged |
| S14 | M7 | 8 | 8 | 0 | 14 | 8 | Completed lines, effects, shape properties, style references, and the split text model with all 28 hashes unchanged and no publication |
| S15 | M7 | 2 | 2 | 0 | 5 | 2 | Completed themes and the stable Word adapter with pinned PowerPoint acceptance, all 28 hashes unchanged, and no publication. The external corpus boundary runs with F-067 at S16 entry |
| S16 | M8 | 4 | 4 | 0 | 12 | 4 | Established the pinned 50-deck corpus and modelled core PresentationML parts and recursive shape trees with all 28 hashes unchanged and no publication |
| S17 | M8 | 4 | 4 | 0 | 10 | 4 | Completed placeholders, pictures, graphic-frame dispatch, and DrawingML tables against all 50 pinned decks with all 28 hashes unchanged and no publication |
| S18 | M8 | 4 | 4 | 0 | 7 | 4 | Completed connectors, alternate-content fallback selection, notes parts, and relationship-id rewriting against all 50 pinned decks with all 28 hashes unchanged and no publication |
| S19 | M8 | 2 | 2 | 0 | 6 | 2 | Completed the rpptx read facade and modelled 50-deck gate with native PowerPoint acceptance, all 28 hashes unchanged, and no publication |
| S20 | M9 | 5 | 5 | 0 | 11 | 5 | Completed placeholder, transform, body, text-style, format-scheme, and typeface inheritance with all 28 hashes unchanged and no publication |
| S21 | M9 | 3 | 3 | 0 | 8 | 3 | Completed M9 with the frozen ResolvedSlide contract, strict all-slide corpus resolution, native PowerPoint acceptance, all 28 hashes unchanged, and no publication |
| S22 | M10 | 4 | 4 | 0 | 9 | 4 | Completed preset provenance, generation, evaluation, fallback, and the unpublished renderer input boundary with all 28 hashes unchanged and no publication |
| S23 | M10 | 5 | 5 | 0 | 10 | 5 | Completed shapes, transforms, arrowheads, cropped and tiled pictures, and backgrounds in the unpublished renderer with all 28 hashes unchanged and no publication |
| S24 | M10 | 8 | 8 | 0 | 14 | 8 | Completed shape text layout, bullets, autofit, and vertical text in the unpublished renderer with all 28 hashes unchanged and no publication |
| S25 | M10 | 3 | 3 | 0 | 10 | 3 | Completed M10 with table rendering, hyperlinks, fields, diagnostics, and the pinned fidelity harness. All 421 slides rendered, the SSIM trend and native PowerPoint evidence were retained, all 28 hashes remained unchanged, and no crate was published |
| S26 | M11 | 4 | 4 | 0 | 10 | 4 | Established slide creation with the bundled template, collision-safe identifiers and media, layout-based slide synthesis, and deterministic validation. Native PowerPoint accepted the generated decks, all 50 pinned decks validated cleanly, all 28 hashes remained unchanged, and no crate was published |
| S27 | M11 | 4 | 4 | 0 | 12 | 4 | Added mutable shape and text handles, schema-ordered shape constructors, atomic picture insertion, native image sizing, and typed text formatting. Pinned PowerPoint accepted the generated decks without repair, python-pptx 1.0.2 sizing parity passed, all 28 hashes remained unchanged, and no crate was published |
| S28 | M11 | 4 | 4 | 0 | 9 | 4 | Completed M11 with mutable tables, slide collection operations, presentation properties, and one SHA-bound ten-slide acceptance deck. PowerPoint, Keynote, Google Slides, and LibreOffice accepted the deck without repair or conversion error, all 28 hashes remained unchanged, and no crate was published |
| S29 | M12 | 3 | 3 | 0 | 12 | 3 | Established the chart data layer with a minimal editable workbook writer, schema-aware ChartML core types, and consistent formula references and caches. Excel and LibreOffice Calc accepted the generated workbook, all 50 pinned decks passed the chart corpus gates, all 28 hashes remained unchanged, and no crate was published |
| S30 | M12 | 4 | 4 | 0 | 12 | 4 | Completed typed axes, data labels, number formats, and all seven v1 plot families with reciprocal axis validation, 50-deck structural gates, SHA-bound viewer evidence, all 28 hashes unchanged, and no crate published |
| S31 | M12 | 3 | 3 | 0 | 12 | 3 | Added atomic editable chart authoring, backend-neutral chart geometry, and deterministic axes, gridlines, labels, and legends. Pinned PowerPoint accepted the generated chart without repair and exposed exact editable workbook values, all 50 pinned chart decks and 28 hashes passed unchanged, and no crate was published |
| S32 | M12 | 2 | 2 | 0 | 3 | 2 | Completed M12 with exact theme-mapped chart colours and source-scoped native or preserved fallbacks. The SHA-bound authored chart opened in pinned PowerPoint without repair, Edit Data exposed exact workbook values, the production renderer emitted native geometry, all 28 hashes remained unchanged, and no crate was published |
| S32.1 | M6 | 4 | 4 | 0 | 7 | 4 | Completed shared publication readiness with exact package inventory, split-family release preparation, stable and incubating tag routing, binding-safe CI gates, and verified local-source publication dry runs. Every generated archive remained below 10 MiB, all 28 hashes remained unchanged, and no crate was published |
| S32.2 | M6 | 8 | 8 | 0 | 10 | 8 | Completed M6 by publishing the shared 0.1.2 family, cutting released rdocx consumers over to shared facades and backends, and documenting the migration. All package archives verified below 10 MiB and all 28 hashes remained unchanged |
| S33 | M13 | 5 | 5 | 0 | 13 | 5 | Established shared Python path and revision support, lazy rdocx handles, formatting, tables, Python-compatible values and errors, and GIL-detached rendering. The cp39-abi3 wheel passed 31 tests, all 28 hashes remained unchanged, and no package was published |
| S34 | M13 | 5 | 5 | 0 | 11 | 5 | Completed typed Python packages, bidirectional python-docx and python-pptx parity, six-target wheel automation, and pull-request binding CI. The hosted every-platform M13 gate remains for later M13 work, all 28 hashes remained unchanged, and no package was published |
| S35 | M13 | 4 | 4 | 0 | 9 | 4 | Completed facade-backed Word and presentation WASM packages, browser PDF export, and two-package WASM CI. Hosted wheels installed and passed parity on every M13 target, all 28 hashes remained unchanged, and no package was published |
| S36 | M13, X | 8 | 8 | 0 | 13 | 8 | Completed both v1 CLIs, local npm package assembly, README doctests, sample-generator deduplication, and concurrent test isolation. All 28 hashes remained unchanged, and no Rust or npm package was published |
| S37 | M13, X | 1 | 1 | 0 | 1 | 1 | Published and verified the complete 14-package rpptx 0.1.3 family through the reviewed release workflow. The M13 hosted wheel gate remains satisfied, all 28 hashes remained unchanged, and no npm package was published |
| S38 | X | 2 | 2 | 0 | 5 | 2 | Integrated PR 25 with Jon Stokes credited, added compile-checked documentation for every stable crate, and published the seven-package stable rdocx 0.5.0 family. All 28 hashes remained unchanged, and no incubating, WASM, Python, or npm package was published |
| S39 | X | 3 | 3 | 0 | 6 | 3 | Added compile-checked README coverage for all 26 workspace crates, published the seven-package stable rdocx 0.6.0 family, and published the complete 14-package rpptx 0.2.0 family. Every crates.io README rendered, all 28 hashes remained unchanged, and no WASM, Python, npm, or PyPI package was published |
| S40 | X | 1 | 1 | 0 | 2 | 1 | Restored a green hosted CI baseline by pinning the Poppler 26.01.0 rendering oracle, the Binaryen 125 identity, uv 0.10.2 and the LibreOffice viewer build behind one bounded installer. All 28 hashes remained unchanged, and no crate, release version, published package or rendering baseline changed |
| S41 | X | 7 | 7 | 0 | 12 | 7 | Landed the surviving half of the external PR 2 rendering contribution as six stories under one umbrella. Footnote lines advance and reserve their own page space, oversized notes split across pages, endnotes moved to the document end with the two note streams keyed apart, kashida justification no longer fails a document open, and body text flows around wrapping drawings. All 28 hashes unchanged throughout, and no crate, package version or published artifact changed |
| S42 | X | 4 | 4 | 0 | 5 | 4 | Refreshed the lockfile, broke the cross-family dependency edge that made the two publication trains mutually dependent, and released both: fourteen incubating crates at 0.3.0 and seven stable crates at 0.7.0, from one reviewed SHA. All 28 hashes unchanged throughout, and zero vulnerabilities across 152 dependencies |
| S43 | X | 5 | 5 | 0 | 8 | 1 | Cleared the five follow-ups S41 and S42 filed. An unmodelled enumerated value now reads as an absent attribute across nine parsers, notes break to the width of the section carrying their reference, and a wrapping drawing anchored to a later paragraph pushes earlier text aside through a second pagination pass. The harness gained a three-part PDF fingerprint per sample, 28 entries to 49, and the story that added it found the PDF writer was not deterministic and fixed that too. /verify now runs the release preflights that publish.yml treats as its publication gate. One declared harness delta, 21 added and 0 changed, and no crate, package version or published artifact changed |
| S44 | X | 4 | 4 | 0 | 6 | 4 | Closed four CI and workflow follow-ups from the S43 review. Pull-request CI now runs the full release-regression module and the golden-PNG gate under pinned Poppler, agent-facing repository claims are regression checked, and path filters preserve a stable aggregate gate. The full close gate passed, all 49 hashes remained unchanged, and no crate, package version, rendering baseline or published artifact changed |
| S45 | M15 | 4 | 4 | 0 | 10 | 4 | Completed one shared chart engine across Word and PowerPoint. Word authoring now saves native editable chart parts and workbooks, `Document::add_chart` authors bar, line, and pie charts, and the Word paginator renders through shared geometry. Microsoft Word 16.104 opened the SHA-bound document without repair and Edit Data changed its workbook values. The cross-family golden had zero differing RGBA pixels, all 49 hashes remained unchanged, and no package was published |
| S46 | M14 | 5 | 5 | 0 | 12 | 5 | Added typed comments, recursive content controls, custom XML binding, bookmarks, and cross-reference resolution. The full close gate passed with all 49 hashes unchanged. The SHA-bound Word comment thread candidate remains recorded for later human UI observation, and no package was published |
| S47 | M14 | 2 | 2 | 0 | 8 | 2 | Added typed revision discovery and atomic accept and reject operations across content, properties, contextual markers, hyperlinks, tables, and controls. Microsoft Word 16.104 normalized-body parity passed, the full close gate passed with all 49 hashes unchanged, and no package was published. M14 continues in S48 for revision display and document protection |
| S48 | M14 | 2 | 2 | 0 | 4 | 2 | Completed M14 with accepted and tracked revision rendering, deterministic change decorations, and typed document-protection intent. The mixed collaboration gate read and wrote revisions, comments, content controls, and bookmarks in one preserved package. The full close gate passed with all 49 hashes unchanged, and no package was published |
| S49 | M16 | 4 | 4 | 0 | 12 | 4 | Established one recursive field grammar, deterministic evaluation against pinned Word results, atomic cache-update policies, and namespace-aware reader preservation. The full close gate passed with all 49 hashes unchanged, every package archive remained below 10 MiB, and the sprint review was clean on pass 1. M16 continues in S50 with template syntax, loops, conditionals, and repeating structures |
| S50 | M16 | 3 | 3 | 0 | 10 | 3 | Added atomic JSON template rendering across run boundaries, nested loops and conditionals, multi-row table repetition, and continuous list numbering. The full close gate passed with all 49 hashes unchanged, all 22 package archives below 10 MiB, and the sprint review clean on pass 3 after four blocking interactions were remediated. M16 continues in S51 with mail merge, comparison, and watermarks |
| S51 | M16, X | 10 | 10 | 0 | 18 | 2 | Completed M16 with mail merge, comparison, watermarks, complete layout provenance, bounded relayout caches, and ordered body access from PR 36. Published the complete rpptx 0.4.0 and stable rdocx 0.8.0 families with reviewed release notes and contributor credit. The full close gate passed with all 49 hashes unchanged and no carries |
| S52 | M17, X | 12 | 12 | 0 | 27 | 1 | Added authenticated Agile encryption reads and Word-compatible writes, exact digital-signature verification, corrected break shaping, and bounded exact editor relayout reuse. PRs 40 and 41 were audited, safely reimplemented, credited, commented, and closed. The full close gate passed with all 49 hashes matching the reviewed baseline and no carries. M17 continues in S53 with signature creation and PDF compliance |
| S53 | M17, X | 7 | 7 | 0 | 16 | 1 | Completed M17 with signature creation, tagged PDF, PDF/A, redaction, and dense-form tables. Published and verified the 15-package incubating 0.5.0 family and seven-package stable 0.9.0 family. The full close gate passed with the reviewed 14-entry PDF-only harness delta and no carries |
| S54 | M18, X | 4 | 4 | 0 | 9 | 1 | Opened M18 with bounded RTF reading and deterministic diagnostic RTF writing, shared PNG, JPEG, TIFF, transparency, and page-range export, plus exact caller-font family aliases. The full close gate passed with all 49 hashes and 7 golden pixel buffers unchanged, all package archives below 10 MiB, and sprint review clean on pass 2 after two gate gaps were fixed. M18 continues in S55 with HTML and ODT input |
| S55 | M18, X | 4 | 4 | 0 | 13 | 1 | Added bounded HTML and ODT input, restored exact relayout performance within the Issue 46 budget, and completed migration and contribution records. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and sprint review clean on pass 2. M18 continues in S56 with ODT writing and outbound formats |
| S56 | M18, X | 7 | 6 | 1 | 15 | 1 | Completed M18 with deterministic ODT writing, EPUB 3 export, searchable SVG pages, and hardened outcomes from PRs 47 through 52. The partial v0.10.0 attempt was archived and recovered by publishing the complete 15-package incubating 0.6.0 family before the complete seven-package stable 0.10.1 family. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and the sprint review clean on the authorized exact-HEAD pass 6 |
| S57 | M20 | 3 | 3 | 0 | 10 | 1 | Established the pinned Word corpus, complete-union SSIM evidence, and thousand-page performance limits. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and sprint review clean on pass 1. M20 continues in S58 with shaping and incremental layout |
| S58 | M20, X | 18 | 17 | 1 | 32 | 4 | Completed M20 with multilingual shaping, automatic hyphenation, bidirectional layout, and bounded incremental pagination. Published the shared 0.7.0 and 0.8.0 families and stable 0.11.1 recovery, then yanked the two incomplete 0.11.0 entries. Hardened Issues 53 and 54 plus PRs 55 through 58, and required the aggregate CI gate on the default branch. F-X060 records the archived partial 0.11.0 attempt. The full close gate passed with 49 of 49 hashes, the approved five-key feature-showcase delta, 22 package archives below 10 MiB, and final sprint review pass 26 clean |
| S59 | M21 | 2 | 2 | 0 | 6 | 1 | Opened M21 with editable comments, replies, sections, notes and handout metadata, plus native password and signature operations with current-state invalidation. Pinned PowerPoint opened the encrypted candidate with the correct password and rejected a wrong password. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and sprint review pass 3 clean. M21 continues in S60 with animation and transitions |
| S60 | M21 | 2 | 2 | 0 | 8 | 2 | Added a typed timing and transition model, deterministic timeline evaluation, ordinary transition rendering, and bounded explicit-name morph composition without changing static output. The nine-case PowerPoint differential passed within 0.96 point and 0.997866 SSIM, while the full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and sprint review pass 2 clean. M21 continues in S61 with media playback and animated export |
| S61 | M21 | 3 | 3 | 0 | 10 | 1 | Added relationship-safe audio and video package mutation, deterministic poster and playback state, and bounded animated GIF and Motion JPEG AVI export. The macOS and Linux arm64 manifests matched exactly. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and close-boundary sprint review pass 7 clean. M21 continues in S62 with embedded objects and SmartArt |
| S62 | M21, X | 4 | 3 | 1 | 16 | 1 | Added relationship-safe OLE, ActiveX, and VBA inventory and mutation, a bounded typed SmartArt inspection and editing model, and hardened outcomes from PRs 61 through 64. F-220 carries to S63 after the ten-pass microscope bound found two remaining exact fail-closed SmartArt validator gaps. Its worker and evidence remain intact, and F-222 waits for its clean completion. The full close gate passed with all 49 hashes unchanged and every package archive below 10 MiB |
| S63 | M21, X | 6 | 6 | 0 | 20 | 1 | Completed authentic SmartArt rendering, bounded ODP interchange, modern presentation package variants, notes and handout export, and the fixes for Issues 65 and 66. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and sprint review pass 1 clean. M21 remains open for its representative-deck gate and later import work |
| S64 | M21, X | 5 | 5 | 0 | 12 | 2 | Completed M21 with bounded editable HTML and PDF import, the Issue 67 pagination correction, and the combined representative-deck PowerPoint gate. Published and independently verified the 15-package rpptx 0.9.0 family and seven-package rdocx 0.12.0 family. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and sprint review pass 9 clean |
| S65 | M22 | 3 | 3 | 0 | 8 | 1 | Opened M22 with typed OfficeMath authoring, deterministic layout and PDF rendering, and bounded MathML and LaTeX conversion. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and sprint review pass 9 clean. M22 continues in S66 with extended field evaluation and dynamic table of contents rebuild |
| S66 | M22 | 2 | 2 | 0 | 8 | 2 | Added structured extended field evaluation and deterministic dynamic table of contents rebuilding with pinned Word parity, final page targets, atomic failure, and source-preserving ownership. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and final sprint review pass 3 clean. M22 continues in S67 with advanced automation and comparison |

## Completed features

| F-ID | Sprint | Size | Est. days | Actual days | Completed | Notes |
|------|--------|------|-----------|-------------|-----------|-------|
| F-001 | S01 | M | 2 | 1 | 2026-07-29 | Deterministic bundled-font path |
| F-002 | S01 | S | 1 | 1 | 2026-07-29 | Rust 1.97.1 toolchain pin |
| F-003 | S01 | L | 4 | 1 | 2026-07-29 | Initial 28-entry hash baseline |
| F-004 | S01 | S | 1 | 1 | 2026-07-29 | Caladea licence and notice |
| F-005 | S01 | S | 1 | 1 | 2026-07-29 | Collision-safe image suffix allocation |
| F-006 | S01 | S | 1 | 1 | 2026-07-29 | Safe JPEG standalone-marker walk |
| F-007 | S02 | S | 1 | 1 | 2026-07-30 | Relationship-based core properties |
| F-008 | S02 | M | 2 | 1 | 2026-07-30 | 61 non-consuming setter twins |
| F-009 | S02 | M | 2 | 1 | 2026-07-30 | Thread-safe two-mode layout cache |
| F-010 | S02 | S | 1 | 1 | 2026-07-30 | Fourteen crates.io names reserved |
| F-011 | S02 | S | 1 | 1 | 2026-07-30 | Unit truncation behavior pinned |
| F-012 | S02 | S | 1 | 1 | 2026-07-30 | Published and tagged rdocx 0.4.1 |
| F-013 | S03 | M | 2 | 1 | 2026-07-30 | Unpublished shared OOXML core |
| F-014 | S03 | M | 2 | 1 | 2026-07-30 | Shared schema unit types |
| F-017 | S03 | M | 2 | 1 | 2026-07-30 | Shared app and custom properties |
| F-018 | S04 | M | 2 | 1 | 2026-07-30 | Unpublished format-neutral OPC package |
| F-019 | S04 | S | 1 | 1 | 2026-07-30 | PresentationML package constants |
| F-020 | S04 | M | 2 | 1 | 2026-07-30 | Code-built PowerPoint OPC proof |
| F-021 | S04 | S | 1 | 1 | 2026-07-30 | Canonical ZIP entry normalization |
| F-023 | S05 | M | 2 | 1 | 2026-07-30 | Dependency-free image format sniffing |
| F-024 | S05 | L | 4 | 1 | 2026-07-30 | Safe image metadata and DPI probing |
| F-025 | S05 | S | 1 | 1 | 2026-07-30 | Collision-free shared media naming |
| F-026 | S05 | S | 1 | 1 | 2026-07-30 | Dependency-free native EMU sizing |
| F-029 | S06 | M | 2 | 1 | 2026-07-31 | Unpublished layout output and font staging |
| F-030 | S06 | L | 4 | 1 | 2026-07-31 | Owned format-neutral line-breaking boundary |
| F-031 | S06 | M | 2 | 1 | 2026-07-31 | Six-coefficient affine transforms |
| F-032 | S07 | M | 2 | 1 | 2026-07-31 | Backend-neutral path geometry |
| F-033 | S07 | M | 2 | 1 | 2026-07-31 | Gradient, tile, and stroke paint model |
| F-034 | S07 | M | 2 | 1 | 2026-07-31 | Nested group and path output arms |
| F-035 | S07 | S | 1 | 1 | 2026-07-31 | Transform-aware nested leaf traversal |
| F-036 | S07 | S | 1 | 1 | 2026-07-31 | Content-addressed staged image keys |
| F-037 | S08 | S | 1 | 1 | 2026-07-31 | Unpublished shared PDF backend staging |
| F-038 | S08 | M | 2 | 1 | 2026-07-31 | Exact deterministic golden-PNG gate |
| F-039 | S08 | L | 4 | 1 | 2026-07-31 | Global page CTM with reviewed pixel delta |
| F-040 | S09 | M | 2 | 1 | 2026-07-31 | Recursive PDF group graphics states |
| F-041 | S09 | M | 2 | 1 | 2026-07-31 | Solid PDF path geometry and paint |
| F-042 | S09 | M | 2 | 1 | Nested font, image, and link collection |
| F-044 | S09 | S | 1 | 1 | Reused PDF ExtGState alpha resources |
| F-043 | S10 | L | 4 | 1 | 2026-07-31 | Deterministic PDF gradient resource graphs |
| F-045 | S10 | L | 4 | 1 | 2026-07-31 | Recursive raster groups, paths, gradients, and dashes |
| F-052 | S12 | S | 1 | 1 | 2026-07-31 | Unpublished DrawingML crate and namespace constants |
| F-053 | S12 | M | 2 | 1 | 2026-07-31 | Schema-boundary raw child ordering |
| F-054 | S12 | M | 2 | 1 | 2026-07-31 | Four DrawingML colour choices with raw preservation |
| F-055 | S12 | L | 4 | 1 | 2026-07-31 | Exact PowerPoint colour transform stack |
| F-056 | S12 | M | 2 | 1 | 2026-07-31 | Master colour-map resolution before theme lookup |
| F-057 | S13 | M | 2 | 1 | 2026-07-31 | DrawingML transforms and exact nested composition |
| F-058 | S13 | L | 4 | 1 | 2026-07-31 | Guide formulas, path evaluation, and arc lowering |
| F-059 | S13 | M | 2 | 1 | 2026-07-31 | Custom geometry XML model and evaluation |
| F-060 | S13 | L | 4 | 1 | 2026-07-31 | DrawingML fill families with raw preservation |
| F-061 | S14 | M | 2 | 1 | 2026-08-01 | DrawingML line properties and preset dash mapping |
| F-062 | S14 | S | 1 | 1 | 2026-08-01 | Outer shadows with unsupported effect preservation |
| F-063 | S14 | M | 2 | 1 | 2026-08-01 | Shape properties and four style-reference forms |
| F-064a | S14 | M | 2 | 1 | 2026-08-01 | Text body properties and typed shell |
| F-064b | S14 | L | 4 | 1 | 2026-08-01 | Paragraphs, runs, fields, breaks, and whitespace |
| F-064c | S14 | S | 1 | 1 | 2026-08-01 | Character, automatic, and no-bullet forms |
| F-064d | S14 | M | 2 | 1 | 2026-08-01 | Fixed nine-level list styles |
| F-064 | S14 | XL | 0 | 1 | 2026-08-01 | Umbrella closed after four child stories and integrated gates |
| F-065 | S15 | L | 4 | 1 | 2026-08-01 | Complete DrawingML theme and pinned PowerPoint default |
| F-066 | S15 | S | 1 | 1 | 2026-08-01 | Stable Word theme projection through the documented edge |
| F-067 | S16 | M | 2 | 1 | 2026-08-01 | Unpublished PresentationML crate and pinned 50-deck corpus harness |
| F-068 | S16 | M | 2 | 1 | 2026-08-01 | Presentation root, sizes, identifiers, and default text style |
| F-069 | S16 | L | 4 | 1 | 2026-08-01 | Slide, layout, master, colour-map, and text-style models |
| F-070 | S16 | L | 4 | 1 | 2026-08-01 | Recursive ordered shape tree with opaque child payloads |
| F-071 | S17 | M | 2 | 1 | 2026-08-01 | Presence-sensitive placeholder keys and typed partial shapes |
| F-072 | S17 | M | 2 | 1 | 2026-08-01 | Typed pictures with crops, relationships, and placeholders |
| F-073 | S17 | M | 2 | 1 | 2026-08-01 | Graphic-frame URI dispatch with typed tables and opaque payloads |
| F-074 | S17 | L | 4 | 1 | 2026-08-01 | DrawingML tables with merges, banding, and preserved content |
| F-075 | S18 | S | 1 | 1 | 2026-08-01 | Typed connectors with optional start and end connections |
| F-076 | S18 | M | 2 | 1 | 2026-08-01 | Raw-preserving alternate content with typed fallback selection |
| F-077 | S18 | M | 2 | 1 | Notes parts and body-placeholder speaker-note extraction |
| F-078 | S18 | M | 2 | 1 | Namespace-aware relationship-id rewriting in preserved XML |
| F-079 | S19 | L | 4 | 1 | 2026-08-02 | Unpublished relationship-resolved rpptx read facade |
| F-080 | S19 | M | 2 | 1 | 2026-08-02 | Seven-root 50-deck modelled round-trip and native PowerPoint gate |
| F-081 | S20 | M | 2 | 1 | 2026-08-02 | Unpublished resolver context and recursive placeholder chain |
| F-082 | S20 | M | 2 | 1 | 2026-08-02 | Typed ordinary-shape properties plus transform and body inheritance |
| F-083 | S20 | L | 4 | 1 | 2026-08-02 | Seven-source, nine-level text-property cascade with safe caching |
| F-084 | S20 | M | 2 | 1 | 2026-08-02 | Typed shape styles and format-scheme resolution with placeholder colours |
| F-085 | S20 | S | 1 | 1 | 2026-08-02 | Major and minor theme-token typeface resolution |
| F-086 | S21 | L | 4 | 1 | 2026-08-02 | Final draw-order flattener with inherited-shape suppression |
| F-087 | S21 | M | 2 | 1 | 2026-08-02 | Frozen owned ResolvedSlide contract with concrete renderer values |
| F-088 | S21 | M | 2 | 1 | 2026-08-02 | Pinned visual differential and native PowerPoint acceptance gates |
| F-089 | S22 | S | 1 | 1 | 2026-08-02 | Licensed ECMA preset geometry provenance decision |
| F-090 | S22 | L | 4 | 1 | 2026-08-02 | Reproducible complete preset geometry table generator |
| F-091 | S22 | M | 2 | 1 | 2026-08-02 | Known preset evaluation and diagnosed bounds fallback |
| F-092 | S22 | M | 2 | 1 | 2026-08-02 | Unpublished scoped relationship and RenderInput boundary |
| F-093 | S23 | L | 4 | 1 | 2026-08-03 | Shape paths with solid, gradient, outline, and visible fallback paint |
| F-094 | S23 | M | 2 | 1 | 2026-08-03 | Exact rotation, centre flip, translation, and parent transform composition |
| F-095 | S23 | S | 1 | 1 | 2026-08-03 | Source-neutral resolved line ends lowered to filled paths |
| F-096 | S23 | M | 2 | 1 | 2026-08-03 | Source-scoped cropped, stretched, and tiled picture rendering |
| F-097 | S23 | S | 1 | 1 | 2026-08-03 | Preserving explicit background projection and concrete paint resolution |
| F-098a | S24 | M | 2 | 1 | 2026-08-05 | Preset text rectangle, insets, bounds fallback, and clamped content box |
| F-098b | S24 | L | 4 | 1 | 2026-08-05 | Concrete resolved runs, script-aware typefaces, fields, and explicit breaks |
| F-098c | S24 | M | 2 | 1 | 2026-08-05 | Paragraph wrapping, spacing, alignment, line stacking, and shared emission |
| F-098d | S24 | S | 1 | 1 | 2026-08-05 | Top, centre, bottom, justified, and distributed text anchoring |
| F-098 | S24 | XL | 0 | 1 | 2026-08-05 | Umbrella closed after four child stories and integrated gates |
| F-099 | S24 | M | 2 | 1 | 2026-08-05 | Character and automatic bullet markers with deterministic numbering |
| F-100 | S24 | M | 2 | 1 | 2026-08-05 | Stored and computed autofit with visible overflow policies |
| F-101 | S24 | S | 1 | 1 | 2026-08-05 | Rotated vertical text with visible diagnosed fallbacks |
| F-102 | S25 | L | 4 | 1 | 2026-08-08 | Styled table resolution and merged-cell rendering with unique borders |
| F-103 | S25 | M | 2 | 1 | 2026-08-08 | Source-scoped hyperlinks, slide-number fields, and visible diagnostics |
| F-104 | S25 | L | 4 | 1 | 2026-08-08 | Complete corpus rendering with retained SSIM trend and native evidence |
| F-105 | S26 | M | 2 | 1 | 2026-08-08 | Bundled zero-slide PowerPoint template and concrete constructor |
| F-106 | S26 | M | 2 | 1 | 2026-08-08 | Recursive shape-id allocation and content-addressed media deduplication |
| F-107 | S26 | L | 4 | 1 | 2026-08-08 | Layout-based slide synthesis with unique package and slide identifiers |
| F-108 | S26 | M | 2 | 1 | 2026-08-08 | Deterministic presentation validation and guarded save boundaries |
| F-109 | S27 | L | 4 | 1 | 2026-08-08 | Borrowed shape mutation handles with preservation-aware setters |
| F-110 | S27 | M | 2 | 1 | 2026-08-08 | Schema-ordered textbox, shape, connector, and group construction |
| F-111 | S27 | M | 2 | 1 | 2026-08-08 | Atomic picture insertion with native sizing and media deduplication |
| F-112 | S27 | L | 4 | 1 | 2026-08-08 | Borrowed text-frame editing with typed paragraph and run formatting |
| F-113 | S28 | L | 4 | 1 | 2026-08-09 | Mutable table construction, formatting, merge, split, and preservation |
| F-114 | S28 | M | 2 | 1 | 2026-08-09 | Validated slide removal, movement, duplication, and graph repair |
| F-115 | S28 | S | 1 | 1 | 2026-08-09 | Slide and presentation properties with valid slideshow output |
| F-116 | S28 | M | 2 | 1 | 2026-08-09 | One SHA-bound ten-slide acceptance deck across four viewers |
| F-117 | S29 | L | 4 | 1 | 2026-08-10 | Minimal deterministic one-sheet SpreadsheetML workbook writer |
| F-118 | S29 | L | 4 | 1 | 2026-08-10 | Schema-aware ChartML core with ordered opaque preservation |
| F-119 | S29 | L | 4 | 1 | 2026-08-10 | Formula references and caches from one source of series data |
| F-120 | S30 | L | 4 | 1 | 2026-08-10 | Four typed ChartML axis forms with reciprocal pairing and ordered preservation |
| F-123 | S30 | M | 2 | 1 | 2026-08-10 | Typed data labels and deterministic number-format projection |
| F-121 | S30 | M | 2 | 1 | 2026-08-10 | Typed bar and line plots with zero-delta viewer evidence |
| F-122 | S30 | L | 4 | 1 | 2026-08-10 | Remaining five v1 plot families with SHA-bound viewer evidence |
| F-124 | S31 | L | 4 | 1 | 2026-08-10 | Atomic editable chart authoring with native PowerPoint evidence |
| F-125 | S31 | L | 4 | 1 | 2026-08-10 | Backend-neutral geometry for all supported chart families |
| F-126 | S31 | L | 4 | 1 | 2026-08-10 | Nice-number axes, gridlines, labels, and legends |
| F-127 | S32 | M | 2 | 1 | 2026-08-10 | Exact direct and theme-mapped chart series colours |
| F-128 | S32 | S | 1 | 1 | 2026-08-10 | Source-scoped native charts and bounded preserved fallbacks |
| F-047 | S32.1 | M | 2 | 1 | 2026-08-11 | Exact oxml-layout package inventory and archive-size gate |
| F-048 | S32.1 | M | 2 | 1 | 2026-08-11 | Preparation-only stable and incubating cargo-release groups |
| F-049 | S32.1 | M | 2 | 1 | 2026-08-11 | Exact split-family publication routing and verified local-source preflight |
| F-050 | S32.1 | S | 1 | 1 | 2026-08-11 | No-default, wasm, prose, skill-sync, and binding-safe CI jobs |
| F-X005 | S32.2 | S | 1 | 1 | 2026-08-11 | Published and verified the complete rpptx 0.1.2 family |
| F-015 | S32.2 | S | 1 | 1 | 2026-08-11 | rdocx-oxml facade over published oxml-core |
| F-016 | S32.2 | S | 1 | 1 | 2026-08-11 | Shared Length re-export with unchanged callers |
| F-022 | S32.2 | S | 1 | 1 | 2026-08-11 | Deprecated rdocx-opc shim and direct shared OPC consumers |
| F-027 | S32.2 | M | 2 | 1 | 2026-08-11 | Byte-first shared media naming and metadata |
| F-028 | S32.2 | S | 1 | 1 | 2026-08-11 | Atomic automatic picture sizing at 72 DPI |
| F-046 | S32.2 | M | 2 | 1 | 2026-08-11 | Shared layout and PDF cutover with unchanged output |
| F-051 | S32.2 | S | 1 | 1 | 2026-08-11 | Root changelog and shared-crate migration guide |
| F-129 | S33 | M | 2 | 1 | 2026-08-12 | Shared Python content paths, revisions, stale errors, and Length conversion |
| F-130 | S33 | L | 4 | 1 | 2026-08-12 | Lazy rdocx Python document, paragraph, and run handles |
| F-131 | S33 | L | 4 | 1 | 2026-08-12 | Tri-state formatting and lazy table handles |
| F-132 | S33 | M | 2 | 1 | 2026-08-12 | Python-compatible units, enums, and exception mapping |
| F-133 | S33 | S | 1 | 1 | 2026-08-12 | GIL-detached rendering with concurrency and Poppler gates |
| F-134 | S34 | M | 2 | 1 | 2026-08-13 | Typed rdocx and rpptx packages with strict live stub validation |
| F-135 | S34 | M | 2 | 1 | 2026-08-13 | Pinned two-way python-docx parity suite |
| F-136 | S34 | L | 4 | 1 | 2026-08-13 | Lazy rpptx bindings with strict global stale-handle semantics |
| F-137 | S34 | M | 2 | 1 | 2026-08-13 | Two-package abi3 wheel matrix and tag-only trusted publication workflow |
| F-138 | S34 | S | 1 | 1 | 2026-08-13 | Pull-request Python binding build and test job |
| F-139 | S35 | L | 4 | 1 | 2026-08-13 | Facade-backed rdocx WASM with complete package preservation |
| F-140 | S35 | S | 1 | 1 | 2026-08-13 | Two-package locked WASM target and Node CI gates |
| F-141 | S35 | M | 2 | 1 | 2026-08-13 | Browser PDF export with embedded bundled fonts |
| F-142 | S35 | M | 2 | 1 | 2026-08-13 | Bounded facade-backed rpptx WASM profiles |
| F-143 | S36 | S | 1 | 1 | 2026-08-13 | Shared bounded range, output-path, and JSON-envelope CLI support |
| F-144 | S36 | L | 4 | 1 | 2026-08-13 | Complete deterministic rpptx command-line interface |
| F-145 | S36 | M | 2 | 1 | 2026-08-13 | Fixed-width thumbnail and recursive outline commands |
| F-146 | S36 | S | 1 | 1 | Locally packable and installable scoped WASM npm tarballs |
| F-X001 | S36 | M | 2 | 1 | Seven command-level rdocx-cli integration gates |
| F-X002 | S36 | S | 1 | 1 | Six compiled README examples and canonical rustdoc runner |
| F-X003 | S36 | S | 1 | 1 | One canonical sample generator with unchanged outputs |
| F-X004 | S36 | S | 1 | 1 | Process-unique integration-test output paths |
| F-X006 | S37 | S | 1 | 1 | Published and verified the complete rpptx 0.1.3 family |
| F-X007 | S38 | L | 4 | 1 | 2026-08-14 | Integrated PR 25, stable crate READMEs, and package-preserving numbering XML |
| F-X008 | S38 | S | 1 | 1 | 2026-08-14 | Published and verified the stable rdocx 0.5.0 family |
| F-X009 | S39 | L | 4 | 1 | 2026-08-14 | Exact README coverage and usage examples for all 26 workspace packages |
| F-X010 | S39 | S | 1 | 1 | 2026-08-14 | Published and verified the stable rdocx 0.6.0 family with rendered crate READMEs |
| F-X011 | S39 | S | 1 | 1 | 2026-08-14 | Published and verified the complete rpptx 0.2.0 family with rendered crate READMEs |
| F-X012 | S40 | M | 2 | 1 | 2026-08-15 | Pinned Poppler, Binaryen, uv, and LibreOffice CI toolchains |
| F-X013a | S41 | S | 1 | 1 | 2026-08-16 | Footnote lines advance, and break width matches the marker indent |
| F-X013b | S41 | L | 3 | 1 | 2026-08-16 | Note area reserved during pagination, notes split across pages |
| F-X013c | S41 | M | 2 | 1 | 2026-08-16 | Endnotes emitted at the document end, note streams keyed apart |
| F-X013 | S41 | M | 0 | 1 | 2026-08-16 | Umbrella closed after three child stories and integrated gates |
| F-X014 | S41 | S | 1 | 1 | 2026-08-16 | Kashida justification values accepted instead of failing the open |
| F-X015 | S41 | M | 2 | 1 | 2026-08-16 | Wrap modes, text distances and anchor alignments read into the model |
| F-X016 | S41 | L | 3 | 1 | 2026-08-16 | Alignment placement, and body text flowing around wrapping drawings |
| F-X020 | S42 | S | 1 | 1 | 2026-08-16 | Sixteen semver-compatible updates, PDF-only delta traced to font-types |
| F-X024 | S42 | M | 2 | 1 | 2026-08-16 | Theme adapter moved to rdocx-oxml, the shared-to-format edge removed |
| F-X022 | S42 | S | 1 | 1 | 2026-08-16 | Incubating train published at 0.3.0, fourteen crates |
| F-X023 | S42 | S | 1 | 1 | 2026-08-16 | Stable train published at 0.7.0, seven crates |
| F-X018 | S43 | M | 2 | 1 | 2026-08-16 | An unmodelled enumerated value reads as an absent attribute |
| F-X017 | S43 | S | 1 | 1 | 2026-08-16 | Notes broken to the width of the section carrying their reference |
| F-X019 | S43 | M | 2 | 1 | 2026-08-16 | Two-pass pagination so paragraph-relative later drawings wrap |
| F-X021 | S43 | L | 2 | 1 | 2026-08-16 | PDF fingerprint in the harness, 28 to 49 entries, and a deterministic writer |
| F-X025 | S43 | S | 1 | 1 | 2026-08-16 | The release preflights run in the local gate, not first on a tag |
| F-X026 | S44 | S | 1 | 1 | 2026-08-16 | Full release regressions run in a named CI job |
| F-X027 | S44 | S | 1 | 1 | 2026-08-16 | Golden-PNG checks run in CI under pinned Poppler |
| F-X028 | S44 | M | 2 | 1 | 2026-08-16 | Agent-facing paths, versions, fonts, and gates agree with the tree |
| F-X029 | S44 | M | 2 | 1 | 2026-08-16 | Path-filtered CI with a stable aggregate gate |
| F-156 | S45 | L | 4 | 1 | 2026-08-17 | Shared oxml-chart crate with a source-compatible rpptx-chart shim |
| F-157 | S45 | M | 2 | 1 | 2026-08-17 | Native Word chart parts with editable embedded workbooks |
| F-158 | S45 | M | 2 | 1 | 2026-08-17 | Atomic inline Document::add_chart authoring |
| F-159 | S45 | M | 2 | 1 | 2026-08-17 | Word paginator chart rendering with exact cross-family pixels |
| F-147 | S46 | M | 2 | 1 | 2026-08-17 | Typed comment part, body anchors, and relationship discovery |
| F-148 | S46 | M | 2 | 1 | 2026-08-17 | Atomic ranged comments, replies, resolution, and removal API |
| F-152 | S46 | L | 4 | 1 | 2026-08-17 | Recursive content controls at all five Word placements |
| F-153 | S46 | M | 2 | 1 | 2026-08-17 | Bounded custom XML data binding with atomic display updates |
| F-154 | S46 | M | 2 | 1 | 2026-08-17 | Typed bookmarks and single-pass REF and PAGEREF resolution |
| F-149 | S47 | L | 4 | 1 | 2026-08-17 | Typed revision metadata, content, prior properties, and ordered traversal |
| F-150 | S47 | L | 4 | 1 | 2026-08-17 | Atomic placement-aware revision acceptance and rejection |
| F-151 | S48 | M | 2 | 1 | 2026-08-17 | Accepted and tracked revision views with deterministic decorations and margin bars |
| F-155 | S48 | M | 2 | 1 | Typed document-protection intent and recorded verification metadata |
| F-160 | S49 | L | 4 | 1 | 2026-08-20 | Recursive source-preserving field instruction grammar for simple and complex fields |
| F-161 | S49 | L | 4 | 1 | 2026-08-20 | Deterministic native field evaluation with pinned Word results and cached fallback |
| F-162 | S49 | M | 2 | 1 | 2026-08-20 | Atomic field cache updates with explicit update-aware save operations |
| F-203 | S49 | M | 2 | 1 | 2026-08-20 | Namespace-aware table-property parsing with absolute schema-slot preservation |
| F-163 | S50 | L | 4 | 1 | 2026-08-21 | Atomic scalar template rendering across Word run boundaries |
| F-164 | S50 | L | 4 | 1 | 2026-08-21 | Nested loops and conditionals over body entries and table rows |
| F-165 | S50 | M | 2 | 1 | 2026-08-21 | Multi-row repetition with preserved table and numbering semantics |
| F-166 | S51 | M | 2 | 1 | 2026-08-21 | Atomic separate and sectioned mail merge with missing fields empty |
| F-167 | S51 | L | 4 | 1 | 2026-08-21 | Deterministic tracked comparison with exact accept and reject results |
| F-168 | S51 | S | 1 | 1 | 2026-08-21 | Preserved VML text and image watermarks behind every applicable page |
| F-X037 | S51 | M | 2 | 1 | 2026-08-21 | Exact Word story and Unicode scalar provenance for positioned glyph runs |
| F-X032 | S51 | S | 1 | 1 | 2026-08-21 | Complete cached and caller-font Word layout bundles for external renderers |
| F-X034 | S51 | S | 1 | 1 | 2026-08-21 | Reviewed changelog notes enforced before publication and published unchanged |
| F-X038 | S51 | L | 4 | 1 | 2026-08-21 | Bounded warm relayout caches with exact cold-layout output and provenance |
| F-X033 | S51 | S | 1 | 1 | 2026-08-21 | Ordered direct Word body items with Pedro Assumpcao's merge record preserved |
| F-X035 | S51 | S | 1 | 1 | 2026-08-21 | Published and verified the complete 15-package rpptx 0.4.0 family |
| F-X036 | S51 | S | 1 | 1 | 2026-08-22 | Published and verified the seven-package stable rdocx 0.8.0 family |
| F-169 | S52 | L | 4 | 1 | 2026-08-22 | Authenticated Agile encrypted package reading with a Word oracle |
| F-171 | S52 | L | 4 | 1 | 2026-08-22 | Exact OPC digital signature verification and coverage reports |
| F-X039 | S52 | M | 2 | 1 | 2026-08-22 | Shared layout payloads and checked reusable-engine transfer |
| F-X041 | S52 | M | 2 | 1 | 2026-08-22 | Break opportunities no longer duplicate glyph vectors |
| F-X042 | S52 | S | 1 | 1 | 2026-08-22 | Public header and footer layout proved through deterministic PDF text |
| F-170 | S52 | M | 2 | 1 | 2026-08-22 | Word-compatible failure-atomic Agile encrypted output |
| F-X040 | S52 | L | 4 | 1 | 2026-08-22 | Exact bounded pagination restart and transactional table cache |
| F-X043 | S52 | M | 2 | 1 | 2026-08-22 | Reusable bundled-fallback caller-font layouts from PRs 40 and 41 |
| F-X044 | S52 | M | 2 | 1 | 2026-08-22 | Editor-scale exact paragraph-cache lookup from PR 41 |
| F-X045 | S52 | M | 2 | 1 | 2026-08-22 | Transactional exact bounded header and footer cache from PR 41 |
| F-X046 | S52 | S | 1 | 1 | 2026-08-22 | Exact bounded substituted-page Arc reuse from PR 41 |
| F-X047 | S52 | S | 1 | 1 | 2026-08-22 | Invisible attributed caret carriers for empty Word paragraphs from PR 41 |
| F-172 | S53 | M | 2 | 1 | 2026-08-23 | Atomic RSA-SHA256 OPC signature creation with Word for Mac recognition |
| F-173 | S53 | L | 4 | 1 | 2026-08-23 | Deterministic tagged PDF structure with PDF/UA validation |
| F-174 | S53 | M | 2 | 1 | 2026-08-23 | Deterministic PDF/A-2b and PDF/A-3b conformance paths |
| F-175 | S53 | M | 2 | 1 | 2026-08-23 | Atomic exact native redaction across document and workbook XML |
| F-X048 | S53 | L | 4 | 1 | 2026-08-23 | Recursive dense-form tables with reviewed one-page geometry |
| F-X049 | S53 | S | 1 | 1 | 2026-08-23 | Published and verified the complete 15-package rpptx 0.5.0 family |
| F-X050 | S53 | S | 1 | 1 | 2026-08-23 | Published and verified the seven-package stable rdocx 0.9.0 family |
| F-176 | S54 | L | 4 | 1 | 2026-08-24 | Bounded RTF reader with Word differential evidence |
| F-177 | S54 | M | 2 | 1 | 2026-08-24 | Deterministic diagnostic RTF writer and round trip |
| F-183 | S54 | S | 1 | 1 | 2026-08-24 | Shared PNG, JPEG, TIFF, transparency, and page-range export |
| F-X051 | S54 | M | 2 | 1 | 2026-08-24 | Bounded caller-font family aliases with exact cache identity |
| F-178 | S55 | L | 4 | 1 | 2026-08-24 | Bounded HTML5 and CSS import into the Word document tree |
| F-179 | S55 | L | 4 | 1 | 2026-08-24 | Bounded ODT reader with pinned LibreOffice structural differential |
| F-X052 | S55 | L | 4 | 1 | 2026-08-24 | Exact shared-block relayout within the Issue 46 performance budget |
| F-X053 | S55 | S | 1 | 1 | 2026-08-24 | Published migration guidance and authenticated contribution closures |
| F-180 | S56 | L | 4 | 1 | 2026-08-25 | Deterministic diagnostic ODT writer with structural round trip |
| F-181 | S56 | M | 2 | 1 | 2026-08-25 | Deterministic EPUB 3 export with pinned EPUBCheck acceptance |
| F-182 | S56 | M | 2 | 1 | 2026-08-25 | Searchable self-contained SVG page export with calibrated pixel parity |
| F-X054 | S56 | L | 4 | 1 | 2026-08-25 | Hardened ordered reader and parser outcomes from PRs 47 through 52 |
| F-X056 | S56 | S | 1 | 1 | 2026-08-25 | Published and verified the complete 15-package rpptx 0.6.0 family |
| F-X057 | S56 | S | 1 | 1 | 2026-08-25 | Published and verified the complete seven-package stable 0.10.1 family |
| F-196 | S57 | M | 2 | 1 | 2026-08-25 | Five-document pinned Word corpus with strict provenance and checksum verification |
| F-201 | S57 | L | 4 | 1 | 2026-08-25 | Thousand-page deterministic layout and PDF performance gate |
| F-197 | S57 | L | 4 | 1 | 2026-08-25 | Complete-union Word SSIM harness against pinned Writer and Poppler |
| F-202 | S58 | L | 4 | 1 | 2026-08-26 | Bounded thousand-page restart pagination with exact warm and fresh equality |
| F-X061 | S58 | S | 1 | 1 | 2026-08-26 | Resumable ordinary and release dependency-prefix checkpoints |
| F-X062 | S58 | M | 2 | 1 | 2026-08-26 | Note-clean restart pagination for unchanged related stories |
| F-X063 | S58 | S | 1 | 1 | 2026-08-26 | Eliminate the redundant 22 MiB retained-font comparison |
| F-X058 | S58 | L | 4 | 1 | 2026-08-26 | Additive multilingual shaping, breaking, direction, fonts, and rich backend output |
| F-X059 | S58 | S | 1 | 1 | 2026-08-27 | Published and verified the complete 15-package rpptx 0.7.0 family |
| F-X064 | S58 | S | 1 | 1 | 2026-08-27 | Hardened PR 55 table measurement parsing with exact whole-decimal tolerance |
| F-X067 | S58 | S | 1 | 1 | 2026-08-27 | Prime the locked Word fidelity graph before the offline harness |
| F-X065 | S58 | S | 1 | 1 | 2026-08-27 | Preserve tracked table grids without changing active layout widths |
| F-X066 | S58 | S | 1 | 1 | 2026-08-28 | Classify strict namespace-aware legacy VML horizontal rules without rendering them |
| F-198 | S58 | L | 4 | 1 | 2026-08-28 | Automatic English hyphenation with exact source spans and pinned Writer parity |
| F-199 | S58 | L | 4 | 1 | 2026-08-28 | Shared rich shaping for Arabic, Devanagari, Thai, and Simplified Chinese Word text |
| F-200 | S58 | M | 2 | 1 | 2026-08-28 | Bidirectional Word layout with logical extraction, visual ordering, and retained vertical approximations |
| F-X068 | S58 | S | 1 | 1 | 2026-08-29 | Published and verified the complete 15-package rpptx 0.8.0 recovery family |
| F-X069 | S58 | S | 1 | 1 | 2026-08-29 | Published and verified the complete seven-package rdocx 0.11.1 recovery family |
| F-X070 | S58 | S | 1 | 1 | 2026-08-29 | Yanked exactly the two incomplete stable 0.11.0 registry entries after verified 0.11.1 recovery |
| F-X031 | S58 | S | 1 | 1 | 2026-08-29 | Required the aggregate CI gate on the default branch with a narrow administrator bypass and two live pull request proofs |
| F-217 | S59 | L | 4 | 1 | 2026-08-30 | Added relationship-safe modern comments, replies, sections, notes headers and footers, and handout settings |
| F-221 | S59 | M | 2 | 1 | 2026-08-30 | Added default-off native presentation encryption and signatures with current-state invalidation |
| F-213 | S60 | L | 4 | 1 | 2026-08-30 | Added typed, schema-ordered animation timing and transitions with lossless unsupported XML preservation |
| F-214 | S60 | L | 4 | 1 | 2026-08-31 | Added deterministic timeline evaluation, transition rendering, and bounded morph composition without changing static output |
| F-215 | S61 | L | 4 | 1 | 2026-08-31 | Added relationship-safe embedded and linked audio and video inspection, atomic mutation, extraction, and exact package preservation |
| F-216 | S61 | M | 2 | 1 | 2026-08-31 | Added deterministic media poster and labelled fallback rendering with synchronized playback state and unchanged legacy entry points |
| F-227 | S61 | L | 4 | 1 | 2026-08-31 | Added bounded deterministic animated GIF and Motion JPEG AVI export with exact cross-platform manifests |
| F-219 | S62 | L | 4 | 1 | 2026-09-01 | Added a bounded typed SmartArt model with atomic editing, graph-safe duplication and transfer, and exact unsupported XML preservation |
| F-218 | S62 | L | 4 | 1 | 2026-09-01 | Added relationship-owned embedded-content inventory, byte-exact extraction, and atomic OLE, ActiveX, and VBA mutation with explicit signature policies |
| F-X071 | S62 | L | 4 | 1 | 2026-09-01 | Integrated and hardened the reader outcomes from PRs 61 through 64 with contributor credit and exact namespace, schema-order, and effective-style preservation |
| F-X072 | S63 | M | 2 | 1 | 2026-09-01 | Kept direct body footnote and endnote references cacheable through complete paragraph keys and exact note-part context |
| F-X073 | S63 | L | 4 | 1 | 2026-09-01 | Restarted ordinary prose from complete block boundaries within the shared aggregate cache budget while keeping unsafe source content fail-closed |
| F-220 | S63 | L | 5 | 1 | 2026-09-01 | Rendered six pinned authentic SmartArt families through bounded typed instruction evaluation and shared DrawingML paths, with an exact guarded cycle1 compatibility profile |
| F-222 | S63 | L | 5 | 1 | 2026-09-01 | Added bounded namespace-aware ODP import and deterministic export for slides, text, rectangles, tables, images, names, and notes with explicit lossy diagnostics |
| F-223 | S63 | M | 2 | 1 | 2026-09-01 | Preserved six modern PresentationML package classes and added staged output-only conversion without changing executable payloads or relationships |
| F-226 | S63 | M | 2 | 1 | 2026-09-01 | Exported relationship-resolved notes pages and all six audience handout layouts through deterministic PDF and PNG paths |
| F-224 | S64 | L | 4 | 1 | 2026-09-02 | Imported bounded HTML and CSS into editable slide shapes, text, tables, images, and links with pinned Chrome parity |
| F-225 | S64 | L | 4 | 1 | 2026-09-02 | Imported PDF pages as preserved graphics or editable text, paths, images, and links with pinned Poppler parity |
| F-X075 | S64 | M | 2 | 1 | 2026-09-02 | Preserved complete-boundary restart pagination across page-spanning prose with exact sourced reuse and authenticated performance evidence |
| F-X074 | S64 | S | 1 | 2 | 2026-09-03 | Published and independently verified the exact 15-package rpptx 0.9.0 family with the reviewed release body and no binding publication |
| F-X076 | S64 | S | 1 | 1 | 2026-09-03 | Published and independently verified the exact seven-package rdocx 0.12.0 stable family with all seven reviewed contribution notifications |
| F-228 | S65 | L | 4 | 1 | 2026-09-03 | Added a typed, source-ordered OfficeMath model with document settings, native authoring, schema-order writing, and unsupported XML preservation |
| F-229 | S65 | M | 2 | 1 | 2026-09-03 | Measured and rendered OfficeMath through shared baseline-aware layout groups with deterministic Word and Poppler oracle geometry |
| F-230 | S65 | M | 2 | 1 | 2026-09-03 | Added bounded native MathML and LaTeX conversion with stable diagnostics and a pinned Pandoc structural differential |
| F-231 | S66 | L | 4 | 1 | 2026-09-03 | Added bounded formula, TOC, TC, mail-control, and barcode evaluation with structured native outcomes and pinned Word parity |
| F-232 | S66 | L | 4 | 1 | 2026-09-04 | Added deterministic native TOC rebuilding with final page targets, owner-aware preservation, and pinned Word parity |
| F-234 | S67 | L | 4 | 1 | 2026-09-04 | Added deterministic full-story comparison, package-wide revision resolution, source-span preservation, and pinned Word parity |
| F-233 | S67 | L | 4 | 1 | 2026-09-04 | Added atomic nested rich mail merge with exact-size images, complete fragment imports, identity remapping, and flat API compatibility |
| F-235 | S67 | M | 2 | 1 | 2026-09-04 | Added deterministic Word and Character comparison plus left-biased formatting, whitespace, field, comment, and story ignore policies |

## Velocity

Recalculated at each sprint close. The backlog assumes about 2 stories per week
sustained, and the whole plan is sized at roughly 390 developer-days. If the
first three sprints diverge from that by more than 30 percent, replan rather
than absorb it.

Stories per week is completed stories divided by actual days, multiplied by
five working days.

| Window | Stories | Days | Stories/week |
|--------|---------|------|--------------|
| S01 | 6 | 6 | 5.00 |
| S02 | 6 | 6 | 5.00 |
| S03 | 3 | 3 | 5.00 |
| S04 | 4 | 4 | 5.00 |
| S05 | 4 | 4 | 5.00 |
| S06 | 3 | 3 | 5.00 |
| S07 | 5 | 5 | 5.00 |
| S08 | 3 | 3 | 5.00 |
| S09 | 4 | 4 | 5.00 |
| S10 | 2 | 2 | 5.00 |
| S11 | 0 | 1 | 0.00 |
| S12 | 5 | 5 | 5.00 |
| S13 | 4 | 4 | 5.00 |
| S14 | 8 | 8 | 5.00 |
| S15 | 2 | 2 | 5.00 |
| S16 | 4 | 4 | 5.00 |
| S17 | 4 | 4 | 5.00 |
| S18 | 4 | 4 | 5.00 |
| S19 | 2 | 2 | 5.00 |
| S20 | 5 | 5 | 5.00 |
| S21 | 3 | 3 | 5.00 |
| S22 | 4 | 4 | 5.00 |
| S23 | 5 | 5 | 5.00 |
| S24 | 8 | 8 | 5.00 |
| S25 | 3 | 3 | 5.00 |
| S26 | 4 | 4 | 5.00 |
| S27 | 4 | 4 | 5.00 |
| S28 | 4 | 4 | 5.00 |
| S29 | 3 | 3 | 5.00 |
| S30 | 4 | 4 | 5.00 |
| S31 | 3 | 3 | 5.00 |
| S32 | 2 | 2 | 5.00 |
| S32.1 | 4 | 4 | 5.00 |
| S32.2 | 8 | 8 | 5.00 |
| S33 | 5 | 5 | 5.00 |
| S34 | 5 | 5 | 5.00 |
| S35 | 4 | 4 | 5.00 |
| S36 | 8 | 8 | 5.00 |
| S37 | 1 | 1 | 5.00 |
| S38 | 2 | 2 | 5.00 |
| S39 | 3 | 3 | 5.00 |
| S40 | 1 | 1 | 5.00 |
| S41 | 7 | 7 | 5.00 |
| S42 | 4 | 4 | 5.00 |
| S43 | 5 | 1 | 25.00 |
| S44 | 4 | 4 | 5.00 |
| S45 | 4 | 4 | 5.00 |
| S46 | 5 | 5 | 5.00 |
| S47 | 2 | 2 | 5.00 |
| S48 | 2 | 2 | 5.00 |
| S49 | 4 | 4 | 5.00 |
| S50 | 3 | 3 | 5.00 |
| S51 | 10 | 2 | 25.00 |
| S52 | 12 | 1 | 60.00 |
| S53 | 7 | 1 | 35.00 |
| S54 | 4 | 1 | 20.00 |
| S55 | 4 | 1 | 20.00 |
| S56 | 6 | 1 | 30.00 |
| S57 | 3 | 1 | 15.00 |
| S58 | 17 | 4 | 21.25 |
| S59 | 2 | 1 | 10.00 |
| S60 | 2 | 2 | 5.00 |
| S61 | 3 | 1 | 15.00 |
| S62 | 3 | 1 | 15.00 |
| S63 | 6 | 1 | 30.00 |
| S64 | 5 | 2 | 12.50 |
| S65 | 3 | 1 | 15.00 |
| S66 | 2 | 2 | 5.00 |

## Escalation record

Logged when an escalation trigger from `.claude/WORKFLOW.md` fires, with what
was done about it. Empty is the expected state.

| Date | Trigger | F-ID or sprint | Response |
|------|---------|----------------|----------|
| 2026-07-30 | Three-sprint velocity variance exceeded 30 percent | S01 to S03 | Reforecast 366 remaining estimated days to 45 to 50 active weeks, retain dependency-defined boundaries, and recalibrate after S06 |
| 2026-07-31 | Sprint estimate variance exceeded 30 percent | S05 | Record 4 actual days against 8 estimated, retain the 45 to 50 active week reforecast, and recalibrate after S06 |
| 2026-07-31 | Sprint estimate variance exceeded 30 percent | S06 | Reforecast 124 remaining stories at the observed five stories per active week to about 25 active weeks, while retaining dependency-defined sprint boundaries |
| 2026-07-31 | Sprint estimate variance exceeded 30 percent | S07 | Record 5 actual days against 8 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-07-31 | Sprint estimate variance exceeded 30 percent | S08 | Record 3 actual days against 7 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-07-31 | Sprint estimate variance exceeded 30 percent | S09 | Record 4 actual days against 7 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-07-31 | Sprint estimate variance exceeded 30 percent | S10 | Record 2 actual days against 8 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-07-31 | Sprint estimate variance exceeded 30 percent | S12 | Record 5 actual days against 11 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-01 | Sprint estimate variance exceeded 30 percent | S13 | Record 4 actual days against 12 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-01 | Sprint estimate variance exceeded 30 percent | S14 | Record 8 actual days against 14 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-01 | Sprint estimate variance exceeded 30 percent | S15 | Record 2 actual days against 5 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-01 | Sprint estimate variance exceeded 30 percent | S16 | Record 4 actual days against 12 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-01 | Sprint estimate variance exceeded 30 percent | S17 | Record 4 actual days against 10 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-02 | Sprint estimate variance exceeded 30 percent | S18 | Record 4 actual days against 7 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-02 | Sprint estimate variance exceeded 30 percent | S19 | Record 2 actual days against 6 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-02 | Sprint estimate variance exceeded 30 percent | S20 | Record 5 actual days against 11 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-02 | Sprint estimate variance exceeded 30 percent | S21 | Record 3 actual days against 8 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-03 | Sprint estimate variance exceeded 30 percent | S22 | Record 4 actual days against 9 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-03 | Sprint estimate variance exceeded 30 percent | S23 | Record 5 actual days against 10 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-05 | Sprint estimate variance exceeded 30 percent | S24 | Record 8 actual days against 14 estimated and retain the about 25 active week reforecast with dependency-defined sprint boundaries |
| 2026-08-08 | Sprint estimate variance exceeded 30 percent | S25 | Record 3 actual days against 10 estimated and reforecast 57 pending stories at the observed five stories per active week to about 12 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-08 | Sprint estimate variance exceeded 30 percent | S26 | Record 4 actual days against 10 estimated and reforecast 53 pending stories at the observed five stories per active week to about 11 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-08 | Sprint estimate variance exceeded 30 percent | S27 | Record 4 actual days against 12 estimated and reforecast 49 pending stories at the observed five stories per active week to about 10 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-09 | Sprint estimate variance exceeded 30 percent | S28 | Record 4 actual days against 9 estimated and reforecast 45 pending stories at the observed five stories per active week to about 9 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-10 | Sprint estimate variance exceeded 30 percent | S29 | Record 3 actual days against 12 estimated and reforecast 42 pending stories at the observed five stories per active week to about 9 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-10 | Sprint estimate variance exceeded 30 percent | S30 | Record 4 actual days against 12 estimated and reforecast 38 pending stories at the observed five stories per active week to about 8 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-10 | Sprint estimate variance exceeded 30 percent | S31 | Record 3 actual days against 12 estimated and reforecast 35 pending stories at the observed five stories per active week to about 7 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-11 | Sprint estimate variance exceeded 30 percent | S32 | Record 2 actual days against 3 estimated and reforecast 33 pending stories at the observed five stories per active week to about 7 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-11 | Sprint estimate variance exceeded 30 percent | S32.1 | Record 4 actual days against 7 estimated and reforecast 29 pending stories at the observed five stories per active week to about 6 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-12 | Sprint estimate variance exceeded 30 percent | S33 | Record 5 actual days against 13 estimated and reforecast 17 pending stories at the observed five stories per active week to about 4 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-13 | Sprint estimate variance exceeded 30 percent | S34 | Record 5 actual days against 11 estimated and reforecast 12 pending stories at the observed five stories per active week to about 3 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-13 | Sprint estimate variance exceeded 30 percent | S35 | Record 4 actual days against 9 estimated and reforecast 8 pending stories at the observed five stories per active week to about 2 active weeks, retaining dependency-defined sprint boundaries |
| 2026-08-13 | Sprint estimate variance exceeded 30 percent | S36 | Record 8 actual days against 13 estimated, retain one pending fresh-version release story, and reforecast it as the remaining active sprint |
| 2026-08-14 | Sprint estimate variance exceeded 30 percent | S38 | Record 2 actual days against 5 estimated. The externally contributed implementation reduced authoring work, while eighteen review passes established the required package-preserving numbering boundary before release |
| 2026-08-14 | Sprint estimate variance exceeded 30 percent | S39 | Record 3 actual days against 6 estimated. Documentation automation and the established release workflow reduced delivery time while preserving exact package, registry, and rendered README verification |
| 2026-08-15 | Sprint estimate variance exceeded 30 percent | S40 | Record 1 actual day against 2 estimated. Reusing the bounded installer across hosted jobs reduced delivery time while preserving the exact pinned-tool and full-workflow evidence |
| 2026-08-16 | Sprint estimate variance exceeded 30 percent | S41 | Record 7 actual days against 12 estimated. Splitting F-X013 into three children at design time meant each arrived with its defect already isolated, and five of the six stories shared one note or drawing subsystem, so later stories reused the first one's investigation rather than repeating it |
| 2026-08-16 | Sprint estimate variance exceeded 30 percent | S43 | Record 1 actual day against 8 estimated. Every story arrived with its defect already isolated by the sprint that filed it, so design read a written-up cause rather than searching for one. Retain the dependency-defined boundaries. The 25.00 stories per week this produces is not a sustainable rate and is not carried into any forecast |
| 2026-08-16 | Sprint estimate variance exceeded 30 percent | S44 | Record 4 actual days against 6 estimated. Three stories reused the existing workflow-contract module, and two reused already pinned CI environments, which reduced implementation time while preserving the dependency-defined boundaries and full close gate |
| 2026-08-17 | Sprint estimate variance exceeded 30 percent | S45 | Record 4 actual days against 10 estimated. The format-neutral chart engine already existed, and the strict dependency chain reused one shared authoring path and one golden fixture across the later stories. Retain the milestone-defined boundaries |
| 2026-08-17 | Sprint estimate variance exceeded 30 percent | S46 | Record 5 actual days against 12 estimated. Three independent roots and two dependent stories reused one typed paragraph, package, and traversal foundation across isolated workers. Retain the dependency-defined boundaries |
| 2026-08-17 | Sprint estimate variance exceeded 30 percent | S47 | Record 2 actual days against 8 estimated. F-150 reused F-149's typed revision model, and the extended sprint review concentrated namespace and ordering investigation into one integrated path. Retain the milestone-defined S48 boundary |
| 2026-08-17 | Sprint estimate variance exceeded 30 percent | S48 | Record 2 actual days against 4 estimated. The two stories used isolated workers, while the integrated review concentrated the milestone interaction proof into one mixed-package gate. Retain the dependency-defined S49 boundary |
| 2026-08-20 | Sprint estimate variance exceeded 30 percent | S49 | Record 4 actual days against 12 estimated. The three dependent field stories reused one recursive source-preserving model, while the independent reader correction and bounded review loop concentrated preservation investigation into the same integrated gate. Retain the dependency-defined S50 boundary |
| 2026-08-21 | Sprint estimate variance exceeded 30 percent | S50 | Record 3 actual days against 10 estimated. The strict feature chain reused one template evaluator across all three stories, while the bounded sprint review concentrated four interaction corrections into the same row and scope ownership path. Retain the dependency-defined S51 boundary |
| 2026-08-22 | Sprint estimate variance exceeded 30 percent | S51 | Record 2 actual days against 18 estimated. Nine isolated workers, reviewed community contributions, and two established release workflows allowed independent implementation and publication work to overlap safely. The resulting 25.00 stories per week is not a sustainable forecast, so retain dependency-defined future sprint boundaries |
| 2026-08-22 | Sprint estimate variance exceeded 30 percent | S52 | Record 1 actual day against 27 estimated. Twelve isolated workers and seven scoped follow-ups from PRs 40 and 41 allowed security, rendering, and cache work to overlap safely. The resulting 60.00 stories per week is not a sustainable forecast, so retain dependency-defined S53 boundaries |
| 2026-08-23 | Sprint estimate variance exceeded 30 percent | S53 | Record 1 actual day against 16 estimated. Seven isolated workers and established security, PDF, and release workflows allowed independent work to overlap safely. The resulting 35.00 stories per week is not a sustainable forecast, so retain the dependency-defined S54 boundary |
| 2026-08-24 | Sprint estimate variance exceeded 30 percent | S54 | Record 1 actual day against 9 estimated. Four isolated workers and established RTF, raster, and layout paths allowed independent work to overlap safely. The resulting 20.00 stories per week is not a sustainable forecast, so retain the dependency-defined S55 boundary |
| 2026-08-24 | Sprint estimate variance exceeded 30 percent | S55 | Record 1 actual day against 13 estimated. Four isolated workers and established import, layout, differential, and contribution workflows allowed independent work to overlap safely. The resulting 20.00 stories per week is not a sustainable forecast, so retain the dependency-defined S56 boundary |
| 2026-08-25 | Sprint estimate variance exceeded 30 percent | S56 | Record 1 actual day against 15 estimated. Isolated workers, established exporter and release workflows, and the v0.10.1 recovery allowed implementation, review, and publication work to overlap safely. The resulting 30.00 stories per week is not a sustainable forecast, so retain the dependency-defined future release checkpoints starting at S58 |
| 2026-08-25 | Sprint estimate variance exceeded 30 percent | S57 | Record 1 actual day against 10 estimated. Isolated workers, established corpus and oracle infrastructure, and test-only reuse of existing render paths reduced delivery time while preserving the dependency-defined S58 shaping boundary. The resulting 15.00 stories per week is not a sustainable forecast, so retain the planned S58 scope |
| 2026-08-29 | Sprint estimate variance exceeded 30 percent | S58 | Record 4 actual days against 32 estimated. Parallel isolated workers, staged release checkpoints, established corpus oracles, and contribution hardening allowed independent work to overlap safely. The resulting 21.25 stories per week is not a sustainable forecast, so retain dependency-defined future sprint boundaries |
| 2026-08-30 | Sprint estimate variance exceeded 30 percent | S59 | Record 1 actual day against 6 estimated. Two isolated workers reused the established PresentationML preservation model and shared package-security implementation, while the bounded reviews concentrated interaction corrections into the same integrated facade. The resulting 10.00 stories per week is not a sustainable rate, so retain the dependency-defined S60 boundary |
| 2026-08-31 | Sprint estimate variance exceeded 30 percent | S60 | Record 2 actual days against 8 estimated. The strict dependency let F-214 reuse the completed F-213 model, while the user-supplied PowerPoint export and bounded review loop concentrated oracle calibration in one integrated path. The resulting 5.00 stories per week matches the long-run delivery velocity, so retain the dependency-defined S61 boundary |
| 2026-08-31 | Sprint estimate variance exceeded 30 percent | S61 | Record 1 actual day against 10 estimated. Three strict dependency waves reused the completed timing, package, and rendering foundations, while the bounded reviews concentrated interaction corrections in one integrated facade and encoder path. The resulting 15.00 stories per week is not a sustainable forecast, so retain the dependency-defined S62 boundary |
| 2026-09-01 | Sprint estimate variance exceeded 30 percent | S62 | Record 1 actual day against 16 estimated. Three isolated workers and the contribution-integration path allowed completed work to overlap, while F-220 reached the ten-pass review bound and carried with its worker intact. The resulting 15.00 stories per week is not a sustainable forecast, so retain the dependency-defined S63 boundary and complete F-220 before F-222 |
| 2026-09-01 | Sprint estimate variance exceeded 30 percent | S63 | Record 1 actual day against 20 estimated. Retained F-220 evidence, established interchange and rendering paths, and tightly scoped cache fixes reduced repeated investigation while the bounded reviews preserved exact behavior. The resulting 30.00 stories per week is not a sustainable forecast, so retain milestone-defined future sprint boundaries |
| 2026-09-03 | Sprint estimate variance exceeded 30 percent | S64 | Record 2 actual days against 12 estimated. Isolated import and release worktrees, established external oracles, and retained milestone evidence allowed implementation, publication, and review to overlap safely. The resulting 12.50 stories per week is not a sustainable forecast, so retain the dependency-defined S65 boundary |
| 2026-09-03 | Sprint estimate variance exceeded 30 percent | S65 | Record 1 actual day against 8 estimated. Isolated workers, the typed equation foundation, and established Word, Poppler, and Pandoc oracles allowed layout and conversion work to overlap after F-228. The resulting 15.00 stories per week is not a sustainable forecast, so retain the dependency-defined S66 boundary |
| 2026-09-04 | Sprint estimate variance exceeded 30 percent | S66 | Record 2 actual days against 8 estimated. The strict dependency let F-232 reuse F-231's completed recursive field grammar, while the established Word differential and deterministic layout boundaries concentrated preservation and pagination corrections in one reviewed path. The resulting 5.00 stories per week matches the long-run delivery velocity, so retain the dependency-defined S67 boundary |
