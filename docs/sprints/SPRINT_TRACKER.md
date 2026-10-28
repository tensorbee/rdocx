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
| S67 | M22 | 3 | 3 | 0 | 10 | 1 | Added advanced rich mail merge, full-story document comparison, and deterministic character, word, and ignore policies. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, prompt-free Word 16.104 parity, and global sprint review pass 2 clean. M22 continues in S68 with embedded content, forms, and building blocks |
| S68 | M22, X | 2 | 2 | 0 | 8 | 1 | Added relationship-owned embedded-content and macro inventory, typed legacy forms, and glossary and building-block editing with atomic source preservation. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and sprint review pass 2 clean after normalized Pack URI handling was corrected and shared XML validation was filed as F-X077. M22 continues in S69 with modern package and web variants |
| S69 | M22, X | 8 | 7 | 1 | 11 | 2 | Completed M22 with shared strict XML validation, Flat OPC and modern Word package variants, bounded MHTML interchange, and restored hosted CI release readiness. Published and verified the shared 0.10.0 and 0.11.0 families and the stable 0.13.1 recovery. F-X078 records the immutable partial v0.13.0 attempt. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and sprint review pass 23 clean |
| S70 | M23, X | 7 | 7 | 0 | 19 | 2 | Established the modern DOCX capability matrix and private-corpus policy, added public from-scratch authoring conformance, replaced volatile README claims with evidence-backed coverage, and closed the three Issue 69 cache and restart gaps. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and sprint review pass 5 clean. M23 continues in S71 with public API completion |
| S71 | M23, X | 10 | 10 | 0 | 35 | 3 | Added complete blank-package, property, theme, font, style, numbering, and deterministic identifier foundations. Folded Kevin Brown's authored Word charts into the sprint, closed Issue 69 with measured evidence and reporter credit, and made the root and all 26 crate READMEs capability-led. The full close gate passed with all 49 reviewed hashes, every package archive below 10 MiB, and final sprint review pass 9 clean. M23 continues in S72 with sections, stories, and content ownership |
| S72 | M23, X | 19 | 19 | 0 | 61 | 7 | Added ordered sections, relationship-safe story editing, transactional fragment import, searchable PDF text, CLI and Python surfaces, hardened outcomes from PRs 77 through 80, and deterministic hosted CI. Published metadata-complete `rdocx 0.13.2` and `rpptx 0.11.0` distributions with reporter and contributor credit. The full close gate passed with all 49 reviewed hashes, the declared seven-entry Word XML delta, every package archive below 10 MiB, and final sprint review pass 24 clean. M23 continues in S73 with tables, rich content, and the private corpus gate |
| S73 | M23, X | 35 | 35 | 0 | 92 | 4 | Closed M23 with complete table and rich-content authoring, layout-backed fields, pure Rust generation of all five private corpus examples, contributor hardening, and verified Rust and Python publication. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and final sprint review pass 9 clean. M24 begins in S74 with complete text, table, section, and settings semantics |
| S74 | M24 | 22 | 22 | 0 | 66 | 3 | Opened M24 with complete paragraph, run, multilingual typography, table, section, settings, and web-settings authoring, plus evidence-backed package documentation. The full close gate passed with all 49 hashes unchanged, every package archive below 10 MiB, and final sprint review pass 3 clean. The predecessor main CI timing failure passed in the exact isolated S74 binding suite and ten consecutive reruns. M24 continues in S75 with rich stories, notes, ranges, fragments, and building blocks |
| S75 | M24, X | 11 | 3 | 8 | 47 | 8 | Restored the hosted Python story gate, integrated PRs 146 through 151, reconciled overlapping PRs 152 through 154, fixed the uncovered issues, and corrected table row breaks. F-X133 and F-271 through F-277 carry to S76 at the approved mid-milestone cutoff, preserving their dependency order. The full close gate passed with 49 matching hashes after PR 148's separately reviewed 15-entry layout delta, seven matching golden buffers, all archives below 10 MiB, and sprint review pass 1 clean. M24 remains open |
| S76 | X | 3 | 3 | 0 | 15 | 2 | Integrated 19 reviewed contributor PRs in three dependency waves covering package and CLI safety, Word story traversal, producer identity, and comparison. The full close gate passed with 49 unchanged hashes, 77 Python tests, all archives below 10 MiB, and a clean final sprint review. The broader issue acceptance matrix remains planned for S78 and S79 |
| S77 | X | 4 | 4 | 0 | 18 | 2 | Integrated 21 reviewed contributor PRs across rendering, layout, Python bindings, Presentation and CLI stories. Two separately reviewed baseline updates ended with all 49 hashes matching. The full close gate, 121 Word and 63 Presentation Python tests, pinned fidelity completeness, 22-crate dry run and clean sprint review passed. Broader issue matrices remain scheduled through S88 |
| S78 | X | 3 | 3 | 0 | 15 | 1 | Integrated six reviewed Word PR behaviors, completed the Issue 157 and 159 identity and picture gates and the Issue 161 redline gate, and retained Issue 160 follow-up scope in F-X161. The full workspace gate, 161 Word Python tests, 49 unchanged hashes, 22-crate dry run and sprint review pass 1 passed. No stories carried |
| S79 | X | 3 | 3 | 0 | 11 | 1 | Integrated five reviewed contribution PRs across Word line geometry, full-story CLI diff and Python comment, numbering and run operations. Issue 162 and Issue 227 acceptance passed, while Issue 226's plain-line symptom remains in F-X163 and Issue 168's remaining checklist remains in F-X147 and F-X160. The full workspace gate, 164 Word Python tests, 49 matching hashes, seven golden buffers, 22-crate dry run and sprint review pass 1 passed. No stories carried |
| S80 | X | 3 | 3 | 0 | 13 | 2 | Completed the Issue 168 Word Python chain with an accepted lxml fallback, repaired Issues 215 and 216, and integrated five Presentation drawing PRs. The full gate passed with 165 Word and 71 Presentation Python tests, pinned LibreOffice drawing evidence, 49 unchanged hashes, all 22 archives below 10 MiB and sprint review pass 1 clean. PowerPoint 16.104 remained unverified under the accepted scope decision. Issue 217 continues in S81. No stories carried |
| S81 | X | 3 | 3 | 0 | 15 | 1 | Integrated slide import, scoped replacement and built-in table styles, then proved the complete Issue 169 and six-item Issue 217 Python deck chains. The full gate passed with 75 Presentation Python tests, pinned python-pptx and LibreOffice evidence, 49 unchanged hashes, all 22 archives below 10 MiB and sprint review pass 1 clean. Reporter fixture acceptance and issue closure evidence continue in S82. No stories carried |
| S82 | X | 3 | 3 | 0 | 13 | 1 | Exercised the original Word report and Presentation deck, completed both matrices and workflows, repaired image comparison and reconciled the original 22 issue criteria. The full gate passed with 168 Word and 76 Presentation Python tests, 49 unchanged hashes, all 22 archives below 10 MiB and sprint review pass 1 clean. Issues 158, 160 and 226 retain follow-up criteria. No stories carried |
| S83 | X | 2 | 2 | 0 | 8 | 1 | Corrected Word package validity, accepted producer style, drawing and decimal measurement forms, and kept F-X160 output stable after F-X159's three declared hash changes. The full gate passed with 169 Word Python tests, 49 matching hashes, all 22 archives below 10 MiB and sprint review pass 1 clean. No stories carried |
| S84 | X | 10 | 10 | 0 | 38 | 2 | Combined the S84 through S89 repair scope in one sprint, including PR 265's mixed-section regression. All nine issue contracts and 32 contribution dispositions have integrated evidence. The full gate passed with 175 Word and 77 Presentation Python tests, 49 matching hashes, 22 package archives below 10 MiB and sprint review pass 1 clean. GitHub closure follows the reviewed main push. No stories carried |
| S85 | X | 1 | 1 | 0 | 1 | 1 | Pinned LibreOffice 26.2.5.2 for the macOS Presentation Python CI cell after the S84 main run exposed the missing viewer. The full integrated local gate passed, including 77 Presentation Python tests, 133 workflow checks, 49 matching hashes, both WASM targets, 22 package archives below 10 MiB and a clear supply-chain scan. Hosted CI and the 32 PR and nine issue dispositions follow the reviewed main push. No stories carried |
| S86 | M24, X | 7 | 7 | 0 | 27 | 2 | Completed rich headers and footers, footnotes and endnotes, namespace preservation, and both family release preparations. The full gate passed with 49 matching hashes, 22 package archives below 10 MiB and clean sprint review pass 2. Hosted build-only run 37165122971 produced all 14 expected artifacts after one targeted Linux arm64 retry. Issue 266 stays open until the separately approved rpptx-v0.13.0 and v0.15.0 releases are published, verified and reported to the requester. No contributed PR was assigned and no stories carried |
| S87 | X | 1 | 1 | 0 | 1 | 1 | Refreshed both CLI archive measurements after S86 Windows stack hardening and moved the planned M24 wave to S88 by approval. The full gate passed with 49 matching hashes, 139 policy tests, 22 package archives below 10 MiB and clean sprint review pass 1. Hosted CI run 37178559664 passed, including Docs and Release regressions. Build-only run 37178579418 produced all 14 expected artifacts after one targeted Linux x86-64 retry. Issue 266 stays open until both separately approved releases are published, verified and reported to the requester. No stories carried |
| S88 | X | 1 | 1 | 0 | 3 | 1 | Repaired the unified release inventory and prepared PowerPoint 0.13.1 beside Word 0.15.0 after the immutable failed PowerPoint tag. The full gate passed with 49 matching hashes, 140 policy tests with two expected skips, exact 22-crate dry run and clean sprint review pass 1. Hosted CI run 37192110554 passed all jobs, including Docs, Release regressions and fidelity. Build-only run 37192115737 passed all 14 wheel and source distribution jobs, and both exact seven-file artifact contracts passed. Issue 266 remains open until both separately approved releases are published and verified. No stories carried |
| S89 | M24, X | 5 | 5 | 0 | 21 | 2 | Completed note policy, cross-story ranges, deterministic fragment remapping and public glossary lifecycle, plus Changjoon's PR 269 fontdb feature-unification fix. Full verification passed with 49 unchanged hashes, 140 policy tests with two expected skips, 22 verified packages below 10 MiB and clean sprint review pass 1. M24 continues through S94. No stories carried |
| S90 | M24, X | 20 | 18 | 2 | 55 | 3 | Completed 18 stories. The reconstructed 55-day estimate comprises 50 estimated days recorded for completed features plus F-X178 at 1 day and the chosen 4-day lower bound for F-283. Actual days are the three elapsed working dates from 2026-10-07 through 2026-10-09 inclusive, not person-days. F-X178 and F-283 carry to S91. F-X178 needs a separate scope decision and Issue264 remains excluded. New F-X192 owns the remaining bibliography catalogue in S91, and Issue281 remains open. Word0.16.0 and shared/PowerPoint0.14.0 are prepared for separately approved publication after close |
| S91 | M24, X | 39 | 29 | 10 | 130 | 7 | Completed 29 top-level stories and four F-X211 delivery children. Child records are not counted twice. The reconstructed 130-day lower bound comprises 97 days for completed scope and 33 for carries, using S1/M2/L4. Actual days are seven worked dates, 2026-10-09 through 2026-10-15, including weekend work, not person-days. The frozen scope contains 40 issues and 30 PRs. Discussions264 and281 stay open. Bibliography, navigation and standalone authoring carry to S92. Local full verification and sprint review pass3 are clean. The exact reviewed-source build-only rehearsal38659673259 passed all27 build jobs, both seven-file Python sets and twelve CLI archives. Fresh hosted-wheel installs passed78 priority cases on each of Python3.9 and3.12 plus strict typing and stub checks. Word0.16.1 and shared/PowerPoint0.14.1 are prepared for publication after close |
| S91.1 | M24, X | 4 | 4 | 0 | 10 | 1 | Four corrective stories completed with clean integrated review and actual final reviewed-commit full verification. Corrected primary CI38708235000 passed, separate owner CI38702668217 qualifies95/100 with42 hard gates,51 native clauses and five no-credit cases. All49 product hashes unchanged. Actual days count the single worked date2026-10-15, not person-effort. Exact-main verification, both prepared publications and acceptance-aware contributor reconciliation follow close. Future planning is preserved separately |
| S92 | M24, X | 6 | 6 | 0 | 15 to 18 | 1 | Six repair stories completed, including explicit Issue378 intake. Actual days count the single worked date2026-10-15, not person-effort. Final full verification at133ce33b passed all14 checks,49 unchanged hashes and22 verified archives. Independent review pass2 is clean. User accepted240-cell fill3.304061s with strict open/edit and linear-work gates retained. Future parity/forms planning is included without delivery credit. M24 continues to S99. No release prepared |
| S93 | M24, X | 7 | 7 | 0 | 21 to 27 | 2 | Seven issue-repair contracts completed. Actual days count worked dates2026-10-15and2026-10-16, not person-effort. Fresh final full verification at55122db6 passes15checks,49hashes,7pixel checks and22verified archives. Native and1666installed Python members have explicit unchanged-input successor authentication. Reviewed18net unique hash changes and4pixel buffers are intentional. Independent review pass2 is clean after the HLD SVG wording correction. Issue365 stays open for pending F-X247/S94. M24 continues to S99. No release prepared |
| S94 | M24, X | 5 | 5 | 0 | 14 to 19 | 1 | Five Python, CLI and finite conversion contracts completed. Actual days count the single worked date 2026-10-16, not person-effort. Independent sprint review pass2 has 0 blocking and 0 should-fix findings, with one optional formatting nit. Fresh reviewed-HEAD full verification at 588f2640 passed all 18 native/repository commands, 49 unchanged hashes, seven deterministic pixels and 22 verified archives, with authenticated installed, hosted and finite oracle evidence. Hosted acceptance remains the authenticated build-only 27-job run at 9cdbea00, with four publication guards skipped. M24 continues to S99. No release prepared |
| S95 | M24, X | 9 | 9 | 0 | 22 to 31 | 1 | Nine contracts completed, including the user-approved Issue386 expansion and Word0.16.2/shared-PowerPoint0.14.2 release preparation. Actual days count the single worked date2026-10-16, not person-effort. Actual6b source has1908 installed tests with no skips, hosted27success4publicationSKIP/26payloads/172 runtime calls, public23, PPT four routes, Issue38619+16 and pinned reporter12strictXPASS/raw1 plus12bodies/raw0. Integrated native18commands,49hashes,7pixels and22verified archives pass at6b7c9798. Formal sprint review remains due. S95.11 scope remains future backlog, not S95 carries. Main merge, sprint close and separately approved publication have not occurred |

| S95.11 | M24, X | 4 | 3 | 1 | 12 to 16 | 2 | F-X250/F-X254 repairs and F-X223 bounded native investigation completed. User accepted documenting Word's native limit, broader genuine multi-face use remains unqualified and fourteen observer refusals remain unsupported. F-X219 full-range opacity remains unfinished and is carried to S96, which has not started. Retained three-story estimate10 to13 developer-days. Actual source-bound verification at 05ba6cc8461a896fa751ca7dc1fe209a8fa70045 joins H3/H5/836/H7 observations, with6616 unique native passes/73ignored,11doctests,955Word/1007PPT installed plus typing/stubs,27 distinct Issue385 bodies,49hashes/7pixels and22verified package union. Actual developer effort not recorded. Close actuals count worked dates2026-10-16and17. No release or new version |

| S95.2 | M24, X | 3 | 2 | 1 | 12 to 15 | 2 | F-283 numbering-aware navigation and F-284 versioned template grammar completed. F-X192 stays unfinished and is carried to S96, which has not started. Retained two-story estimate is 8 to 10 developer-days. Original three-story estimate was 12 to 15. Actual days are the worked dates 2026-10-17 and 2026-10-18, not person-effort. Full verification at b240cb3509a9641fe265ebfc605c987a3f35e143 passed with 49 unchanged hashes, seven pixel cases, 22 package inventories under 10 MiB, both WASM targets, installed Python typing and stubs, and pinned Word 16.114 plus LibreOffice 26.2.5.2 oracles. Actual developer effort not recorded. Issues 264 and 281 stay open. No release prepared. M24 continues through S99. |

| S95.3 | M24, X | 8 | 8 | 0 | 34 to 49 | 3 | Eight contracts complete, no carries. Actual days are worked dates 2026-10-18, 19 and 22, not person-effort. Current source-qualified native 6,734/73 ignored, installed Word 1,066 and applicable PPT 1,028, 49 hashes/seven pixels, 22 verified archives and current hosted 39830343186 with 27 successes/four publication skips, 26 payloads, 172 priorities and six matrix calls pass. Historical failures remain. Final review and closure outcomes follow the actual reports and sprint tag. M24 continues to S99. No publication |
| S95.4 | M24, X | 9 | 9 | 0 | see scope notes | 6 | Nine contracts complete, no carries. Six worked dates 2026-10-23 through 28, not person-effort. Actual full source c49 with explicit old-producer applicability, all 79 Word gates and memory, PowerPoint/native qualification, 22 verified archives/49 hashes and hosted 40608035966 with 27 success/four publication skips/26 payloads/172 priorities pass. Fresh 0.16.4/0.14.4 prepared, publication held. Final review/close follow actual records |

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
| F-236 | S68 | L | 4 | 1 | 2026-09-05 | Added relationship-owned OLE, ActiveX, and VBA inventory with exact hashes, byte-preserving mutation, and explicit signature policy |
| F-237 | S68 | L | 4 | 1 | 2026-09-05 | Added typed legacy form editing and relationship-owned glossary, AutoText, and building-block inventory and replacement |
| F-X077 | S69 | M | 2 | 1 | 2026-09-05 | Consolidated strict XML 1.0 lexical validation in oxml-core while preserving glossary, embedded, and package-story error and rollback contracts |
| F-239 | S69 | M | 2 | 1 | 2026-09-05 | Added bounded MHTML import and deterministic export with contained resources, stable diagnostics, and pinned Word parity |
| F-X080 | S69 | S | 1 | 1 | 2026-09-05 | Restored release-ready CI package inventory, bounded authenticated Pandoc extraction, and exhaustive Python error mapping |
| F-X079 | S69 | S | 1 | 1 | 2026-09-05 | Published and independently verified the exact 15-package rpptx 0.10.0 family with no external contribution notifications |
| F-238 | S69 | M | 2 | 1 | 2026-09-05 | Added strict Flat OPC import and export plus source-preserving DOCX, DOCM, DOTX, and DOTM package identity and conversion |
| F-X081 | S69 | S | 1 | 1 | 2026-09-06 | Published and independently verified the exact 15-package rpptx 0.11.0 recovery family required by stable package-class handling |
| F-X082 | S69 | S | 1 | 1 | 2026-09-06 | Published and independently verified the exact seven-package rdocx 0.13.1 stable recovery family against shared 0.11.0 |
| F-240 | S70 | L | 4 | 1 | 2026-09-06 | Closed the 85-row modern DOCX capability matrix, anonymous private-corpus requirements, and M23 and M24 roadmap ownership |
| F-X083 | S70 | S | 1 | 1 | 2026-09-06 | Closed confirmed Issue 67 with release evidence and mapped the three Issue 69 mechanisms to independent credited stories |
| F-241 | S70 | L | 4 | 1 | 2026-09-07 | Added public-only from-scratch DOCX conformance and ignored required-private package, modeled, privacy, and visual evidence |
| F-X084 | S70 | M | 2 | 1 | 2026-09-07 | Preserved ordinary paragraph cache reuse across note-only changes while retaining exact gates for note-bearing and related caches |
| F-242 | S70 | M | 2 | 1 | 2026-09-07 | Replaced volatile root README claims with matrix-backed capabilities, official comparison evidence, and metadata-derived validation |
| F-X085 | S70 | M | 2 | 1 | 2026-09-07 | Shared one lazy exact-identity memo across restart scans and publication with bounded transient-memory evidence |
| F-X086 | S70 | L | 4 | 1 | 2026-09-07 | Reused safe sourced prefix checkpoints across body-length changes while rebuilding shifted tails for exact provenance |
| F-243 | S71 | L | 4 | 1 | 2026-09-07 | Added deterministic Word-compatible DOCX, DOCM, DOTX, and DOTM creation profiles with complete owned package graphs and an explicit compact profile |
| F-249 | S71 | M | 2 | 1 | 2026-09-07 | Centralized deterministic Word package identifiers with collision-safe staged allocation and case-insensitive OPC identity validation |
| F-245 | S71 | L | 4 | 2 | 2026-09-08 | Added public theme and font-table authoring, licensed embedded-font ownership, and deterministic authored-font layout resolution |
| F-244 | S71 | L | 4 | 2 | 2026-09-08 | Added atomic settings plus core, application, and custom property authoring with deterministic removal and raw XML preservation |
| F-246 | S71 | L | 4 | 2 | 2026-09-08 | Added atomic paragraph, character, and table style graphs with effective resolution and pinned Word formatting parity |
| F-247 | S71 | L | 4 | 1 | 2026-09-09 | Added complete numbering definitions, levels, instances, and overrides with atomic graph validation, exact XML preservation, and non-coercing export behavior |
| F-248 | S71 | L | 4 | 1 | 2026-09-09 | Added atomic style-linked numbering plus result-local counters shared by body, table, TOC, and REF output |
| F-X087 | S71 | L | 4 | 1 | 2026-09-09 | Folded Kevin Brown's PR 71 into portable editable Word charts with atomic theme ownership and verified Word and Pages interoperability |
| F-X088 | S71 | S | 1 | 1 | 2026-09-09 | Verified all Issue 69 fixes with six regressions and 21-sample timing evidence, credited the reporter, and closed the issue |
| F-X089 | S71 | L | 4 | 1 | 2026-09-09 | Reframed the root and all 26 crate READMEs around implemented outcomes, checked examples, and bounded official comparisons |
| F-253 | S72 | L | 4 | 2 | 2026-09-10 | Added one deterministic story identity, traversal, and staged text-mutation model across body, cell, related, note, comment, and text-box content |
| F-250 | S72 | L | 4 | 1 | 2026-09-10 | Added ordered mutable section handles, atomic boundary edits, effective story inheritance, and safe orphan pruning |
| F-251 | S72 | L | 4 | 1 | 2026-09-10 | Added checked section geometry, source-preserving section XML mutation, and restart-aware displayed page numbering |
| F-254 | S72 | L | 4 | 1 | 2026-09-10 | Added transactional generic content insertion, removal, cloning, and same-owner movement with exact XML preservation |
| F-255 | S72 | M | 2 | 1 | 2026-09-10 | Added story-scoped picture, hyperlink, lookup, and validation operations with part-local relationship ownership and deterministic global drawing allocation |
| F-X090 | S72 | S | 1 | 1 | 2026-09-13 | Accepted reused producer drawing identities across physical XML parts while retaining package-global authored allocation |
| F-X091 | S72 | M | 2 | 1 | 2026-09-13 | Allowed safe serialization of unused root defaults and added atomic fallible text replacement for the CLI |
| F-252 | S72 | L | 4 | 1 | 2026-09-13 | Added transactional rich default, first, and even header and footer lifecycle operations for every section with exact Word selection parity |
| F-X093 | S72 | M | 2 | 1 | 2026-09-13 | Preserved exact drawing wrappers and package graphs through comparison, save, reopen, accept, and reject |
| F-256 | S72 | L | 4 | 1 | 2026-09-13 | Added transactional package-authoritative main-body fragment import with complete selected dependency closure and deterministic conflict policies |
| F-X094a | S72 | M | 2 | 1 | 2026-09-14 | Exposed schema-versioned Word comments, revision resolution, comparison, and TOC rebuilding through atomic CLI commands |
| F-X094b | S72 | L | 4 | 1 | 2026-09-14 | Added structured accepted-view text, direct body layout fragments, and expected-count guarded replacement to the Word CLI |
| F-X094c | S72 | L | 4 | 1 | 2026-09-14 | Added frozen typed Python snapshots for Word comparison, comments, deterministic layout, page lookup, and TOC rebuilding |
| F-X094d | S72 | L | 4 | 1 | 2026-09-14 | Added ordered typed Python snapshots for Word sections, styles, rich stories, header and footer variants, and relationship-scoped hyperlinks |
| F-X094e | S72 | L | 4 | 1 | 2026-09-14 | Added deterministic PowerPoint and notes rendering plus frozen modern comment snapshots and atomic collaboration mutation to the Python binding |
| F-X092 | S72 | L | 4 | 1 | 2026-09-14 | Preserved complete logical reading order in generated Word and PowerPoint PDFs without changing painted output or the deterministic hash harness |
| F-X095 | S72 | L | 4 | 1 | 2026-09-14 | Integrated hardened equivalents of PRs 77 through 80, restored pinned Presentation fidelity CI, and preserved no-op comparison handles |
| F-X096 | S72 | M | 2 | 1 | 2026-09-14 | Aligned each Python distribution with its native crate version and gave rdocx and rpptx independent PyPI tag routes |
| F-X094f | S72 | M | 2 | 2 | 2026-09-15 | Published and verified metadata-complete rdocx 0.13.2 and rpptx 0.11.0 Python distributions, then notified and closed every included issue and pull request |
| F-X097 | S73 | M | 2 | 1 | 2026-09-15 | Preserved ancestor-scoped drawing namespaces and exact complex-field ownership through comparison, accept, and reject |
| F-X098 | S73 | M | 2 | 1 | 2026-09-15 | Preserved attributes and ordered child payloads on modeled content-control type elements until explicit type replacement |
| F-X099 | S73 | M | 2 | 1 | 2026-09-15 | Added checked direct-body ownership to native and frozen Python story-item snapshots without changing recursive paths |
| F-X100 | S73 | S | 1 | 1 | 2026-09-15 | Preserved explicit false row-header, row-split, and cell no-wrap values from contributor PR 101 |
| F-X102 | S73 | M | 2 | 1 | 2026-09-15 | Resolved body, header, and footer pictures against their owning relationship scope from contributor PR 102 |
| F-X103 | S73 | M | 2 | 1 | 2026-09-15 | Accepted Word-default TOC web layout switches and exposed ordered native and Python rebuild diagnostics from contributor PR 101 |
| F-X104 | S73 | M | 2 | 1 | 2026-09-15 | Rendered namespace-aware DrawingML picture transparency through every presentation output backend from contributor PR 105 |
| F-X105 | S73 | M | 2 | 1 | 2026-09-15 | Separated slide-owned latent placeholders from source-specific layout and master header-footer policy |
| F-X110 | S73 | S | 1 | 1 | 2026-09-15 | Added lossless native and Python control of the Word update-fields-on-open setting from contributor PR 104 |
| F-X111 | S73 | L | 4 | 1 | 2026-09-15 | Added selected-family six-target CLI release archives, exact checksums, cargo-binstall metadata, and pre-publication asset validation |
| F-X106a | S73 | L | 4 | 1 | 2026-09-15 | Added owner-checked indexed Python content mutation and counted replacement from contributor PRs 109 and 111 |
| F-260 | S73 | L | 4 | 1 | 2026-09-15 | Added ordered mixed run content authoring with schema-valid field segmentation and exact raw boundary preservation |
| F-X101 | S73 | M | 2 | 1 | 2026-09-15 | Honored run-level page breaks in physical pagination and every page-dependent output from contributor PR 102 |
| F-X106b | S73 | M | 2 | 1 | 2026-09-15 | Added paragraph and run formatting mutation with named Word highlights and separate shading from contributor PR 108 |
| F-X106c | S73 | L | 4 | 1 | 2026-09-15 | Added typed Python story mutation, scoped hyperlinks, revision resolution, field updates, exact XML snapshots, and Issue 94 hardening from contributor PRs 109 and 110 |
| F-X108 | S73 | M | 2 | 1 | 2026-09-15 | Added atomic story-scoped image replacement with shared-target copy-on-write and format-aware package cleanup from contributor PR 107 |
| F-X109 | S73 | M | 2 | 1 | 2026-09-15 | Added atomic native and Python run splitting at Unicode scalar offsets from contributor PR 112 |
| F-257 | S73 | L | 4 | 1 | 2026-09-15 | Completed checked native table properties and atomic full-grid mutation with typed reopen and exact producer XML preservation |
| F-258 | S73 | L | 4 | 1 | 2026-09-16 | Completed checked row and cell authoring with grid and merge validation, nested tables, typed reopen, and Word parity |
| F-X120 | S73 | S | 1 | 1 | 2026-09-16 | Integrated and hardened contributor PR 122 with exact fractional Word line-spacing normalization and deterministic layout parity |
| F-X107 | S73 | M | 2 | 1 | 2026-09-16 | Integrated and hardened contributor PR 113 with atomic native and Python formatted-row cloning and removal |
| F-X113 | S73 | M | 2 | 1 | 2026-09-16 | Preserved one, two, and three appended comparison paragraphs through exact accept and reject reconstruction |
| F-259 | S73 | M | 2 | 1 | 2026-09-16 | Added pure deterministic paragraph and table measurement at caller width with ordered diagnostics and exact equal-height layout proof |
| F-261 | S73 | L | 4 | 1 | 2026-09-16 | Added atomic rich HTML fragment insertion across body, cell, header, and footer stories with explicit images and ordered diagnostics |
| F-262 | S73 | L | 4 | 1 | 2026-09-16 | Added typed crop and floating picture controls, story text boxes with modeled compatibility, and section-aware text watermarks |
| F-X115 | S73 | M | 2 | 1 | 2026-09-16 | Preserved modern comment content types, stable identities, thread links, resolved state, and optional native and Python dates |
| F-X116 | S73 | L | 4 | 1 | 2026-09-16 | Made Python story snapshots linear and complete with one accepted-view inventory and recursive live run paths |
| F-X117 | S73 | M | 2 | 1 | 2026-09-16 | Premultiplied raster image pixels and rendered large presentation pictures within a checked 64 MiB decode ceiling |
| F-X118 | S73 | L | 4 | 1 | 2026-09-16 | Rendered producer-valid notes graphs and added guarded counted slide and notes replacement across native, Python, and CLI surfaces |
| F-X121 | S73 | S | 1 | 1 | 2026-09-16 | Integrated and hardened Kevin Brown's PR 123 with portable authored line-chart title, marker, and smoothing defaults |
| F-X119 | S73 | L | 4 | 1 | 2026-09-17 | Completed round-three Python authoring and inspection across clean rdocx and rpptx abi3 wheels |
| F-263 | S73 | L | 4 | 1 | 2026-09-17 | Added atomic layout-backed field caches and passed the five-document pure Rust private corpus gate |
| F-X114 | S73 | M | 2 | 1 | 2026-09-17 | Rebuilt TOC entries with localized built-in styles, section-derived tab stops, and structural numbering suffix tabs from Issue 116 |
| F-X122 | S73 | M | 2 | 1 | 2026-09-17 | Recovered the immutable rpptx release attempt with newline-safe asset validation and a coherent 0.12.1 package family |
| F-X112 | S73 | L | 4 | 2 | 2026-09-18 | Published and verified all four S73 Rust and Python package families, notified every included contributor record, and closed every fully addressed open record |
| F-X123 | S74 | S | 1 | 1 | 2026-09-18 | Accepted producer TOC style lists with a trailing separator and duplicate style identifiers without weakening strict style mutation validation |
| F-X124 | S74 | M | 2 | 1 | 2026-09-18 | Made Python content cloning linear through one owned body inventory and added explicit source and destination type errors |
| F-X125 | S74 | M | 2 | 1 | 2026-09-18 | Compared gained, lost, and resized table grids as exact tracked table replacements with clean accept and reject outcomes |
| F-X127 | S74 | S | 1 | 1 | 2026-09-18 | Collapsed adjacent trailing run and paragraph page-break requests while preserving visible and structural boundaries |
| F-267 | S74 | L | 5 | 1 | 2026-09-19 | Completed table style conditional layers and fixed three precedence defects that resolved bands and chains unlike Word |
| F-269 | S74 | L | 5 | 1 | 2026-09-19 | Made authored columns reach pagination and typed seven preserved section children without moving a harness entry |
| F-265 | S74 | L | 5 | 1 | 2026-09-19 | Typed sixteen run children and closed three theme-attribute data-loss paths, deferring the visual effect render to F-312 |
| F-270 | S74 | L | 5 | 1 | 2026-09-19 | Modeled the remaining settings behind one order table and made web settings a modeled part |
| F-264 | S74 | L | 5 | 1 | 2026-09-19 | Typed nine paragraph children and shipped the full paragraph facade, with border retention and w:bidi as cross-story contracts |
| F-266b | S74 | L | 5 | 1 | 2026-09-19 | Modeled ruby as typed paragraph content over a base-run span and projected emphasis marks into layout |
| F-268b | S74 | M | 3 | 1 | 2026-09-19 | Placed floating tables through the existing wrap machinery and found the gate that kept floats off the two-pass path, closing DOCX-035 |
| F-266a | S74 | L | 5 | 1 | 2026-09-19 | Gave Hangul and Kana script identity, made w:rFonts resolve per slot, and bundled the three deterministic subsets the golden gate needs |
| F-268a | S74 | L | 5 | 1 | 2026-09-19 | Modeled the advanced table grammar and made gridBefore, gridAfter and tblLayout mean something, closing DOCX-034 |
| F-X132 | S74 | S | 1 | 1 | 2026-09-19 | Stopped a redundant namespace declaration making its element an owner, clearing the second F-X128 regression on the Word corpus |
| F-X131 | S74 | S | 1 | 1 | 2026-09-19 | Stopped modeled roots rebinding a namespace their scope already owns, clearing five pre-existing F-X128 failures |
| F-X129 | S74 | S | 1 | 1 | 2026-09-18 | Skipped unmatched notes-slide placeholder overlays with ordered diagnostics while preserving strict ownership failures |
| F-X128 | S74 | M | 2 | 1 | 2026-09-18 | Preserved ordered paragraph, run, and section root identities across save, reopen, and typed mutation |
| F-X126 | S74 | M | 2 | 1 | 2026-09-18 | Normalized drawing namespace ownership in comparison-only story projections while preserving exact story-scoped package payloads |
| F-266c | S74 | L | 5 | 1 | 2026-09-19 | Modeled character grids and East Asian run layout, then rendered vertical table and section text with deterministic geometry |
| F-266 | S74 | L | 0 | 0 | 2026-09-19 | Closed the split typography parent after all three reviewed child gates passed and DOCX-033 reached its final classification |
| F-268 | S74 | L | 0 | 0 | 2026-09-19 | Closed the split advanced-table parent after both reviewed child gates passed and DOCX-035 reached its final classification |
| F-X130 | S74 | L | 5 | 1 | 2026-09-20 | Rewrote all 27 package pages around final authoring depth and added reproducible archive footprint and bounded speed evidence |
| F-X134 | S75 | S | 1 | 1 | 2026-09-20 | Batched hyperlink namespace scopes and bounded story link text extraction so Python snapshots remain linear |
| F-X135 | S75 | L | 5 | 6 | 2026-09-26 | Integrated PRs 146 through 151, resolved five uncovered issues, audited late PRs 153 and 154, and preserved unchanged package bytes |
| F-X136 | S75 | L | 5 | 1 | 2026-09-26 | Split ordinary Word table rows across page boundaries, respected cantSplit and exact heights, and removed the source-built footer-only page |
| F-X137 | S76 | L | 5 | 1 | 2026-09-30 | Integrated five package and CLI safety PRs with atomic saves, safe output refusal, namespace binding and package class selection |
| F-X138 | S76 | L | 5 | 1 | 2026-09-30 | Integrated eight Word story PRs with direct body coordinates, comment anchoring, content control traversal and source-preserving replacement |
| F-X139 | S76 | L | 5 | 1 | 2026-09-30 | Integrated six Word identity and comparison PRs with producer attribute preservation, marker ordering, comparison options and CLI comment dates |

| F-X140 | S77 | L | 5 | 1 | 2026-09-30 | Integrated seven rendering and layout PRs with reviewed ligature PDF hashes, presentation preservation and Word pagination regressions |
| F-X141 | S77 | L | 5 | 1 | 2026-09-30 | Integrated six Word Python PRs, added native common styles and refreshable TOCs, and reviewed a separate 16-key baseline delta |
| F-X142 | S77 | L | 5 | 1 | 2026-10-01 | Integrated six Presentation Python PRs and completed the Issue 169 API, comment anchor and built-in table style checklist |
| F-X143 | S77 | M | 3 | 1 | 2026-10-01 | Integrated two revision and CLI PRs with all-story listing, related-part warnings and strict validation |

| F-X151 | S78 | L | 5 | 1 | 2026-10-01 | Integrated six Word PR behaviors with scoped review fixes, unchanged hash output and 121 passing Python tests |
| F-X144 | S78 | L | 5 | 1 | 2026-10-01 | Completed 17 identity rows and 11 producer rows across operations, repaired picture and styles preservation, and kept 49 hash entries unchanged |
| F-X145 | S78 | L | 5 | 1 | 2026-10-01 | Completed Issue 161 comment, TOC and marker redline cases, preserved run defaults, and kept 49 hash entries unchanged |
| F-X146 | S79 | L | 5 | 1 | 2026-10-01 | Integrated PRs 222, 225 and 237 with Word line pitch, inline picture, rich-line, tab and TOC geometry and a reviewed 22-key hash delta |
| F-X152 | S79 | M | 3 | 1 | 2026-10-01 | Integrated PR 236 with every supported story in CLI diff, bounded matching, one count per changed paragraph and incomplete-story reporting |
| F-X153 | S79 | M | 3 | 1 | 2026-10-01 | Integrated PR 220 with text-anchored CLI comments, inherited numbering validation and atomic native and Python run removal |
| F-X154 | S80 | M | 3 | 1 | 2026-10-01 | Repaired DrawingML body property preservation and presentation line-separator layout for Issues 215 and 216 |
| F-X147 | S80 | L | 5 | 1 | 2026-10-02 | Completed the Issue 168 Word Python editing chain with atomic counted replacement, rich footers and a documented lxml package-write fallback |
| F-X155 | S80 | L | 5 | 1 | 2026-10-02 | Integrated five drawing PRs with slide jumps, line ends, shadows, preset geometry and theme effect control under schema-order and viewer checks |
| F-X156 | S81 | L | 5 | 1 | 2026-10-02 | Integrated slide import, scoped counted replacement and built-in table backgrounds with relationship, Python and viewer checks |
| F-X148 | S81 | L | 5 | 1 | 2026-10-02 | Completed the Issue 169 production deck chain with a pinned Python reader and LibreOffice viewer, and documented one table-style package fallback |
| F-X157 | S81 | L | 5 | 1 | 2026-10-02 | Proved all six Issue 217 authoring items in one saved Python deck with pinned reader, validator and per-item viewer checks |
| F-X158 | S82 | L | 5 | 1 | 2026-10-02 | Exercised the SHA-bound reporter deck through the complete Python workflow, package validation, pinned reader and 72 DPI viewer comparisons |
| F-X149 | S82 | L | 5 | 1 | 2026-10-02 | Ran the original report workflow and both matrices, and repaired replacement-image comparison with accepted and rejected media checks |
| F-X150 | S82 | M | 3 | 1 | 2026-10-02 | Reconciled all 22 original issue criteria with 18 passing rows, four unresolved rows and post-main issue decisions |
| F-X159 | S83 | M | 3 | 1 | 2026-10-02 | Corrected fresh Word AppVersion and nested-table validity, opened the showcase in Word and reviewed a three-entry deterministic hash delta |
| F-X160 | S83 | L | 5 | 1 | 2026-10-02 | Accepted producer style, drawing ID and decimal measurement variants with unchanged 49-entry hash output |
| F-X169 | S84 | L | 5 | 1 | 2026-10-02 | Classified all 31 live PRs and eight issues, tied 19 older PRs to S76 integration commits, and aligned the whole-number roadmap assertion |
| F-X161 | S84 | L | 5 | 1 | 2026-10-02 | Integrated PR 251's compact Word part serializers and namespace preservation with 20 reviewed Word XML baseline entries |
| F-X165 | S84 | M | 3 | 1 | 2026-10-02 | Integrated PR 256's Python and CLI accepted or tracked view selectors with pinned PDF text evidence and invalid-selector checks |
| F-X162 | S84 | M | 3 | 1 | 2026-10-02 | Integrated PR 241's governing section layout with two reviewed feature showcase PDF hash changes |
| F-X163 | S84 | M | 3 | 1 | 2026-10-02 | Integrated PR 242's plain line hanging-space fit with the complete Issue 226 width matrix and reviewed PDF and pixel baselines |
| F-X164 | S84 | M | 3 | 1 | 2026-10-02 | Integrated PR 252's theme style for new shapes and SmartArt text colour precedence with pinned PowerPoint and LibreOffice checks |
| F-X166 | S84 | L | 5 | 2 | 2026-10-03 | Integrated PRs 257, 258, 260 and 261 with 80 Word for Mac accept and reject outcomes, final-block and picture regressions, and unchanged hashes |
| F-X170 | S84 | M | 3 | 1 | 2026-10-03 | Added high-level Word style formatting for Issue 264 with public examples, round-trip resolution and unchanged hashes |
| F-X167 | S84 | L | 5 | 1 | 2026-10-03 | Integrated accepted-view paragraph joins and deleted-row projection across text, CLI JSON, HTML, Markdown and PDF with 175 Python binding cases and unchanged hashes |
| F-X168 | S84 | M | 3 | 1 | 2026-10-03 | Proved both Issue 158 fixture workflows, the identity and producer matrices, nine issue contracts and PR 265 mixed-section pagination on the integrated repair prefix |
| F-X171 | S85 | S | 1 | 1 | 2026-10-03 | Pinned LibreOffice 26.2.5.2 in the macOS presentation Python CI cell, verified the reviewed DMG digest and restored the complete Issue 158 deck oracle gate |
| F-271 | S86 | L | 5 | 1 | 2026-10-03 | Added rich content and staged comments across all header and footer variants, with scoped relationship and preservation checks |
| F-X133 | S86 | S | 1 | 1 | 2026-10-03 | Scoped canonical root bindings in retained Word paragraphs while preserving standalone declarations and local shadows |
| F-272 | S86 | L | 5 | 1 | 2026-10-03 | Added atomic rich footnote lifecycle and occurrence labels, with pinned Word continuation and package preservation evidence |
| F-X172 | S86 | L | 5 | 1 | 2026-10-03 | Unified CLI and Python release assets under one family tag with producer attestations, complete checksums and closed-main publication gates |
| F-273 | S86 | L | 5 | 1 | 2026-10-03 | Added atomic rich endnote authoring with independent IDs, preserved part XML and scoped assets, and recorded the pinned Word policy divergence |
| F-X173 | S86 | M | 3 | 1 | 2026-10-03 | Prepared rpptx 0.13.0 across 15 Rust crates, CLI and Python with family notes, package evidence and Python 3.9 wheel smoke repair |
| F-X174 | S86 | M | 3 | 1 | 2026-10-04 | Prepared rdocx 0.15.0 across seven Rust crates, CLI and Python with family notes, clean wheel smoke, package evidence and 49 matching hashes |
| F-X175 | S87 | S | 1 | 1 | 2026-10-04 | Refreshed both CLI archive footprints after Windows stack hardening and aligned the S87 roadmap check, with 49 matching hashes and a clean 22-crate dry run |
| F-X176 | S88 | M | 3 | 1 | 2026-10-04 | Repaired the unified tag package inventory against dry-run directories and prepared rpptx 0.13.1 beside rdocx 0.15.0, with 49 matching hashes and a clean 22-crate dry run |
| F-274 | S89 | L | 5 | 1 | 2026-10-04 | Added checked note numbering and placement policies, custom markers and authored continuation records with pinned Word evidence and 49 matching hashes |
| F-275 | S89 | L | 5 | 1 | 2026-10-04 | Added atomic bookmark, comment, permission and proofing ranges across eight story owners and nested controls with 49 matching hashes |
| F-276 | S89 | L | 5 | 2 | 2026-10-05 | Added deterministic fragment closure across 64 owner pairs with exact dependency remaps, atomic failures and 49 matching hashes |
| F-X177 | S89 | S | 1 | 1 | 2026-10-05 | Integrated changjoon-park PR 269 fontdb feature-unification fix, with four feature configurations and unchanged hashes |
| F-277 | S89 | L | 5 | 1 | 2026-10-05 | Added checked glossary creation, updates, insertion, removal and placeholder binding with raw XML preservation and atomic package closure |

| F-X179 | S90 | S | 1 | 1 | 2026-10-07 | Integrated hadim PR 271 against every Issue 270 criterion, fixed nested legacy reply removal and retained all 49 hash entries |

| F-278 | S90 | L | 5 | 1 | 2026-10-07 | Added checked simple and complex fields, recursive operands, ordered caches and locks, with clean microscope and 49 unchanged hashes |
| F-279 | S90 | L | 5 | 1 | 2026-10-07 | Added five-kind pagination materialization across modeled stories, authenticated Word ownership, clean implementation and integration reviews, 49 unchanged hashes and 22 verified packages |

| F-280 | S90 | M | 2 | 1 | 2026-10-08 | Added atomic caption and cross-reference authoring, shared physical sequences and numbering, bounded note and annotation refresh, clean worker and integration reviews and unchanged49 worker hashes |

| F-281 | S90 | L | 5 | 1 | 2026-10-08 | Added atomic indexes, figures and authority tables, namespace-safe owned caches and one post-insertion snapshot, clean worker and integration reviews and49 unchanged integrated hashes |

| F-X182 | S90 | M | 2 | 1 | 2026-10-08 | Integrated hadim PR279 for Issue277, 480 alignment combinations, clean delta and ALL reviews, five explained hash changes and invoice-only golden delta |

| F-X184 | S90 | L | 5 | 1 | 2026-10-08 | Safe complete-thread removal and atomic partial-cut refusal for hadim Issue282, clean ALL review, native1750 and final rebuilt Python85+2, hash49 unchanged |

| F-X185 | S90 | M | 2 | 1 | 2026-10-08 | Checked typed comment anchors and nested paragraph snapshots for hadim Issues283/289, clean ALL, native1761/XML629/CLI64 and current rebuilt Python93, hash49 unchanged |

| F-X250 | S95.11 | L | 4 to 5 | not recorded | 2026-10-17 | Direct-main-body chart snapshots with finite data fields, display math, list restart/font and note style/removal. Independent clean feature reviews, five unchanged Word reporter bodies, independent XML and bounded native/own appearance. Shared accepted 05ba6cc8461a896fa751ca7dc1fe209a8fa70045 evidence retains H3/H5/836/H7 producer epochs and actual receipts, no duplicated full-suite or native claim |
| F-X254 | S95.11 | L | 4 to 5 | not recorded | 2026-10-17 | Atomic PowerPoint value/XML/CLI refusals, ten unchanged reporter cases, native PowerPoint save/reopen and independent XML, affected native/CLI GREEN, explicit user-approved ValidationIssue enum exhaustive-match compatibility boundary. Shared 05ba6cc8461a896fa751ca7dc1fe209a8fa70045 verification and current H7 CLI genuine archive proof, no release |
| F-X223 | S95.11 | M | 2 to 3 | not recorded | 2026-10-17 | User accepted documented Word limit after27 actual phases,13 finite observations/14unsupported refusals. Genuine properBold in3phases, pairedItalic Regular outlines with shear, native normalization and sensitivities retained. Bounded investigation complete, broader genuine paired/four-face consumption unqualified. No speculative writer repair or F-X223 carry |

| F-283 | S95.2 | L | 4 to 5 | not recorded | 2026-10-18 | Numbered REF and STYLEREF follow the captured Word 16.114 context rules, including literal full context, unnumbered STYLEREF level results, and cross-instance sharing only for one definition without a start override. Clean microscope pass 2. 49 hashes unchanged |
| F-284 | S95.2 | L | 4 to 5 | not recorded | 2026-10-18 | Opt-in grammar 1 renders one payload across body, table, header, footer, note, and text box, with legacy calls unchanged. Clean microscope pass 1. Word 16.114 and LibreOffice 26.2.5.2 preserved the declared text. 49 hashes unchanged |

| F-X259 | S95.3 | M | 2 to 3 | not recorded | 2026-10-22 | Accepted bounded provider implementation and actual qualification. Clean feature microscopes.49hash/sevenpixel unchanged. F-X263 and final sprint gate remain due |
| F-X260 | S95.3 | L | 4 to 5 | not recorded | 2026-10-22 | Accepted bounded provider implementation and actual qualification. Clean feature microscopes.49hash/sevenpixel unchanged. F-X263 and final sprint gate remain due |
| F-X261 | S95.3 | L | 4 to 5 | not recorded | 2026-10-22 | Accepted bounded provider implementation and actual qualification. Clean feature microscopes.49hash/sevenpixel unchanged. F-X263 and final sprint gate remain due |
| F-X262 | S95.3 | L | 4 to 5 | not recorded | 2026-10-22 | Accepted bounded provider implementation and actual qualification. Clean feature microscopes.49hash/sevenpixel unchanged. F-X263 and final sprint gate remain due |
| F-285 | S95.3 | L | 4 to 5 | not recorded | 2026-10-22 | Accepted bounded provider implementation and actual qualification. Clean feature microscopes.49hash/sevenpixel unchanged. F-X263 and final sprint gate remain due |
| F-286 | S95.3 | L | 4 to 5 | not recorded | 2026-10-22 | Accepted bounded provider implementation and actual qualification. Clean feature microscopes.49hash/sevenpixel unchanged. F-X263 and final sprint gate remain due |
| F-287 | S95.3 | L | 4 to 5 | not recorded | 2026-10-22 | Accepted bounded provider implementation and actual qualification. Clean feature microscopes.49hash/sevenpixel unchanged. F-X263 and final sprint gate remain due |

| F-X263 | S95.3 | L | 3 to 5 additional | not recorded | 2026-10-22 | Exact patch families prepared, clean microscope, genuinely verified22archives and actual hosted27success/fourpublicationSKIP/26payloads. Downloaded priority172calls and strict typing/stubs pass. No publication |
| F-288 | S95.4 | M | shared authoring 12 to 17 | not recorded | 2026-10-28 | Approved contract accepted in final source-qualified union at c49, original producers/failures and limitations retained. See final AS_BUILT |
| F-289 | S95.4 | L | shared authoring 12 to 17 | not recorded | 2026-10-28 | Approved contract accepted in final source-qualified union at c49, original producers/failures and limitations retained. See final AS_BUILT |
| F-290 | S95.4 | M | shared authoring 12 to 17 | not recorded | 2026-10-28 | Approved contract accepted in final source-qualified union at c49, original producers/failures and limitations retained. See final AS_BUILT |
| F-X264 | S95.4 | L | 3 to 5 | not recorded | 2026-10-28 | Approved contract accepted in final source-qualified union at c49, original producers/failures and limitations retained. See final AS_BUILT |
| F-X265 | S95.4 | M | 2 to 3 | not recorded | 2026-10-28 | Approved contract accepted in final source-qualified union at c49, original producers/failures and limitations retained. See final AS_BUILT |
| F-X266 | S95.4 | L | 4 to 6 | not recorded | 2026-10-28 | Approved contract accepted in final source-qualified union at c49, original producers/failures and limitations retained. See final AS_BUILT |
| F-X267 | S95.4 | L | 6 to 10 | not recorded | 2026-10-28 | Approved contract accepted in final source-qualified union at c49, original producers/failures and limitations retained. See final AS_BUILT |
| F-X268 | S95.4 | L | 12 to 21 scope judgment | not recorded | 2026-10-28 | Approved contract accepted in final source-qualified union at c49, original producers/failures and limitations retained. See final AS_BUILT |
| F-X270 | S95.4 | L | 4 to 5 provisional | not recorded | 2026-10-28 | Actual requested release preparation, build-only artifacts and installed qualification. Publication held |

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
| S67 | 3 | 1 | 15.00 |
| S68 | 2 | 1 | 10.00 |
| S69 | 7 | 2 | 17.50 |
| S70 | 7 | 2 | 17.50 |
| S71 | 10 | 3 | 16.67 |
| S72 | 19 | 7 | 13.57 |
| S73 | 35 | 4 | 61.25 |
| S74 | 22 | 3 | 36.67 |
| S75 | 3 | 8 | 1.88 |
| S76 | 3 | 2 | 7.50 |
| S77 | 4 | 2 | 10.00 |
| S78 | 3 | 1 | 15.00 |
| S79 | 3 | 1 | 15.00 |
| S80 | 3 | 2 | 7.50 |
| S81 | 3 | 1 | 15.00 |
| S82 | 3 | 1 | 15.00 |
| S83 | 2 | 1 | 10.00 |
| S84 | 10 | 2 | 25.00 |
| S85 | 1 | 1 | 5.00 |
| S86 | 7 | 2 | 17.50 |
| S87 | 1 | 1 | 5.00 |
| S88 | 1 | 1 | 5.00 |
| S89 | 5 | 2 | 12.50 |
| S90 | 18 | 3 | 30.00 |
| S91 | 29 | 7 | 20.71 |
| S91.1 | 4 | 1 | 20.00 |
| S92 | 6 | 1 | 30.00 |
| S93 | 7 | 2 | 17.50 |
| S94 | 5 | 1 | 25.00 |
| S95 | 9 | 1 | 45.00 |
| S95.11 | 3 | 2 | 7.50 |
| S95.2 | 2 | 2 | 5.00 |
| S95.3 | 8 | 3 | 13.33 |
| S95.4 | 9 | 6 | 7.50 |

S95.2 records two worked dates, 2026-10-17 and 2026-10-18. Its 5.00 stories per week counts two completed stories over those dates. It is calendar throughput, not person effort. F-X192 remains carried and is not counted as completed.

S90 velocity measures calendar throughput for a parallel team. It does not
measure person effort or establish a sustainable single-developer forecast.
Later per-feature effort was not recorded, and those actuals remain unchanged.

S91 counts29 completed top-level contracts once. Four completed F-X211 delivery children remain separate records and add neither stories nor estimates to that total. Seven actual dates include worked Saturday10 and Sunday11 October. The20.71 stories per week measures calendar throughput for parallel work. Per-feature actual effort remains unrecorded.

S91.1 records one worked date,2026-10-15. Its20.00 stories per week measures
parallel calendar throughput. Per-feature person-effort remains unrecorded.

S92 records one worked date,2026-10-15. Its30.00 stories per week is parallel
calendar throughput, not a sustainable single-developer forecast. The last
three sprint windows, S91, S91.1 and S92, total39 stories over9 worked-date
units, or21.67 stories per week. Shared calendar dates are not person-days.

S93 records two worked dates,2026-10-15and2026-10-16. Its17.50 stories per week measures parallel calendar throughput. The last three windows, S91.1, S92 and S93, total17stories over4worked-date units, or21.25stories per week. Shared dates are not person-days, and per-feature actual effort remains unrecorded. The41.67percent decrease from S92 exceeds the30percent escalation trigger. Retain developer-effort estimates and dependency gates. Split the eleven-story S95 plan before opening it rather than treating parallel throughput as a sustainable forecast.

S94 records one worked date, 2026-10-16. Its 25.00 stories per week measures parallel calendar throughput, not person-effort or a sustainable single-developer forecast. The last three windows, S92, S93 and S94, total 18 stories over 4 worked-date units, or 22.50 stories per week. Shared dates are not person-days, and per-feature actual effort remains unrecorded. The 42.86 percent increase from S93 exceeds the 30 percent close-report escalation threshold. Retain the approved 14 to 19 developer-day effort estimate and the capacity-reviewed S95 through S95.4 repair and carryover boundaries. The new rolling window is 5.88 percent above the preceding 21.25 rate, but remains unsuitable as a person-effort forecast.

S95 records one worked date,2026-10-16. Its45.00 stories per week measures parallel calendar throughput, not person-effort or a sustainable single-developer forecast. The last three windows,S93,S94andS95,total21stories over4worked-date units,or26.25stories per week. Shared dates are not person-days, and per-feature actual effort remains unrecorded. The80.00percent increase from S94 exceeds the30percent close-report escalation threshold. The rolling rate is16.67percent above the preceding22.50 rate. Retain the approved22to31 developer-day estimate for seven medium and two large stories. The user explicitly expanded S95 to include Issue386 before release preparation. Preserve the dependent S95.11 and later roadmap boundaries and their effort estimates. This summary does not claim sprint review, main merge, sprint close, tags or publication have occurred.

S95.3 records three worked dates,2026-10-18,19and22. Eight stories over three date units give13.33stories/week,166.67percent above S95.2 and beyond the30percent escalation threshold. S95.11/S95.2/S95.3 total13completed stories over7worked-date units, or9.29stories/week. Shared calendar dates are not person-days. Feature effort is not recorded. Preserve the34to49developer-day expanded estimate, later sprint boundaries and explicit dependency gates. Final review and close remain due.

S95.4 records six worked dates 2026-10-23 through 28. Its 7.50 stories/week is parallel calendar throughput, not person-effort. This is 43.75 percent below S95.3's 13.33 rate and exceeds the 30 percent escalation trigger. Intake expanded to eight product contracts and the requested ninth release-preparation contract. The original three-authoring-story 12 to 17 and four-repair-story 15 to 24 estimates, later combined F-X268 12 to 21 judgment and release size L category remain historical scope disclosures. Overlapping judgments are not re-summed into a capacity promise. Actual per-feature effort remains unrecorded. Preserve dependency-defined future boundaries and use smaller coherent verification batches. The last three windows S95.2/S95.3/S95.4 total 19 stories over 11 worked-date units, 8.64 stories/week, rather than a sustainable developer-effort forecast.

## Escalation record

Logged when an escalation trigger from `.claude/WORKFLOW.md` fires, with what
was done about it. Empty is the expected state.

| Date | Trigger | F-ID or sprint | Response |
|------|---------|----------------|----------|
| 2026-10-22 | Sprint throughput variance exceeds30percent | S95.3 | Record13.33stories/week,166.67percent above S95.2, as calendar throughput across three worked dates. Retain34to49estimated developer-days and later dependency/capacity boundaries. Do not infer person effort or a sustainable forecast |
| 2026-10-17 | Sprint throughput variance exceeds30percent | S95.11 | Record7.50stories/week,83.33percent below S95, as parallel calendar throughput across two worked dates. Preserve S95.2/S95.3/S95.4 dependency boundaries and developer estimates. No person-effort or elapsed-time forecast is inferred |
| 2026-10-16 | Sprint throughput variance exceeds30percent | S95 | Record45.00stories/week,80.00percent above S94,as parallel calendar throughput. The last three windows yield26.25stories/week. Preserve22to31developer-day effort estimates and the explicit user-approved expansion to nine stories. Retain S95.11 and later dependency boundaries without converting throughput into elapsed delivery promises. Formal review,close and publication remain separate |
| 2026-10-16 | Sprint throughput variance exceeds30percent | S93 | Record17.50stories/week,41.67percent below S92, as parallel calendar throughput. The last three windows yield21.25stories/week, above the2story/week planning assumption. Preserve effort estimates and require the existing S95 capacity split before opening that sprint |
| 2026-10-15 | Three-sprint throughput differs from the2-stories/week assumption by983percent | S91, S91.1, S92 | Record21.67 stories/week as parallel calendar throughput. Retain dependency-defined roadmap boundaries and developer-effort estimates. Do not convert this observation into elapsed delivery promises or invent person-effort actuals |
| 2026-10-28 | Throughput variance exceeds30percent | S95.4 | Record7.50stories/week over six worked dates,43.75percent below S95.3. Preserve unrecorded person-effort, approved scope limits and smaller batched verification. Release remains separately approved |
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
| 2026-09-04 | Sprint estimate variance exceeded 30 percent | S67 | Record 1 actual day against 10 estimated. Isolated workers, the completed field and comparison foundations, and the established Word oracle concentrated review and verification across one integrated automation boundary. The resulting 15.00 stories per week is not a sustainable forecast, so retain the dependency-defined S68 boundary |
| 2026-09-05 | Sprint estimate variance exceeded 30 percent | S68 | Record 1 actual day against 8 estimated. Two isolated workers reused the completed package-signature, field, and preservation foundations, while exhaustive feature reviews and the integrated sprint review concentrated fail-closed corrections in the owned Word package paths. The resulting 10.00 stories per week is not a sustainable forecast, so retain the dependency-defined S69 boundary |
| 2026-09-06 | Sprint estimate variance exceeded 30 percent | S69 | Record 2 actual days against 11 estimated. Isolated workers, established package and release workflows, and two recovery releases allowed implementation, review, and publication work to overlap safely. The resulting 17.50 stories per week is not a sustainable forecast, so retain the dependency-defined S70 decision boundary |
| 2026-09-07 | Sprint estimate variance exceeded 30 percent | S70 | Record 2 actual days against 19 estimated. Isolated workers and established conformance, documentation, and cache verification paths allowed independent audit and implementation work to overlap safely. The resulting 17.50 stories per week is not a sustainable forecast, so retain the dependency-defined S71 boundary |
| 2026-09-09 | Sprint estimate variance exceeded 30 percent | S71 | Record 3 actual days against 35 estimated. Isolated workers, dependency-prefix reviews, and established package, Word, Pages, layout, and README gates allowed independent work and contribution hardening to overlap safely. The resulting 16.67 stories per week is not a sustainable forecast, so retain the dependency-defined S72 boundary |
| 2026-09-15 | Sprint estimate variance exceeded 30 percent | S72 | Record 7 actual days against 61 estimated. Parallel isolated workers, dependency-prefix reviews, established release workflows, and issue and contribution hardening allowed independent work and publication checks to overlap safely. The resulting 13.57 stories per week is not a sustainable forecast, so retain the dependency-defined S73 boundary |
| 2026-09-18 | Sprint estimate variance exceeded 30 percent | S73 | Record 4 actual days against 92 estimated. Parallel isolated workers, reviewed contributor pull requests, established corpus and release workflows, and one consolidated publication boundary allowed implementation, review, and registry verification to overlap safely. The resulting 61.25 stories per week is not a sustainable forecast, so retain the dependency-defined S74 boundary |
| 2026-09-20 | Sprint estimate variance exceeded 30 percent | S74 | Record 3 actual days against 66 estimated. Parallel isolated workers, split parent stories, established Word corpus and package gates, and dependency-prefix reviews allowed independent implementation and review work to overlap safely. The resulting 36.67 stories per week is not a sustainable forecast, so retain the dependency-defined S75 boundary |
| 2026-09-26 | Sprint estimate variance exceeded 30 percent | S75 | Record 8 actual days against 47 originally estimated, with 36 estimated days moved to S76 in the approved eight-story carry. The three delivered stories were estimated at 11 days and took 8. This is a deliberate mid-milestone scope cutoff, not evidence that the remaining related-story work became cheaper. The resulting 1.88 stories per week is close to the long-run assumption, so retain the dependency-defined S76 boundary and split it before implementation if capacity requires |
| 2026-09-30 | Sprint estimate variance exceeded 30 percent | S76 | Record 2 elapsed workdays against 15 estimated for three completed contribution waves. Stacked PRs reused established implementation and oracle paths, while dependency-prefix review and full verification remained explicit. The resulting 7.50 stories per week is a short contribution-intake result rather than a sustainable forecast, so retain the dependency-defined S77 boundary |
| 2026-10-01 | Sprint estimate variance exceeded 30 percent | S77 | Record 2 elapsed workdays against 18 estimated for four completed contribution waves. Reviewed stacked changes and established binding, corpus and package gates compressed implementation time while the full integrated gate and sprint review remained explicit. The resulting 10.00 stories per week is a short contribution-intake result rather than a sustainable forecast, so retain the dependency-defined S78 through S88 repair boundaries |
| 2026-10-01 | Sprint estimate variance exceeded 30 percent | S78 | Record 1 elapsed workday against 15 estimated for three completed dependency waves. Reviewed contributor changes, source-built acceptance matrices and the established full gate compressed elapsed time. The resulting 15.00 stories per week is not a sustainable forecast, so retain the dependency-defined S79 through S88 repair boundaries |
| 2026-10-01 | Sprint estimate variance exceeded 30 percent | S79 | Record 1 elapsed workday against 11 estimated for three completed contribution stories. Reviewed PR increments and existing geometry, story and binding gates reduced elapsed integration time while the full verification and sprint review remained explicit. The resulting 15.00 stories per week is not a sustainable forecast, so retain the dependency-defined S80 through S88 repair boundaries |
| 2026-10-02 | Sprint estimate variance exceeded 30 percent | S80 | Record 2 elapsed workdays against 13 estimated for three completed contribution stories. Reviewed contributor increments, dependency-prefix integration and established binding and viewer gates reduced authoring time while the full verification and sprint review remained explicit. The resulting 7.50 stories per week is a short contribution-intake result, so retain the dependency-defined S81 through S88 repair boundaries |
| 2026-10-02 | Sprint estimate variance exceeded 30 percent | S81 | Record 1 elapsed workday against 15 estimated for three completed dependency waves. Reviewed contributor increments and established Python, viewer and package gates compressed elapsed work while full integrated verification and sprint review remained explicit. The resulting 15.00 stories per week is a short contribution-intake result, so retain the dependency-defined S82 through S88 acceptance and repair boundaries |
| 2026-10-02 | Sprint estimate variance exceeded 30 percent | S82 | Record 1 elapsed workday against 13 estimated for three completed acceptance and evidence stories. Existing test infrastructure and pinned fixture inputs reduced elapsed work while full integrated verification and sprint review remained explicit. The resulting 15.00 stories per week is a short acceptance-wave result, so retain the dependency-defined S83 through S88 repair boundaries |
| 2026-10-02 | Sprint estimate variance exceeded 30 percent | S83 | Record 1 elapsed workday against 8 estimated for two completed dependency waves. Reviewed contributions and established Word, binding and package gates compressed elapsed work while full integrated verification and sprint review remained explicit. The resulting 10.00 stories per week is a short contribution-intake result, so retain the dependency-defined S84 through S88 repair boundaries |
| 2026-10-03 | Sprint estimate variance exceeded 30 percent | S84 | Record 2 elapsed workdays against 38 estimated for ten completed stories. This deliberately combined the S84 through S89 repair scope to pay the full-test cost once. Isolated workers, reviewed contribution increments and one integrated acceptance gate reduced elapsed time without changing the 38-day planning estimate. The resulting 25.00 stories per week is not a sustainable forecast. Keep the dependency-defined S85 feature boundary |
| 2026-10-04 | Sprint estimate variance exceeded 30 percent | S86 | Record 2 elapsed workdays against 27 estimated for seven completed stories. Isolated workers reused established authoring and release infrastructure while the integrated gate, clean review and hosted build evidence remained explicit. The resulting 17.50 stories per week is a short sprint result, so retain the dependency-defined S87 boundary |
| 2026-10-04 | Sprint estimate variance exceeded 30 percent | S88 | Record 1 elapsed workday against 3 estimated for the focused release repair. The failed tag had isolated the package inventory cause, and established release checks were reused. The resulting 5.00 stories per week is a short repair result, so retain the dependency-defined S89 boundary |
| 2026-10-05 | Sprint estimate variance exceeded 30 percent | S89 | Record 2 elapsed workdays against 21 estimated. Dependency waves reused the completed related-story and package foundations, and PR 269 supplied a focused reproducer and fix. The resulting 12.50 stories per week is not a sustainable forecast, so retain the dependency-defined S90 boundary |
| 2026-10-09 | Sprint estimate variance exceeded 30 percent | S90 | Record 3 elapsed working dates against the reconstructed 55-day planning estimate, including 50 estimated days for completed features. Nominal variance is 94.55 percent against 55 or 94.00 percent against 50. Parallel assistance, contribution reuse and exclusive implementation waves make elapsed time different from person effort. The resulting 30.00 stories per week is calendar throughput, not a sustainable forecast. Retain the S91 navigation and bibliography boundaries and recalibrate with measured effort |
| 2026-10-15 | Sprint estimate and velocity variance exceeded30 percent | S91 | Record seven worked dates against the reconstructed130-day lower bound, comprising97 days for completed scope and33 carried. Nominal variance is94.62 percent against130 or92.78 percent against97. Throughput is20.71 stories per week,30.95 percent below S90. Rolling contributions and parallel workers make calendar dates different from person effort. Retain the ten S92 carries and recalibrate with measured effort |
| 2026-10-15 | Sprint estimate variance exceeded30 percent | S91.1 | Record one worked date against the10-day lower-bound estimate, a90percent estimate variance. Isolated workers, existing provider repairs and parallel evidence preparation compressed elapsed delivery. Retain dependency and capacity review for future sprints. The20.00 stories per week is calendar throughput, not a sustained person-effort forecast |
| 2026-10-16 | Sprint estimate and velocity variance exceeded 30 percent | S94 | Record one worked date against 14 to 19 estimated developer-days, a nominal elapsed-versus-effort variance of 92.86 to 94.74 percent. The 25.00 stories per week calendar throughput is 42.86 percent above S93. The S92 through S94 window is 18 stories over 4 worked-date units, or 22.50 stories per week. Parallel workers and existing native providers do not establish measured person-effort savings. Retain the reviewed S95 through S95.4 capacity split and dependency boundaries, and recalibrate with measured effort |

## S79 contribution reconciliation

The five assigned PRs, 220, 222, 225, 236 and 237, are superseded by the reviewed F-X153, F-X146 and F-X152 integrations on `main`. Close each with contributor credit and a link to the integrated sprint merge. Close Issues 162 and 227 after their complete layout and diff reproducers pass. Close Issues 163 and 172 after their direct body coordinate and inline control comment entry point regressions pass. Keep Issue 226 open for the plain-line trailing-space fit in F-X163. Keep Issue 168 open for the remaining production checklist in F-X147 and tolerant reads in F-X160. Comment on both open issues with the completed S79 work and the remaining acceptance scope.

## S80 contribution reconciliation

PRs 218 and 223 were closed as superseded by F-X154. PRs 219, 221, 224, 230 and 234 were closed as superseded by F-X155. Each received an individual comment thanking hadim and linking the pushed S80 merge at `bc700efb`.

Issues 215 and 216 were closed after their preservation and line-separator regressions passed. Issue 168 was closed after the complete Word Python chain passed and the reporter-accepted lxml package-write fallback was documented and exercised. Issue 217 remains open for F-X156 slide import and scoped replacement, followed by F-X157 complete deck-chain acceptance in S81. The accepted S80 LibreOffice and python-pptx evidence does not claim a PowerPoint 16.104 observation.

## S81 contribution reconciliation

PRs [231](https://github.com/tensorbee/rdocx/pull/231#issuecomment-5950014500), [235](https://github.com/tensorbee/rdocx/pull/235#issuecomment-5950024422) and [238](https://github.com/tensorbee/rdocx/pull/238#issuecomment-5950035288) were closed as superseded by the pushed S81 merge at `58e41809`. Each received an individual note thanking hadim, naming the integrated result and explaining its remaining scope.

Issues [169](https://github.com/tensorbee/rdocx/issues/169#issuecomment-5950044534) and [217](https://github.com/tensorbee/rdocx/issues/217#issuecomment-5950052508) remain open. Their S81 comments cite the production and six-item deck-chain evidence on `main`. F-X158 in S82 still owns the reporter's exported deck fixture and criterion-level closure check. The accepted viewer evidence uses pinned LibreOffice 26.2.5.2 and python-pptx 1.0.2 and does not claim a PowerPoint 16.104 observation.

## S82 issue reconciliation

No contributed PR was assigned to the S82 fixture and evidence wave. The verified [S82 main merge](https://github.com/tensorbee/rdocx/commit/470e3781811f34044d9a3acb2f5cb8af2d09a65b) carries all three completed F-IDs. Issues [156](https://github.com/tensorbee/rdocx/issues/156#issuecomment-5957215809), [164](https://github.com/tensorbee/rdocx/issues/164#issuecomment-5957230429), [166](https://github.com/tensorbee/rdocx/issues/166#issuecomment-5957233694), [169](https://github.com/tensorbee/rdocx/issues/169#issuecomment-5957237108), [170](https://github.com/tensorbee/rdocx/issues/170#issuecomment-5957245288) and [217](https://github.com/tensorbee/rdocx/issues/217#issuecomment-5957248701) received individual evidence and reporter-credit comments and were closed.

Issues [158](https://github.com/tensorbee/rdocx/issues/158#issuecomment-5957255636), [160](https://github.com/tensorbee/rdocx/issues/160#issuecomment-5957258109) and [226](https://github.com/tensorbee/rdocx/issues/226#issuecomment-5957260298) remain open. F-X161 in S84 owns the compact Word part and namespace gap in Issue 160. F-X163 was scheduled in S86 at S82 close and is now in S84 for the plain-line spacing gap in Issue 226. Issue [163](https://github.com/tensorbee/rdocx/issues/163#issuecomment-5957261891) remains closed after its missing bookmark-after-table evidence passed on the S82 main merge. The original 22-issue ledger now has 19 passing rows and three unresolved rows.

## S83 contribution and issue reconciliation

The verified [S83 main merge](https://github.com/tensorbee/rdocx/commit/ea3fdbe1dd534a088c2be36f146fa945ee81cd9e) carries F-X159 and F-X160. PRs [240](https://github.com/tensorbee/rdocx/pull/240#issuecomment-5959494283), [248](https://github.com/tensorbee/rdocx/pull/248#issuecomment-5959495610), [249](https://github.com/tensorbee/rdocx/pull/249#issuecomment-5959496595) and [250](https://github.com/tensorbee/rdocx/pull/250#issuecomment-5959497876) received individual contributor-credit comments and were closed as superseded by the integrated work. PR 249's broader `insert_document` remapping question remains outside the same-part drawing-ID fix.

Issues [243](https://github.com/tensorbee/rdocx/issues/243#issuecomment-5959533687), [246](https://github.com/tensorbee/rdocx/issues/246#issuecomment-5959534625) and [247](https://github.com/tensorbee/rdocx/issues/247#issuecomment-5959535490) received reporter-credit comments citing the integrated regression and direct Python or CLI evidence and were closed. Both Issue 243 reporter packages passed. Issue 160 remains open for F-X161's compact edited-part and namespace criteria in S84, and Issue 245 remains open for that same serialization work.

## S89 contribution reconciliation

PR [269](https://github.com/tensorbee/rdocx/pull/269), contributed by Changjoon (`changjoon-park`), is superseded by the F-X177 integration at `8cced41d`. The original memmap compile reproducer fails before the fix and passes after it. Default, no-default, memmap and no-default plus memmap configurations passed on the integrated tree. Close the PR after pushing the S89 main merge, with an individual thank-you comment linking the integration and the sprint merge. No separate issue was assigned to S89. Issue 264's declarative authoring discussion remains open and outside this sprint's acceptance scope.

## S90 contribution reconciliation

PR [271](https://github.com/tensorbee/rdocx/pull/271), contributed by Hadrien
Mary (`hadim`), is integrated through F-X179 with an additional reviewed fix
for nested legacy reply removal. Every [Issue 270](https://github.com/tensorbee/rdocx/issues/270)
criterion passes, including XML parent and done rows and multiline authoring
through all existing entry points. Reconcile both records only after the
verified S90 merge is pushed through `/close-sprint`, with contributor thanks
and links to the feature integration and main merge. Issue 264 is excluded.


Issue [281](https://github.com/tensorbee/rdocx/issues/281), raised by Hadrien
Mary (`hadim`), is assessed in the existing F-184 decision at S95. The go
requires a material combined-lifecycle gap and an explicit preserve, edit and
execute boundary. The roadmap stages reading and loss-aware editing before
calculation and pivots. Rendering and distribution remain scheduled for S104.
The request for a unified CLI/Python workflow and an earlier usable package
adds integration and release tradeoffs to that decision. The existing
[maintainer response](https://github.com/tensorbee/rdocx/issues/281#issuecomment-6058792996)
answers those questions and invites concrete workflows and preservation needs.
Keep the discussion open for F-184 and reconcile this existing response during
verified S90 sprint close.

The [reporter follow-up](https://github.com/tensorbee/rdocx/issues/281#issuecomment-6063901190)
adds concrete Google Sheets review requirements: byte-preserve unmodeled
comment metadata, validation, conditional formatting and names, read typed
cells and formula caches, edit with styles and atomic output, and keep shared
formulas intact. Explicit stale-cache handling is required. The first-cut
alternatives are recalculate-on-load or a bounded basic evaluator. Unified
installation and API shape, Rust performance and safety, and later rendering
are the differentiators. Chart authoring is not requested. HLD14 and the S95
plan retain these inputs without starting conditional spreadsheet work in S90
or changing the reviewed release boundary. Keep Issue281 open for F-184.

| F-X186 | S90 | M | 2 | 1 | 2026-10-08 | Identity-preserving native/Python/CLI comment movement for hadim Issue284 and adapted PR287, ALL pass4 clean, facade1770/CLI65/Python94, hash49 unchanged |

| F-X187 | S90 | M | 2 | 1 | 2026-10-08 | Owned native and Python scoped replacement for hadim Issue285 and adapted PR286, direct-control correction from PR287, ALL pass2 clean, native2422/Python97/CLI11, hash49 unchanged |

| F-X188 | S90 | L | 4 | not recorded | 2026-10-09 | Issues288/292 whole-story ownership and qualified glossary omission, ALL pass3 clean, native1878/Python99, hash49 unchanged. Actual effort was not measured |

| F-X180 | S90 | S | 1 | not recorded | 2026-10-09 | All7 Issue272 nil/none topologies, ALL pass2 clean, durable final scoped attestation, hash49 unchanged and declared separate dense-form segment delta |
| F-X181 | S90 | S | 1 | not recorded | 2026-10-09 | Full Issue273 cell/field/break controls, ALL pass1 clean,19 retained logs rehashed, hash49 unchanged |
| F-X183 | S90 | M | 2 | not recorded | 2026-10-09 | Issues276/278 bounded legacy/default policies,600cases and accepted-empty-row proof, ALL pass2 plus three DELTA reviews,14 changed keys, qualified combined regression receipts |

| F-X190 | S90 | M | 2 | not recorded | 2026-10-09 | Issue291 existing snapshot repair, ALL pass2 clean after measured namespace lifetime correction, native1822/Python105, hash49 unchanged |

| F-X191 | S90 | S | 1 | not recorded | 2026-10-09 | PR293 namespace-only numbering reader correction, ALL pass1 clean, native1825, all12 scoped stages passed, hash49 unchanged |

| F-282 | S90 | L | 4 | not recorded | 2026-10-09 | Revised measured native bibliography contract, ALL pass4 clean, qualified native4160/all1340 APA and current layout320/Python178/CLI66, verified22 packages, hash49 unchanged. Full catalogue F-X192 and navigation F-283 in S91 |

| F-X189 | S90 | L | 4 | not recorded | 2026-10-09 | Word0.16.0 and shared/PowerPoint0.14.0 local preparation, ALL pass3 clean, actual22crate archives, full native5708 and installed Word178/PPT77 per interpreter, WASM/npm/size and golden7 passed, hash49 match. Final HEAD review, hosted rehearsal, close and publication remain separate |

## Next sprint replan, 2026-10-09

Insert S91 before the template and form programme. F-X192 and F-283 retain
their S91 carryover destination. F-X193 through F-X197 cover PR294 and
Issues295 through300, bringing the new boundary to seven pending stories and
20 to 25 estimated developer-days. Issues264 and281 remain excluded. F-X178's
next eligible planning target becomes S92 with its separate scope decision.
Issue281 remains open for F-184 at S96.

The previous S91 through S104 roadmap shifts to S92 through S105. Historical
S90 estimates, delivered rows and contribution reconciliation remain unchanged.
The S90 summary's original F-X178 destination is superseded by this replan.
No S91 completion, measured velocity, contribution closure or release is
recorded by planning. AS_BUILT.md remains the append-only delivered record.

S91 contribution intake stays open at the user's instruction. Refresh issues
and PRs at design, implementation/resume checkpoints, integration and final
verification/close. Attach later contributor PRs to their existing issue's
F-ID, inspect updated heads before duplicating implementation, and reuse
accepted work against the full contract. Update the canonical plan, backlog,
HLD14 and active sprint record for added scope. Issues264 and281 stay excluded.
The rolling intake procedure is defined in SPRINT_PLAN.md under S91.

At S91 opening, the rolling refresh adds hadim Issue301 through F-X198. The
scope is eight pending stories, estimated at 22 to 28 developer-days. No
implementation or acceptance is claimed by this intake.

| F-X196 | S91 | S | 1 | not recorded | 2026-10-09 | Dependency-prefix completion, private qualified CLI scan, ALL pass2 clean,4unit/67integration and141policy passes, hash49 unchanged, exact integrated source congruence. Final sprint gate remains due |

| F-X198 | S91 | M | 2 | not recorded | 2026-10-09 | Dependency-prefix completion, qualified paired-box/cell atomic refusal,40native/106installed controls, clean production/integration reviews, recovered unchanged consumer and remaining targets, hash49 unchanged, final actual archives/source bytes. Full sprint gate remains due |

| F-X193 | S91 | S | 1 | not recorded | 2026-10-10 | Scoped completion, PR294 semantic mark inheritance, clean production/integration reviews, 11HTML/514library/364integration controls across qualified retries/2331regression/2doctests, policy terminal141 with unavailable tool exit, hash49 unchanged, current F-X198 archives/source congruence. Full sprint gate remains due |

| F-X199 | S91 | S | 1 | not recorded | 2026-10-10 | Scoped prerequisite completion, PR302 qualified empty margins,54table/public21cell/639OXML/515library/2325regression/2doctest controls and qualified consumer recovery, policy141 and22patched publication dry-runs, clean independent integration review, hash49 unchanged and actual combined archives. Full sprint gate remains due |

| F-X194 | S91 | M | 2 | not recorded | 2026-10-10 | Scoped prerequisite completion, editable cell text and signed indentation, installed199/86 plus11 focused controls, native and typing gates, WASM, verified22 packages, policy141 with2 publication-dependent skips, five clean reviews, exact source congruence and hash49 unchanged. Full sprint gate remains due |

| F-X212 | S91 | M | 2 | not recorded | 2026-10-10 | Scoped prerequisite completion, qualified native/CLI visibility and field caches, clean independent reviews, OXML647/Word521/regression2331/doc2/CLI3+68 and qualified365 integration coverage, integrated focused9 and observed49 match, reviewed3-key delta, verified3 affected packages and policy141 with2 publication-dependent skips. Full sprint gate remains due |

| F-X197 | S91 | L | 4 | not recorded | 2026-10-10 | Scoped prerequisite completion, shared RGB validation/live colour/theme/brightness, installed652 and typing/stubs, actual Word354 R14 reuse and fresh PPT583, native4/support7/WASM, unchanged49 and verified22 packages, policy141 with2 publication-dependent skips, clean reviews. Integrated seven-command prefix passed with ROOT authentication. Full sprint gate remains due |

## S91 scope reduction, user decision

The user parks bibliography F-X192 and navigation F-283 at S92 and keeps all
frozen issue and PR repairs in S91. Twenty active stories remain, seven done,
three in progress and ten approved pending. Nineteen PRs and twenty-six
included issues retain their existing acceptance owners. Issues264 and281
remain excluded and open. No new contribution intake is authorized.
Preserve the dirty F-X192 worker branch, worktree and required evidence.
No completion, velocity, issue closure or release is recorded by this replan.
The earlier S91 carryover destination and rolling-intake instructions above
are historical and superseded by this decision and the existing intake freeze.
S92 is the planned next sprint and is not opened by this scope change.

## S91 rolling intake reopened, latest user decision

The user reopened issue and PR intake while retaining F-X192 and F-283 at
S92. The authenticated inventory has thirty open issues and twenty-one PRs,
with twenty-eight included issues after264 and281 are excluded. New339/342
story-text work extends existing F-X203. New340/341 TOC topology repair is
being designed separately from the parked full navigation contract. Assess
immutable source once before adoption, and reconcile later arrivals through
final close. No completion or contribution closure is recorded by intake.
The prior freeze instructions are superseded by this explicit user decision.

### S91 new contribution ownership, 2026-10-11

The reopened inventory contains28 included issues and21 PRs. New
Issue339/PR342 extend F-X203. New Issue340/PR341 form approved F-X213,
a bounded TOC repair separate from parked F-283. S91 has21 active stories,
seven complete, three running and eleven approved pending implementation.
Full immutable source has been assessed, with preservation gaps recorded
for correction and local acceptance. No new contribution is marked complete.
Issues264 and281 remain excluded and open.

| F-X195 | S91 | L | 4 | not recorded | 2026-10-11 | Scoped prerequisite completion, caller-font and variable-style rendering, installed104/caller19/eight SVG owners, declared risk gates and clean reviews. Integrated native6 and typing/provenance5 passed, unchanged49 and nine exact archive inventories. Historical source and ignore qualifications retained. Full sprint gate remains due |

| F-X210 | S91 | L | 4 | not recorded | 2026-10-12 | Scoped prerequisite completion, approved font and symbol assets with bounded compatibility and same-render diagnostics. Worker native/installed/typing/paint/package gates qualified. Integrated seven checks and49 matching hashes pass, final package measurement within64 bytes, reviews29/30 clean. Full sprint gate remains due |

| F-X213 | S91 | M | 2 | not recorded | 2026-10-12 | Scoped completion, entry-owned TOC boundaries and exact opaque preservation. Native21 and full affected suites, installed2 and typing, policy141, verified22 archives and clean integration review17. Composed TOC2 and Word/CLI checks pass on completed F-X210 prefix. Full sprint gate remains due |

| F-X206 | S91 | L | 4 | not recorded | 2026-10-12 | Scoped completion, shape authoring and causal live handles, native affine and source-preserving XML. Current11-crate checks, installed8, typing, verified8 archives, policy successors and49 hashes pass. Integrated source exact and reviews53 to59 clean. Full sprint gate remains due |

| F-X200 | S91 | L | 4 | not recorded | 2026-10-12 | Scoped completion, Word formatting and list parity, genuine Base and scoped current gates, clean independent reviews, exact source/archive composition and seven actual integrated checks including49 approved hashes. Full sprint gate remains due |

| F-X207 | S91 | L | 4 | not recorded | 2026-10-12 | Scoped completion, PowerPoint explicit autofit and fit_text, genuine Base and scoped current gates, clean independent reviews, exact source/archive composition and seven actual integrated checks including49 approved hashes. Full sprint gate remains due |

| F-X208 | S91 | L | 4 | not recorded | 2026-10-12 | Scoped completion, legal footer/field/theme/transition authoring, genuine Base and qualified native/installed/archive gates, clean integration review14, nine actual integrated checks including49 approved hashes. Original stalled attempts remain failed. Full sprint gate remains due |

| F-X201 | S91 | L | 4 | not recorded | 2026-10-12 | Scoped completion, live section stories and checked page authoring, native Word and LibreOffice evidence, clean integration review23, integrated11 regressions and49 matching hashes. Composite policy139 passes and2 existing skips. Full sprint gate remains due |

| F-X205 | S91 | L | 4 | not recorded | 2026-10-12 | Scoped completion, PowerPoint parity and exports,40 installed controls, seven verified archives, clean integration review20 and10 integrated rows including49 matching hashes. Original failures retained. Full sprint gate remains due |

| F-X203 | S91 | L | 4 | not recorded | 2026-10-12 | Scoped completion, shared Word operations and issue repairs, installed31 composite and typing, native367 integration, verified four archives and clean integration review17. Composed native/binding checks, exact chart paint and49 hashes pass. Policy139 qualified with2 existing skips and21 README cases compiled. Original failures retained. Full sprint gate remains due |

| F-X209 | S91 | L | 4 | not recorded | 2026-10-12 | Scoped completion, physical master/layout/theme editing, explicit native and public composites, current installed logo and typing, three verified archives and clean independent integration review30. Composed six-crate check and49 matching hashes pass. Original failures retained. Full sprint gate remains due |

| F-X202 | S91 | L | 4 | not recorded | 2026-10-12 | Scoped completion, tables and physical pictures, supported heading owners and synchronized alternate copies. Current79 public controls and typing, native Office acceptance, two verified archives and policy139 qualified passes with2 existing skips. Integrated product exact, review14 clean, six-crate check and49 hashes pass. Full sprint gate remains due |

| F-X214 | S91 | L | not recorded | not recorded | 2026-10-12 | Scoped completion, shared depth32 parsing and iterative traversal, exact opaque source and atomic refusal. Native956 qualified passes and34 existing ignores, installed controls and typing, native PowerPoint acceptance, four verified archives and policy139 passes with2 existing skips. Product exact, integration review5 zero defects and smells, six-crate check and49 hashes pass. Full sprint gate remains due |

| F-X204 | S91 | L | not recorded | not recorded | 2026-10-12 | Scoped prerequisite completion, ordered counted maps, complete native and Python field reports, strict timestamps and atomic CLI publication. Native2396 regression passes, installed allocation-failure rollback and typing, four unattended LibreOffice controls, five verified archive dry runs and exact policy successors. Three clean reviews, eight integrated checks and49 unchanged hashes. Full sprint gate remains due |

| F-X222 | S91 | M | not recorded | not recorded | 2026-10-12 | Scoped completion of Issue362 and PR363, operation-owned inventory and preserved qualified empty owners. Native567 library and371 integration controls,2396 regressions, four unattended LibreOffice successors, fresh Word wheel and CLI, typing, WASM, verified archive and policy141. Independent full review clean, six integrated checks and49 unchanged hashes. Full sprint gate remains due |

| F-X217 | S91 | L | not recorded | not recorded | 2026-10-13 | Scoped completion of Issues350/351 and PR359, HTML/CommonMark import through native, installed Python and CLI. Source-qualified gates, exact parser approval, clean reviews and49 unchanged hashes. Both native Word combined/template open/save/normal-reopen/local-render cases pass.3211 integrated source files checked, product exact. Final full sprint gate remains due |

| F-X211b | S91 | L | not recorded | not recorded | 2026-10-13 | Scoped PowerPoint XML child completion, nine checked physical owners,24 native and31 installed controls, clean whole/composition/integration reviews, five verified composed archives and actual integrated compile/release-owner/hash49/prose checkpoint. Original debug abort and historical absent oracles retained. Parent211 and final full sprint gate remain due |

| F-X220 | S91 | L | not recorded | not recorded | 2026-10-13 | Scoped licensed font embedding across native, Python and CLI families. Native ten-case evidence retains approved Word limit/F223 carryover, current62 installed tests plus typing/fullstubs, eight verified archives, both WASM and49 unchanged hashes. Independent integration and metadata reviews clean. Original failures retained. Full sprint gate remains due |


| F-X211a | S91 | L | not recorded | not recorded | 2026-10-13 | Scoped dependency-prefix completion. Five Word checked XML owners and concrete source-qualified publication, bounded ordinary-stack parser lifecycle, opaque namespace preservation and owner-local child retirement. Prepared returned Paragraph capture composes with chart creation, imports and fonts. Parent/c and final full sprint gates remain due |


| F-X215 | S91 | L | not recorded | not recorded | 2026-10-13 | Scoped dependency-prefix completion. Both-family Python chart creation and preservation-aware data replacement retain actual cache/workbook graphs, styles and owned relationships. Shared targets use copy-on-write. Chart views retire after checked frame XML replacement, and Word returned Paragraph allocation uses prepared owner authority. Native chart source/render acceptance remains attributed to its actual inputs. Parent/c and final full sprint gates remain due |


| F-X218 | S91 | L | not recorded | not recorded | 2026-10-13 | Scoped dependency-prefix completion. PowerPoint semantic HTML/Markdown export preserves supported source order, list/style/link/table/image/notes/hidden-slide semantics with immutable detached diagnostics/assets. Native atomic publication and one bounded image validator back Python and CLI. Two native source cycles qualify source semantics separately from HTML pixel parity. Parent/c and final full sprint gates remain due |

| F-X216 | S91 | L | not recorded | not recorded | 2026-10-14 | Scoped dependency-prefix completion. Complete slide graphs, source formatting and deck merge with admitted native source comparisons, graph survival and save/reopen acceptance. Independent integration review clean. Four historical oracle assets uncredited in S92. Final union gates remain due |


| F-X211d | S91 | L | not recorded | not recorded | 2026-10-14 | Scoped Word compatibility dependency completion. Installed32 and typing, ordinary2MiB12, pinned native complements, verified22 archives, full141 policy and README qualified. Independent completion and integration reviews clean. Parent unchanged100 and final sprint gates remain due |

| F-X211c | S91 | L | not recorded | not recorded | 2026-10-14 | Scoped completion,95 qualified files from unchanged100, all42hard51native gates and24documentation cases, five no-credit divergences. Paired controls, nine raster and native Word morphology, verified22 archives, policy154pass2skip and clean whole/integration reviews. Final full sprint gate remains due |

| F-X211 | S91 | L | not recorded | not recorded | 2026-10-14 | Aggregate completion of a/b/c/d. Original Base17/100 and admitted final95/100 remain distinct. All42 hard gates,51 native clauses and24 documentation examples qualify. Independent aggregate review clean. Release preparation and final sprint gates remain due |

| F-X221 | S91 | L | not recorded | not recorded | 2026-10-15 | Prepared Word0.16.1 and shared/PowerPoint0.14.1. Final local full gate passes: stable native core, Python778/854, font104/native-eight, policy158pass2skip,22 archives, README and deny. Admitted95of100 retains42hard51native/24documentation and five no-credit cases. Component reviews clean. Subsequent bare-wheel routing amendment passes exact37/41 priority runtime and policy159pass2skip with four strengthened controls. Final integrated successor review, hosted rehearsal, close, both publications and humanized remote reconciliation remain separate gates |

| F-X224 | S91.1 | M | 2 | not recorded | 2026-10-15 | Pinned native HB/animation and fractional-tag routing, independently reviewed three-value F210 external raster propagation. Corrected hosted receiptad03f29ef28fe9a5174cdb7cb30a8811c4cb25fa237eac4df1b8a0d63d388792. Final sprint/main/release gates remain due |

| F-X225 | S91.1 | M | 2 | not recorded | 2026-10-15 | Private parser storage preserves public/XML contracts and passes original plus populated default1MiB Node callers. Scoped native/WASM/archive/hash/policy and independent reviews qualified |

| F-X226 | S91.1 | M | 2 | not recorded | 2026-10-15 | Targeted checked-owner relationship retirement preserves shared/opaque graphs, debug validation and atomicity. Actual debug and optimized public regression plus scoped gates qualified |

| F-X227 | S91.1 | L | 4 | not recorded | 2026-10-15 | Actual ownerCI38702668217 qualifies95/100,42hard51native24docs and5no-credit with safe one-job cleanup. Exact source215 rehearsal artifacts and installed156/typing4 qualify their recorded source. Final gates/publication pending |

## External benchmark sprint insertion, 2026-10-15

Insert S96 before the conditional spreadsheet programme. F-313 through F-318
answer Issue 377 with six pending stories estimated at 21 developer-days. The
previous S96 through S105 spreadsheet roadmap shifts to S97 through S106.
Issue281 remains open for F-184, now at S97. Historical delivered rows and
earlier replan records remain unchanged. No completion, velocity or release is
recorded by planning.

## Repair sprint insertion, 2026-10-15

Plan S92 and S93 for the contribution and issue-repair wave, with nine new
pending F-IDs F-X228 throughF-X236 and existing F-X178. Shift the previous
S92 throughS106 future roadmap by two. Preserve the separately added external
Word benchmark sprint and F-313 throughF-318 at S98, before Excel starts at
S99. The roadmap now ends at S108. S91.1 scope and release preparation remain
unchanged. Historical sprint IDs, tags and delivered rows are unchanged.
Issues264,281 and377 remain open. This records planning, not delivery.


| F-X229 | S92 | L | 4 | not recorded | 2026-10-15 | Correctness-only PR374 subset repairs hyperlink, empty paragraph, table-adjacent and trailing deletion reconstruction. Scoped worker gates and canonical focused reconciliation qualify dependency-prefix completion. Final sprint gate remains due |


## Issue378 execution intake, 2026-10-15

Add pending F-X237, size L estimated4 developer-days, to S92 at the user's explicit request. The six-story sprint estimate is now15 to18 developer-days. This records scope intake only, with no completion or velocity credit. All original repair acceptance remains intact.

## Issue379 through Issue381 sprint insertion, 2026-10-15

Insert S94, Python and CLI parity, for Issues379 and380 as F-X238 through
F-X241, about 12 developer-days. Insert S96, Fillable PDF forms, for Issue381
as F-X242 through F-X244, about 13 developer-days. Planned S94 through S98
work moves to S95 and S97 through S100. The conditional Excel programme moves
to S101 through S110, and Issue281's F-184 decision is at S101. The active S92
wave and S93 plan are unchanged. No completion, velocity or release is
recorded by planning.

| F-X178 | S92 | S | 1 | not recorded | 2026-10-15 | Approved repair contract, scoped regression and preservation gates, clean independent feature/integration review, final integrated native and installed risk riders, verified22 archives and49 unchanged hashes. Actual evidence retains source epochs and user-accepted F-X237 fill tradeoff |

| F-X228 | S92 | M | 2 to 3 | not recorded | 2026-10-15 | Approved repair contract, scoped regression and preservation gates, clean independent feature/integration review, final integrated native and installed risk riders, verified22 archives and49 unchanged hashes. Actual evidence retains source epochs and user-accepted F-X237 fill tradeoff |

| F-X230 | S92 | M | 2 to 3 | not recorded | 2026-10-15 | Approved repair contract, scoped regression and preservation gates, clean independent feature/integration review, final integrated native and installed risk riders, verified22 archives and49 unchanged hashes. Actual evidence retains source epochs and user-accepted F-X237 fill tradeoff |

| F-X231 | S92 | M | 2 to 3 | not recorded | 2026-10-15 | Approved repair contract, scoped regression and preservation gates, clean independent feature/integration review, final integrated native and installed risk riders, verified22 archives and49 unchanged hashes. Actual evidence retains source epochs and user-accepted F-X237 fill tradeoff |

| F-X237 | S92 | L | 4 | not recorded | 2026-10-15 | Approved repair contract, scoped regression and preservation gates, clean independent feature/integration review, final integrated native and installed risk riders, verified22 archives and49 unchanged hashes. Actual evidence retains source epochs and user-accepted F-X237 fill tradeoff |

| F-X234 | S93 | M | 2 to 3 | not recorded | 2026-10-15 | Scoped native640/380/2418/2, policy182with2expected skips, pinned Pandoc3.11, clean independent reviews and focused prefix integration pass. Worker49hashes unchanged. Final full sprint gate and review remain due |
| F-X232 | S93 | L | 4 to 5 | not recorded | 2026-10-16 | Legal paint/cache/native-CLI-Python and exact checked XML qualification, approved Rust source changes, clean feature/metadata review. Two reviewed underline hashes overlap border union. Native/Python accepted at16f1fc1, full Rust accepted at16f1fc1, sprint review pending |
| F-X233 | S93 | M | 2 to 3 | not recorded | 2026-10-16 | Native component-qualified border frame/flow/wrap repair on inferred inputs, missing original retained. Reviewed18entry union and4pixel buffers, D1 warm-cache12point gap repaired, regression RED then layout352+1doctest, independent cache pass2clean and original external3pairs equal. Native/Python accepted at16f1fc1, full Rust accepted at16f1fc1, sprint review pending |
| F-X235 | S93 | L | 4 to 5 | not recorded | 2026-10-16 | Six matching-font native equation controls, unchanged1point/.99bounds, composite mutation rejects5of6. Historical identities and49worker hashes retained. Clean source/metadata/integration reviews. Native/Python accepted at16f1fc1, full Rust accepted at16f1fc1, sprint review pending |
| F-X236 | S93 | S | 1 | not recorded | 2026-10-16 | Six selective Run inline annotations supported by real non-LTO nine-pair benchmark, assembly and ten-case saved/reopened semantic controls. No end-to-end speedup claim, clean reviews,49worker hashes unchanged. Native/Python accepted at16f1fc1, full Rust accepted at16f1fc1, sprint review pending |
| F-X245 | S93 | L | 4 to 5 | not recorded | 2026-10-16 | Relationship-resolved chart/workbook accept/reject reconstruction, original fixed-member observer retained separately, shared/cycle/bound/atomicity controls and clean reviews.49worker hashes unchanged. Native/Python accepted at16f1fc1, full Rust accepted at16f1fc1, sprint review pending |
| F-X246 | S93 | L | 4 to 5 | not recorded | 2026-10-16 | Both-family legend rendering and Python APIs, native component-only qualification, ownership/opaque/atomicity controls and clean reviews.49worker hashes unchanged. Native/Python accepted at16f1fc1, full Rust accepted at16f1fc1, sprint review pending |

## S93 close evidence

The reviewed sprint HEAD is55122db683aa62991a345db855c477947416c1a6. Its fresh full-gate receipt is /private/tmp/S93-full-closure/receipt.json, SHA256 2092ffc87cef3de035546d2b614739b682091ec4559d4c27d015489365f7986a. Native successor authentication is /private/tmp/S93-native-final/final-head-native-closure-55122db683aa62991a345db855c477947416c1a6.json, SHA256 5c2a07594268b36a9b71de595991f6c4fbc411b00fc830687e524c7d909bbdb8. Installed Python successor authentication is /private/tmp/S93-python-final-head-closure/run-55122db683aa62991a345db855c477947416c1a6/receipt.json, SHA256 2d250640fba42924fe2ae7bb6c5b658c28f31bf31ef9bfc7dc84d4c05fe2ad14. These closures retain native/Word actual source16f1fc1 and PowerPoint actual source12fa271, not fresh external execution at the documentation successor. The final Rust gate is fresh at the reviewed HEAD.

The close workflow adds only this tracker summary, velocity and evidence record after preflight. Its commit and the main merge/tree comparison are recorded in the local close receipt. No product source, harness, dependency, public surface or acceptance contract changes in that close-only step. All seven clean worker directories were removed after integration and their branches are retained until the main and s93 pushes succeed. Dirty, carried and unrelated worktrees remain protected. Docker inventory contains no rdocx stack or matching volume. Cache retirement preserves source, native proof artifacts and required binaries, with an observed130.526GiB free-space increase from the two audited canonical incremental directories.

The issue and PR reconciliation follows the successful main push. PR375 is superseded by the qualified F-X232 integration. Issues373,366,367,382and383 have their bounded evidence and stated limits. Issue365 retains pending F-X247 in S94 for its broader73input conversion matrix. Actual remote comment URLs and open/closed decisions are recorded after posting in the close receipt, rather than claimed in advance here. No release tag or publication is prepared by S93.

| F-X238 | S94 | L | 4 to 5 | not recorded | 2026-10-16 | Named Word compatibility, exact corpus qualifications and section module. Shared actual 1,782 installed tests, 23 artifacts and 27 hosted jobs and complete native gate accepted with documented producer and continuation boundaries. Full native receipt 594b2f351f2effb28bae678136732ab4dfadb117b95d7f29231f7c62d3a10992. Sprint review and final successor gate remain due |
| F-X239 | S94 | M | 2 to 3 | not recorded | 2026-10-16 | Owned notes, enum aliases and exact silent-zero refusal, with pinned controls and qualified S93 proof correction. Full native receipt 594b2f351f2effb28bae678136732ab4dfadb117b95d7f29231f7c62d3a10992. Sprint review and final successor gate remain due |
| F-X240 | S94 | M | 2 to 3 | not recorded | 2026-10-16 | Current-state DER/PEM signing in both ordinary Python and CLI families, approved ZIP 9, all 27 hosted gates, precise test-only amendment and fresh verified 22 archives. Full native receipt 594b2f351f2effb28bae678136732ab4dfadb117b95d7f29231f7c62d3a10992. Sprint review and final successor gate remain due |
| F-X241 | S94 | M | 2 to 3 | not recorded | 2026-10-16 | Five-kind form filling and snapshots, atomic JSON/source admission and current fill/sign ownership controls. Full native receipt 594b2f351f2effb28bae678136732ab4dfadb117b95d7f29231f7c62d3a10992. Sprint review and final successor gate remain due |
| F-X247 | S94 | L | 4 to 5 | not recorded | 2026-10-16 | Finite 73-input conversion, actual Pandoc 72-input emission plus degree fallback, seven bounded native positives/mutations and qualified historical/nonregression limits. Full native receipt 594b2f351f2effb28bae678136732ab4dfadb117b95d7f29231f7c62d3a10992. Sprint review and final successor gate remain due |

## Issue384 and Issue385 repair intake, 2026-10-16

At the user's request, prioritize seven new pending repairs over the eleven carryovers. S95 has five M stories estimated at 10 to 15 developer-days. S95.1 contains the remaining two new repairs and two native interoperability carryovers, followed by the bibliography/navigation and template boundaries at S95.2 through S95.4. The full 18-story programme totals 54 to 72 developer-days. S96 onward and existing M24 story totals remain unchanged. Issue384 belongs to S95. Issue385 remains open until its S95 and S95.1 acceptance union passes. GitHub refresh3 found no newer issue or PR. This is planning, with no completion or velocity credit for the new repairs.

## S94 close evidence

Reviewed sprint HEAD: 588f26407253655f9e85cf9f53499a9148b8deb2. Committed sprint review pass2 audits parent 76000663 and records 0 blocking, 0 should-fix and one retained optional formatting nit. Fresh full verification at the reviewed HEAD passed all 18 native/repository commands. Receipt `/private/tmp/S94-full-closure-588f2640-1/receipt.json` has SHA-256 `3cdceed380b24c22d6a66326a92e5fdb6d007423ef47d382634b24284c241eef`. Root and independent authentication confirm 5,493 default passes with 68 ignored, 6,579 overlapping all-feature passes with 73 ignored, 149 no-default font passes with one ignored, 183 policy tests with two expected skips, 49 unchanged hashes, seven pixel buffers, 22 genuinely verified publication dry runs and 22 retained archives below 10 MiB. The actual failed 3bb policy run is retained. Its decimal-sprint parsing error was repaired in the single existing policy test at 76000663 and the complete policy suite reran successfully. The private authentication helper needed workspace-inherited version resolution, with its failed predecessor retained, without changing product code or the successful gate.

This close-only tracker successor follows successful preflight and gates. Its commit, main merge and exact tree comparison will be recorded in the root close receipt after those operations. It changes no product source, public surface, dependency, baseline or acceptance contract. Historical per-feature rows retain their original evidence qualifications. Installed 877 Word and 905 PowerPoint tests, 23 new public artifacts, strict typing and 27 hosted jobs remain actual producer 9cdbea, with four publication guards skipped. The authenticated 3,533-input successor inventory explicitly carries the prior test-only, archive-metadata and new policy-test exceptions rather than claiming universal source equality. All five contracts are done with no carries. New Issues384 and385 retain their S95 and S95.1 acceptance homes, without S94 delivery credit. The latest intake refresh found no newer open item or PR. All five clean worker directories were already removed after integration. Their recorded branches are cleanup targets only after main and s94 push successfully. Dirty, carried and unrelated worktrees remain protected. Issue reconciliation follows the successful main push, with exact decisions and comment links recorded then. No release or publication is part of this close.

## S95 release dependency-prefix checkpoint

Four integrated prerequisites complete through scoped evidence so the end-S95 release preparation can follow all five repairs. F-X252 remains in implementation and F-X255 pending. No full sprint gate, sprint review, publication or measured person-effort credit is claimed. The new release story extends the estimate to 14 to 20 developer-days without moving later sprint boundaries.

| F-ID | Sprint | Size | Estimate days | Actual days | Completed | Evidence |
|------|--------|------|---------------|-------------|-----------|----------|
| F-X248 | S95 | M | 2 to 3 | not recorded | 2026-10-16 | Scoped acceptance, independent clean feature review and integrated source authentication. Receipt SHA256 5cad8051b7f35b202c9c88ae9c452ce349cbb0b870ca2544df05da56ebd334fd. Final full sprint gate and review remain due |
| F-X249 | S95 | M | 2 to 3 | not recorded | 2026-10-16 | Scoped acceptance, independent clean feature review and integrated source authentication. Receipt SHA256 5d095be06d08d3e74acdac972f682f27a8354cff8464269b9b3117827798c5f2. Final full sprint gate and review remain due |
| F-X251 | S95 | M | 2 to 3 | not recorded | 2026-10-16 | Scoped acceptance, independent clean feature review and integrated source authentication. Receipt SHA256 6ce4af6b199c3c337a72f44f9927e8d245633b81aa1f6f8baf3d07a85621d6c3. Final full sprint gate and review remain due |
| F-X253 | S95 | M | 2 to 3 | not recorded | 2026-10-16 | Scoped acceptance, independent clean feature review and integrated source authentication. Receipt SHA256 54a4fc73d167e4b00e36a24d98dd4a7517c38a85e7927f86c539024468023a87. Final full sprint gate and review remain due |

## Issue386 intake and capacity split, 2026-10-16

Read-only GitHub refresh found new Issue386 and no open PRs. Its new Word unit,
remaining Word names/legacy diagnostic and PowerPoint name/color/language
contracts are pending F-X256, F-X257 and F-X258 in S95.1, totaling 8 to 11
estimated developer-days. The former S95.1 F-X250/F-X254/F-X219/F-X223 group
moves intact to S95.11, at 12 to 16 developer-days, before S95.2. This preserves
all acceptance groups and the user's new-repair priority without extending
S95 or making its release wait for future repairs. No implementation, runtime,
completion or velocity credit is claimed. Historical references to the prior
S95.1 allocation retain their original planning date. The canonical roadmap
and backlog carry the revised homes. Issue385 and Issue386 remain open for
their respective remaining acceptance unions. Exact input and source facts
are retained in `/private/tmp/S95-issue386-intake-triage-1/triage.md` and its
receipt SHA256 `8024b68da12ba4f7a1d7c12265ad7b3a80935954d74a16cc133822fb470b0fa0`.

| F-X252 | S95 | M | 2 to 3 | not recorded | 2026-10-16 | Scoped prerequisite acceptance: native3480/28ignored, installed45, typing/stub, policy181pass2skip, genuine affected publication, hash49/pixels7 unchanged. Independent correctness3 and numeric packaging4 clean. Integrated94cd8abb preserves all12 qualified files. Final six-story full gate and sprint review remain due |

All five repair prerequisites are now complete. F-X255 can design and start
release preparation for Word0.16.2 and shared/PowerPoint0.14.2. The actual
full integrated gate, hosted artifacts and independent sprint review still
follow the final versioned source. No publication has started.

## S95 expanded dependency-prefix checkpoint, 2026-10-16

The user expanded active S95 to nine stories. F-X256 and F-X258 now complete
through scoped evidence and exact integrated product/test reconciliation.
F-X257 follows the Word prefix, then F-X255 reconciles its preserved release
preparation. Final integrated full verification and sprint review remain due.
Historical S95.1 allocation statements above retain their original date.

| F-ID | Sprint | Size | Estimate days | Actual days | Completed | Evidence |
|------|--------|------|---------------|-------------|-----------|----------|
| F-X256 | S95 | M | 2 to 3 | not recorded | 2026-10-16 | Scoped Word933 and focused33 at successor3, current typing-only applicability and stub checks, clean independent pass5, exact integration f3291339. Policy181pass2skip, worker hash49/pixels7 unchanged. Final sprint gate due |
| F-X258 | S95 | M | 2 to 3 | not recorded | 2026-10-16 | Scoped complete PPT961 with no skips, selected6 and typing/stub checks, clean independent pass1, exact integration8c9fb137. Policy181pass2skip, worker hash49/pixels7 unchanged. Final sprint gate due |

| F-X257 | S95 | L | 4 to 5 | not recorded | 2026-10-16 | Scoped complete Word947/0skip, focused14, strict typing/stub8modules and selected6 observations. Two clean reviews,27 root-authenticated references, exact integration1afb1006. Policy181pass2skip, hash49/pixels7 unchanged. Final integrated gate remains due |

All eight S95 repairs now complete through scoped dependency-prefix acceptance.
Release preparation F-X255 remains in-progress for revised design and preserved
WIP reconciliation. The final nine-story union, actual hosted acceptance and
sprint review remain due. No publication or measured person-effort credit.

| F-X255 | S95 | L | 4 to 5 | not recorded | 2026-10-16 | Exact Word0.16.2/shared-PPT0.14.2 preparation. Clean source/numeric reviews, scoped verified22 archives and actual6b final installed1908/no skips, hosted27success4publicationSKIP/26payloads/172 priority calls, public23, PPT four routes, Issue38619+16 and reporter12strictXPASS/raw1 plus12bodies/raw0. Integrated native18commands,49hashes,7pixels and22verified archives pass at6b7c9798. Formal sprint review and publication remain separate |

## S95.11 close evidence

Planned four contracts, completed F-X250/F-X254/F-X223 and carried F-X219 unfinished
to S96. Original estimate12 to16 developer-days, retained completed scope10 to13.
Actual worked calendar dates are2026-10-16and17, two dates, authenticated by
`git log main..HEAD --format=%cs`. Feature developer-effort actuals remain not
recorded. Existing calendar-throughput velocity is3/2*5 =7.50 stories/week,
83.33percent below S95's45.00. This variance records observed parallel calendar
throughput. Preserve the S95.2/S95.3/S95.4 capacity split and developer estimates,
without deriving elapsed-time promises from this figure. The
user's "document word limitation" disposition completes F-X223's bounded native
investigation without broadening its genuine-face qualification or erasing14
unsupported observer outcomes. No S96 implementation has started.

Accepted full source 05ba6cc8461a896fa751ca7dc1fe209a8fa70045 binds actual H3/H5/836/H7 observations, rather than
pretending all commands ran at one new metadata HEAD. Native union32 binaries
contains6616 unique passes and73ignored, plus11complete doctests. Actual ordinary
installed suites pass955Word and1007PowerPoint with strict typing and stubtest.
Issue385's27 distinct bodies join12prior and15new, each epoch separately recording
strict XPASS/raw1 and body PASS/raw0. Hash49/pixels7 unchanged. No-default fonts,
WASM, documentation, README and deny observations remain source-applicable.
Genuine package verification is21 applicable H5 crates plus current H7 CLI.
Source/archive/recovered-wrapper distinctions and actual failures remain intact.
F250 bounded Word/LO/own appearances and source XML are separate from F254 native
PowerPoint custody/XML and F223's documented native limitation.

Receipt identities: installed cff97a3605b2b32601881e5727cbc926b392294f66c79e02062fa0b6747d7805,
reporter15 6109789870f72fa7b7a4892b27ce8f7adf583fa3b8c97b717201f54f06306070,
runtime/doc partial99787cf314315b25f5e66d89812c15cc37423f133fd68b52b314b0a9b1e2dcda,
build-tail partial921e51101707792aba8d34bbc8792af35e36906986ad988319dcfc47c12cdca2,
and current CLI b656c2c672dc5ee6f3d18d6a82d457373d86715b8c6421d69c0d467f58789ea9.
Partial receipts retain passed:false with their explicit stage flags. The final qualification receipt SHA256
6bd76990051cb4e095fdbd673692aca7dfbc1398c3b7b2905f1dda29571b09d1
records passed:true at05ba6cc8461a896fa751ca7dc1fe209a8fa70045 by joining
actual external evidence, source correspondence and the user limitation
disposition, without replaying any producer. Policy aggregate is188 tests,
186 passes and two retained skips, with its original failed command retained. Formal sprint review and close operate against the
actual final metadata/source correspondence. No release tag, publication or new
version is part of this close. Fresh registry/publisher checks and separate final
approval at the exact reviewed main SHA remain release obligations.


## S96 planning reorganization, 2026-10-28

Read-only GitHub intake at 2026-10-28T03:03:24.502962+00:00 finds2open PRs and 8 open issues. S95.4 is already closed and immutable. The unopened old S96 feature bundle has no started story. S96 now reserves four pending F-X271 through F-X274 repair contracts, estimated12 to 17 developer-days including the integrated gate, explicitly larger than a normal sprint. No implementation, completion, actual days or velocity credit is recorded.

Old S96 splits into S97 AcroForms and S98 native/bibliography recovery to expose its prior19 to 21day estimate plus the unestimated F-X269 addition. Pending old S97 throughS110shift to S99 through S112. Historical allocations above retain their original identities. The existing draft F-X192/F-X219 contracts move current scheduling to S98without rewriting their old carry evidence. All issue381acceptance stays on F-X242/243/244. Issues 264/281/377 remain open conversations. No source, commit, remote state, tag or main merge changes in this planning run.
