# Backlog

Live status table for every F-ID. The detailed story descriptions
(acceptance gates, dependencies, sizes, test gates) live in
`docs/hld/14-development-backlog.md`. This file is the **execution-time
tracker** keyed by F-ID.

Statuses: `pending`, `in-progress`, `done`, `archived`.

Updated by `/complete-feature` (single-row updates) and `/sync-status`
(consistency audit). The counts inside the AUTOGEN sentinels are
regenerated, never hand-edited.

<!-- AUTOGEN:backlog-summary START -->
## Summary

| Milestone | F-IDs | Done | In Progress | Pending |
|-----------|-------|------|-------------|---------|
| M1, Preparation and safety net | 12 | 12 | 0 | 0 |
| M2, Shared infrastructure extraction | 10 | 10 | 0 | 0 |
| M3, Media | 6 | 6 | 0 | 0 |
| M4, Layout primitives | 8 | 8 | 0 | 0 |
| M5, PDF backend | 9 | 9 | 0 | 0 |
| M6, Shared publication and rdocx cutover | 6 | 6 | 0 | 0 |
| M7, DrawingML | 19 | 19 | 0 | 0 |
| M8, PresentationML | 14 | 14 | 0 | 0 |
| M9, Inheritance resolver | 8 | 8 | 0 | 0 |
| M10, Renderer | 20 | 20 | 0 | 0 |
| M11, Write API | 12 | 12 | 0 | 0 |
| M12, Charts | 12 | 12 | 0 | 0 |
| M13, Bindings and tooling | 18 | 18 | 0 | 0 |
| M14, Word collaboration layer | 9 | 9 | 0 | 0 |
| M15, Charts beyond PowerPoint | 4 | 4 | 0 | 0 |
| M16, Document automation | 10 | 10 | 0 | 0 |
| M17, Security and compliance | 7 | 7 | 0 | 0 |
| M18, Format breadth | 8 | 8 | 0 | 0 |
| M19, Advanced spreadsheets | 21 | 0 | 0 | 21 |
| M20, Fidelity at scale | 13 | 7 | 0 | 6 |
| M21, Presentation depth | 15 | 15 | 0 | 0 |
| M22, Word depth | 12 | 12 | 0 | 0 |
| M23, From-scratch business documents | 24 | 24 | 0 | 0 |
| M24, Modern DOCX authoring completeness | 54 | 32 | 0 | 22 |
| X, Cross-cutting | 289 | 275 | 3 | 7 |
| **Total** | **620** | **557** | **3** | **56** |
<!-- AUTOGEN:backlog-summary END -->

## All F-IDs

### M1, Preparation and safety net

<!-- AUTOGEN:backlog-M1 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-001 | Deterministic font mode                      | S01 | M | done |
| F-002 | rust-toolchain.toml                          | S01 | S | done |
| F-003 | Output-stability hash harness                | S01 | L | done |
| F-004 | Caladea licence and the false OFL claim      | S01 | S | done |
| F-005 | Fix the image counter                        | S01 | S | done |
| F-006 | Fix the JPEG standalone-marker walk          | S01 | S | done |
| F-007 | Resolve core properties through the rel      | S02 | S | done |
| F-008 | Non-consuming setter twins                   | S02 | M | done |
| F-009 | Cache the layout result                      | S02 | M | done |
| F-010 | Reserve crate names                          | S02 | S | done |
| F-011 | Pin unit truncation behaviour                | S02 | S | done |
| F-012 | Tag v0.4.1                                   | S02 | S | done |
<!-- AUTOGEN:backlog-M1 END -->

### M2, Shared infrastructure extraction

<!-- AUTOGEN:backlog-M2 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-013 | Create oxml-core                             | S03 | M | done |
| F-014 | New unit types                               | S03 | M | done |
| F-015 | rdocx-oxml becomes a facade                  | S32.2 | S | done |
| F-016 | Length re-export                             | S32.2 | S | done |
| F-017 | App and custom properties                    | S03 | M | done |
| F-018 | Create oxml-opc                              | S04 | M | done |
| F-019 | PresentationML relationship and content types| S04 | S | done |
| F-020 | oxml-opc reads a pptx                        | S04 | M | done |
| F-021 | Zip-slip hardening tests                     | S04 | S | done |
| F-022 | rdocx-opc deprecation shim                   | S32.2 | S | done |
<!-- AUTOGEN:backlog-M2 END -->

### M3, Media

<!-- AUTOGEN:backlog-M3 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-023 | oxml-media format sniffing                   | S05 | M | done |
| F-024 | Image probing and DPI                        | S05 | L | done |
| F-025 | MediaNamer                                   | S05 | S | done |
| F-026 | native_size with explicit DPI                | S05 | S | done |
| F-027 | rdocx adopts oxml-media                      | S32.2 | M | done |
| F-028 | add_picture_auto                             | S32.2 | S | done |
<!-- AUTOGEN:backlog-M3 END -->

### M4, Layout primitives

<!-- AUTOGEN:backlog-M4 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-029 | Create oxml-layout                           | S06 | M | done |
| F-030 | Decouple line.rs                             | S06 | L | done |
| F-031 | Transform                                    | S06 | M | done |
| F-032 | Path and PathCommand                         | S07 | M | done |
| F-033 | Paint and Stroke                             | S07 | M | done |
| F-034 | Path and Group arms                          | S07 | M | done |
| F-035 | The walk helper                              | S07 | S | done |
| F-036 | MediaId                                      | S07 | S | done |
<!-- AUTOGEN:backlog-M4 END -->

### M5, PDF backend

<!-- AUTOGEN:backlog-M5 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-037 | Create oxml-pdf                              | S08 | S | done |
| F-038 | Golden-PNG harness                           | S08 | M | done |
| F-039 | Global CTM flip                              | S08 | L | done |
| F-040 | Group rendering                              | S09 | M | done |
| F-041 | Path rendering                               | S09 | M | done |
| F-042 | Rewrite the three collection passes on walk  | S09 | M | done |
| F-044 | ExtGState alpha                              | S09 | S | done |
| F-043 | Gradient shading dictionaries                | S10 | L | done |
| F-045 | Rasteriser: groups, paths, gradients, dashes | S10 | L | done |
<!-- AUTOGEN:backlog-M5 END -->

### M6, Shared publication and rdocx cutover

<!-- AUTOGEN:backlog-M6 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-046 | rdocx layout and PDF cutover                 | S32.2 | M | done |
| F-047 | Packaging include and size gate              | S32.1 | M | done |
| F-048 | Automate split-family release preparation   | S32.1 | M | done |
| F-049 | Extend publish.yml to the extracted workspace| S32.1 | M | done |
| F-050 | CI matrix additions                          | S32.1 | S | done |
| F-051 | CHANGELOG and migration notes                | S32.2 | S | done |
<!-- AUTOGEN:backlog-M6 END -->

### M7, DrawingML

<!-- AUTOGEN:backlog-M7 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-052 | Create oxml-drawing and namespace constants  | S12 | S | done |
| F-053 | OrderedRawChildren                           | S12 | M | done |
| F-054 | Colour choices                               | S12 | M | done |
| F-055 | The colour transform stack                   | S12 | L | done |
| F-056 | Colour map resolution                        | S12 | M | done |
| F-057 | a:xfrm                                       | S13 | M | done |
| F-058 | Guide evaluator                              | S13 | L | done |
| F-059 | a:custGeom                                   | S13 | M | done |
| F-060 | Fills                                        | S13 | L | done |
| F-061 | Lines                                        | S14 | M | done |
| F-062 | Effects                                      | S14 | S | done |
| F-063 | Shape properties and style references        | S14 | M | done |
| F-064 | DrawingML text model                         | S14 | XL | done |
| F-064a | Text body properties and shell               | S14 | M | done |
| F-064b | Text paragraphs and runs                     | S14 | L | done |
| F-064c | Text bullets                                 | S14 | S | done |
| F-064d | Nine-level list styles                       | S14 | M | done |
| F-065 | Theme read and write                         | S15 | L | done |
| F-066 | The rdocx Theme adapter                      | S15 | S | done |
<!-- AUTOGEN:backlog-M7 END -->

### M8, PresentationML

<!-- AUTOGEN:backlog-M8 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-067 | Create rpptx-oxml and the corpus harness     | S16 | M | done |
| F-068 | presentation.xml                             | S16 | M | done |
| F-069 | Slide, layout and master parts               | S16 | L | done |
| F-070 | The shape tree                               | S16 | L | done |
| F-071 | Placeholders                                 | S17 | M | done |
| F-072 | Pictures                                     | S17 | M | done |
| F-073 | Graphic frames                               | S17 | M | done |
| F-074 | DrawingML tables                             | S17 | L | done |
| F-075 | Connectors                                   | S18 | S | done |
| F-076 | mc:AlternateContent                          | S18 | M | done |
| F-077 | Notes slides and notes master                | S18 | M | done |
| F-078 | relmap rewrite_rel_ids                       | S18 | M | done |
| F-079 | The rpptx read facade                        | S19 | L | done |
| F-080 | Modelled round-trip gate                     | S19 | M | done |
<!-- AUTOGEN:backlog-M8 END -->

### M9, Inheritance resolver

<!-- AUTOGEN:backlog-M9 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-081 | ResolveCtx skeleton and placeholder chain    | S20 | M | done |
| F-082 | Effective transform and body properties      | S20 | M | done |
| F-083 | The seven-step list style merge              | S20 | L | done |
| F-084 | Format scheme reference resolution           | S20 | M | done |
| F-085 | Typeface resolution                          | S20 | S | done |
| F-086 | Draw order and the flattener                 | S21 | L | done |
| F-087 | ResolvedSlide contract                       | S21 | M | done |
| F-088 | Visual differential tests                    | S21 | M | done |
<!-- AUTOGEN:backlog-M9 END -->

### M10, Renderer

<!-- AUTOGEN:backlog-M10 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-089 | Resolve the preset geometry licensing question | S22 | S | done |
| F-090 | Preset table generator                       | S22 | L | done |
| F-091 | Preset evaluation and fallback               | S22 | M | done |
| F-092 | rpptx-render skeleton and RenderInput        | S22 | M | done |
| F-093 | Shape geometry, fills and lines              | S23 | L | done |
| F-094 | Rotation, flips and groups                   | S23 | M | done |
| F-095 | Arrowheads                                   | S23 | S | done |
| F-096 | Pictures with crop and tile                  | S23 | M | done |
| F-097 | Backgrounds                                  | S23 | S | done |
| F-098 | Shape text layout                            | S24 | XL | done |
| F-098a | Text content box                             | S24 | M | done |
| F-098b | Paragraph inline resolution                  | S24 | L | done |
| F-098c | Line stacking                                | S24 | M | done |
| F-098d | Text anchoring                               | S24 | S | done |
| F-099 | Bullets                                      | S24 | M | done |
| F-100 | Autofit                                      | S24 | M | done |
| F-101 | Vertical text                                | S24 | S | done |
| F-102 | Table rendering                              | S25 | L | done |
| F-103 | Hyperlinks, fields and diagnostics           | S25 | M | done |
| F-104 | SSIM fidelity harness                        | S25 | L | done |
<!-- AUTOGEN:backlog-M10 END -->

### M11, Write API

<!-- AUTOGEN:backlog-M11 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-105 | Bundled default.pptx                         | S26 | M | done |
| F-106 | ShapeIdAllocator and MediaStore              | S26 | M | done |
| F-107 | add_slide                                    | S26 | L | done |
| F-108 | validate()                                   | S26 | M | done |
| F-109 | Shape mutation facade                        | S27 | L | done |
| F-110 | add_textbox, add_shape, add_connector, group | S27 | M | done |
| F-111 | add_picture                                  | S27 | M | done |
| F-112 | Text frame mutation                          | S27 | L | done |
| F-113 | Table facade                                 | S28 | L | done |
| F-114 | remove_slide, move_slide, duplicate_slide    | S28 | M | done |
| F-115 | Slide and presentation properties            | S28 | S | done |
| F-116 | Cross-viewer acceptance                      | S28 | M | done |
<!-- AUTOGEN:backlog-M11 END -->

### M12, Charts

<!-- AUTOGEN:backlog-M12 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-117 | oxml-sml workbook writer                     | S29 | L | done |
| F-118 | ChartML core types                           | S29 | L | done |
| F-119 | Series and data references                   | S29 | L | done |
| F-120 | Axes                                         | S30 | L | done |
| F-121 | Bar and line plots                           | S30 | M | done |
| F-122 | Pie, doughnut, area, scatter and radar plots | S30 | L | done |
| F-123 | Data labels and number formats               | S30 | M | done |
| F-124 | add_chart                                    | S31 | L | done |
| F-125 | Chart rendering: geometry                    | S31 | L | done |
| F-126 | Chart rendering: axes, gridlines and labels  | S31 | L | done |
| F-127 | Chart colour resolution                      | S32 | M | done |
| F-128 | Preserved chart fallback                     | S32 | S | done |
<!-- AUTOGEN:backlog-M12 END -->

### M13, Bindings and tooling

<!-- AUTOGEN:backlog-M13 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-129 | oxml-py-support                              | S33 | M | done |
| F-130 | rdocx-py core                                | S33 | L | done |
| F-131 | rdocx-py formatting and tables               | S33 | L | done |
| F-132 | Python enums, units and exceptions           | S33 | M | done |
| F-133 | rdocx-py rendering with allow_threads        | S33 | S | done |
| F-134 | Type stubs and py.typed                      | S34 | M | done |
| F-135 | python-docx parity suite                     | S34 | M | done |
| F-136 | rpptx-py                                     | S34 | L | done |
| F-137 | wheels.yml                                   | S34 | M | done |
| F-138 | PR-time Python job                           | S34 | S | done |
| F-139 | Rewrite rdocx-wasm                           | S35 | L | done |
| F-140 | wasm CI job                                  | S35 | S | done |
| F-141 | to_pdf in the browser                        | S35 | M | done |
| F-142 | rpptx-wasm                                   | S35 | M | done |
| F-143 | oxml-cli-support                             | S36 | S | done |
| F-144 | rpptx-cli                                    | S36 | L | done |
| F-145 | rpptx-cli thumbnail and outline              | S36 | M | done |
| F-146 | npm publication                              | S36 | S | done |
<!-- AUTOGEN:backlog-M13 END -->

### M14, Word collaboration layer

<!-- AUTOGEN:backlog-M14 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-147 | Comment model and part                       | S46  | M | done |
| F-148 | Comment API                                  | S46  | M | done |
| F-149 | Revision model                               | S47  | L | done |
| F-150 | Accept and reject revisions                  | S47  | L | done |
| F-151 | Revision display in the renderer             | S48  | M | done |
| F-152 | Content control model                        | S46  | L | done |
| F-153 | Content control binding                      | S46  | M | done |
| F-154 | Bookmarks and cross-references               | S46  | M | done |
| F-155 | Document protection                          | S48  | M | done |
<!-- AUTOGEN:backlog-M14 END -->

### M15, Charts beyond PowerPoint

<!-- AUTOGEN:backlog-M15 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-156 | Extract oxml-chart                           | S45  | L | done |
| F-157 | Word chart part and embedded workbook        | S45  | M | done |
| F-158 | Document::add_chart                          | S45  | M | done |
| F-159 | Chart rendering in the Word paginator        | S45  | M | done |
<!-- AUTOGEN:backlog-M15 END -->

### M16, Document automation

<!-- AUTOGEN:backlog-M16 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-160 | Field instruction parser                     | S49  | L | done |
| F-161 | Field evaluation engine                      | S49  | L | done |
| F-162 | Field update policy                          | S49  | M | done |
| F-203 | Reader compatibility corrections             | S49  | M | done |
| F-163 | Template syntax                              | S50  | L | done |
| F-164 | Loops and conditionals                       | S50  | L | done |
| F-165 | Repeating table rows and lists               | S50  | M | done |
| F-166 | Mail merge                                   | S51  | M | done |
| F-167 | Document comparison                          | S51  | L | done |
| F-168 | Watermarks                                   | S51  | S | done |
<!-- AUTOGEN:backlog-M16 END -->

### M17, Security and compliance

<!-- AUTOGEN:backlog-M17 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-169 | Agile encryption, read                       | S52  | L | done |
| F-170 | Agile encryption, write                      | S52  | M | done |
| F-171 | Digital signature verification               | S52  | L | done |
| F-172 | Digital signature creation                   | S53  | M | done |
| F-173 | Tagged PDF structure tree                    | S53  | L | done |
| F-174 | PDF/A conformance                            | S53  | M | done |
| F-175 | Redaction                                    | S53  | M | done |
<!-- AUTOGEN:backlog-M17 END -->

### M18, Format breadth

<!-- AUTOGEN:backlog-M18 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-176 | RTF reader                                   | S54  | L | done |
| F-177 | RTF writer                                   | S54  | M | done |
| F-178 | HTML import                                  | S55  | L | done |
| F-179 | ODT reader                                   | S55  | L | done |
| F-180 | ODT writer                                   | S56  | L | done |
| F-181 | EPUB export                                  | S56  | M | done |
| F-182 | SVG page export                              | S56  | M | done |
| F-183 | Image export options                         | S54  | S | done |
<!-- AUTOGEN:backlog-M18 END -->

### M19, Advanced spreadsheets

<!-- AUTOGEN:backlog-M19 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-184 | Advanced spreadsheet go or no-go             | S103 | S | pending |
| F-185 | Workbook and worksheet model                 | S103 | L | pending |
| F-186 | Shared strings, styles and number formats    | S104 | L | pending |
| F-187 | Reader                                       | S106 | L | pending |
| F-188 | Writer                                       | S106 | L | pending |
| F-189 | Formula parser                               | S104 | L | pending |
| F-190 | Calculation engine                           | S107 | L | pending |
| F-191 | Charts in spreadsheets                       | S107 | M | pending |
| F-192 | Conditional formatting and data validation   | S107 | M | pending |
| F-193 | Pivot cache and table model                  | S108 | L | pending |
| F-194 | Sheet rendering                              | S112 | L | pending |
| F-195 | rxlsx distribution                           | S112 | L | pending |
| F-204 | Spreadsheet corpus and compatibility matrix | S103 | M | pending |
| F-205 | Excel tables and structured references      | S104 | L | pending |
| F-206 | Advanced worksheet objects                  | S105 | L | pending |
| F-207 | Pivot recalculation engine                  | S108 | L | pending |
| F-208 | Slicers, pivot charts, and Data Model boundary | S108 | L | pending |
| F-209 | Power Query package and M language          | S109 | L | pending |
| F-210 | Power Query execution and connectors        | S110 | L | pending |
| F-211 | Office Scripts artifacts and ExcelScript surface | S111 | L | pending |
| F-212 | Sandboxed Office Scripts runtime            | S111 | L | pending |
<!-- AUTOGEN:backlog-M19 END -->

### M20, Fidelity at scale

<!-- AUTOGEN:backlog-M20 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-196 | Word corpus                                  | S57  | M | done |
| F-197 | Word SSIM harness                            | S57  | L | done |
| F-198 | Hyphenation                                  | S58  | L | done |
| F-199 | Complex script shaping                       | S58  | L | done |
| F-200 | Vertical and bidirectional text              | S58  | M | done |
| F-201 | Large document performance                   | S57  | L | done |
| F-202 | Incremental layout                           | S58  | L | done |
| F-313 | External Word benchmark reproduction loop    | S102 | S | pending |
| F-314 | Benchmark conversion failures and startup latency | S102 | M | pending |
| F-315 | Word tracked-change markup print view        | S102 | L | pending |
| F-316 | Word comment and revision balloon pane       | S102 | L | pending |
| F-317 | Benchmark pagination and layout fidelity wave | S102 | L | pending |
| F-318 | Benchmark release, adapter update and README badge | S102 | M | pending |
<!-- AUTOGEN:backlog-M20 END -->

### M21, Presentation depth

<!-- AUTOGEN:backlog-M21 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-213 | Animation and transition timing model       | S60 | L | done |
| F-214 | Timeline evaluation and transition rendering | S60 | L | done |
| F-215 | Audio and video package model                | S61 | L | done |
| F-216 | Media poster and playback rendering          | S61 | M | done |
| F-217 | Presentation collaboration and navigation model | S59 | L | done |
| F-218 | Embedded object and macro inventory          | S62 | L | done |
| F-219 | SmartArt typed model                         | S62 | L | done |
| F-220 | SmartArt layout and rendering                | S63 | L | done |
| F-221 | Presentation encryption and signatures      | S59 | M | done |
| F-222 | ODP read and write                           | S63 | L | done |
| F-223 | Modern presentation package variants        | S63 | M | done |
| F-224 | HTML slide content import                    | S64 | L | done |
| F-225 | PDF page content import                      | S64 | L | done |
| F-226 | Notes and handout export                     | S63 | M | done |
| F-227 | Animated GIF and video export                | S61 | L | done |
<!-- AUTOGEN:backlog-M21 END -->

### M22, Word depth

<!-- AUTOGEN:backlog-M22 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-228 | OfficeMath model and authoring               | S65 | L | done |
| F-229 | OfficeMath layout and PDF rendering          | S65 | M | done |
| F-230 | MathML and LaTeX conversion                  | S65 | M | done |
| F-231 | Extended field evaluation                    | S66 | L | done |
| F-232 | Dynamic table of contents rebuild            | S66 | L | done |
| F-233 | Advanced mail merge                          | S67 | L | done |
| F-234 | Full-story document comparison               | S67 | L | done |
| F-235 | Comparison granularity and ignore policy     | S67 | M | done |
| F-236 | Embedded object and macro inventory          | S68 | L | done |
| F-237 | Forms, glossary, and building blocks         | S68 | L | done |
| F-238 | Flat OPC and modern Word package variants    | S69 | M | done |
| F-239 | MHTML import and export                      | S69 | M | done |
<!-- AUTOGEN:backlog-M22 END -->

### M23, From-scratch business documents

<!-- AUTOGEN:backlog-M23 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-240 | Modern DOCX completeness audit and private corpus matrix | S70 | L | done |
| F-241 | Public authoring conformance harness | S70 | L | done |
| F-242 | Root README product and capability overview | S70 | M | done |
| F-243 | Word-compatible fresh package profiles | S71 | L | done |
| F-244 | Corpus settings and document properties | S71 | L | done |
| F-245 | Corpus themes, font tables, and embedded fonts | S71 | L | done |
| F-246 | Corpus style authoring | S71 | L | done |
| F-247 | Complete numbering level and instance model | S71 | L | done |
| F-248 | Style-linked numbering, counters, TOC, and REF | S71 | L | done |
| F-249 | Deterministic package identifier allocation | S71 | M | done |
| F-250 | Ordered mutable section facade | S72 | L | done |
| F-251 | Complete section and page geometry | S72 | L | done |
| F-252 | Rich per-section headers and footers | S72 | L | done |
| F-253 | Container-neutral story editing | S72 | L | done |
| F-254 | Generic insert, move, clone, and remove operations | S72 | L | done |
| F-255 | Part-scoped assets, links, and relationships | S72 | M | done |
| F-256 | Transactional cross-document fragment import | S72 | L | done |
| F-257 | Complete M23 table authoring | S73 | L | done |
| F-258 | Complete M23 row and cell authoring | S73 | L | done |
| F-259 | Container measurement and equal-height layout | S73 | M | done |
| F-260 | Ordered run content authoring | S73 | L | done |
| F-261 | Rich HTML fragments in arbitrary containers | S73 | L | done |
| F-262 | Corpus drawings, text boxes, and watermarks | S73 | L | done |
| F-263 | Layout-backed fields and M23 corpus gate | S73 | L | done |
<!-- AUTOGEN:backlog-M23 END -->

### M24, Modern DOCX authoring completeness

<!-- AUTOGEN:backlog-M24 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-264 | Complete paragraph property authoring | S74 | L | done |
| F-265 | Complete run property and inline authoring | S74 | L | done |
| F-266 | International and vertical typography | S74 | L | done |
| F-266a | Script identity and font slot resolution | S74 | L | done |
| F-266b | Ruby and emphasis marks | S74 | L | done |
| F-266c | Character grid and vertical text | S74 | L | done |
| F-267 | Complete table style and conditional formatting authoring | S74 | L | done |
| F-268 | Floating and advanced table layout | S74 | L | done |
| F-268a | Advanced table authoring and geometry | S74 | L | done |
| F-268b | Floating table placement and wrap | S74 | M | done |
| F-269 | Complete section page semantics | S74 | L | done |
| F-270 | Complete settings and web settings authoring | S74 | L | done |
| F-271 | Uniform rich header and footer editing | S86 | L | done |
| F-272 | Rich footnote authoring | S86 | L | done |
| F-273 | Rich endnote authoring | S86 | L | done |
| F-274 | Note separators, markers, and restart policy | S89 | L | done |
| F-275 | Cross-story bookmarks, ranges, and annotations | S89 | L | done |
| F-276 | Complete fragment conflict and dependency policy | S89 | L | done |
| F-277 | Glossary and building-block creation | S89 | L | done |
| F-278 | General simple and complex field builder | S90 | L | done |
| F-279 | Pagination field materialization across stories | S90 | L | done |
| F-280 | Captions, sequences, and complete cross-references | S90 | M | done |
| F-281 | Indexes and tables of figures and authorities | S90 | L | done |
| F-282 | Citations and bibliography authoring | S90 | L | done |
| F-283 | Complete numbering-aware navigation fields | S95.2 | L | done |
| F-284 | Stable container-wide template grammar | S95.2 | L | done |
| F-285 | Content control creation and lifecycle | S95.3 | L | done |
| F-286 | Rich, repeating, and typed content controls | S95.3 | L | done |
| F-287 | Custom XML stores and data binding authoring | S95.3 | L | done |
| F-288 | Legacy form field creation | S95.4 | M | done |
| F-289 | Modern Word form authoring | S95.4 | L | done |
| F-290 | Mail-merge package and data-source authoring | S95.4 | M | done |
| F-291 | Tracked insertion and deletion authoring | S99 | L | pending |
| F-292 | Property revisions and move ranges | S99 | L | pending |
| F-293 | Complete comments and modern comment metadata | S99 | L | pending |
| F-294 | Permission ranges and protection integration | S99 | M | pending |
| F-295 | Comparison output as complete revisions | S99 | L | pending |
| F-296 | Collaboration identity and deterministic time policy | S99 | M | pending |
| F-297 | Complete Word drawing anchor and effect authoring | S100 | L | pending |
| F-298 | Shapes, text boxes, groups, and connectors in Word | S100 | L | pending |
| F-299 | AlternateContent, VML, and SVG compatibility authoring | S100 | L | pending |
| F-300 | Charts at arbitrary Word insertion points | S100 | M | pending |
| F-301 | SmartArt and diagram authoring in Word | S100 | L | pending |
| F-302 | Embedded objects, icons, and alternative-format parts | S100 | M | pending |
| F-303 | Drawing and embedded-content layout completion | S100 | L | pending |
| F-304 | Typed package extensibility facade | S101 | L | pending |
| F-305 | Attached templates, web extensions, and task panes | S101 | L | pending |
| F-306 | Executable compatibility attachment and signature rules | S101 | M | pending |
| F-307 | Complete accessibility authoring and audit | S101 | L | pending |
| F-308 | Fully modeled and losslessness diagnostics | S101 | L | pending |
| F-309 | Strict, transitional, and repair-free conformance | S101 | L | pending |
| F-310 | Determinism, resource limits, bindings, and stability gate | S101 | L | pending |
| F-311 | Positioned paragraph frame placement | S101 | M | pending |
| F-312 | Run visual effect render projection | S101 | L | pending |
<!-- AUTOGEN:backlog-M24 END -->

### X, Cross-cutting

Rolling intake PR315 maps to F-X195 for Issues296 and297, also consumed by
Issue313. It adds no F-ID or status change. See its existing approved plan and
HLD14 contract before adopting contributor changes.


<!-- AUTOGEN:backlog-MX START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-X001 | rdocx-cli tests                             | S36 | M | done |
| F-X002 | README example correctness                  | S36 | S | done |
| F-X003 | Deduplicate the sample generators           | S36 | S | done |
| F-X004 | Fix the shared temp path in the test suite  | S36 | S | done |
| F-X005 | Tag rpptx-v0.1.2                            | S32.2 | S | done |
| F-X006 | Tag the expanded rpptx family               | S37 | S | done |
| F-X007 | Integrate PR 25 and stable crate documentation | S38 | L | done |
| F-X008 | Tag v0.5.0                                  | S38 | S | done |
| F-X009 | README coverage for every workspace crate   | S39 | L | done |
| F-X010 | Tag v0.6.0                                  | S39 | S | done |
| F-X011 | Tag rpptx-v0.2.0                            | S39 | S | done |
| F-X012 | Restore pinned CI toolchains                | S40 | M | done |
| F-X013 | Footnote and endnote placement             | S41 | M | done |
| F-X013a | Footnote line advance                     | S41 | S | done |
| F-X013b | Footnote reservation and splitting        | S41 | L | done |
| F-X013c | Endnotes at the document end              | S41 | M | done |
| F-X014 | Kashida justification values               | S41 | S | done |
| F-X015 | Anchored drawing wrap and alignment model  | S41 | M | done |
| F-X016 | Floating drawing placement and text wrapping | S41 | L | done |
| F-X017 | Notes broken to their own section's width   | S43 | S | done |
| F-X018 | Unknown enumerated values must not fail open | S43 | M | done |
| F-X019 | Paragraph-relative later drawings should wrap | S43 | M | done |
| F-X020 | Refresh the dependency lockfile             | S42 | S | done |
| F-X021 | Hash harness should cover PDF output       | S43 | M | done |
| F-X025 | /verify must run the release regressions   | S43 | S | done |
| F-X024 | Move the theme adapter into rdocx-oxml     | S42 | M | done |
| F-X022 | Tag rpptx-v0.3.0                           | S42 | S | done |
| F-X023 | Tag v0.7.0                                 | S42 | S | done |
| F-X026 | CI must run the release regressions too     | S44 | S | done |
| F-X027 | Wire the golden-PNG gate into something     | S44 | S | done |
| F-X028 | Repair the agent-facing documentation drift | S44 | M | done |
| F-X029 | Path-filtered CI jobs                       | S44 | M | done |
| F-X030 | Decouple the npm package versions           | -   | S | archived |
| F-X031 | Require the CI gate in branch protection    | S58 | S | done |
| F-X032 | Expose complete Word layout results         | S51 | S | done |
| F-X033 | Integrate PR 36 ordered body items          | S51 | S | done |
| F-X034 | Reviewed release notes for every release    | S51 | S | done |
| F-X035 | Tag rpptx-v0.4.0                            | S51 | S | done |
| F-X036 | Tag v0.8.0                                  | S51 | S | done |
| F-X037 | Trace Word glyphs to source paragraphs     | S51 | M | done |
| F-X038 | Cache relayout work across document edits  | S51 | L | done |
| F-X039 | Share layout payloads and transfer reusable engines | S52 | M | done |
| F-X040 | Restart pagination and cache table blocks  | S52 | L | done |
| F-X041 | Remove duplicated glyphs at break opportunities | S52 | M | done |
| F-X042 | Prove headers and footers in PDF output    | S52 | S | done |
| F-X043 | Reuse bundled-fallback caller-font layouts | S52 | M | done |
| F-X044 | Scale paragraph-cache lookup for editors   | S52 | M | done |
| F-X045 | Cache headers and footers transactionally  | S52 | M | done |
| F-X046 | Reuse substituted pages exactly            | S52 | S | done |
| F-X047 | Attribute empty Word paragraphs            | S52 | S | done |
| F-X048 | Dense form table fidelity                  | S53 | L | done |
| F-X049 | Tag rpptx-v0.5.0                           | S53 | S | done |
| F-X050 | Tag v0.9.0                                 | S53 | S | done |
| F-X051 | Honor caller-supplied font family aliases  | S54 | M | done |
| F-X052 | Restore interactive relayout performance   | S55 | L | done |
| F-X053 | Complete layout migration and contribution records | S55 | S | done |
| F-X054 | Integrate PRs 47 through 52                | S56 | L | done |
| F-X055 | Tag v0.10.0                               | S56 | S | archived |
| F-X056 | Tag rpptx-v0.6.0                          | S56 | S | done |
| F-X057 | Tag v0.10.1                               | S56 | S | done |
| F-X058 | Shared multilingual text substrate        | S58 | L | done |
| F-X059 | Tag rpptx-v0.7.0                          | S58 | S | done |
| F-X060 | Tag v0.11.0                               | S58 | S | archived |
| F-X061 | Support staged dependency checkpoints in run-sprint | S58 | S | done |
| F-X062 | Reuse restart pagination with notes and headers | S58 | M | done |
| F-X063 | Avoid duplicate caller-font byte comparisons | S58 | S | done |
| F-X064 | Accept whole-valued decimal table measurements | S58 | S | done |
| F-X065 | Expose tracked table grid changes             | S58 | S | done |
| F-X066 | Classify legacy VML horizontal rules          | S58 | S | done |
| F-X067 | Prime Word fidelity Cargo dependencies        | S58 | S | done |
| F-X068 | Tag rpptx-v0.8.0                          | S58 | S | done |
| F-X069 | Tag v0.11.1                               | S58 | S | done |
| F-X070 | Yank incomplete v0.11.0 packages          | S58 | S | done |
| F-X071 | Integrate PRs 61 through 64                | S62 | L | done |
| F-X072 | Keep paragraph caching across note references | S63 | M | done |
| F-X073 | Restart ordinary-prose pagination within the aggregate cache | S63 | L | done |
| F-X074 | Tag rpptx-v0.9.0                          | S64 | S | done |
| F-X075 | Preserve restart pagination across page-spanning paragraphs | S64 | M | done |
| F-X076 | Tag v0.12.0                               | S64 | S | done |
| F-X077 | Share strict XML lexical validation       | S69 | M | done |
| F-X078 | Tag v0.13.0                               | S69 | S | archived |
| F-X079 | Tag rpptx-v0.10.0                         | S69 | S | done |
| F-X080 | Restore CI release readiness              | S69 | S | done |
| F-X081 | Tag rpptx-v0.11.0                         | S69 | S | done |
| F-X082 | Tag v0.13.1                               | S69 | S | done |
| F-X083 | Close confirmed Issue 67 and intake Issue 69 | S70 | S | done |
| F-X084 | Narrow note-part paragraph cache invalidation | S70 | M | done |
| F-X085 | Memoize restart body identities once per layout | S70 | M | done |
| F-X086 | Provenance-safe restart after body-length changes | S70 | L | done |
| F-X087 | Portable authored Word charts from PR 71 | S71 | L | done |
| F-X088 | Verify and close Issue 69 after S70 fixes | S71 | S | done |
| F-X089 | Capability-led README family | S71 | L | done |
| F-X090 | Accept part-local producer drawing identities | S72 | S | done |
| F-X091 | Serialize unused root default namespaces safely | S72 | M | done |
| F-X092 | Preserve logical reading order in generated PDFs | S72 | L | done |
| F-X093 | Preserve drawings through document comparison staging | S72 | M | done |
| F-X094a | Expose Word collaboration and redline commands in rdocx-cli | S72 | M | done |
| F-X094b | Structured CLI text and layout plus guarded replacement | S72 | L | done |
| F-X094c | Priority rdocx Python collaboration, comparison, layout, and TOC | S72 | L | done |
| F-X094d | rdocx Python sections, styles, rich stories, and hyperlinks | S72 | L | done |
| F-X094e | rpptx Python rendering, comments, and notes | S72 | L | done |
| F-X094f | Prepare the version-aligned Python release paths | S72 | M | done |
| F-X095 | Integrate PRs 77 through 80 and restore deterministic CI | S72 | L | done |
| F-X096 | Align Python distribution versions and release tags | S72 | M | done |
| F-X097 | Preserve namespace-scoped drawings and complex fields in comparison | S73 | M | done |
| F-X098 | Preserve content-control type payloads | S73 | M | done |
| F-X099 | Expose direct body ownership for story items | S73 | M | done |
| F-X100 | Preserve explicit false table toggles | S73 | S | done |
| F-X101 | Honor run-level page breaks during pagination | S73 | M | done |
| F-X102 | Resolve header and footer pictures in their story scope | S73 | M | done |
| F-X103 | Accept standard TOC switches and report rebuild diagnostics | S73 | M | done |
| F-X104 | Render DrawingML picture transparency | S73 | M | done |
| F-X105 | Separate slide-owned placeholders from master header flags | S73 | M | done |
| F-X106a | Expose indexed content mutation and counted replacement in Python | S73 | L | done |
| F-X106b | Expose paragraph and run formatting mutations in Python | S73 | M | done |
| F-X106c | Expose story mutation, hyperlinks, revisions, fields, and XML in Python | S73 | L | done |
| F-X107 | Clone and remove existing table rows | S73 | M | done |
| F-X108 | Replace an existing picture atomically | S73 | M | done |
| F-X109 | Split text runs at Unicode character offsets | S73 | M | done |
| F-X110 | Control field updates on document open | S73 | S | done |
| F-X111 | Attach portable CLI binaries to Rust releases | S73 | L | done |
| F-X112 | Publish the complete S73 package families | S73 | L | done |
| F-X113 | Preserve appended paragraphs in document comparison | S73 | M | done |
| F-X114 | Rebuild TOC entries with document styles and geometry | S73 | M | done |
| F-X115 | Preserve modern comment metadata and identity | S73 | M | done |
| F-X116 | Make Python story reads linear and complete | S73 | L | done |
| F-X117 | Render transparent and large raster pictures safely | S73 | M | done |
| F-X118 | Make notes rendering and replacement safe | S73 | L | done |
| F-X119 | Complete round-three Python authoring and inspection | S73 | L | done |
| F-X120 | Accept fractional DOCX line spacing values | S73 | S | done |
| F-X121 | Adopt PR 123 authored line-chart portability | S73 | S | done |
| F-X122 | Recover the immutable rpptx 0.12.0 release attempt | S73 | M | done |
| F-X123 | Accept producer TOC style variants | S74 | S | done |
| F-X124 | Make content cloning linear and explicit | S74 | M | done |
| F-X125 | Compare table grid changes | S74 | M | done |
| F-X126 | Preserve drawings through comparison acceptance | S74 | M | done |
| F-X127 | Collapse adjacent page break requests | S74 | S | done |
| F-X128 | Preserve Word paragraph and revision identities | S74 | M | done |
| F-X129 | Tolerate unmatched notes placeholders | S74 | S | done |
| F-X130 | Show package depth, footprint, and speed | S74 | L | done |
| F-X131 | Retain only the namespace declarations a root attribute uses | S74 | S | done |
| F-X132 | Match a retained namespace owner by structure, not by identity | S74 | S | done |
| F-X133 | Stop rebinding a canonical prefix on every retained element | S86 | S | done |
| F-X134 | Keep Python story hyperlink snapshots linear | S75 | S | done |
| F-X135 | Integrate PRs 146 through 151 and resolve unassigned reports | S75 | L | done |
| F-X136 | Fix table row breaks and footer-only pages | S75 | L | done |
| F-X137 | Package and CLI safety contribution wave | S76 | L | done |
| F-X138 | Word story and content contribution wave | S76 | L | done |
| F-X139 | Word identity and comparison contribution wave | S76 | L | done |
| F-X140 | Rendering and layout contribution wave | S77 | L | done |
| F-X141 | Word Python contribution wave | S77 | L | done |
| F-X142 | Presentation Python contribution wave | S77 | L | done |
| F-X143 | Revision listing and CLI story contribution wave | S77 | M | done |
| F-X144 | Identity and producer matrices across operations | S78 | L | done |
| F-X145 | Comparison options and redline completion | S78 | L | done |
| F-X146 | Word line height and inline picture spacing | S79 | L | done |
| F-X147 | Complete rdocx Python production checklist | S80 | L | done |
| F-X148 | Complete rpptx Python production checklist | S81 | L | done |
| F-X149 | Word fixture and workflow acceptance gate | S82 | L | done |
| F-X150 | Reconcile issue closure evidence | S82 | M | done |
| F-X151 | Word preservation and comparison PR intake | S78 | L | done |
| F-X152 | Full-story CLI diff and count repair | S79 | M | done |
| F-X153 | Word Python supplemental contribution | S79 | M | done |
| F-X154 | Presentation text and preservation repair | S80 | M | done |
| F-X155 | Presentation drawing API contribution | S80 | L | done |
| F-X156 | Presentation slide and table contribution | S81 | L | done |
| F-X157 | Complete deck-chain authoring checklist | S81 | L | done |
| F-X158 | Presentation fixture and workflow acceptance gate | S82 | L | done |
| F-X159 | Word document validity baseline | S83 | M | done |
| F-X160 | Tolerant style, drawing and measurement reads | S83 | L | done |
| F-X161 | Compact Word XML and namespace preservation | S84 | L | done |
| F-X162 | Per-paragraph section width and pagination | S84 | M | done |
| F-X163 | Plain-line trailing-space fit | S84 | M | done |
| F-X164 | Added shape theme style | S84 | M | done |
| F-X165 | Python and CLI tracked revision view | S84 | M | done |
| F-X166 | Picture and final-block comparison revisions | S84 | L | done |
| F-X167 | Accepted-view exporters and readers | S84 | L | done |
| F-X168 | Current issue and contribution closure evidence | S84 | M | done |
| F-X169 | Reconcile live contributions and open issue contracts | S84 | L | done |
| F-X170 | High-level Word style formatting API | S84 | M | done |
| F-X171 | Pin LibreOffice for macOS Python acceptance | S85 | S | done |
| F-X172 | Unified release artifacts and provenance | S86 | L | done |
| F-X173 | Prepare unified rpptx 0.13.0 family | S86 | M | done |
| F-X174 | Prepare unified rdocx 0.15.0 family | S86 | M | done |
| F-X175 | Refresh CLI archive evidence after release hardening | S87 | S | done |
| F-X176 | Repair unified release inventory and respin PowerPoint | S88 | M | done |
| F-X177 | Accept unified fontdb source features | S89 | S | done |
| F-X178 | Clearable direct run formatting setters | S92 | S | done |
| F-X179 | Correct multi-paragraph comment threads from PR 271 | S90 | S | done |
| F-X180 | Correct cell nil and none border precedence | S90 | S | done |
| F-X181 | Ignore page and column breaks inside table cells | S90 | S | done |
| F-X182 | Honor direct table alignment | S90 | M | done |
| F-X183 | Correct table margins and legacy positioning | S90 | M | done |
| F-X184 | Safe comment ownership during content removal | S90 | L | done |
| F-X185 | Expose comment anchor text and story location | S90 | M | done |
| F-X186 | Move comment anchors without losing threads | S90 | M | done |
| F-X187 | Scoped paragraph and cell text replacement | S90 | M | done |
| F-X188 | Preserve comment ownership when replacing or removing whole stories | S90 | L | done |
| F-X190 | Preserve cached text in multi-run complex field story snapshots | S90 | M | done |
| F-X192 | Complete the remaining Word bibliography catalogue | S98 | L | pending |
| F-X191 | Ignore namespace declarations in numbering reader completeness | S90 | S | done |
| F-X189 | Prepare Word 0.16.0 and PowerPoint 0.14.0 families | S90 | L | done |
| F-X193 | Preserve semantic mark highlights through nested HTML runs | S91 | S | done |
| F-X194 | Complete Python table text and indentation APIs | S91 | M | done |
| F-X195 | Honor caller fonts and variable styles across render paths | S91 | L | done |
| F-X196 | Validate visible paragraph content without false empty warnings | S91 | S | done |
| F-X197 | Unify Python colour inputs and font colour formats | S91 | L | done |
| F-X198 | Preserve alternate text-box copies during scoped replacement | S91 | M | done |
| F-X199 | Model namespace-qualified empty table cell margins | S91 | S | done |
| F-X200 | Complete Python Word formatting and list parity | S91 | L | done |
| F-X201 | Complete Python section stories and page authoring | S91 | L | done |
| F-X202 | Complete Word tables pictures and internal links | S91 | L | done |
| F-X203 | Complete Python Word document operations and validation | S91 | L | done |
| F-X204 | Complete shared CLI one-shot editing commands | S91 | L | done |
| F-X205 | Complete Python PowerPoint text fill and export parity | S91 | L | done |
| F-X206 | Complete PowerPoint shape authoring and live handles | S91 | L | done |
| F-X207 | Persist PowerPoint autofit and implement fit_text | S91 | L | done |
| F-X208 | Author PowerPoint footer fields themes and transitions | S91 | L | done |
| F-X209 | Edit PowerPoint masters layouts and themes | S91 | L | done |
| F-X210 | Extend metric-compatible fonts and bullet fallbacks | S91 | L | done |
| F-X211 | Complete Python script compatibility and safe XML access | S91 | L | done |
| F-X211a | Word checked XML owners and resource admission | S91 | L | done |
| F-X211b | PowerPoint checked XML owners | S91 | L | done |
| F-X211c | Integrated Python compatibility and causal errors | S91 | L | done |
| F-X211d | Repair Word compatibility owners and aliases | S91 | L | done |
| F-X212 | Align native accessibility with visible paragraph content | S91 | M | done |
| F-X213 | Preserve TOC entry boundary paragraphs during rebuild | S91 | M | done |
| F-X214 | Bound recursive PowerPoint parsing and traversal | S91 | L | done |
| F-X215 | Create charts in Python and replace existing chart data | S91 | L | done |
| F-X216 | Import complete slides and merge decks with source formatting | S91 | L | done |
| F-X217 | Import HTML and Markdown through public Word surfaces | S91 | L | done |
| F-X218 | Export presentations as semantic HTML and Markdown | S91 | L | done |
| F-X219 | Qualify full-range text watermark interoperability | S98 | M | pending |
| F-X220 | Embed fonts through both native and public package families | S91 | L | done |
| F-X221 | Prepare Word 0.16.1 and PowerPoint 0.14.1 patch releases | S91 | L | done |
| F-X222 | Restore linear Word story range inventory and empty owners | S91 | M | done |
| F-X223 | Qualify native Word embedded multi-face interoperability | S95.11 | M | done |
| F-X224 | Repair hosted release tool and stack routing | S91.1 | M | done |
| F-X225 | Keep DrawingML text-style parsing within the default WASM stack | S91.1 | M | done |
| F-X226 | Release superseded relationships after checked PowerPoint XML edits | S91.1 | M | done |
| F-X227 | Execute the unchanged first-draft compatibility corpus in owner CI | S91.1 | L | done |
| F-X228 | Qualify faster CI without weakening gates | S92 | M | done |
| F-X229 | Repair comparison reconstruction correctness | S92 | L | done |
| F-X230 | Group comparison changes for readable revisions | S92 | M | done |
| F-X231 | Preserve PowerPoint HTML import whitespace | S92 | M | done |
| F-X232 | Render legal revisions and qualify checked XML repros | S93 | L | done |
| F-X233 | Repair paragraph border geometry | S93 | M | done |
| F-X234 | Preserve supported LaTeX conversion structure | S93 | M | done |
| F-X235 | Qualify equation layout against native Word | S93 | L | done |
| F-X236 | Measure public setter inlining before changing policy | S93 | S | done |
| F-X237 | Restore linear Python table owner inventory | S92 | L | done |
| F-X238 | Restore python-docx Word names from the first-draft corpus | S94 | L | done |
| F-X239 | python-pptx notes names and refusal of silent zero values | S94 | M | done |
| F-X240 | Python and CLI signing and signature verification | S94 | M | done |
| F-X241 | Python and CLI form reading and filling | S94 | M | done |
| F-X242 | Form control geometry and AcroForm field model | S97 | L | pending |
| F-X243 | Field appearance streams, full-font resources and conformance | S97 | L | pending |
| F-X244 | Signature field, bindings and CLI option | S97 | M | pending |
| F-X245 | Preserve chart data changes in Word comparison | S93 | L | done |
| F-X246 | Honor chart legend position and overlay during rendering | S93 | L | done |
| F-X247 | Preserve the common LaTeX inputs reported in Issue365 | S94 | L | done |
| F-X248 | Validate identified paragraphs inside block content controls | S95 | M | done |
| F-X249 | Set typed content control values as Word does | S95 | M | done |
| F-X250 | Complete everyday Word chart equation list and note APIs | S95.11 | L | done |
| F-X251 | Refuse invalid Word values before publication | S95 | M | done |
| F-X252 | Refuse newly invalid Word raw XML atomically | S95 | M | done |
| F-X253 | Honor imported chart types and hidden master shapes | S95 | M | done |
| F-X254 | Refuse invalid PowerPoint values XML and empty edits | S95.11 | L | done |
| F-X255 | Prepare Word 0.16.2 and PowerPoint 0.14.2 patch releases | S95 | L | done |
| F-X256 | Refuse ambiguous Word numeric units and preserve line spacing | S95 | M | done |
| F-X257 | Complete remaining Word names and legacy dropdown diagnostics | S95 | L | done |
| F-X258 | Complete remaining PowerPoint names color and language contracts | S95 | M | done |
| F-X259 | Repair reviewed consistency defects and strengthen acceptance gates | S95.3 | M | done |
| F-X260 | Restore live handles and bound structural edit cost | S95.3 | L | done |
| F-X261 | Resolve Word package preservation and inspection contributions | S95.3 | L | done |
| F-X262 | Resolve PowerPoint authoring and table-render contributions | S95.3 | L | done |
| F-X263 | Prepare Word 0.16.3 and PowerPoint 0.14.3 patch releases | S95.3 | L | done |
| F-X264 | Repair Word defaults package inspection and comment projection | S95.4 | L | done |
| F-X265 | Preserve fragment relationship targets and duplicate styles | S95.4 | M | done |
| F-X266 | Bound Word body reads and structural edit cost | S95.4 | L | done |
| F-X267 | Bound PowerPoint mutation cost across public publishers | S95.4 | L | done |
| F-X268 | Bound Word owner inventory and unchanged package staging | S95.4 | L | done |
| F-X269 | Qualify PowerPoint shadow falloff and wipe paint domains | S98 | M | pending |
| F-X270 | Prepare Word 0.16.4 and PowerPoint 0.14.4 patch releases | S95.4 | L | done |
| F-X271 | Restore pinned viewer downloads and planning policy | S96 | M | in-progress |
| F-X272 | Preserve existing revisions during document comparison | S96 | L | in-progress |
| F-X273 | Repair header body clearance and complete TOC navigation | S96 | L | done |
| F-X274 | Honor section and field page-number formats | S96 | M | in-progress |
| F-X275 | Prepare Word 0.16.5 and PowerPoint 0.14.5 patch releases | S96 | L | pending |
<!-- AUTOGEN:backlog-MX END -->
