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
| M1, Preparation and safety net              | 12 | 12 | 0 | 0  |
| M2, Shared infrastructure extraction        | 10 | 10 | 0 | 0  |
| M3, Media                                   | 6  | 6 | 0 | 0  |
| M4, Layout primitives                       | 8  | 8 | 0 | 0  |
| M5, PDF backend                             | 9  | 9 | 0 | 0  |
| M6, Shared publication and rdocx cutover     | 6  | 6 | 0 | 0  |
| M7, DrawingML                               | 19 | 19 | 0 | 0  |
| M8, PresentationML                          | 14 | 14 | 0 | 0  |
| M9, Inheritance resolver                    | 8  | 8 | 0 | 0  |
| M10, Renderer                               | 20 | 20 | 0 | 0  |
| M11, Write API                              | 12 | 12 | 0 | 0  |
| M12, Charts                                 | 12 | 12 | 0 | 0  |
| M13, Bindings and tooling                   | 18 | 18 | 0 | 0  |
| M14, Word collaboration layer                  | 9  | 9 | 0 | 0  |
| M15, Charts beyond PowerPoint                  | 4  | 4 | 0 | 0  |
| M16, Document automation                       | 10 | 10 | 0 | 0  |
| M17, Security and compliance                   | 7  | 7 | 0 | 0  |
| M18, Format breadth                            | 8  | 5 | 3 | 0  |
| M19, Advanced spreadsheets                     | 21 | 0 | 0 | 21 |
| M20, Fidelity at scale                         | 7  | 7 | 0 | 0  |
| M21, Presentation depth                        | 15 | 15 | 0 | 0  |
| M22, Word depth                                | 12 | 0 | 1 | 11 |
| X, Cross-cutting (opportunistic)            | 79 | 71 | 3 | 2  |
| **Total** | **326** | **282** | **7** | **34** |
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
| F-184 | Advanced spreadsheet go or no-go             | S70  | S | pending |
| F-185 | Workbook and worksheet model                 | S70  | L | pending |
| F-186 | Shared strings, styles and number formats    | S71  | L | pending |
| F-187 | Reader                                       | S73  | L | pending |
| F-188 | Writer                                       | S73  | L | pending |
| F-189 | Formula parser                               | S71  | L | pending |
| F-190 | Calculation engine                           | S74  | L | pending |
| F-191 | Charts in spreadsheets                       | S74  | M | pending |
| F-192 | Conditional formatting and data validation   | S74  | M | pending |
| F-193 | Pivot cache and table model                  | S75  | L | pending |
| F-194 | Sheet rendering                              | S79  | L | pending |
| F-195 | rxlsx distribution                           | S79  | L | pending |
| F-204 | Spreadsheet corpus and compatibility matrix | S70  | M | pending |
| F-205 | Excel tables and structured references      | S71  | L | pending |
| F-206 | Advanced worksheet objects                  | S72  | L | pending |
| F-207 | Pivot recalculation engine                  | S75  | L | pending |
| F-208 | Slicers, pivot charts, and Data Model boundary | S75 | L | pending |
| F-209 | Power Query package and M language          | S76  | L | pending |
| F-210 | Power Query execution and connectors        | S77  | L | pending |
| F-211 | Office Scripts artifacts and ExcelScript surface | S78 | L | pending |
| F-212 | Sandboxed Office Scripts runtime            | S78  | L | pending |
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
| F-228 | OfficeMath model and authoring               | S65 | L | in-progress |
| F-229 | OfficeMath layout and PDF rendering          | S65 | M | pending |
| F-230 | MathML and LaTeX conversion                  | S65 | M | pending |
| F-231 | Extended field evaluation                    | S66 | L | pending |
| F-232 | Dynamic table of contents rebuild            | S66 | L | pending |
| F-233 | Advanced mail merge                          | S67 | L | pending |
| F-234 | Full-story document comparison               | S67 | L | pending |
| F-235 | Comparison granularity and ignore policy     | S67 | M | pending |
| F-236 | Embedded object and macro inventory          | S68 | L | pending |
| F-237 | Forms, glossary, and building blocks         | S68 | L | pending |
| F-238 | Flat OPC and modern Word package variants    | S69 | M | pending |
| F-239 | MHTML import and export                      | S69 | M | pending |
<!-- AUTOGEN:backlog-M22 END -->

### X, Cross-cutting

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
<!-- AUTOGEN:backlog-MX END -->
