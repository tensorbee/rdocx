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
| M2, Shared infrastructure extraction        | 10 | 0 | 0 | 10 |
| M3, Media                                   | 6  | 0 | 0 | 6  |
| M4, Layout primitives                       | 8  | 0 | 0 | 8  |
| M5, PDF backend                             | 9  | 0 | 0 | 9  |
| M6, rdocx 0.3.0 release                     | 6  | 0 | 0 | 6  |
| M7, DrawingML                               | 15 | 0 | 0 | 15 |
| M8, PresentationML                          | 14 | 0 | 0 | 14 |
| M9, Inheritance resolver                    | 8  | 0 | 0 | 8  |
| M10, Renderer                               | 16 | 0 | 0 | 16 |
| M11, Write API                              | 12 | 0 | 0 | 12 |
| M12, Charts                                 | 12 | 0 | 0 | 12 |
| M13, Bindings and tooling                   | 18 | 0 | 0 | 18 |
| X, Cross-cutting (opportunistic)            | 4  | 0 | 0 | 4  |
| **Total** | **150** | **12** | **0** | **138** |
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
| F-013 | Create oxml-core                             | S03 | M | pending |
| F-014 | New unit types                               | S03 | M | pending |
| F-015 | rdocx-oxml becomes a facade                  | S03 | S | pending |
| F-016 | Length re-export                             | S03 | S | pending |
| F-017 | App and custom properties                    | S03 | M | pending |
| F-018 | Create oxml-opc                              | S04 | M | pending |
| F-019 | PresentationML relationship and content types| S04 | S | pending |
| F-020 | oxml-opc reads a pptx                        | S04 | M | pending |
| F-021 | Zip-slip hardening tests                     | S04 | S | pending |
| F-022 | rdocx-opc deprecation shim                   | S04 | S | pending |
<!-- AUTOGEN:backlog-M2 END -->

### M3, Media

<!-- AUTOGEN:backlog-M3 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-023 | oxml-media format sniffing                   | S05 | M | pending |
| F-024 | Image probing and DPI                        | S05 | L | pending |
| F-025 | MediaNamer                                   | S05 | S | pending |
| F-026 | native_size with explicit DPI                | S05 | S | pending |
| F-027 | rdocx adopts oxml-media                      | S05 | M | pending |
| F-028 | add_picture_auto                             | S05 | S | pending |
<!-- AUTOGEN:backlog-M3 END -->

### M4, Layout primitives

<!-- AUTOGEN:backlog-M4 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-029 | Create oxml-layout                           | S06 | M | pending |
| F-030 | Decouple line.rs                             | S06 | L | pending |
| F-031 | Transform                                    | S06 | M | pending |
| F-032 | Path and PathCommand                         | S07 | M | pending |
| F-033 | Paint and Stroke                             | S07 | M | pending |
| F-034 | Path and Group arms                          | S07 | M | pending |
| F-035 | The walk helper                              | S07 | S | pending |
| F-036 | MediaId                                      | S07 | S | pending |
<!-- AUTOGEN:backlog-M4 END -->

### M5, PDF backend

<!-- AUTOGEN:backlog-M5 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-037 | Create oxml-pdf                              | S08 | S | pending |
| F-038 | Golden-PNG harness                           | S08 | M | pending |
| F-039 | Global CTM flip                              | S08 | L | pending |
| F-040 | Group rendering                              | S09 | M | pending |
| F-041 | Path rendering                               | S09 | M | pending |
| F-042 | Rewrite the three collection passes on walk  | S09 | M | pending |
| F-044 | ExtGState alpha                              | S09 | S | pending |
| F-043 | Gradient shading dictionaries                | S10 | L | pending |
| F-045 | Rasteriser: groups, paths, gradients, dashes | S10 | L | pending |
<!-- AUTOGEN:backlog-M5 END -->

### M6, rdocx 0.3.0 release

<!-- AUTOGEN:backlog-M6 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-046 | rdocx-pdf deprecation shim                   | S11 | S | pending |
| F-047 | Packaging include and size gate              | S11 | M | pending |
| F-048 | Replace release.sh with cargo-release        | S11 | M | pending |
| F-049 | Rework publish.yml                           | S11 | M | pending |
| F-050 | CI matrix additions                          | S11 | S | pending |
| F-051 | CHANGELOG and migration notes                | S11 | S | pending |
<!-- AUTOGEN:backlog-M6 END -->

### M7, DrawingML

<!-- AUTOGEN:backlog-M7 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-052 | Create oxml-drawing and namespace constants  | S12 | S | pending |
| F-053 | OrderedRawChildren                           | S12 | M | pending |
| F-054 | Colour choices                               | S12 | M | pending |
| F-055 | The colour transform stack                   | S12 | L | pending |
| F-056 | Colour map resolution                        | S12 | M | pending |
| F-057 | a:xfrm                                       | S13 | M | pending |
| F-058 | Guide evaluator                              | S13 | L | pending |
| F-059 | a:custGeom                                   | S13 | M | pending |
| F-060 | Fills                                        | S13 | L | pending |
| F-061 | Lines                                        | S14 | M | pending |
| F-062 | Effects                                      | S14 | S | pending |
| F-063 | Shape properties and style references        | S14 | M | pending |
| F-064 | DrawingML text model                         | S14 | XL | pending |
| F-065 | Theme read and write                         | S15 | L | pending |
| F-066 | The rdocx Theme adapter                      | S15 | S | pending |
<!-- AUTOGEN:backlog-M7 END -->

### M8, PresentationML

<!-- AUTOGEN:backlog-M8 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-067 | Create rpptx-oxml and the corpus harness     | S16 | M | pending |
| F-068 | presentation.xml                             | S16 | M | pending |
| F-069 | Slide, layout and master parts               | S16 | L | pending |
| F-070 | The shape tree                               | S16 | L | pending |
| F-071 | Placeholders                                 | S17 | M | pending |
| F-072 | Pictures                                     | S17 | M | pending |
| F-073 | Graphic frames                               | S17 | M | pending |
| F-074 | DrawingML tables                             | S17 | L | pending |
| F-075 | Connectors                                   | S18 | S | pending |
| F-076 | mc:AlternateContent                          | S18 | M | pending |
| F-077 | Notes slides and notes master                | S18 | M | pending |
| F-078 | relmap rewrite_rel_ids                       | S18 | M | pending |
| F-079 | The rpptx read facade                        | S19 | L | pending |
| F-080 | Modelled round-trip gate                     | S19 | M | pending |
<!-- AUTOGEN:backlog-M8 END -->

### M9, Inheritance resolver

<!-- AUTOGEN:backlog-M9 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-081 | ResolveCtx skeleton and placeholder chain    | S20 | M | pending |
| F-082 | Effective transform and body properties      | S20 | M | pending |
| F-083 | The seven-step list style merge              | S20 | L | pending |
| F-084 | Format scheme reference resolution           | S20 | M | pending |
| F-085 | Typeface resolution                          | S20 | S | pending |
| F-086 | Draw order and the flattener                 | S21 | L | pending |
| F-087 | ResolvedSlide contract                       | S21 | M | pending |
| F-088 | Visual differential tests                    | S21 | M | pending |
<!-- AUTOGEN:backlog-M9 END -->

### M10, Renderer

<!-- AUTOGEN:backlog-M10 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-089 | Resolve the preset geometry licensing question | S22 | S | pending |
| F-090 | Preset table generator                       | S22 | L | pending |
| F-091 | Preset evaluation and fallback               | S22 | M | pending |
| F-092 | rpptx-render skeleton and RenderInput        | S22 | M | pending |
| F-093 | Shape geometry, fills and lines              | S23 | L | pending |
| F-094 | Rotation, flips and groups                   | S23 | M | pending |
| F-095 | Arrowheads                                   | S23 | S | pending |
| F-096 | Pictures with crop and tile                  | S23 | M | pending |
| F-097 | Backgrounds                                  | S23 | S | pending |
| F-098 | Shape text layout                            | S24 | XL | pending |
| F-099 | Bullets                                      | S24 | M | pending |
| F-100 | Autofit                                      | S24 | M | pending |
| F-101 | Vertical text                                | S24 | S | pending |
| F-102 | Table rendering                              | S25 | L | pending |
| F-103 | Hyperlinks, fields and diagnostics           | S25 | M | pending |
| F-104 | SSIM fidelity harness                        | S25 | L | pending |
<!-- AUTOGEN:backlog-M10 END -->

### M11, Write API

<!-- AUTOGEN:backlog-M11 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-105 | Bundled default.pptx                         | S26 | M | pending |
| F-106 | ShapeIdAllocator and MediaStore              | S26 | M | pending |
| F-107 | add_slide                                    | S26 | L | pending |
| F-108 | validate()                                   | S26 | M | pending |
| F-109 | Shape mutation facade                        | S27 | L | pending |
| F-110 | add_textbox, add_shape, add_connector, group | S27 | M | pending |
| F-111 | add_picture                                  | S27 | M | pending |
| F-112 | Text frame mutation                          | S27 | L | pending |
| F-113 | Table facade                                 | S28 | L | pending |
| F-114 | remove_slide, move_slide, duplicate_slide    | S28 | M | pending |
| F-115 | Slide and presentation properties            | S28 | S | pending |
| F-116 | Cross-viewer acceptance                      | S28 | M | pending |
<!-- AUTOGEN:backlog-M11 END -->

### M12, Charts

<!-- AUTOGEN:backlog-M12 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-117 | oxml-sml workbook writer                     | S29 | L | pending |
| F-118 | ChartML core types                           | S29 | L | pending |
| F-119 | Series and data references                   | S29 | L | pending |
| F-120 | Axes                                         | S30 | L | pending |
| F-121 | Bar and line plots                           | S30 | M | pending |
| F-122 | Pie, doughnut, area, scatter and radar plots | S30 | L | pending |
| F-123 | Data labels and number formats               | S30 | M | pending |
| F-124 | add_chart                                    | S31 | L | pending |
| F-125 | Chart rendering: geometry                    | S31 | L | pending |
| F-126 | Chart rendering: axes, gridlines and labels  | S31 | L | pending |
| F-127 | Chart colour resolution                      | S32 | M | pending |
| F-128 | Preserved chart fallback                     | S32 | S | pending |
<!-- AUTOGEN:backlog-M12 END -->

### M13, Bindings and tooling

<!-- AUTOGEN:backlog-M13 START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-129 | oxml-py-support                              | S33 | M | pending |
| F-130 | rdocx-py core                                | S33 | L | pending |
| F-131 | rdocx-py formatting and tables               | S33 | L | pending |
| F-132 | Python enums, units and exceptions           | S33 | M | pending |
| F-133 | rdocx-py rendering with allow_threads        | S33 | S | pending |
| F-134 | Type stubs and py.typed                      | S34 | M | pending |
| F-135 | python-docx parity suite                     | S34 | M | pending |
| F-136 | rpptx-py                                     | S34 | L | pending |
| F-137 | wheels.yml                                   | S34 | M | pending |
| F-138 | PR-time Python job                           | S34 | S | pending |
| F-139 | Rewrite rdocx-wasm                           | S35 | L | pending |
| F-140 | wasm CI job                                  | S35 | S | pending |
| F-141 | to_pdf in the browser                        | S35 | M | pending |
| F-142 | rpptx-wasm                                   | S35 | M | pending |
| F-143 | oxml-cli-support                             | S36 | S | pending |
| F-144 | rpptx-cli                                    | S36 | L | pending |
| F-145 | rpptx-cli thumbnail and outline              | S36 | M | pending |
| F-146 | npm publication                              | S36 | S | pending |
<!-- AUTOGEN:backlog-M13 END -->

### X, Cross-cutting

<!-- AUTOGEN:backlog-MX START -->
| F-ID | Title | Sprint | Size | Status |
|------|-------|--------|------|--------|
| F-X001 | rdocx-cli tests                             | -   | M | pending |
| F-X002 | README example correctness                  | -   | S | pending |
| F-X003 | Deduplicate the sample generators           | -   | S | pending |
| F-X004 | Fix the shared temp path in the test suite  | -   | S | pending |
<!-- AUTOGEN:backlog-MX END -->
