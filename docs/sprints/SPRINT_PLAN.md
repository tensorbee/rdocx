# Sprint Plan

Sprint-by-sprint roadmap for the shared OOXML infrastructure, the Word and
PowerPoint families, and the conditional advanced spreadsheet programme.
Sprints are dependency and review boundaries, not fixed two-week containers.
Sprint clocks start at the first `/start-feature` of that sprint, not at a fixed
calendar date.

The active roadmap runs through S79, with earlier deferred cutover boundaries
retained in place. The sizing rationale and compression options are in
`docs/hld/14-development-backlog.md`.

M1 to M5 stage the extraction without changing released rdocx dependencies.
M7 to M12 build rpptx. Deferred M6 publication and rdocx cutover follow M12,
then M13 ships the bindings for both.

## Capacity calibration after S03

S01 through S03 completed 15 stories representing 24 estimated developer-days
in 15 recorded actual days. That is 5.00 stories per working week and 1.6
estimated days per actual day. The story-count rate is not used alone because
the remaining plan contains substantially more L and XL work than the first
three sprints.

The remaining 366 estimated developer-days therefore forecast to about 46
active working weeks at the observed weighted rate. Use 45 to 50 weeks as the
current planning range. Keep the existing dependency-defined sprint and
milestone boundaries, and size each implementation wave by dependencies,
exclusive resources, and estimated days rather than filling a fixed calendar
box. Recalculate after S06, when the evidence includes the first post-baseline
L story and the higher-risk layout extraction.

## Goals per sprint

### M1, Preparation and safety net

#### Sprint S01, The safety net

**Goal**: rendering is reproducible across machines, a byte-level baseline
exists for every sample, and the three shipped defects found during the audit
are fixed. Nothing has moved yet.

| F-ID | Title | Size |
|------|-------|------|
| F-001 | Deterministic font mode                      | M |
| F-002 | rust-toolchain.toml                          | S |
| F-003 | Output-stability hash harness                | L |
| F-004 | Caladea licence and the false OFL claim      | S |
| F-005 | Fix the image counter                        | S |
| F-006 | Fix the JPEG standalone-marker walk          | S |

F-001 gates F-003: a baseline recorded against system fonts would not reproduce
on another machine, which would make the harness worthless.

#### Sprint S02, Prerequisites and the pre-churn tag

**Goal**: everything the later milestones depend on is in place, and a
known-good published state is tagged immediately before the extraction begins.

| F-ID | Title | Size |
|------|-------|------|
| F-007 | Resolve core properties through the rel      | S |
| F-008 | Non-consuming setter twins                   | M |
| F-009 | Cache the layout result                      | M |
| F-010 | Reserve crate names                          | S |
| F-011 | Pin unit truncation behaviour                | S |
| F-012 | Tag v0.4.1                                   | S |

F-008 is required by M13 and improves the Rust API independently. F-011 must
land before anyone is tempted to change truncation to rounding.

### M2, Shared infrastructure extraction

#### Sprint S03, oxml-core

**Goal**: the generic types leave `rdocx-oxml` and 323 call sites do not change.

| F-ID | Title | Size |
|------|-------|------|
| F-013 | Create oxml-core                             | M |
| F-014 | New unit types                               | M |
| F-017 | App and custom properties                    | M |

#### Sprint S04, oxml-opc

**Goal**: the package layer is format-neutral and proven against a real pptx.

| F-ID | Title | Size |
|------|-------|------|
| F-018 | Create oxml-opc                              | M |
| F-019 | PresentationML relationship and content types| S |
| F-020 | oxml-opc reads a pptx                        | M |
| F-021 | Zip-slip hardening tests                     | S |

F-015 and F-016 carried from S03, and F-022 joined their deferred cutover. They
are rescheduled to S32.2 after PowerPoint development and shared-crate
publication readiness. Passing the rdocx 0.5.0 release boundary protects that
release but does not make an unpublished implementation available to later
package dry-runs. F-020 converts the plan's central package assumption into a
test.

### M3, Media

#### Sprint S05, oxml-media

**Goal**: stage and prove one crate that owns image sniffing, dimensions, and
naming without changing released rdocx dependencies.

| F-ID | Title | Size |
|------|-------|------|
| F-023 | oxml-media format sniffing                   | M |
| F-024 | Image probing and DPI                        | L |
| F-025 | MediaNamer                                   | S |
| F-026 | native_size with explicit DPI                | S |

F-027 and F-028 move to S32.2. F-027 retains a focused package regression for
the intentional change from trusted extensions to sniffed content types. The
existing hash harness remains unchanged because it does not collect package
content types or media relationship targets.

### M4, Layout primitives

#### Sprint S06, oxml-layout and the line.rs decoupling

**Goal**: the format-neutral layout types are staged in isolation, including
the one file that needs genuine API design.

| F-ID | Title | Size |
|------|-------|------|
| F-029 | Create oxml-layout                           | M |
| F-030 | Decouple line.rs                             | L |
| F-031 | Transform                                    | M |

F-030 is the highest drift risk in the extraction. Own PR, own review, gated
hard on the hash harness.

#### Sprint S07, The PositionedElement extension

**Goal**: the staged shared element type can express a rotated, clipped,
gradient-filled shape without changing released rdocx construction sites.

| F-ID | Title | Size |
|------|-------|------|
| F-032 | Path and PathCommand                         | M |
| F-033 | Paint and Stroke                             | M |
| F-034 | Path and Group arms                          | M |
| F-035 | The walk helper                              | S |
| F-036 | MediaId                                      | S |

Two new arms, not ten new fields. F-035 exists specifically to prevent the
recursion hazard that S09 then tests for.

### M5, PDF backend

#### Sprint S08, The coordinate system

**Goal**: the renderer moves to one global CTM with zero pixel change.

| F-ID | Title | Size |
|------|-------|------|
| F-037 | Create oxml-pdf                              | S |
| F-038 | Golden-PNG harness                           | M |
| F-039 | Global CTM flip                              | L |

F-039 is the single highest-risk change in the plan. It lands before
PresentationML rendering code exists, so a regression has only one possible
cause.

#### Sprint S09, Groups, paths and the recursion fix

**Goal**: nested content renders, and the three collection passes see inside
groups.

| F-ID | Title | Size |
|------|-------|------|
| F-040 | Group rendering                              | M |
| F-041 | Path rendering                               | M |
| F-042 | Rewrite the three collection passes on walk  | M |
| F-044 | ExtGState alpha                              | S |

F-042 is the R3 regression gate. Its three tests are the only thing standing
between this design and PDFs that silently lose fonts, images or links.

#### Sprint S10, Gradients and the rasteriser

**Goal**: both backends render everything the element types can express.

| F-ID | Title | Size |
|------|-------|------|
| F-043 | Gradient shading dictionaries                | L |
| F-045 | Rasteriser: groups, paths, gradients, dashes | L |

F-045 also fixes the dash pattern that all PNG output currently discards.

### M6, deferred shared publication and rdocx cutover

#### Sprint S11, Staged extraction gate

**Goal**: verify the isolated shared crates and continue PowerPoint development
without publishing them or changing released rdocx dependencies.

No publication or consumer-cutover story runs at S11. F-046 through F-051 are
rescheduled to S32.1 and S32.2 after PowerPoint development. S11 is the staged
extraction validation boundary before DrawingML construction begins.

### M7, DrawingML

#### Sprint S12, Colour

**Goal**: a theme colour with a transform stack resolves to the exact RGB
PowerPoint produces.

| F-ID | Title | Size |
|------|-------|------|
| F-052 | Create oxml-drawing and namespace constants  | S |
| F-053 | OrderedRawChildren                           | M |
| F-054 | Colour choices                               | M |
| F-055 | The colour transform stack                   | L |
| F-056 | Colour map resolution                        | M |

F-055's test gate is a table of 40 pairs sampled from real renders. Getting
`lumMod` wrong makes an entire deck the wrong shade.

#### Sprint S13, Geometry and fills

**Goal**: any shape's outline and fill can be described.

| F-ID | Title | Size |
|------|-------|------|
| F-057 | a:xfrm                                       | M |
| F-058 | Guide evaluator                              | L |
| F-059 | a:custGeom                                   | M |
| F-060 | Fills                                        | L |

F-058 is what makes the preset table a data problem in M10 rather than a code
problem.

#### Sprint S14, Lines, effects and text

**Goal**: the DrawingML text vocabulary is modelled.

| F-ID | Title | Size |
|------|-------|------|
| F-061 | Lines                                        | M |
| F-062 | Effects                                      | S |
| F-063 | Shape properties and style references        | M |
| F-064 | DrawingML text model                         | XL |
| F-064a | Text body properties and shell               | M |
| F-064b | Text paragraphs and runs                     | L |
| F-064c | Text bullets                                 | S |
| F-064d | Nine-level list styles                       | M |

F-064 is the umbrella gate. Its implementation is split into F-064a through
F-064d, and the parent closes only after every child closes.

#### Sprint S15, Theme

**Goal**: themes read and write, and rdocx adopts the shared type without
changing behaviour.

| F-ID | Title | Size |
|------|-------|------|
| F-065 | Theme read and write                         | L |
| F-066 | The rdocx Theme adapter                      | S |

F-066's test gate is the hash harness being unchanged. The Word tint and shade
path is deliberately left alone.

### M8, PresentationML

#### Sprint S16, Parts and the shape tree

**Goal**: the corpus round-trips with everything opaque, then with the core
parts modelled.

| F-ID | Title | Size |
|------|-------|------|
| F-067 | Create rpptx-oxml and the corpus harness     | M |
| F-068 | presentation.xml                             | M |
| F-069 | Slide, layout and master parts               | L |
| F-070 | The shape tree                               | L |

F-067's raw round-trip proves the OPC layer and the corpus harness before any
XML modelling exists.

#### Sprint S17, Placeholders, pictures and tables

| F-ID | Title | Size |
|------|-------|------|
| F-071 | Placeholders                                 | M |
| F-072 | Pictures                                     | M |
| F-073 | Graphic frames                               | M |
| F-074 | DrawingML tables                             | L |

#### Sprint S18, The long tail

| F-ID | Title | Size |
|------|-------|------|
| F-075 | Connectors                                   | S |
| F-076 | mc:AlternateContent                          | M |
| F-077 | Notes slides and notes master                | M |
| F-078 | relmap rewrite_rel_ids                       | M |

F-078 is what makes deep copy safe in M11. Without it a duplicated slide's
SmartArt points at the source slide's relationships.

#### Sprint S19, The read facade

**Goal**: open any deck and read it.

| F-ID | Title | Size |
|------|-------|------|
| F-079 | The rpptx read facade                        | L |
| F-080 | Modelled round-trip gate                     | M |

**This is the M8 gate**: all 50 decks round-trip and every one opens in
PowerPoint without a repair prompt.

### M9, Inheritance resolver

#### Sprint S20, The chains

**Goal**: every inherited property resolves.

| F-ID | Title | Size |
|------|-------|------|
| F-081 | ResolveCtx skeleton and placeholder chain    | M |
| F-082 | Effective transform and body properties      | M |
| F-083 | The seven-step list style merge              | L |
| F-084 | Format scheme reference resolution           | M |
| F-085 | Typeface resolution                          | S |

#### Sprint S21, Draw order and the contract

**Goal**: `ResolvedSlide` is frozen and correct.

| F-ID | Title | Size |
|------|-------|------|
| F-086 | Draw order and the flattener                 | L |
| F-087 | ResolvedSlide contract                       | M |
| F-088 | Visual differential tests                    | M |

F-086's test gate is that a rendered slide contains no "Click to edit Master
title style". Placeholders on layouts and masters are templates, never drawn.

### M10, Renderer

#### Sprint S22, Geometry

| F-ID | Title | Size |
|------|-------|------|
| F-089 | Resolve the preset geometry licensing question | S |
| F-090 | Preset table generator                       | L |
| F-091 | Preset evaluation and fallback               | M |
| F-092 | rpptx-render skeleton and RenderInput        | M |

F-089 is a decision, not code, and it blocks F-090. LibreOffice's table is
MPL-2.0 and cannot be used.

#### Sprint S23, Shapes

| F-ID | Title | Size |
|------|-------|------|
| F-093 | Shape geometry, fills and lines              | L |
| F-094 | Rotation, flips and groups                   | M |
| F-095 | Arrowheads                                   | S |
| F-096 | Pictures with crop and tile                  | M |
| F-097 | Backgrounds                                  | S |

Ships slides with shapes but no text.

#### Sprint S24, Text

| F-ID | Title | Size |
|------|-------|------|
| F-098 | Shape text layout                            | XL |
| F-098a | Text content box                             | M |
| F-098b | Paragraph inline resolution                  | L |
| F-098c | Line stacking                                | M |
| F-098d | Text anchoring                               | S |
| F-099 | Bullets                                      | M |
| F-100 | Autofit                                      | M |
| F-101 | Vertical text                                | S |

**The milestone that makes the project real.** F-098 is the umbrella gate for
F-098a through F-098d, which implement the content box, paragraph inline
resolution, line stacking, and anchoring. The parent closes only after every
child closes.

#### Sprint S25, Tables and the fidelity gate

| F-ID | Title | Size |
|------|-------|------|
| F-102 | Table rendering                              | L |
| F-103 | Hyperlinks, fields and diagnostics           | M |
| F-104 | SSIM fidelity harness                        | L |

**This is the M10 gate** and the natural point to cut an early
read-plus-render release if the schedule needs compressing.

### M11, Write API

#### Sprint S26, Slides

| F-ID | Title | Size |
|------|-------|------|
| F-105 | Bundled default.pptx                         | M |
| F-106 | ShapeIdAllocator and MediaStore              | M |
| F-107 | add_slide                                    | L |
| F-108 | validate()                                   | M |

F-108 will save more debugging time than any other story in the backlog.

#### Sprint S27, Shapes and text

| F-ID | Title | Size |
|------|-------|------|
| F-109 | Shape mutation facade                        | L |
| F-110 | add_textbox, add_shape, add_connector, group | M |
| F-111 | add_picture                                  | M |
| F-112 | Text frame mutation                          | L |

#### Sprint S28, Tables and acceptance

| F-ID | Title | Size |
|------|-------|------|
| F-113 | Table facade                                 | L |
| F-114 | remove_slide, move_slide, duplicate_slide    | M |
| F-115 | Slide and presentation properties            | S |
| F-116 | Cross-viewer acceptance                      | M |

**This is the M11 gate**: a generated deck opens clean in PowerPoint, Keynote,
Google Slides and LibreOffice.

### M12, Charts

#### Sprint S29, The data layer

| F-ID | Title | Size |
|------|-------|------|
| F-117 | oxml-sml workbook writer                     | L |
| F-118 | ChartML core types                           | L |
| F-119 | Series and data references                   | L |

F-119's caches are what actually render. A chart written without them is empty
in most viewers.

#### Sprint S30, Axes and plots

| F-ID | Title | Size |
|------|-------|------|
| F-120 | Axes                                         | L |
| F-121 | Bar and line plots                           | M |
| F-122 | Pie, doughnut, area, scatter and radar plots | L |
| F-123 | Data labels and number formats               | M |

#### Sprint S31, Authoring and rendering

| F-ID | Title | Size |
|------|-------|------|
| F-124 | add_chart                                    | L |
| F-125 | Chart rendering: geometry                    | L |
| F-126 | Chart rendering: axes, gridlines and labels  | L |

#### Sprint S32, Chart polish

| F-ID | Title | Size |
|------|-------|------|
| F-127 | Chart colour resolution                      | M |
| F-128 | Preserved chart fallback                     | S |

### Deferred shared publication and rdocx cutover

#### Sprint S32.1, Shared publication readiness

**Goal**: make the completed shared crates packageable, fully gated, and ready
for an explicitly approved publication without publishing from this sprint.

| F-ID | Title | Size |
|------|-------|------|
| F-047 | Packaging include and size gate              | M |
| F-048 | Automate split-family release preparation    | M |
| F-049 | Extend publish.yml to the extracted workspace| M |
| F-050 | CI matrix additions                          | S |

After S32.1, publication runs only through a separate reviewed release plan
with explicit approval. S32.2 cannot start until the registry contains the
approved shared-crate versions and a clean consumer resolves those versions.

#### Sprint S32.2, Released rdocx cutover

**Goal**: after the real shared crates are published through their approved
release plan, move released rdocx consumers onto them and document the cutover.

| F-ID | Title | Size |
|------|-------|------|
| F-X005 | Tag rpptx-v0.1.2                           | S |
| F-015 | rdocx-oxml becomes a facade                  | S |
| F-016 | Length re-export                             | S |
| F-022 | rdocx-opc deprecation shim                   | S |
| F-027 | rdocx adopts oxml-media                      | M |
| F-028 | add_picture_auto                             | S |
| F-046 | rdocx layout and PDF cutover                 | M |
| F-051 | CHANGELOG and migration notes                | S |

**This is the deferred M6 release gate.** The shared crates are real published
dependencies, released rdocx packages pass archive verification, and the hash
harness remains unchanged while a focused regression proves the F-027
content-type change.

### M13, Bindings and tooling

#### Sprint S33, rdocx-py

**Goal**: the handle design is validated against the settled API before it is
reused.

| F-ID | Title | Size |
|------|-------|------|
| F-129 | oxml-py-support                              | M |
| F-130 | rdocx-py core                                | L |
| F-131 | rdocx-py formatting and tables               | L |
| F-132 | Python enums, units and exceptions           | M |
| F-133 | rdocx-py rendering with allow_threads        | S |

#### Sprint S34, Wheels and rpptx-py

| F-ID | Title | Size |
|------|-------|------|
| F-134 | Type stubs and py.typed                      | M |
| F-135 | python-docx parity suite                     | M |
| F-136 | rpptx-py                                     | L |
| F-137 | wheels.yml                                   | M |
| F-138 | PR-time Python job                           | S |

#### Sprint S35, WASM

**Goal**: the wasm crates wrap the real facades and are watched by CI.

| F-ID | Title | Size |
|------|-------|------|
| F-139 | Rewrite rdocx-wasm                           | L |
| F-140 | wasm CI job                                  | S |
| F-141 | to_pdf in the browser                        | M |
| F-142 | rpptx-wasm                                   | M |

F-139 fixes a shipped defect that silently discards every package part except
two. F-140 is why it will not happen again.

#### Sprint S36, CLIs and local packaging

| F-ID | Title | Size |
|------|-------|------|
| F-143 | oxml-cli-support                             | S |
| F-144 | rpptx-cli                                    | L |
| F-145 | rpptx-cli thumbnail and outline              | M |
| F-146 | npm publication                              | S |
| F-X001 | rdocx-cli tests                             | M |
| F-X002 | README example correctness                  | S |
| F-X003 | Deduplicate the sample generators           | S |
| F-X004 | Fix the shared temp path in the test suite  | S |

**This is the v1 implementation gate.** The CLI and local package surfaces are
complete. Registry publication is explicitly deferred. The expanded
incubating Rust family requires a fresh version and reviewed release through
F-X006. npm registry publication requires a separate future story.

#### Sprint S37, Expanded Rust family release

**Goal**: prepare a fresh common incubating version and publish the complete
14-package Rust family through the reviewed release workflow after separate
final approval.

| F-ID | Title | Size |
|------|-------|------|
| F-X006 | Tag the expanded rpptx family               | S |

#### Sprint S38, Contributor integration and stable release

**Goal**: land PR 25 with its contributor credit intact, harden the new Word
composition APIs against package and table-geometry regressions, and make the
stable crate family documentation useful at the point of publication. Prepare
and publish the complete stable family at the breaking pre-1.0 0.5.0 boundary
only after the integrated result passes review and receives separate release
approval.

| F-ID | Title | Size |
|------|-------|------|
| F-X007 | Integrate PR 25 and stable crate documentation | L |
| F-X008 | Tag v0.5.0                                      | S |

F-X008 depends on F-X007. The GitHub PR is merged into the sprint branch so
the contributor remains credited in the pull request and merge record. Only
`/close-sprint` later merges the reviewed sprint to `main`.

#### Sprint S39, Workspace crate documentation and next-minor releases

**Goal**: every one of the 26 Cargo workspace packages declares a useful
README. Each README explains the crate's role, intended audience, relationship
to neighbouring packages, and includes a concrete example in the language or
command surface that users actually consume. Publish those READMEs for every
crates.io-eligible package through separate next-minor stable and incubating
release tags.

| F-ID | Title | Size |
|------|-------|------|
| F-X009 | README coverage for every workspace crate | L |
| F-X010 | Tag v0.6.0 | S |
| F-X011 | Tag rpptx-v0.2.0 | S |

The existing README runner becomes an exact 26-package contract. Published
archives must each contain one intended README, while unpublished Python and
WASM packages receive accurate local usage examples without gaining any new
publication authority. F-X010 publishes the seven stable crates first. F-X011
then publishes the fourteen incubating crates. Each tag has its own full
verification, clean review, and immediate release approval boundary.

#### Sprint S40, Restore pinned CI toolchains

**Goal**: restore a green hosted CI baseline after runner and package-manager
updates exposed unpinned or incorrectly validated external tools. Keep the
reviewed Poppler 26.01.0 rendering oracle and Binaryen 125 optimizer boundary
without changing product output or recorded rendering baselines.

| F-ID | Title | Size |
|------|-------|------|
| F-X012 | Restore pinned CI toolchains | M |

The story installs checksum-pinned Poppler 26.01.0 for every job that executes
its oracle-dependent tests, validates the official Binaryen 125 Linux version
string, and proves the complete pull-request workflow on a hosted runner. It
does not change a crate, release version, published package, or rendering
baseline.

#### Sprint S41, Footnote placement and floating drawing wrapping

**Goal**: land the parts of the external PR 2 contribution that current `main`
still lacks, rebuilt on the anchor architecture that superseded the
contributor's own. Fix the two footnote placement defects, then give anchored
drawings a real wrap model and make body text flow around them.

| F-ID | Title | Size |
|------|-------|------|
| F-X013 | Footnote and endnote placement | M |
| F-X013a | Footnote line advance | S |
| F-X013b | Footnote reservation and splitting | L |
| F-X013c | Endnotes at the document end | M |
| F-X014 | Kashida justification values | S |
| F-X015 | Anchored drawing wrap and alignment model | M |
| F-X016 | Floating drawing placement and text wrapping | L |

The note work lands first, because it is independent of the drawing work and
each child carries its own baseline delta. F-X013 was planned as a single M and
split into three children during its design, when splitting oversized notes and
correcting endnote placement were taken into scope. F-X014 is a one-line parser
widening carried in the same wave because it comes from the same contribution.
F-X015 adds the wrap and alignment surface without changing placement, which
keeps the harness flat and makes F-X016 the single story that owns the rendering
delta for wrapped drawings.

#### Sprint S42, Dependency refresh

**Goal**: take the outstanding semver-compatible dependency updates and measure
what they do to rendered output. Nothing here is a security fix, since the
advisory scan is already clean, so the value is in not letting the lockfile
drift far enough that a later update becomes a large unexplained delta.

| F-ID | Title | Size |
|------|-------|------|
| F-X020 | Refresh the dependency lockfile | S |
| F-X024 | Move the theme adapter into rdocx-oxml | M |
| F-X022 | Tag rpptx-v0.3.0 | S |
| F-X023 | Tag v0.7.0 | S |

Two of the sixteen pending updates are in the font and image decoding path, so
the hash harness decides whether the refresh is a no-op or a declared rendering
delta. It runs first and alone, because a refresh that moved a baseline should
not compete with a release to explain the same delta.

F-X024 then removes the reason the release order was impossible. Scoping the two
release stories exposed a cycle between the trains: `rdocx-layout` depends on
`oxml-layout`, and `oxml-drawing` depends on `rdocx-oxml` through the one
documented architecture exception. With both trains carrying breaking changes,
neither could publish first. Moving the theme adapter into `rdocx-oxml` inverts
that edge, so the dependency runs one way and incubating always publishes
first.

The two release stories then carry S41's work to crates.io in that order.

#### Sprint S43, Robustness and gate coverage

**Goal**: clear the follow-ups S41 and S42 filed. Three are defects that a real
document can reach, and two close gaps in the gates that let the other three
survive as long as they did.

| F-ID | Title | Size |
|------|-------|------|
| F-X018 | Unknown enumerated values must not fail a document open | M |
| F-X017 | Notes broken to their own section's width | S |
| F-X019 | Paragraph-relative later drawings should wrap | M |
| F-X021 | The hash harness should cover PDF output | M |
| F-X025 | /verify must run the release regressions | S |

F-X018 leads because it is the only one where a document fails to open rather
than rendering imperfectly. F-X014 fixed the three kashida values because a real
contribution reached them, and eight more parsers have the same shape.

F-X017 and F-X019 are the two limitations S41 recorded rather than hid, both
narrow and both reachable by a real document. F-X021 and F-X025 are the gates:
one gives the harness PDF coverage it has never had, the other makes `/verify`
run the release regressions that `publish.yml` treats as its publication gate.
The gate stories land last because neither blocks the three defect fixes, and
putting them first would delay work that users can actually hit.

#### Sprint S44, Gate coverage and specification repair

**Goal**: finish the job S43 started. S43 closed two gaps in the gates and
found, in passing, that the records describing those gates had drifted from
them. This sprint puts the two remaining gates where CI can see them and repairs
the documentation that tells every future session what is true.

| F-ID | Title | Size |
|------|-------|------|
| F-X026 | CI must run the release regressions too | S |
| F-X027 | Wire the golden-PNG gate into something | S |
| F-X029 | Path-filtered CI jobs | M |
| F-X028 | Repair the agent-facing documentation drift | M |

Every implementation milestone is closed, so this sprint carries no feature
work. Three of the four exist because S43 went looking at the instruments rather
than the product. F-X029 came out of a review of whether the workspace should be
split into separate repositories. That review also produced F-X030, archived
before the sprint started because the WASM packages are deliberately
unpublished and its stated problem does not exist.

No pair has a hard dependency, so the order is a preference. The one soft
coupling is F-X026 and F-X029, which both edit `ci.yml`. F-X026 first, because it is the narrower
half of a gap S43 half-closed: `/verify` runs the release preflights now and CI
still does not, so a contributor who skips the local gate can move a version
carrier and see a green pull request. F-X027 next, because the golden-PNG gate
is fully specified and wired into nothing, and deciding where it belongs needs a
judgement about the pinned Poppler build that F-X026 does not need.

F-X028 lands last and is the largest, because it is the only one that touches
`CLAUDE.md`, and a story that rewrites the file every other session reads first
should land against a tree the other two have already settled.

## Post-v1 roadmap, S45 onward

v1 shipped at S43 and every implementation milestone closed. S44 repairs the
gates and the records. From S45 the plan resumes at milestone granularity
against `14-development-backlog.md` M14 through M22.

The order is deliberate and each boundary is a stopping point. Stopping after
S45 leaves one chart engine serving both families. Stopping after S51 leaves a
document-automation product. Stopping after S69 leaves every planned Word and
PowerPoint capability complete. The advanced spreadsheet programme starts only
after that boundary and only if its feasibility gate confirms a gap worth
filling in the Rust ecosystem.

| Sprints | Milestone | Stories | Days |
|---|---|---|---|
| S45 | M15, charts beyond PowerPoint | 4 | 12 |
| S46 to S48 | M14, Word collaboration layer | 9 | 28 |
| S49 to S51 | M16, document automation | 15 | 39 |
| S52 to S53 | M17, security and compliance | 7 | 23 |
| S54 to S56 | M18, format breadth | 8 | 26 |
| S57 to S58 | M20, fidelity at scale | 7 | 27 |
| S59 to S64 | M21, presentation depth | 15 | 60 |
| S65 to S69 | M22, Word depth | 12 | 44 |
| S70 to S79 | M19, advanced spreadsheets | 21 | 85 |

### M15, Charts beyond PowerPoint

#### Sprint S45, One chart engine

**Goal**: make the chart engine serve Word as well as PowerPoint. It is the
cheapest milestone on the roadmap and the only one whose engine already exists
on the format-neutral side of the crate graph.

| F-ID | Title | Size |
|------|-------|------|
| F-156 | Extract oxml-chart | L |
| F-157 | Word chart part and embedded workbook | M |
| F-158 | Document::add_chart | M |
| F-159 | Chart rendering in the Word paginator | M |

F-156 is a file move and nothing else. The hash harness must be byte-identical
across it, and folding a behaviour change into it is forbidden. The other three
are strictly ordered, since each needs the part the one before it writes.

### M14, Word collaboration layer

#### Sprint S46, Comments and content controls

**Goal**: open the collaboration layer at both ends. Comments are the most
requested missing API in this space and content controls are the primitive every
document-assembly product is built on.

| F-ID | Title | Size |
|------|-------|------|
| F-147 | Comment model and part | M |
| F-148 | Comment API | M |
| F-152 | Content control model | L |
| F-153 | Content control binding | M |
| F-154 | Bookmarks and cross-references | M |

Two independent pairs, so they parallelise cleanly. F-154 joins this sprint
rather than the next because F-161 needs bookmarks to resolve `REF` and
`PAGEREF`, and that is two sprints away.

#### Sprint S47, Tracked changes

**Goal**: read, write and resolve revisions. The single most demanded enterprise
capability in this space and the one with no open-source answer in any language.

| F-ID | Title | Size |
|------|-------|------|
| F-149 | Revision model | L |
| F-150 | Accept and reject revisions | L |

Two stories, both L, and deliberately alone in a sprint. F-150 has to reproduce
what Word produces from the same input, which is the kind of correctness that
takes the time it takes.

#### Sprint S48, Revision display and protection

**Goal**: close M14. Show revisions, and read the author's intent about who may
change what.

| F-ID | Title | Size |
|------|-------|------|
| F-151 | Revision display in the renderer | M |
| F-155 | Document protection | M |

A short sprint on purpose. It carries the M14 end-of-milestone gate, which
covers four subsystems built across three sprints.

### M16, Document automation

#### Sprint S49, Fields

**Goal**: evaluate the field codes real documents are full of. Everything in
M16 rests on this.

| F-ID | Title | Size |
|------|-------|------|
| F-160 | Field instruction parser | L |
| F-161 | Field evaluation engine | L |
| F-162 | Field update policy | M |
| F-203 | Reader compatibility corrections | M |

F-160 through F-162 are strictly ordered. F-203 is an independent corrective
story for reader preservation. F-161 also depends on F-154 from S46, which is
why bookmarks landed early.

#### Sprint S50, Templating

**Goal**: turn substitution into generation.

| F-ID | Title | Size |
|------|-------|------|
| F-163 | Template syntax | L |
| F-164 | Loops and conditionals | L |
| F-165 | Repeating table rows and lists | M |

F-163 leads because the tag-split-across-runs problem is the one every naive
implementation gets wrong, and the two after it inherit whatever it decides.

#### Sprint S51, Automation milestone and community release

**Goal**: close M16, add the requested native reader and editor surfaces,
establish a custom reviewed release-notes ceremony, and publish the coherent
incubating and stable trains.

| F-ID | Title | Size |
|------|-------|------|
| F-166 | Mail merge | M |
| F-167 | Document comparison | L |
| F-168 | Watermarks | S |
| F-X032 | Expose complete Word layout results | S |
| F-X033 | Integrate PR 36 ordered body items | S |
| F-X034 | Reviewed release notes for every release | S |
| F-X035 | Tag rpptx-v0.4.0 | S |
| F-X036 | Tag v0.8.0 | S |
| F-X037 | Trace Word glyphs to source paragraphs | M |
| F-X038 | Cache relayout work across document edits | L |

F-167 is the flagship of every commercial library in this category. It is scoped
to body text, tables and list structure, with formatting-only differences
recorded as a diagnostic, which is what keeps it one story. The community API
surface combines complete layouts, source provenance, bounded relayout caches,
and direct ordered body items. The incubating release precedes stable 0.8.0
because `oxml-chart`, the low-level provenance types, and shared font-cache
work must exist on crates.io before the stable dependency graph can publish.
Both release tags retain separate final approval boundaries.

### M17, Security and compliance

#### Sprint S52, Encryption and renderer follow-ups

**Goal**: open the files that currently cannot be opened at all, then close the
community-reported rendering and interactive-layout gaps with exact public
regressions.

| F-ID | Title | Size |
|------|-------|------|
| F-169 | Agile encryption, read | L |
| F-170 | Agile encryption, write | M |
| F-171 | Digital signature verification | L |
| F-X039 | Share layout payloads and transfer reusable engines | M |
| F-X040 | Restart pagination and cache table blocks | L |
| F-X041 | Remove duplicated glyphs at break opportunities | M |
| F-X042 | Prove headers and footers in PDF output | S |
| F-X043 | Reuse bundled-fallback caller-font layouts | M |
| F-X044 | Scale paragraph-cache lookup for editors | M |
| F-X045 | Cache headers and footers transactionally | M |
| F-X046 | Reuse substituted pages exactly | S |
| F-X047 | Attribute empty Word paragraphs | S |

Reading comes first and matters most. A password-protected document is a hard
stop for a user today, where an unsigned one is only a missing assurance.
F-X039 establishes shared ownership and a checked editor handoff before F-X040
retains pagination tails. F-X041 fixes Issue 23 at the common layout layer, and
F-X042 closes Issue 15 with a public DOCX-to-PDF gate. F-X043 through F-X047
close PRs 40 and 41 after retaining their useful editor behavior behind exact
context identity, transactional publication, and bounded memory. The sprint
does not import unchecked engine setters, hash-authoritative reuse, or
unbounded caches from the draft branches.

#### Sprint S53, Security milestone and community release

**Goal**: close M17, harden the dense-form layout reported by the community,
and publish the complete incubating and stable Rust families with reviewed
contributor credit.

| F-ID | Title | Size |
|------|-------|------|
| F-172 | Digital signature creation | M |
| F-173 | Tagged PDF structure tree | L |
| F-175 | Redaction | M |
| F-X048 | Dense form table fidelity | L |
| F-174 | PDF/A conformance | M |
| F-X049 | Tag rpptx-v0.5.0 | S |
| F-X050 | Tag v0.9.0 | S |

F-173 is the one a LibreOffice-based pipeline cannot do well, and the layout
engine already knows the document semantics it needs, because
`audit_accessibility` reads them. F-X048 reimplements Issue 42 and PR 43 on the
current transactional, bounded cache model rather than importing the stacked
draft branch. The two release stories run last and retain separate final
approval boundaries. Incubating 0.5.0 publishes before stable 0.9.0 so every
stable dependency pin resolves from crates.io. Both reviewed changelog sections
credit and link every included external issue and pull request, and each record
receives a maintainer release note for its contributor.

### M18, Format breadth

#### Sprint S54, RTF and caller font aliases

**Goal**: open M18 with the inbound format that blocks the most corpora, and
honour the document-facing family names supplied with caller fonts.

| F-ID | Title | Size |
|------|-------|------|
| F-176 | RTF reader | L |
| F-177 | RTF writer | M |
| F-183 | Image export options | S |
| F-X051 | Honor caller-supplied font family aliases | M |

F-183 rides along because it is a day of work against an entry point every
format in this milestone shares. F-X051 closes Issue 44 and PR 45 on the
hardened reusable-engine path from F-X043. It remains independent of the RTF
reader and writer.

#### Sprint S55, HTML and ODT in

**Goal**: add the two remaining inbound formats, restore the v0.9 editor path
to the reviewed relayout budget, and close the associated migration and
contribution records.

| F-ID | Title | Size |
|------|-------|------|
| F-178 | HTML import | L |
| F-179 | ODT reader | L |
| F-X052 | Restore interactive relayout performance | L |
| F-X053 | Complete layout migration and contribution records | S |

F-178, F-179, and F-X052 are independent. HTML import is the most requested
inbound conversion in every comparable library's tracker, and ODT is
procurement-mandated across European public bodies and read by nothing in
Rust. F-X052 closes the separate performance residual reported in Issue 46
without weakening the exact, transactional, bounded cache contract. F-X053
runs after it, records the `MarkedContent` migration requirement, and closes
Issue 44, Issue 46, and PR 45 with the exact F-X051 and F-X052 evidence.
Issues 39 and 42 remain closed because their reporter confirmed that note
operations recovered and F-X048 fully replaced the dense-form patches.

#### Sprint S56, ODT, EPUB and SVG out

**Goal**: close M18, integrate the six pending reader contributions, and
publish the stable family at its next pre-1.0 beta boundary.

| F-ID | Title | Size |
|------|-------|------|
| F-180 | ODT writer | L |
| F-181 | EPUB export | M |
| F-182 | SVG page export | M |
| F-X054 | Integrate PRs 47 through 52 | L |
| F-X055 | Tag v0.10.0 | S |
| F-X056 | Tag rpptx-v0.6.0 | S |
| F-X057 | Tag v0.10.1 | S |

F-181 and F-182 both fall out of work that exists: EPUB from the outline API,
SVG from the same `PageFrame` the PDF and PNG backends already consume.
F-X054 audits the ordered reader APIs and parser fidelity changes proposed by
PRs 47 through 52 against current main, retaining six distinct contribution
records and authenticated credit to `@pedroassumpcao`. F-X055 runs only after
all four implementation stories complete. Its immutable v0.10.0 attempt
published only `rdocx-opc` and `rdocx-oxml` before package verification
stopped. F-X057 owns the complete stable publication at v0.10.1 with reviewed
notes, compatibility guidance, direct record links, and contributor credit.
It also owns all nine notifications and the six authorized pull-request
closures after the complete publication and release body verify.

The immutable v0.10.0 attempt published `rdocx-opc` and `rdocx-oxml`, then
proved that the stable source graph requires shared layout APIs newer than the
published incubating 0.5.0 family. F-X055 is retained as an archived failed
release attempt. F-X056 publishes the complete shared and Presentation family
at 0.6.0 while stable packages remain at 0.10.0. F-X057 then publishes the
complete stable family at 0.10.1 against those verified shared dependencies.
Each tag retains its own reviewed SHA, notes, full gate, and separate final
approval.

### M20, Fidelity at scale

#### Sprint S57, The Word corpus

**Goal**: measure the Word renderer against documents nobody here wrote. This is
the largest untested surface in the workspace.

| F-ID | Title | Size |
|------|-------|------|
| F-196 | Word corpus | M |
| F-197 | Word SSIM harness | L |
| F-201 | Large document performance | L |

PowerPoint has 50 fetched decks and an SSIM harness. Word has seven samples this
project generates itself, so it can only catch a regression against its own
output and can never catch a disagreement with Word.

#### Sprint S58, Text shaping and incremental layout

**Goal**: close M20 and finish every planned non-spreadsheet capability before
the advanced spreadsheet programme begins.

| F-ID | Title | Size |
|------|-------|------|
| F-198 | Hyphenation | L |
| F-199 | Complex script shaping | L |
| F-200 | Vertical and bidirectional text | M |
| F-202 | Incremental layout | L |
| F-X061 | Support staged dependency checkpoints in run-sprint | S |
| F-X062 | Reuse restart pagination with notes and headers | M |
| F-X063 | Avoid duplicate caller-font byte comparisons | S |
| F-X058 | Shared multilingual text substrate | L |
| F-X059 | Tag rpptx-v0.7.0 | S |
| F-X064 | Accept whole-valued decimal table measurements | S |
| F-X067 | Prime Word fidelity Cargo dependencies | S |
| F-X065 | Expose tracked table grid changes | S |
| F-X066 | Classify legacy VML horizontal rules | S |
| F-X060 | Tag v0.11.0 | S |
| F-X068 | Tag rpptx-v0.8.0 | S |
| F-X069 | Tag v0.11.1 | S |
| F-X070 | Yank incomplete v0.11.0 packages | S |
| F-X031 | Require the CI gate in branch protection | S |

F-198 changes line breaking and therefore every line after the first hyphenated
one, so it lands after the corpus exists to measure it. Expect a declared hash
harness delta, and expect it to be large.

F-X061 makes ordinary and release dependency checkpoints resumable inside one
sprint. F-X062 and F-X063 close the two independently reproduced editor
performance cliffs from Issues 53 and 54. F-X058 then establishes the complete
shared text contract needed by F-198,
F-199, and F-200. F-X059 publishes that contract as the incubating 0.7.0 family
before the stable Word consumers verify against crates.io. F-X064 lands the
first external reader fix, and F-X067 primes the locked Word fidelity graph
before F-X065 and F-X066 land the remaining reader fixes. Together these are
the hardened or directly adopted outcomes of PRs 55 through 58 before the
paused Word text stories resume. F-X060 records the immutable partial v0.11.0
attempt, which published two packages before registry verification exposed the
missing shared source boundary. F-X068 publishes that shared boundary at
0.8.0, and F-X069 publishes the complete stable recovery at 0.11.1. F-X070
then yanks the two incomplete 0.11.0 registry entries after separate approval
without moving the tag. Each release retains the separate final approval
required by `/release`.

F-X031 remains at the non-spreadsheet boundary after the recovery and cleanup.
F-X029 creates the stable
repository-side `ci-gate` in S44. This operational story makes it a required
GitHub check after the Word fidelity and release gates have settled and before
the larger spreadsheet programme starts.

### M21, Presentation depth

#### Sprint S59, Presentation collaboration and security

**Goal**: deepen the package and collaboration surfaces before adding dynamic
rendering behavior.

| F-ID | Title | Size |
|------|-------|------|
| F-217 | Presentation collaboration and navigation model | L |
| F-221 | Presentation encryption and signatures | M |

The sprint adds comments, replies, sections, footer metadata, protected package
handling, and signature policy. Executable payloads are preserved and
inspectable, never run. Binary `.ppt` remains out of scope.

#### Sprint S60, Animation and transitions

**Goal**: turn preserved timing XML into deterministic, bounded behavior.

| F-ID | Title | Size |
|------|-------|------|
| F-213 | Animation and transition timing model | L |
| F-214 | Timeline evaluation and transition rendering | L |

Static rendering remains unchanged. The new timeline path evaluates supported
entrance, exit, emphasis, motion, transition, and morph behavior at explicit
timestamps, with unsupported extensions preserved and diagnosed.

#### Sprint S61, Audio, video and animated export

**Goal**: complete the media timeline from package relationships to output.

| F-ID | Title | Size |
|------|-------|------|
| F-215 | Audio and video package model | L |
| F-216 | Media poster and playback rendering | M |
| F-227 | Animated GIF and video export | L |

Codec handling is bounded to reviewed backends. Unsupported media remains
extractable and visible through poster frames and diagnostics rather than being
dropped.

#### Sprint S62, SmartArt, embedded content, and reader contributions

**Goal**: model the two major opaque PresentationML surfaces without executing
untrusted content.

| F-ID | Title | Size |
|------|-------|------|
| F-218 | Embedded object and macro inventory | L |
| F-219 | SmartArt typed model | L |
| F-X071 | Integrate PRs 61 through 64 | L |

OLE, ActiveX, and VBA remain inventory, extraction, replacement, removal, and
preservation surfaces. SmartArt gains typed inspection and editing. F-X071
independently adopts the reviewed Word
reader contributions from PRs 61 through 64 and hardens namespace, schema-order,
and effective-style edge cases before integration.

#### Sprint S63, ODP, notes and handouts

**Goal**: broaden modern presentation interchange, complete the presenter and
audience output surfaces, and remove two reporter-confirmed Word editor cache
cliffs.

| F-ID | Title | Size |
|------|-------|------|
| F-220 | SmartArt layout and rendering | L |
| F-222 | ODP read and write | L |
| F-223 | Modern presentation package variants | M |
| F-226 | Notes and handout export | M |
| F-X072 | Keep paragraph caching across note references | M |
| F-X073 | Restart ordinary-prose pagination within the aggregate cache | L |

F-220 completes the carried SmartArt validator boundary before F-222 consumes
SmartArt through ODP interchange. ODP uses a declared LibreOffice differential
boundary. Notes pages and handouts reuse the existing master hierarchy and
shared PDF and image backends. Modern macro, template, and slide-show variants
build on the embedded-content inventory from S62.

F-X072 and F-X073 address issues 65 and 66 sequentially. The first keeps
paragraph-cache reads enabled after note-reference paragraphs. The second lets
ordinary prose publish bounded restart checkpoints and charges them against the
existing aggregate cache budget.

#### Sprint S64, HTML and PDF slide import

**Goal**: turn common modern content sources into editable or explicitly
preserved slide content.

| F-ID | Title | Size |
|------|-------|------|
| F-224 | HTML slide content import | L |
| F-225 | PDF page content import | L |
| F-X074 | Tag rpptx-v0.9.0 | S |
| F-X075 | Preserve restart pagination across page-spanning paragraphs | M |
| F-X076 | Tag v0.12.0 | S |

HTML maps a bounded DOM and CSS subset into shapes. PDF imports either a
preserved page graphic or the declared editable text, image, path, and link
subset. Neither path promises arbitrary browser or PDF-engine compatibility.
F-X074 publishes the completed M21 boundary as the exact 15-package
incubating 0.9.0 family after review and separate final approval.
F-X075 is a late cross-cutting correction for Issue 67. It preserves the
recorded pagination pass across page-spanning ordinary prose while retaining
only complete-block restart checkpoints and every existing unsafe-state
exclusion. Its stable package release follows the presentation release under a
separate reviewed release story. F-X076 publishes that exact stable family at
0.12.0 with the reviewed PR 61 through 64 and Issue 65 through 67 contribution
inventory.

### M22, Word depth

#### Sprint S65, OfficeMath

**Goal**: author, convert, lay out, and render modern Word equations.

| F-ID | Title | Size |
|------|-------|------|
| F-228 | OfficeMath model and authoring | L |
| F-229 | OfficeMath layout and PDF rendering | M |
| F-230 | MathML and LaTeX conversion | M |

The equation model preserves unsupported siblings and remains independent of
legacy Equation Editor objects.

#### Sprint S66, Fields and dynamic tables of contents

**Goal**: update the field-driven structures real reports depend on.

| F-ID | Title | Size |
|------|-------|------|
| F-231 | Extended field evaluation | L |
| F-232 | Dynamic table of contents rebuild | L |

TOC, TC, formula, mail-merge control, and barcode fields retain unsupported
instructions and cached results rather than substituting guessed values.

#### Sprint S67, Advanced automation and comparison

**Goal**: deepen the two flagship document-automation workflows.

| F-ID | Title | Size |
|------|-------|------|
| F-233 | Advanced mail merge | L |
| F-234 | Full-story document comparison | L |
| F-235 | Comparison granularity and ignore policy | M |

Mail merge gains nested regions, multiple data sources, images, fragments, and
format hooks. Comparison expands across related stories and adds explicit word,
character, and ignore policies.

#### Sprint S68, Embedded content, forms and building blocks

**Goal**: expose modern package content that is currently preserved but opaque.

| F-ID | Title | Size |
|------|-------|------|
| F-236 | Embedded object and macro inventory | L |
| F-237 | Forms, glossary, and building blocks | L |

Executable objects and macros remain non-executing inventory surfaces. Legacy
form fields are modeled only where they occur inside modern OOXML packages.

#### Sprint S69, Modern Word package and web variants

**Goal**: close Word depth without opening a legacy `.doc` programme.

| F-ID | Title | Size |
|------|-------|------|
| F-238 | Flat OPC and modern Word package variants | M |
| F-239 | MHTML import and export | M |
| F-X077 | Share strict XML lexical validation | M |
| F-X080 | Restore CI release readiness | S |
| F-X079 | Tag rpptx-v0.10.0 | S |
| F-X078 | Tag v0.13.0 | S |

Flat OPC, DOCM, DOTX, DOTM, and bounded MHTML share the current document model.
Binary `.doc` and Word 2003 XML remain permanent non-goals. The shared strict
XML lexical validator removes the three copies exposed by S68 without changing
their fail-closed owner contracts. F-X080 restores the hosted package, Pandoc,
and Python binding gates before either release begins. F-X079 publishes the new
shared API as the incubating family at rpptx-v0.10.0 before F-238 consumes it
through the stable graph. The M22 end gate then runs only modern package
fixtures. F-X078 runs last and prepares and publishes the exact seven-package
stable family at v0.13.0 only after that gate is clean. Each publication uses
`/release` and its own separate final approval at the reviewed SHA.

### M19, Advanced spreadsheets

#### Sprint S70, Spreadsheet decision, corpus and core model

**Goal**: decide whether a material Rust ecosystem gap still exists, then build
the ownership model only if that decision is affirmative.

| F-ID | Title | Size |
|------|-------|------|
| F-184 | Advanced spreadsheet go or no-go | S |
| F-204 | Spreadsheet corpus and compatibility matrix | M |
| F-185 | Workbook and worksheet model | L |

F-184 is a true go or no-go gate. It reassesses Calamine,
`rust_xlsxwriter`, `umya-spreadsheet`, `xls`, and any credible successor at S70,
then classifies each proposed feature as preserved, modeled and editable, or
executable. If the ecosystem provides the complete required lifecycle by then,
M19 is archived rather than implemented.

#### Sprint S71, Styles, tables and structured references

**Goal**: model the indexed formatting and structured data semantics that
ordinary business workbooks rely on.

| F-ID | Title | Size |
|------|-------|------|
| F-186 | Shared strings, styles and number formats | L |
| F-189 | Formula parser | L |
| F-205 | Excel tables and structured references | L |

F-189 lands here because structured table references are part of the formula
grammar rather than a string convention. The sprint does not yet calculate
formulas.

#### Sprint S72, Advanced worksheet objects

**Goal**: cover the visible and interactive worksheet surface before the
streaming package boundary freezes it.

| F-ID | Title | Size |
|------|-------|------|
| F-206 | Advanced worksheet objects | L |

This includes comments, hyperlinks, rich text, drawings, grouping, panes,
sparklines, page breaks, and modern image cells. External content stays offline
unless an explicit bounded policy allows retrieval.

#### Sprint S73, Streaming read and write

**Goal**: prove that advanced workbooks remain bounded at the package boundary.

| F-ID | Title | Size |
|------|-------|------|
| F-187 | Reader | L |
| F-188 | Writer | L |

Both carry an asserted memory ceiling rather than a hoped-for one. A 100 MB
fixture is the gate, not a smoke test. Unsupported package parts and
relationships remain attached through unrelated typed edits.

#### Sprint S74, Calculation and sheet features

**Goal**: calculate ordinary and modern formulas, then expose the features that
depend on their results.

| F-ID | Title | Size |
|------|-------|------|
| F-190 | Calculation engine | L |
| F-191 | Charts in spreadsheets | M |
| F-192 | Conditional formatting and data validation | M |

F-190 is differential against the values Excel stored in the pinned corpus,
including dynamic arrays, spill ranges, and structured references. Unsupported
functions retain cached values with diagnostics. F-191 reuses `oxml-chart`
rather than creating a spreadsheet-only chart engine.

#### Sprint S75, Pivots and the Data Model boundary

**Goal**: move PivotTables from opaque preservation to typed local refresh and
define the boundary around proprietary analytical models.

| F-ID | Title | Size |
|------|-------|------|
| F-193 | Pivot cache and table model | L |
| F-207 | Pivot recalculation engine | L |
| F-208 | Slicers, pivot charts, and Data Model boundary | L |

Worksheet and table-backed pivots refresh locally and regenerate their cache
and visible cells. Slicers and pivot charts follow that refresh. OLAP, Power
Pivot, VertiPaq, and DAX state is preserved and inspectable but is not executed
under this milestone.

#### Sprint S76, Power Query language and package model

**Goal**: understand and evaluate M independently of external data access.

| F-ID | Title | Size |
|------|-------|------|
| F-209 | Power Query package and M language | L |

The package model preserves queries, connections, load destinations, and
refresh metadata. The language boundary covers the pure transformations needed
by the corpus before credentials, connectors, or network policy enter the
runtime.

#### Sprint S77, Power Query execution

**Goal**: refresh a bounded, useful connector set without weakening privacy or
offline determinism.

| F-ID | Title | Size |
|------|-------|------|
| F-210 | Power Query execution and connectors | L |

The first allowlist covers workbook tables, CSV, JSON, and HTTP. Credentials,
privacy levels, source combination, timeouts, byte limits, caching, and query
folding are explicit contracts. Proprietary and tenant-bound connectors remain
preserved with diagnostics.

#### Sprint S78, Office Scripts-compatible automation

**Goal**: automate the same workbook model through a versioned and sandboxed
TypeScript surface.

| F-ID | Title | Size |
|------|-------|------|
| F-211 | Office Scripts artifacts and ExcelScript surface | L |
| F-212 | Sandboxed Office Scripts runtime | L |

Scripts remain external artifacts associated with workbooks. The runtime does
not pretend to provide OneDrive, SharePoint, Power Automate, or Microsoft tenant
identity. It does provide bounded workbook, range, table, chart, pivot, and
query automation with atomic failure.

#### Sprint S79, Rendering, distribution and advanced end gate

**Goal**: close M19 as a headless advanced spreadsheet engine rather than a
file-format crate.

| F-ID | Title | Size |
|------|-------|------|
| F-194 | Sheet rendering | L |
| F-195 | rxlsx distribution | L |

The end gate runs one representative advanced workbook through read, edit,
formula and pivot recalculation, selected Power Query refresh, sandboxed script
automation, save, reopen, and deterministic PDF rendering. Every unavailable
execution surface remains preserved and diagnosed. Distribution follows the
shape M13 established only after that complete lifecycle passes.

## Future release boundaries

These are the planned points where a coherent crate family may publish. They
do not authorize a tag or publication by themselves. The sprint that reaches a
boundary adds an explicit release F-ID, selects the exact version from the
reviewed public API diff, and uses `/release` with its separate final approval
at the fully verified SHA.

| Boundary | Eligible publication | Why this point is coherent |
|---|---|---|
| End of S58, M20 | Stable `rdocx` family. Publish the incubating shared family first only if a shared crate version or stable dependency pin moved. | The Word corpus, shaping, pagination, and incremental layout gates have all settled, so the stable family can describe one measured fidelity boundary. |
| End of S64, M21 | Incubating `rpptx` family. Publish the stable family too only when the reviewed dependency diff requires new shared pins. | Collaboration, security, timing, media, SmartArt, ODP, handouts, HTML import, and PDF import form one complete presentation-depth boundary. |
| End of S69, M22 | Stable `rdocx` family at v0.13.0 through F-X078. Publish the incubating shared family first only if a shared crate version or stable dependency pin moved. | OfficeMath, fields, dynamic TOC, automation, comparison, embedded content, and modern package variants complete the planned Word-depth boundary. |
| End of S79, conditional M19 | The new `rxlsx` distribution family defined by F-195, plus only the existing families whose reviewed dependency pins moved. | F-195 is the first point where the conditional spreadsheet programme has a complete facade, CLI, WASM, Python, rendering, and advanced lifecycle gate. |

No intermediate sprint publishes merely because one subsystem compiles. A
security fix may still justify a separately planned patch release, but ordinary
feature work waits for the next boundary above. If F-184 archives M19, the S79
boundary disappears with the programme and no spreadsheet release namespace is
created.

## Cross-cutting

F-X001 through F-X004 are scheduled in S36 as the final cross-cutting v1
hardening wave. F-X007 and F-X008 handle the external Word contribution and
its stable-family release without rewriting the completed milestone history.
F-X013 through F-X016 carry the surviving half of the external PR 2 rendering
contribution, whose anchored-drawing placement was overtaken by the M7 anchor
work before it could land.
F-X031 carries the external branch-protection mutation deferred from F-X029 to
the Word fidelity boundary before the two depth milestones begin.

## Future ideas

These are discovery candidates, not scheduled work. They have no F-IDs, sizes,
dependency promises or delivery order. Each candidate needs a scope decision,
design plan and acceptance boundary before it can enter the backlog.

### Spreadsheet breadth

- Extend the advanced spreadsheet milestone with CSV, TSV, JSON and ODS output
  and broader structured-data import. Legacy binary XLS remains permanently
  excluded. XLSB requires its own future demand and architecture decision.
- Add worksheet copy and move operations, advanced paste options, named-range
  mutation and structured data import and export.
- Treat VBA and XLM as compatibility surfaces. Preserve, inventory, extract and
  optionally remove their projects and signatures without executing them.
- Consider DAX and VertiPaq execution only after the M19 Data Model preservation
  boundary has a licensed representative corpus and a separate security and
  performance decision.
- Add OLAP pivot refresh, proprietary Power Query connectors, custom connector
  SDK support, and broader query folding only against named consumers and
  reproducible service fixtures.

### Modern spreadsheet extensibility

- Preserve and inspect web-extension, task-pane and content-add-in parts,
  manifests, permissions and external resource locations. HTML, CSS and
  JavaScript assets should remain external web application content.
- Preserve namespaced custom-function formulas, cached results and add-in
  associations. The calculation engine should report unavailable custom
  functions instead of replacing their cached values.
- Preserve Python cell formulas, source, result previews and service metadata.
  Keep cloud execution outside the core library and expose the dependency as a
  diagnostic.
- Broaden the Office Scripts-compatible API only through versioned conformance
  fixtures. Microsoft-hosted script storage, tenant identity, Power Automate,
  and administrative policy remain integrations rather than local runtime
  claims.
- Add trusted custom-function hosts only after a separate capability and
  isolation design. Until then, retain namespaced formulas and cached results
  and report the unavailable function.

### Conversion and output

- Define a format-priority decision that weighs corpus prevalence, procurement
  needs, implementation risk and maintenance cost before adding another reader
  or writer.
- Consider XPS and PCL only when named users or a representative corpus justify
  them. Legacy binary Office formats remain excluded.
- Add a platform-neutral print-layout contract before considering operating
  system printer integration. PDF and image output should remain the portable
  default.
- Add per-page and per-shape rendering with explicit size, resolution, quality,
  compression, transparency and page-range controls across document families.

### Product surfaces

- Explore a self-hostable HTTP service for conversion, rendering, inspection
  and validation while keeping every operation available through local crates.
- Consider optional adapters for document summarisation, translation, grammar
  checking and presentation localisation. External model providers must remain
  outside the core document model.
- Add a browser viewer and lightweight editing surface only if the WASM APIs can
  remain the single implementation rather than creating a second document
  engine.

### Guardrails

Future breadth must preserve the properties that distinguish this workspace:

- Pure Rust implementations with no Office runtime or hidden conversion
  service.
- First-class native, CLI, Python and WASM entry points where the feature is
  technically available.
- Deterministic bundled-font rendering and declared output-baseline changes.
- Verbatim preservation of unmodelled XML and schema-correct child order.
- Shared OPC, DrawingML, chart, layout and rendering infrastructure rather than
  one implementation per document family.
- Offline operation by default, with network access isolated behind explicit
  adapters.
