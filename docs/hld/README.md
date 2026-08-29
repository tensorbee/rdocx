# High-level design

The spec set for this repository. Numbered so the reading order is the
learning order, and so a Word merge produces a coherent document.

## Precedence

When two documents disagree, the lower number wins on scope and intent, the
higher number wins on mechanism. `02-scope-and-non-goals.md` decides whether
something is in v1 at all. `03-architecture.md` decides which crate owns it.
Everything from `04` onward decides how it works.

`.claude/WORKFLOW.md` wins over every document here on process questions. This
set describes the product, not the way it is built.

## The set

| # | Document | What it settles |
|---|---|---|
| 00 | [Vision](00-vision.md) | Why rpptx, why one workspace, what "done" means |
| 01 | [Glossary](01-glossary.md) | OOXML vocabulary, units, the placeholder/layout/master triangle |
| 02 | [Scope and non-goals](02-scope-and-non-goals.md) | The v1 feature matrix against python-pptx, and what is deferred |
| 03 | [Architecture](03-architecture.md) | The crate families and the dependency DAG |
| 04 | [OPC and packaging](04-opc-and-packaging.md) | The ZIP container, relationships, content types, part naming, media |
| 05 | [DrawingML model](05-drawingml-model.md) | The `a:` namespace: colour, transforms, geometry, fills, text, theme |
| 06 | [PresentationML model](06-presentationml-model.md) | The `p:` namespace: parts, shape tree, placeholders, passthrough |
| 07 | [Inheritance and resolution](07-inheritance-and-resolution.md) | The chains from slide to layout to master to theme |
| 08 | [Rendering spec](08-rendering-spec.md) | `PositionedElement`, the PDF backend, preset geometry, shape text |
| 09 | [Charts spec](09-charts-spec.md) | ChartML and the embedded workbook |
| 10 | [Bindings spec](10-bindings-spec.md) | PyO3 handles, WASM, the CLIs |
| 11 | [Migration plan](11-migration-plan.md) | Extracting `oxml-*` without breaking rdocx |
| 12 | [Testing strategy](12-testing-strategy.md) | The harnesses, the corpus, the gates |
| 13 | [Risks and open questions](13-risks-and-open-questions.md) | What could go wrong, and what is still undecided |
| 14 | [Development backlog](14-development-backlog.md) | Every story, sized, with a test gate |
| 15 | [Build and toolchain](15-build-and-toolchain.md) | Toolchain pinning, deterministic rendering, CI, packaging |

## Living status

This set is written once and changed rarely. It describes intent. The
execution-time record lives in `docs/sprints/`, and the current state of the
code lives in the code.

A story that contradicts this set is a bug in one of them. Resolve it before
implementing, not after.

<!-- F-X031 disposable docs-only protection proof. -->
