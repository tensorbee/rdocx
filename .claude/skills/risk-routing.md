---
description: Route a design plan to the mandatory reading and the extra verification riders its diff earns. Matched rows add obligations and never remove the normal test plan.
---

# Skill: risk-routing

Some diffs in this workspace are ordinary and some can break output on a
machine you do not own without failing to compile. This router names the second
kind.

Use it during `/design`, during `/run-sprint` wave planning, and when building
the consolidated sprint gate. **Load only the rows the proposed diff triggers.**
A matched row adds reading and adds checks. It never replaces the story's test
plan or lowers the `/verify` floor.

## The table

| Trigger | Read before editing | Mandatory concern |
|---|---|---|
| Unit conversion, `Twips`, `Emu`, `Points`, `Inches` | `docs/hld/01-glossary.md` units, `CLAUDE.md` "Things that are deliberately wrong" | Constructors truncate with `as i64` and that is pinned by tests. Rounding shifts every twip, which shifts layout. Declare the harness delta |
| Theme colour, tint, shade, colour mapping | `docs/hld/05-drawingml-model.md` | `rdocx_oxml::theme::apply_tint_shade` is deliberately naive. Add spec-correct work under a new name in `oxml-drawing`. Do not correct the old function |
| Layout, pagination, line breaking, text shaping | `docs/hld/08-rendering-spec.md` | Deterministic font mode for every baseline. A baseline recorded against system fonts is worthless. Re-record deliberately, never incidentally |
| Any parser or serialiser | `docs/hld/04-opc-and-packaging.md`, `06-presentationml-model.md` | `xsd:sequence` child order on write. Prefix-tolerant on read, fixed prefix on write. Round-trip test proving `capture_element` preserved the unmodelled subtree byte for byte |
| Crate dependency graph, a new `use` across families | `docs/hld/03-architecture.md` | `oxml-*` must not depend on `rdocx-*` or `rpptx-*`. The single documented exception is `oxml-drawing -> rdocx-oxml` for the `Theme` adapter |
| Bundled fonts, `crates/rpptx/assets/` | `docs/hld/15-build-and-toolchain.md` | Every font family needs its licence file, and the stated licence must be the real one. A bundled asset outside the crate directory is absent from the published tarball |
| Public API of a published crate | `docs/hld/10-bindings-spec.md`, `CLAUDE.md` structural rules | Semver impact stated. `cargo publish --dry-run` and the `.crate` size assertion. No surface no story asked for |
| WASM or PyO3 bindings | `docs/hld/10-bindings-spec.md` | `cargo check --target wasm32-unknown-unknown`. Workspace test runs need `--exclude rdocx-py --exclude rpptx-py`, because `pyo3/extension-module` makes a test binary fail to link |
| A new feature flag or a change to `default` | `CLAUDE.md` structural rules | A named consumer that exists today. `cargo test -p oxml-layout --no-default-features` is the only thing exercising bundled fonts being off |
| A new trait, generic parameter, crate, module or file | `CLAUDE.md` structural rules | Name the second implementer or the second instantiation that exists **today**. A new crate, module or file needs an explicit ask first |
| An external oracle comparison | `.claude/skills/differential-testing.md` | Pin the oracle version and record it. An unpinned oracle turns its upgrade into your regression |
| Release scripting, version strings | `.claude/commands/release.md`, `docs/hld/15-build-and-toolchain.md` | Inspect every manifest, lockfile and README version diff. Require a clean full gate and a separate final approval before tagging |
| A file move or rename with no behaviour change | `.claude/WORKFLOW.md`, the hash harness section | The harness must be byte-identical across the move. Never fold a behaviour change into a file move |

## Recording the result

Every design plan carries a `## Risk routing` section. Write the matched
triggers and the exact extra checks each one adds. If nothing matches, write
`none`. That is a common and valid answer for a story that only adds tests.

Cite the linked reference, do not copy it into the plan. The plan stays about
the decision and the acceptance evidence.

`/run-sprint` takes the union of every matched row across the sprint when it
builds the consolidated gate, so a rider one story earned runs once for the
whole sprint rather than once per worker.

## Why a router instead of running everything

`/verify --full` is the floor and it already runs on every completion. The rows
above are the checks that are either slow, or manual, or impossible to express
as a workspace-wide command, such as opening a corpus deck in PowerPoint. The
router exists so those are chosen deliberately at design time rather than
remembered at review time.

## Related

- `.claude/commands/design.md`, which writes `## Risk routing`.
- `.claude/commands/verify.md`, the floor every story clears regardless.
- `.claude/commands/run-sprint.md`, which unions the matched rows.
