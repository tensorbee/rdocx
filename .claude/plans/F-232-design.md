# F-232, Dynamic table of contents rebuild

**Status**: approved
**Sprint**: S66
**Size**: L
**Depends on**: F-154, F-231

## Problem

Paragraphs already expose style mutation and direct outline levels in
`crates/rdocx/src/paragraph.rs:456` and `crates/rdocx/src/paragraph.rs:850`,
and F-154 provides bookmark ownership. The facade has no operation that reads
an existing `TOC` instruction, discovers the selected heading and `TC`
sources, rebuilds its entries, or substitutes final displayed page numbers.

A safe rebuild must change only the owned TOC result region. Replacing whole
paragraphs or rebuilding every field would lose unrelated run formatting,
field source scaffolding, bookmarks, and unmodelled XML. It also needs a
declared layout mode because displayed page numbers depend on the font input
used to paginate the document.

## Spec reference

- `docs/hld/03-architecture.md`, the field grammar, bookmark ownership, pure
  evaluation, and cache update paragraphs beginning with "The Word text model
  also projects bookmark starts".
- `docs/hld/04-opc-and-packaging.md`, the package mutation and fail-closed
  relationship identity rules in the mail-merge section.
- `docs/hld/08-rendering-spec.md`, "Word bookmark field pagination" and "The
  renderer's input".
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability".
- `docs/hld/12-testing-strategy.md`, the Word field regression matrix and Word
  render fidelity sections.
- `docs/hld/14-development-backlog.md`, "F-232, Dynamic table of contents rebuild".

## Approach

Add one native `Document` operation in the existing `field.rs` owner that
rebuilds each supported existing TOC field in document order. Parse its
instruction through the F-231 boundary to select built-in heading levels,
custom style mappings, direct outline levels, `TC` entries, hyperlink and page
number behavior, and any approved switches. Do not create a TOC when none
exists.

Discover eligible body paragraphs without changing their content. Allocate
collision-free hidden bookmarks for targets that need one, preserving valid
existing bookmarks and references. Build TOC entries with level-appropriate
styles, tab leaders, internal hyperlinks, and `PAGEREF` fields. Replace only
the owned cached result range of the existing TOC, preserving its instruction,
unrelated field formatting, surrounding raw XML, section structure, and
package parts.

Stage the candidate document and package. Paginate it with deterministic
bundled fonts, substitute final page targets through the existing
post-pagination field mechanism, serialize, reopen, and commit only after
every heading, style, TC source, bookmark, relationship, and generated field
validates. Invalidate layout caches once after commit.

Expose `Document::rebuild_toc()` and return a compact native report containing
entry, bookmark, and diagnostic counts. Python, WASM, and CLI remain unchanged.
No new module, trait, generic parameter, feature flag, or crate is introduced.

## Rejected alternatives

- Generate a new TOC at a guessed location. The story requires rebuilding an
  existing TOC and preserving its surrounding content.
- Replace complete TOC paragraphs from strings. That loses producer XML and
  run formatting outside the owned result range.
- Reuse stored page numbers. Heading and style mutations require page targets
  from the final candidate layout.
- Add a second pagination algorithm to the facade. The existing layout and
  post-pagination substitution boundary already owns displayed page numbers.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| differential | `dynamic_toc_rebuild_matches_the_pinned_word_update` | Source-built heading, custom-style, outline-level, and TC mutations produce the same ordered entries, levels, links, and displayed page numbers as Microsoft Word 16.104 build 16.104.25121423. |
| integration | `toc_rebuild_uses_final_deterministic_page_targets` | A mutation that changes pagination yields page numbers from the rebuilt candidate, including entries sourced inside tables and content controls where supported. |
| round-trip | `toc_rebuild_preserves_unowned_field_and_package_xml` | Save and reopen preserves the TOC instruction, unrelated field formatting, neighbouring raw XML, relationships, and untouched package parts byte for byte at their source boundaries. |
| regression | `toc_rebuild_rejects_ambiguous_or_malformed_sources_atomically` | Duplicate or invalid bookmarks, malformed TOC ownership, unsupported switches, and layout or serialization failure leave the live document and caches unchanged with stable diagnostics. |

The **test gate** is differential. Heading, style, and TC mutations produce the
same entries, links, levels, and page numbers as the pinned Word update.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/04-opc-and-packaging.md`
- `docs/hld/08-rendering-spec.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`

## Risk routing

- Layout, pagination, line breaking, or text shaping. Read
  `docs/hld/08-rendering-spec.md`. Use deterministic bundled-font mode for
  every baseline and re-record no baseline incidentally.
- Parser or serializer. Read `docs/hld/04-opc-and-packaging.md` and the
  WordprocessingML preservation parts of `docs/hld/06-presentationml-model.md`.
  Prove schema order, prefix-tolerant reads, fixed-prefix writes, and byte
  preservation of unmodelled subtrees in a round-trip test.
- Public API of a published crate. Read `docs/hld/10-bindings-spec.md` and the
  `CLAUDE.md` structural rules. State semver impact, run
  `cargo publish --dry-run -p rdocx`, and assert the generated `.crate` remains
  below 10 MiB.
- External oracle comparison. Read
  `.claude/skills/differential-testing.md`. Pin Microsoft Word 16.104 build
  16.104.25121423 and the locale in source metadata, keep the oracle out of
  published dependencies, and use only source-built fixtures.

## Hash harness

Expected to be unchanged. Rebuild fixtures are source-built and no harness
sample is added or changed. Any delta is unexplained and blocks integration.

## Implementation checklist

- [ ] Add deterministic `Document::rebuild_toc()` and its approved report surface.
- [ ] Discover and validate existing TOC ownership plus approved instruction switches.
- [ ] Discover headings, custom styles, direct outline levels, and TC entries in document order.
- [ ] Reuse or allocate collision-free bookmarks and build styled linked entry content.
- [ ] Replace only the TOC-owned cached result range and preserve unrelated formatting and raw XML.
- [ ] Paginate the staged candidate and substitute final displayed page numbers.
- [ ] Serialize, reopen, commit atomically, and invalidate both layout caches once.
- [ ] Add pinned Word differential, deterministic layout, preservation, and failure regressions to existing test binaries.
- [ ] Update exactly the listed HLD files and run the full risk-rider union.

## Open questions

None. `Document::rebuild_toc()` uses deterministic bundled fonts and returns a
report containing entry, bookmark, and diagnostic counts.
