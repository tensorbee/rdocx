# F-232, correctness, pass 9

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 13 files and 4,493 changed lines, with 4,383 insertions and 110 deletions. All 35 focused `toc_` regression tests, all 242 `rdocx-layout` unit tests, its doc test, all 368 `rdocx-oxml` unit tests, its doc test, `cargo check -p rdocx --all-targets`, `cargo fmt --all --check`, and `git diff --check` pass.
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, comments and processing instructions still desynchronize owner raw slots
`crates/rdocx/src/field.rs:1046`
`crates/rdocx-oxml/src/text.rs:2420`

The ownership scanner advances `paragraph_raw_before` only while handling
element start and empty events. Its fallback ignores comments and processing
instructions. The typed paragraph parser retains both as non-whitespace raw
children, so they contribute to `raw_xml_count_at` and shift the raw slot of a
following content control or revision. For example, put a comment before an
accepted insertion that contains a selected TC field, then the TOC begin,
instruction, and separator. The scanner records the owner at `Raw(0)`, while
source discovery records it at `Raw(1)`. The pre-separator TC then compares
after the synthetic result start and is incorrectly omitted. Pass-8 D1 is
therefore still open for non-element preserved raw children.

### D2, a malformed direct simple field still advances the scanner run boundary
`crates/rdocx/src/field.rs:1331`
`crates/rdocx-oxml/src/text.rs:1167`

The scan assigns a modeled run position to every direct Word `fldSimple`
without checking whether the typed parser can project it. The typed parser
retains a simple field without `w:instr`, or with an empty parsed instruction,
as raw XML instead. Place that malformed field before a same-paragraph complex
TOC and a selected TC immediately after the separator. The scanner shifts the
result start by one run, while source discovery does not, so the TC appears
outside the old result and becomes a spurious rebuilt entry. The raw hyperlink
case from pass 8 is fixed, but the direct malformed shape named by that finding
is not.

### D3, a hyperlink revision immediately before a direct end marker shares its position
`crates/rdocx/src/field.rs:1325`
`crates/rdocx/src/field.rs:1496`
`crates/rdocx/src/field.rs:2492`

For a modeled hyperlink with direct runs, the scanner classifies a nonterminal
revision as `AfterRaw`. A direct run following that revision receives the same
run boundary, `AfterRaw`, and nested order zero. The accepted projection later
assigns one total nested order to both the revision runs and the direct run, but
the span's scanned end position does not use that order. Put a selected TC in
an accepted hyperlink revision immediately before a direct run containing the
TOC end marker. The TC and end marker compare equal instead of the TC comparing
before the end, so the old-result TC is retained as a new source even though
its XML is removed by result replacement.

### D4, boundary-paragraph heading titles include text from the old result
`crates/rdocx/src/field.rs:2355`
`crates/rdocx/src/field.rs:2436`

Only paragraphs strictly between the begin and end paragraphs are skipped.
Heading discovery on either boundary paragraph then concatenates every
accepted run without filtering runs through `toc_source_position_is_owned`.
For example, give the end paragraph a selected heading style, put an old cached
result run before the TOC end marker, and put the real heading text after it.
The paragraph is a valid post-end source, but its generated title contains both
the deleted old result and the real heading. Bookmark insertion correctly
starts after the end run, which leaves the displayed title and hyperlink target
covering different text ranges.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-8 D1 element and terminal-hyperlink triggers: modeled owners after a
  preserved raw element use the typed raw slot, and a post-separator source in
  a terminal hyperlink revision is excluded. D1 and D3 cover adjacent
  coordinate shapes.
- Pass-8 D2 hyperlink trigger: a simple field retained as hyperlink raw XML no
  longer advances the modeled run boundary. D2 covers a direct malformed
  simple field.
- Pass-8 D3: a direct control-owned run with a run-property revision now marks
  its tracked paragraph and renders one change bar. Both control and revision
  nesting orders retain the same tracked projection behavior.
- Contract and public surface: the additive native rebuild operation and report
  remain within the approved plan. Python, WASM, and CLI surfaces are unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion, recursion,
  or arithmetic panic was found.
- OOXML generation and preservation: no fresh expanded-name, generated-child
  order, wrapper-balance, fixed-prefix, whitespace, unmodelled-subtree, or
  package-byte preservation defect was found beyond the coordinate mismatches
  in D1 through D3 and heading filtering in D4.
- Facade and layout projection parity: accepted nested controls, accepted
  insertion and move-destination text, positioned tables, final page targets,
  and tracked visibility remain aligned outside the cited boundary cases.
- Diagnostics: supported and unsupported complex TOCs plus direct and accepted
  revision simple TOCs retain stable counts. No new diagnostic omission was
  found.
- Test gate: the pinned differential metadata and exact entry, link, level, and
  distinct-page assertions remain intact. Missing mutation-sensitive shapes
  are identified in D1 through D4.
- Bookmark scope, stable whole-paragraph reuse, collision-safe owned
  substitution, nested TOC rejection, maximum bookmark-id allocation, checked
  outline conversion, case-insensitive sequence identity, and atomic staged
  commit remain correct.
- Structure and dependencies: no unjustified trait, generic, forwarding
  wrapper, module, feature flag, crate, dependency, or published binding
  surface was introduced.
