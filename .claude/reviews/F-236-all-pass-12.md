# F-236, all, pass 12

**Reviewed**: Pass-12 uncommitted implementation diff against `dbb5ab1`, excluding the eleven earlier review artifacts, 7 files and 6,155 changed lines, comprising 6,149 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all eleven prior reviews and their closure evidence
**Verdict**: 7 defects, 0 smells, 0 nitpicks

## Defects

### D1, the remaining empty MC attribute lists make valid owners opaque
`crates/rdocx/src/embedded.rs:1794`
`crates/rdocx/src/embedded.rs:1825`
`crates/rdocx/src/embedded.rs:1834`

The pass-11 fix accepts an empty `ProcessContent`, but the other MC list
attributes still require at least one token. MCE permits zero-prefix
`Ignorable` and `MustUnderstand` values and empty `PreserveElements` and
`PreserveAttributes` QName lists. Each is equivalent to omitting that rule.
Here an empty `Ignorable` sets `mc_rules_valid` false, an empty
`MustUnderstand` fails its explicit nonempty check, and either empty preserve
value fails the shared QName-list checker. Otherwise schema-positioned owners
beneath any of those harmless valid attributes disappear from inventory.

### D2, ignorable extension children invalidate a conforming AlternateContent
`crates/rdocx/src/embedded.rs:1698`
`crates/rdocx/src/embedded.rs:1731`

MCE permits elements from an effectively ignorable non-MC namespace as
preceding, intervening, or trailing children of `mc:AlternateContent`. The
scanner classifies every such element as `Other`, then unconditionally marks
the container grammar invalid without consulting its namespace or effective
`Ignorable` set. A conforming container such as one with an ignorable
`x:extension` before a valid `mc:Choice` therefore hides every embedded owner
in that choice instead of keeping the wrapper transparent.

### D3, non-ignorable qualified attributes are accepted on MC elements
`crates/rdocx/src/embedded.rs:1849`
`crates/rdocx/src/embedded.rs:1869`

Qualified attributes on `mc:AlternateContent`, `mc:Choice`, and `mc:Fallback`
must belong either to the MC namespace or to a namespace declared ignorable.
The attribute validator rejects XML-namespace, unqualified, and unknown
MC-namespace attributes, but accepts every other bound namespace without
checking effective `Ignorable`. A choice carrying `x:producer="value"` with a
bound but non-ignorable `x` prefix remains a valid ownership path, so mutation
can act through a non-conforming MC branch instead of failing closed.

### D4, entity references bypass AlternateContent character-content grammar
`crates/rdocx/src/embedded.rs:1640`
`crates/rdocx/src/embedded.rs:1663`

The story reader now validates general references, but returns from that event
arm without applying the surrounding MC content rules. Raw non-whitespace text
directly under `mc:AlternateContent` invalidates the container, while the
equivalent predefined reference `&amp;` does not. The following choice and its
embedded owner remain actionable even though the resolved non-whitespace
character content violates the AlternateContent child grammar. Character
references need to participate in the same whitespace and grammar decision as
ordinary text.

### D5, forbidden XML 1.0 character references are treated as legal
`crates/rdocx/src/embedded.rs:1079`
`crates/rdocx/src/embedded.rs:1084`

The new character-reference check accepts the entire range from U+E000 through
U+10FFFF. XML 1.0 ends that lower range at U+FFFD before resuming at U+10000,
so U+FFFE and U+FFFF are forbidden. A relationship-owning story or ActiveX
properties part containing `&#xFFFE;` therefore passes the new helper and is
trusted by inventory and mutation despite not being well-formed XML.

### D6, grouped WordprocessingCanvas paths use the wrong group child
`crates/rdocx/src/embedded.rs:2189`
`crates/rdocx/src/embedded.rs:2195`

A WordprocessingCanvas can contain a `wpg:wgp`, which can then contain nested
`wpg:grpSp` and `wps:wsp` shapes. The state machine accepts `wpg:wgp` only
directly below group graphic data, not below a canvas. It instead accepts
`wpg:grpSp` directly below the canvas even though that is not the canvas child
defined by the schema. Thus the valid path
`wpc:wpc/wpg:wgp/wps:wsp/wps:txbx/w:txbxContent` is invisible, while the
schema-invalid direct `wpc:wpc/wpg:grpSp` counterpart can make an owner
actionable. The new canvas regression covers only a direct `wps:wsp` child.

### D7, DrawingML text boxes beneath a Word object remain invisible
`crates/rdocx/src/embedded.rs:1985`
`crates/rdocx/src/embedded.rs:2113`
`crates/rdocx/src/embedded.rs:2119`

The pass-11 remediation makes a run-owned `w:object` a valid start for VML
text-box paths only. `w:drawing` is also a valid child of `w:object`, but the
drawing state still requires that element itself to be an immediate run child.
Consequently a valid
`w:r/w:object/w:drawing/wp:inline/a:graphic/.../wps:txbx/w:txbxContent`
path resets to `Other` at `w:drawing`. Embedded owners in that nested story are
omitted from inventory and cannot be extracted, replaced, or removed.

## Smells

None.

## Nitpicks

None.

## Not found

All 47 findings from passes 1 through 11 are closed for their cited
reproductions. In particular, overlapping owners of either the same or mixed
kinds now fail closed, undeclared general entities are rejected while ordinary
references survive, source and target Pack names are normalized, direct
WordprocessingCanvas text boxes work, VML text boxes beneath `w:object` work,
and an empty `ProcessContent` remains actionable. The earlier graph, signature
MIME and incoming-edge, relationship singleton, root anchoring, owner
cardinality, story MIME, MC vocabulary, raw preservation, and grouped
DrawingML cases also remain closed.

No additional findings were found in signature invalidation or removal,
relationship reachability, failure atomicity, hashing and exact extraction,
panic safety, dependency direction, public API shape, or repository structure.
All 45 focused `word_embedded_` regressions pass with default features and with
all features. `cargo check -p rdocx --all-targets`,
`cargo fmt --all --check`, and `git diff --check dbb5ab1` pass.
