# F-236, all, pass 7

**Reviewed**: Pass-7 uncommitted implementation diff against `dbb5ab1`, 7 files and 4,175 changed lines, comprising 4,169 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all six prior reviews and their closure evidence
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, non-whitespace outside the story root does not invalidate owner discovery
`crates/rdocx/src/embedded.rs:1086`, `crates/rdocx/src/embedded.rs:1295`

The scanner decides that the first element event is the document element even
when a preceding text event contained non-whitespace, and its fallback event
arm also ignores non-whitespace after that element closes. A source such as
`junk<w:hdr>...<w:object>...</w:object>...</w:hdr>` therefore treats the
header as a valid document root and inventories or removes its owner. This is
not a single-root XML story and must stay opaque under the root-anchoring and
fail-closed contracts.

### D2, duplicate OLE children with the same relationship id are collapsed
`crates/rdocx/src/embedded.rs:1220`

The owner finalization sorts and deduplicates the collected relationship ids
before checking their count. A `w:object` containing two `o:OLEObject`
children that both use the same `r:id` is consequently accepted as one owner.
Removal then deletes the complete object instead of rejecting the malformed
duplicate-owner shape. Relationship identity equality does not make two
schema children unambiguous.

### D3, a story part without a relationship set is never checked for broken owners
`crates/rdocx/src/embedded.rs:289`

Owner discovery iterates `package.part_rels`, not the package's XML parts. A
header, footer, note, comment, or glossary story containing a schema-positioned
`w:object` or `w:control` whose relationship set is entirely absent is never
scanned. Inventory therefore succeeds while silently omitting the malformed
executable owner instead of rejecting its missing relationship before any
mutation.

### D4, compatibility-rule attributes are accepted without validating their values
`crates/rdocx/src/embedded.rs:1407`

The MC attribute check accepts `Ignorable`, `MustUnderstand`,
`ProcessContent`, `PreserveElements`, and `PreserveAttributes` solely by local
name. It never validates the required namespace-prefix or QName lists against
the in-scope bindings. For example, an otherwise ordinary wrapper carrying
`mc:Ignorable="unbound"` is treated as valid ancestry, so an embedded owner
beneath malformed compatibility markup becomes actionable rather than staying
opaque.

## Smells

None.

## Nitpicks

None.

## Not found

Pass 6 D1 through D4 are closed for their cited reproductions. ActiveX owners
must now be run-owned `w:object` or `w:pict` elements, recognized story roots
must be the first XML element, nested `wpg:grpSp` paths retain text-box story
ownership, and prohibited XML or unknown MC attributes make compatibility
ancestry opaque. All 17 focused `word_embedded` regressions pass.

All findings from passes 1 through 5 remain closed for their cited
reproductions. No additional findings were found in target normalization,
shared-target reachability, package or VBA signature cleanup and content-type
validation, mutation atomicity, panic safety, additive public API shape,
dependency direction, or repository structure. `cargo check -p rdocx
--all-targets`, `cargo fmt --all --check`, and `git diff --check dbb5ab1` pass.
