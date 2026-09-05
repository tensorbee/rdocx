# F-236, all, pass 6

**Reviewed**: Pass-6 uncommitted implementation diff against `dbb5ab1`, 7 files and 3,753 changed lines, comprising 3,747 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all five prior reviews and their closure evidence
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, ActiveX inventory recognizes the wrong Word owner position
`crates/rdocx/src/embedded.rs:1110`, `crates/rdocx/tests/regression_test.rs:15992`

The scanner accepts `w:control` only when it is a direct child of a run. In the
Word schema, `w:control` is a leaf below the run-owned `w:object` or `w:pict`
element, so a conforming ActiveX owner such as
`w:r/w:object/w:control[@r:id]` is omitted from inventory and cannot be
extracted, replaced, or removed. The shared fixture puts `w:control` directly
under `w:r`, which is why every ActiveX regression passes while exercising the
invalid position instead of the promised schema-positioned owner.

### D2, a story-looking descendant can still bypass the source XML root
`crates/rdocx/src/embedded.rs:1428`

The pass-5 nested-root fix selects the first recognized Word story root, but it
does not require that node to be the actual document element. A part rooted at
an unrecognized producer element can contain
`x:wrapper/w:hdr/w:p/w:r/w:object`, and the validator discards `x:wrapper`
before accepting the remaining path. The same logic accepts a recognized story
element after a prior closed root. Inventory and removal can therefore act on a
story-shaped descendant of malformed or non-story XML instead of keeping it
opaque. The new regression closes the cited nested `w:hdr` under a recognized
outer `w:hdr`, but not this root-anchoring bypass.

### D3, nested DrawingML groups still hide valid text-box stories
`crates/rdocx/src/embedded.rs:1524`

The text-box state machine recognizes the top-level `wpg:wgp` below
`a:graphicData`, but no recursive `wpg:grpSp` child. `wpg:grpSp` has the same
word-processing group content model and can contain another `wps:wsp`. A valid
path such as
`a:graphicData/wpg:wgp/wpg:grpSp/wps:wsp/wps:txbx/w:txbxContent` resets to
`Other` at `wpg:grpSp`, so executable owners in that text box disappear from
inventory. The pass-5 regression proves only a shape directly inside the
top-level group.

### D4, compatibility validation still ignores prohibited MC attributes
`crates/rdocx/src/embedded.rs:1107`, `crates/rdocx/src/embedded.rs:1320`, `crates/rdocx/src/embedded.rs:1357`

The compatibility grammar checks child order and the unqualified `Requires`
attribute, but it never validates attributes on `mc:AlternateContent` and
ignores every attribute on `mc:Choice` or `mc:Fallback` except unqualified
`Requires`. ECMA markup compatibility forbids `xml:lang` and `xml:space` on all
three elements and restricts MC-namespace attributes there. An owner below an
otherwise well-formed branch carrying one of those prohibited attributes is
still treated as schema-positioned and removable. The pass-5 tests cover
missing, empty, and unbound `Requires` plus branch ordering, but not the full MC
element grammar required for fail-closed ancestry.

## Smells

None.

## Nitpicks

None.

## Not found

Pass 1 D1 through D5, both pass 2 defects, pass 3 D1, and pass 4 D1 through D3
remain closed for their cited reproductions. ActiveX binary multiplicity,
shared properties, package-signature incoming edges, synchronized signature
invalidation, compatibility-wrapper byte preservation, invalid target modes,
the prior invalid ancestry paths, text-box discovery, and shared VBA targets
have concrete implementation and regression coverage.

Pass 5 D4 is closed. Legacy and Agile VBA signatures, the package signature
origin, and package XML signatures must resolve to their exact MIME types before
inventory or mutation. Pass 5 D1 through D3 are closed for their exact
reproductions, but D2 through D4 above show remaining root, recursive group,
and MC grammar cases within those broader requirements.

No additional findings were found in internal-target normalization,
relationship identity and reachability, package or VBA signature policy,
byte-range removal, staged mutation atomicity, panic safety, public API shape,
dependency direction, or repository structure. All 13 focused
`word_embedded` regressions pass with default features and with all features.
