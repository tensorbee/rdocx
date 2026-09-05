# F-236, all, pass 5

**Reviewed**: Pass-5 uncommitted implementation diff against `dbb5ab1`, 7 files and 3,253 changed lines, comprising 3,247 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all four prior reviews and their closure evidence
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, a nested story root can reset and bypass invalid outer ancestry
`crates/rdocx/src/embedded.rs:1244`

`valid_story_owner_path` starts at the last node whose local name resembles a
Word story root and discards every earlier ancestor. A path such as
`w:hdr/w:pPr/w:hdr/w:p/w:r/w:object` therefore passes from the inner `w:hdr`,
even though that element is itself inside an unsupported `w:pPr` subtree. The
scanner can inventory and remove that object instead of keeping the malformed
subtree opaque. This is another bypass of the complete root-to-owner condition
from pass 3.

### D2, compatibility validation still accepts incomplete branch grammar
`crates/rdocx/src/embedded.rs:1073`, `crates/rdocx/src/embedded.rs:1451`

The scanner collapses both `mc:Choice` and `mc:Fallback` into one `Branch`
state, then checks only that a branch immediately follows
`mc:AlternateContent`. It consequently accepts an alternate-content container
whose first and only branch is `mc:Fallback`, and it accepts an `mc:Choice`
without its required `Requires` attribute. Both are malformed MC ancestry, but
an embedded owner below either path is still actionable. Pass 4 D1 required a
valid branch structure, and the new regression covers standalone and nested
branches but not these accepted malformed forms.

### D3, grouped DrawingML text-box stories remain invisible
`crates/rdocx/src/embedded.rs:1339`

The modern text-box state machine recognizes `wps:wsp` only when its immediate
parent state is `a:graphicData`. A valid grouped Word drawing instead uses
`a:graphicData/wpg:wgp/wps:wsp/wps:txbx/w:txbxContent`. The unrecognized
`wpg:wgp` resets the path, so an OLE object or control inside that text box is
omitted from inventory and cannot be mutated or removed. The pass-4 regression
exercises only the legacy VML text-box form.

### D4, signature graph validation ignores required content types
`crates/rdocx/src/embedded.rs:1573`, `crates/rdocx/src/embedded.rs:1728`

The VBA and package signature validators prove relationship type and target
part existence, but never prove that the project signature, signature origin,
and XML signature targets resolve to their required content types. A
signature-typed relationship targeting a part with a missing or unrelated
content type is therefore reported as present, and
`RemoveInvalidatedSignatures` treats it as validated signature infrastructure
and deletes it. The approved malformed-graph contract requires missing or
wrong signature metadata to fail closed before mutation.

## Smells

None.

## Nitpicks

None.

## Not found

Pass 4 D3 is closed. VBA removal checks for another incoming relationship to
the selected project before either signature policy mutates the staged graph,
and the named regression covers both policies with byte-for-byte atomicity.

The pass 1 ActiveX multiplicity and shared-properties defects, package-signature
incoming-edge defect, synchronized signature invalidation defect, and raw MC
preservation gap remain closed. Both pass 2 defects and the exact pass 3
invalid story-path reproducer also remain closed.

No additional findings were found in target normalization, relationship-id
identity, byte-range removal, staged commit atomicity, panic safety, public API
shape, dependency direction, or repository structure. The 9 focused
`word_embedded` regressions pass both with default features and all features,
and `cargo check -p rdocx --all-targets` passes.
