# F-X272 core capture and raw textbox, Root bounded final pass 1

Reviewed the nine-path2599-line incremental patch against accepted row source,
canonical contractc5377908 and its operative core trim contractf60bc864.
Manifest SHA256 is
`303f7e644d42d0b8bbb6d32b04e4a97327a0454c8cf2ddc05961eb445de14e4b`.
Patch SHA256 is
`aecdf8df76c08b9593830bf246ba6880c44fba0ed76e594ed78079690e50333e`.
Citations refer to the frozen final snapshot, not the moving worker.
Verdict: one defect, zero smells, zero nitpicks. This pass changes no product.

## Defects

### D1, raw rewrite copies a generated stretch before finding any match

`crates/rdocx-oxml/src/placeholder.rs:3188`
`crates/rdocx-oxml/src/placeholder.rs:3207`

The nonempty-replacement branch creates an empty String for every stretch and
calls read.append for every Text before matching. Unlike discovery, it does
not move the first uniquely owned parsed Text. append unconditionally reserves
the copied length at `crates/rdocx-oxml/src/document.rs:4772`, and zero allowance
therefore produces a sticky refusal even when no replacement matches.

A concrete supported source is the generated PAGE result already constructed
by `crates/rdocx-oxml/src/document.rs:26042`. Its actual emitted text is1, while
both selected Word allowance and copy demand are zero. Wrap that existing
carrier as a raw textbox using compose_sources at
`crates/rdocx-oxml/src/document.rs:5518`, preserving the Field and generated
roles, then invoke the new literal consumer with a nonmatching token and empty
JSON dependencies. Discovery uses owned_walk_stretches and can move1. Rewrite
instead tries to reserve a new byte before it can discover the zero matches.
The legacy operation accepts this no-op. The approved raw contract forbids a
new conservative refusal and explicitly keeps generated text from becoming
input credit.

This is a source-traced finding, not an executed historical RED claim. The
existing generated-result test exercises only template_replaceable_texts, and
the new raw tests provide real input text that funds this extra copy. Neither
tests this cross-consumer boundary. Require an actual same-carrier no-match
regression against the legacy output, zero added input credit and a separate
real unfunded-copy refusal. Preserve the existing matching walk, validation,
source identities and replacement behavior. Do not fix it with a grant or a
no-validation early return.

## Smells

None found in the bounded delta. The concrete ordinary/template edit enum
serves two actual consumers. No new trait, generic parameter, module, public
type or independent quota was introduced.

## Nitpicks

None.

## Not found

No additional defect found in the core same-loop transport, namespace copy
recipe, recovered-run companion alignment, glossary publication handback or
copied publication lifetime materialization examined here. The core disabled
path preserves the existing substantive capture loop without a transcript
allocation. Actual numeric consumed windows establish copy candidates, with
byte validation and separate Generated gaps for normalized syntax. Unsupported
transport keeps captured bytes and returns no transcript. This supplies no
complete Field classification.

Recovered Started runs share the successful parser output and parallel slot
push with ordinary runs. The legacy recovery projection-boundary distinction
is retained. Real returned paragraph, table and control owners commit positive
publication ranges after successful construction. Copied current-coordinate
ranges materialize while the collector is live, with no retained source buffer
or new input baseline. The changed glossary uses its actual output source map.

## Evidence and exclusions

The packet reports eight core, eight raw, seven affected row and37disjoint
related passing tests,60unique, with nine scoped gates and two metadata checks
exiting0. These observations do not exercise D1. Mechanical authentication is
separate support and does not replace this semantic review.

The retained consumed-root failure is an introduced scoped regression, not an
old baseline defect. A one-slot optional facts vector contained None. Actual
fallback parser handback and mixed ordinary/recovered/empty slot controls
address that omission. Earlier allocator evidence applies to its identified
pre-trim source, not the final source. Compilation, fixture, probe-link and
lint failures stay separate from semantic RED evidence. The actual downstream
paragraph parser resets trim mode, so the earlier alleged raw trim-funding
failure is not established.

Native production migration, repeat and both imported operations, complete
versus unknown Field ownership, inherited deferred raw/model roles, nested
requested Text refresh, current and constructor pre-growth, child integration
and every final integrated rider remain open. This bounded review neither
qualifies the provider nor changes the peer-review cap.
