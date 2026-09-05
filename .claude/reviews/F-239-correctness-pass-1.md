# F-239, correctness, pass 1

**Reviewed**: working diff, 4 files, 1,547 insertions and 11 deletions
**Verdict**: 4 defects, 1 smell, 0 nitpicks

## Defects

### D1, resource preflight misses CSS and responsive image references

`crates/rdocx/src/html.rs:1330`

The preflight selects only one URL-bearing attribute per listed element. A
`srcset` candidate or a CSS `url()` value can therefore name an external or
unresolved subresource without causing the atomic rejection required by the
design contract.

### D2, MIME relationship declarations and URL grammar are under-validated

`crates/rdocx/src/html.rs:1169`

The parser ignores a conflicting multipart `type` parameter and an unsupported
HTML root charset. The location validator at
`crates/rdocx/src/html.rs:1026` also accepts embedded ASCII whitespace. These
inputs should fail before projection.

### D3, delimiters after the closing multipart boundary are accepted

`crates/rdocx/src/html.rs:901`

The final-boundary check examines only the last recognized delimiter. An
earlier closing delimiter followed by another delimiter can bypass the
epilogue rejection, leaving ambiguous content silently ignored.

### D4, MHTML export cannot report any document-side loss

`crates/rdocx/src/html.rs:547`

The writer always returns an empty diagnostic vector even when the HTML emitter
drops body content controls or preserved body XML. This contradicts the stable
ordered diagnostic contract and can hide an unsupported fact beside supported
siblings.

## Smells

### S1, the differential mutation matrix omits image and diagnostic dimensions

`crates/rdocx/tests/integration_test.rs:145`

The record includes an image count but no image-bearing mutation is supplied,
and diagnostics are absent from the record. Two acceptance dimensions named by
the approved test plan are therefore not proved mutation-sensitive.

## Nitpicks

None.

## Not found

No additional contract, panic-safety, OOXML-order, or structural-rule findings.
The implementation adds no crate, module, file, dependency, trait, generic
parameter, feature flag, or forwarding wrapper. Default HTML output remains
covered as byte-identical, and the generated DOCX is reopened before results
are returned.
