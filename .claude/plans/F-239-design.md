# F-239, MHTML import and export

**Status**: approved
**Sprint**: S69
**Size**: M
**Depends on**: F-178

## Problem

The native facade exports HTML but has no diagnostic-bearing MHTML byte or path
surface near `crates/rdocx/src/document.rs:5427`. Existing HTML import accepts
only markup, projects images to alternate text near
`crates/rdocx/src/html.rs:1082`, and drops hyperlink targets near
`crates/rdocx/src/html.rs:1122`. It therefore cannot preserve embedded
resources or links through the bounded two-way MHTML contract.

MHTML adds MIME root selection, transfer decoding, Content-ID and
Content-Location resolution, resource identity, and deterministic multipart
writing. Unsafe, malformed, external, unresolved, or over-limit resources
must fail before a partial document or file is published.

## Spec reference

- `docs/hld/02-scope-and-non-goals.md`, bounded modern interchange and
  permanent legacy-format exclusions.
- `docs/hld/03-architecture.md`, "Why these seams" and "Facade conventions".
- `docs/hld/04-opc-and-packaging.md`, "The package", "Media", and "Package
  integrity".
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability", "Packaging",
  and "WASM".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy", "Binding tests", and
  "What CI runs".
- `docs/hld/14-development-backlog.md`, "F-239, MHTML import and export".
- `docs/hld/15-build-and-toolchain.md`, oracle and published-package gates.
- RFC 2045, RFC 2046, RFC 2387, RFC 2392, and RFC 2557 for transfer encoding,
  multipart relationships, `cid:` URLs, and Content-Location resolution.

## Approach

Extend the existing private `crates/rdocx/src/html.rs` owner and expose
additive native pre-1.0 values and methods:

```rust
#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub struct MhtmlDiagnostic {
    pub location: String,
    pub property: Option<String>,
    pub message: String,
}

pub struct MhtmlReadResult {
    pub document: Document,
    pub diagnostics: Vec<MhtmlDiagnostic>,
}

pub struct MhtmlWriteResult {
    pub bytes: Vec<u8>,
    pub diagnostics: Vec<MhtmlDiagnostic>,
}

impl Document {
    pub fn from_mhtml_bytes(bytes: &[u8]) -> Result<MhtmlReadResult>;
    pub fn open_mhtml<P: AsRef<Path>>(path: P) -> Result<MhtmlReadResult>;
    pub fn to_mhtml_bytes(&self) -> Result<MhtmlWriteResult>;
    pub fn save_mhtml<P: AsRef<Path>>(
        &self,
        path: P,
    ) -> Result<Vec<MhtmlDiagnostic>>;
}
```

Add one concrete `Error::Mhtml` variant with optional part identity, byte
offset, and message. Parse a bounded `multipart/related` entity with header
folding, a required boundary, optional `start`, unique Content-ID and normalized
Content-Location keys, and base64, quoted-printable, 7bit, or 8bit transfer
forms. Index every part with checked size and count limits before projection.
Select exactly one HTML root. Resolve its resource references only from exact
`cid:` values or normalized contained locations. Perform no network or ambient
filesystem fetch.

Extend the existing HTML projection internally so the MHTML path can insert
validated image bytes and carry safe hyperlink targets. The normal
`Document::from_html` behavior remains unchanged. An image is accepted only
when the declared MIME type and `oxml-media` sniff agree. External anchor links
remain navigation and are retained after URI validation. External or
unresolved subresource references are input failures, not navigation links,
and cause atomic rejection.

Export through the existing `rdocx-html` emitter with an internal relationship
to `cid:` source mapping. Default HTML output remains byte-identical. Emit
deterministic CRLF headers, stable part order, 76-column base64, source-ordered
referenced images only, and a collision-safe content-derived boundary. Reparse
the MHTML and reopen its projected DOCX before returning it. Path output uses
the existing atomic writer. Stable ordered diagnostics record every supported
loss boundary.

Keep all code in existing modules and reuse current dependencies. Add no crate,
module, file, feature, trait, generic parameter, public intermediate MIME
model, network authority, binding method, or new test binary.

## Rejected alternatives

- Put MIME import in `rdocx-html`. That would create a facade cycle or a second
  document model.
- Add a new MHTML crate or module. Existing HTML ownership is cohesive and
  avoids a new approved file.
- Use data URIs only. That would not exercise the required Content-ID and
  Content-Location semantics.
- Fetch missing resources. Network and ambient filesystem access are outside
  this story.
- Compare container bytes with Word. MIME formatting trivia is not the
  supported semantic contract.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| unit | `mhtml_parser_rejects_ambiguous_unsafe_and_over_limit_resources_before_projection` | Header, boundary, root, identity, transfer, path, resolution, size, count, and diagnostic failures are atomic. |
| unit | `mhtml_transfer_decoding_and_resource_resolution_are_exact` | Header folding, quoted-printable, base64, `cid:`, relative Content-Location, MIME parameters, and root selection are deterministic. |
| unit | `mhtml_writer_is_deterministic_bounded_and_collision_safe` | CRLF, headers, part order, base64 wrapping, boundary selection, referenced-only resources, repeated bytes, and output caps are exact. |
| integration | `mhtml_import_and_export_preserve_supported_word_structure` | Body order, formatting, list identity, table grid and spans, images, links, and ordered diagnostics survive both directions and reopen. |
| regression | `mhtml_loss_records_do_not_hide_supported_siblings` | Every unsupported safe fact produces one stable path-aware loss record while supported siblings survive. |
| differential | `mhtml_conversions_match_the_pinned_word_structure` | Source-built MHTML and DOCX conversions agree with pinned Word structure and reject body, formatting, table, list, image, link, and diagnostic perturbations. |

The **test gate is differential**. Microsoft Word 16.104 build
16.104.25121423 is the pinned import and export oracle already used by the Word
acceptance harness. Inputs remain source-built, normalized semantic trees and
resource bytes are compared instead of MIME or DOCX bytes, and the ignored
regeneration test authenticates the exact oracle identity.

## HLD impact

- `docs/hld/02-scope-and-non-goals.md`
- `docs/hld/03-architecture.md`
- `docs/hld/04-opc-and-packaging.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`
- `docs/hld/14-development-backlog.md`
- `docs/hld/15-build-and-toolchain.md`

## Risk routing

- **Unit conversion**. Preserve truncating CSS pixel and point conversions to
  EMU and add exact 96 DPI dimension assertions.
- **Any parser or serialiser**. Bound MIME parsing and writing, reparse MHTML,
  save and reopen generated DOCX, prove schema order, and keep default HTML
  serialization unchanged.
- **Public API of a published crate**. State additive pre-1.0 impact, run
  rustdoc with warnings denied, run the patched publish dry run, and assert
  `rdocx` and `rdocx-html` archives remain below 10 MiB.
- **External oracle comparison**. Verify Word 16.104 exactly, use common
  source-built inputs, compare normalized semantics, document intentional
  differences, and prove every acceptance dimension is mutation-sensitive.

## Hash harness

Expected unchanged across all 49 entries. Existing samples call default HTML
conversion, not MHTML. Any delta blocks the story.

## Implementation checklist

- [ ] Add the native result, diagnostic, error, re-export, and method surfaces.
- [ ] Implement bounded MIME header, multipart, transfer, and resource indexing
      in the existing HTML module.
- [ ] Add safe contained resource resolution and MHTML-only image and link
      projection.
- [ ] Implement deterministic MHTML writing and the internal emitter resource
      mapping while keeping default HTML bytes unchanged.
- [ ] Reparse MHTML, save and reopen DOCX, and publish only complete results.
- [ ] Add source-built unit, regression, integration, differential, limit, and
      perturbation coverage to existing binaries.
- [ ] Run focused checks, the pinned oracle gate, every routed rider, and full
      verification.

## Open questions

None. The sprint definition requires fail-closed external and unresolved
subresources. Safe absolute hyperlinks are retained navigation rather than
fetched resources. The existing pinned Word 16.104 harness is the independent
oracle because `python-docx` has no MHTML boundary.
