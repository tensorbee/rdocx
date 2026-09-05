# F-X077, Share strict XML lexical validation

**Status**: completed
**Sprint**: S69
**Size**: M
**Depends on**: F-236, F-237

## Problem

The embedded-content scanner, glossary parser, and package-story scanner each
carry a separate XML 1.0 lexical validation stack at
`crates/rdocx/src/embedded.rs:1080`,
`crates/rdocx-oxml/src/glossary.rs:480`, and
`crates/rdocx/src/field.rs:7951`. S68 sprint review finding S1 records that a
security correction now requires three implementations and three reviews.

The checks cover declarations, literal characters, qualified and processing
instruction names, namespace declarations and bindings, duplicate expanded
attributes, and character and entity references. Their enclosing scanners
have different document roots, schema rules, error variants, and diagnostic
labels, which must remain local and unchanged.

## Spec reference

- `docs/hld/03-architecture.md`, "Three families, one workspace", "The
  dependency rule", "What stays put", and "Crate-level conventions".
- `docs/hld/04-opc-and-packaging.md`, "Package integrity".
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability" and
  "Packaging".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy", the legacy glossary and
  executable-content gates, and "The hash harness".
- `docs/hld/14-development-backlog.md`, "F-X077, Share strict XML lexical
  validation".

## Approach

Put the format-neutral policy in the existing `crates/oxml-core/src/xml.rs`,
the lowest existing crate already used directly by both `rdocx-oxml` and
`rdocx`. Add one concrete error value and one public entry point:

```rust
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum XmlLexicalError {
    InvalidUtf8,
    InvalidDeclaration(String),
    ForbiddenLiteralCharacter,
    InvalidName(String),
    InvalidNamespace(String),
    DuplicateExpandedAttribute,
    InvalidReference(String),
    InvalidProcessingInstruction(String),
    InvalidComment(String),
}

pub fn validate_strict_xml_1_0(
    xml: &[u8],
) -> std::result::Result<(), XmlLexicalError>;
```

The shared validator owns UTF-8 and literal-character checks, declaration
pseudo-attribute grammar and version, QName and Name grammar, namespace
declarations and bindings, duplicate expanded attributes, entity and
character references, comment lexical validity, and processing instruction
targets. It uses one namespace-aware `quick-xml` pass and exposes no parser
model.

Each consumer adds a small local mapping from `XmlLexicalError` to its current
error variant and wording. Embedded scans retain
`Error::InvalidEmbeddedMutation`, package stories retain `Error::Other`, and
glossary parsing retains `OxmlError::InvalidValue`. Root identity, schema
position, declaration placement, doctype prohibition, element-only grammar,
and semantic whitespace remain in the owner-specific pass. The embedded
UTF-8 declaration restriction also remains local where it is stricter than
the generic XML grammar.

Delete the three duplicated helper stacks only after the complete malformed
matrices and exact error assertions pass. Add no crate, module, file,
dependency edge, feature, trait, generic parameter, wrapper, or recovery path.
The new `oxml-core` API is additive before 1.0. F-X079 publishes it at the
required shared-family boundary before a stable package consumes it.

## Rejected alternatives

- Put the helper in `rdocx-oxml`. Format-neutral XML security policy belongs
  in `oxml-core`, and both consumers already depend on it.
- Expose event-by-event validator state. It would enlarge the public API and
  force every owner to coordinate shared parser state.
- Use a trait or generic error callback. There is no second implementation,
  and local explicit error mapping is clearer.
- Move owner roots, schema rules, or permissive recovery into the helper.
  Those contracts are deliberately different and remain local.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| unit | `strict_xml_1_0_validator_rejects_every_shared_lexical_class` | Declaration, literal character, reference, name, namespace, duplicate expanded attribute, comment, and processing instruction branches are independently sensitive. |
| regression | existing glossary malformed XML matrix | Shared rejection retains the exact `OxmlError` surface and glossary placement rules. |
| regression | existing embedded malformed XML matrix | ActiveX and package-reference scans retain exact operation labels, errors, and byte-identical failed mutation rollback. |
| regression | existing package-story malformed XML matrix | Headers, footers, notes, and other stories retain their exact `Error::Other` messages and schema ownership. |
| round-trip | existing glossary and package-story preservation cases | Unmodelled subtrees still pass through `capture_element` byte-for-byte after valid input. |

The **test gate is regression**. The existing embedded, glossary, and
package-story malformed XML matrices all execute through one shared lexical
validator, retain their current error surfaces and byte-identical rollback,
and collectively make every shared lexical branch mutation-sensitive.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/12-testing-strategy.md`

## Risk routing

- **Any parser or serialiser**. Preserve owner schema rules, prefix-tolerant
  reads, fixed-prefix writes, and byte-exact `capture_element` round trips.
- **Crate dependency graph and cross-family use**. Keep the helper in
  `oxml-core`, add no manifest edge, and run the dependency-direction
  regression proving no shared crate depends on a format crate.
- **Public API of a published crate**. State the additive pre-1.0 impact, run
  rustdoc with warnings denied, run the exact patched workspace publish dry
  run, and enforce every archive size limit.

## Hash harness

Expected unchanged. This refactors rejection policy without changing valid
serialization or rendering. Any output delta blocks the story.

## Implementation checklist

- [x] Add the concrete shared error and validator to the existing
      `oxml-core/src/xml.rs` file.
- [x] Add one mutation-sensitive shared unit matrix covering every lexical
      branch.
- [x] Map shared failures back to the exact glossary error surface.
- [x] Map shared failures back to the exact embedded scanner operation and
      error surface.
- [x] Map shared failures back to the exact package-story error surface.
- [x] Remove all three duplicated helper stacks while retaining local roots,
      schema positions, doctypes, declaration placement, and whitespace rules.
- [x] Prove byte-identical failed mutation rollback and valid raw subtree
      preservation.
- [x] Run focused crate checks, every routed rider, and full verification.

## Open questions

None. `oxml-core` is the required lowest shared owner. The resulting published
API dependency requires F-X079 and its separate release approval before the
stable family can consume it.
