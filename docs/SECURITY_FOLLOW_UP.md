# Deferred security checks and methods

The deeper security assessment below is deferred at the user's request on
5 November 2026. Continue ordinary implementation, regression tests and release
validation. Keep existing guards enabled. These checks are future work, and
their outcomes must not be described as passed.

The completed copy-reservation and BOM fixes remain in the sprint. The items
below distinguish wider assurance work from those specific tested repairs.

## Remaining checks

| Check | Method | Evidence needed to close it |
|---|---|---|
| End-to-end copy accounting | Map every actual allocation, copy and move from checked ingress through model preparation, serialization, namespace replay, parsing and publication. Include separately owned legacy and Source buffers. Exercise repeated use of the same input and two output occurrences sharing one origin. | Every applicable copy reserves the existing budget before growth. A move incurs no copy charge. No double debit, renewed input allowance or unpaid materialization. |
| Remaining Native producer handoffs | Trace namespace-changing replay, prepared second writers, bookmark/TOC/drawing canonicalization, refreshes, imports and related-story writers. Retain the actual producer's source correspondence through each change. Test a real changed document through the complete caller chain. | Actual correspondence reaches the consumer. Equal bytes and cache eligibility do not substitute for producer identity. Unknown regions remain explicit. |
| Comment and related-part emission | Exercise dirty and unchanged comments, footnotes, endnotes and glossary parts containing Fields. Follow actual dirty flags, clones, namespace insertion and cache refresh. Include repeated dependency copying. | Shared work reaches every relevant writer. A prepared Field that is later discarded is not mistaken for a published one. Existing imported parts retain their required bytes and relationships. |
| Original, current and derived input roles | Run the same physical input through clone, grouping, repeated source selection, aliases, cache refresh and reparse. Compare input totals before and after each step. Include genuine callback delivery and generated replacement data. | Original/current allowances register only at their actual approved ingress or delivery. Derived data, equal strings and generated XML cannot mint another baseline. |
| Returned-reader pairing | After real document, paragraph and note reads, test intentionally mismatched same-shaped models, same-length text changes and shifted note/body indices. Keep legitimate producer-directed edits and remaps as positive controls. Distinguish unsupported caller misuse from actual Native handoffs. | The pairing contract and caller responsibilities are explicit. Unrelated or stale model facts cannot acquire new input authority through equal values or shape. Record unproved guarantees separately from actual production behavior. |
| Complete Field ownership coverage | Cover authored, parsed simple and parsed complex Fields, shared runs, nested instructions, ordered results, revisions, Ruby, Drawing and stored legacy alternatives. Use cases with preserved outside text and partial captures. | Only positive Field-owned images are counted as Field output. Outside text and generated gaps retain their proper roles. Retired or discarded alternatives do not publish facts. |
| Metadata bounds and amplification | Count namespace records, URI bytes, references and transient clone work over direct, nested, sibling, alias and shared-owner cases. Compare the actual quantities with the existing reference formula before testing aggregate limits. | The decomposition agrees term by term. There is no repeated input credit or hidden multiplicative growth. Valid raw-fallback acceptance remains intact. |
| Deep checked model operations | Construct argument/switch-only chains through public model APIs. Test actual checked attachment, preparation, comparison and refusal separately in isolated child processes with stated stack and memory settings. Separate fixture cleanup from operation execution. | Each checked stage returns or refuses safely. Stage markers and child exits identify where work occurred. A fixture destructor or an earlier semantic refusal cannot be reported as an equality failure. |
| Failed preparation and sticky refusal | Force existing copy or constructor budgets to refuse, then attempt further real work on the same operation. Also test malformed plans, foreign-operation Sources, overlap and checked-length failures before reservation. | Refusal stays effective for subsequent applicable work. Failed preparation does not spend a reservation or publish model, package, cache or binding changes. |
| XML and namespace confusion | Use generated XML with shadowed canonical prefixes, default namespaces, aliases, MCE alternatives, unknown subtrees and unbound or malformed names. Compare typed output, untouched round-trip and the existing schema validator. | Generated names keep their intended expanded names. Unknown XML is preserved under the established contract. A namespace rewrite cannot silently reinterpret retained content. |
| Formatter and imported-value escaping | Put quotes, ampersands, angle brackets, CR/LF, non-ASCII text and forbidden XML characters through actual formatter and imported-value entry points. Include empty values and repeated callback aliases. | Escaping occurs exactly once. Invalid XML characters refuse atomically. Generated values do not become trusted raw XML or another input allowance. |
| Package parser robustness | Extend the existing generated ZIP/XML corpus with truncated records, conflicting EOCD/index data, duplicate entries, invalid lengths, oversized expansion and malformed referenced parts. Run public open and validation entry points in bounded child processes. | Existing package limits and refusal policies hold before excessive allocation. Public callers agree, and partial documents are not published after failure. |
| Relationship and filesystem boundaries | Review OPC part-name normalization, relative relationship targets, external targets and export filenames. Use generated packages with traversal-like and malformed URI/path cases. Observe filesystem and network access during ordinary open/render/export. | Package targets stay within the intended package or explicit output destination. External resources are handled only through the documented supported path. Ordinary loading does not unexpectedly fetch or read them. |
| Ownership-marker corruption | Extend comparison histories with missing, duplicated, moved and imported ownership bookmarks, selective resolution and unrelated preexisting references. Retain controls for valid tombstones and original unused notes. | Identifiable corruption refuses atomically. Valid histories remain usable. The approved ambiguity between a tombstone and a lost counterpart is documented rather than claimed detectable. |
| Non-emitting scalar validation projection | Compare the borrowed validation path with the actual emitting formatter on small inputs with sufficient existing copy allowance. Cover joined empty/null operands, separator size, case-filter depth, boundary whitespace and final illegal-character positions. Then increase dimensions separately in bounded child processes and inspect character-counter arithmetic and validation work without generating replacement Strings. | Type and XML diagnostics agree where both modes validate. No derived scalar credit appears. Deferred arithmetic and complexity claims have measured evidence rather than being inferred from the functional regression cohort. |
| Cumulative memory and work growth | After functional fixes, compare ordinary and instrumented builds over increasing nesting, siblings and repetitions. Measure quiet-process peak memory and actual work counters, including temporary buffers and failed attempts. | Results support the existing bounds without inventing new limits or treating instrumentation behavior as ordinary-build evidence. Keep raw samples, exact build identities and unsuccessful runs. |

## Execution and recording

Use generated fixtures and isolated processes. Begin with small positive and
negative controls, then increase one relevant dimension at a time. Preserve
source, compiler/features, command, fixture hash, child exit, raw output and the
stage reached. Establish a real predecessor failure where the check concerns
a repair. A build failure, zero selected tests or a mistaken fixture premise
does not establish a product failure.

Batch checks that share a binary. Review the actual caller chain and a different
discriminating case after a repair. Keep unsupported or unproved outcomes
explicit, including any required implementation that the assessment exposes.
Use the existing sprint records for delivery status rather than adding another
tracker here.

## Existing context

- [C1 contracts and preserved decisions](../.claude/plans/F-X272-design.md)
- [First Native writer and copy-boundary review](../.claude/reviews/F-X272-first-main-writer-implementation-codex-pass-1.md)
- [Checked w14 declaration and BOM review](../.claude/reviews/F-X272-checked-w14-implementation-codex-pass-1.md)
- [Current sprint](sprints/CURRENT_SPRINT.md)

The narrow approved native footnote exception and the valid-tombstone decision
retain their existing meanings. This follow-up does not turn either into a
passing observation. Ordinary schema/native compatibility, installed bindings,
CI, packaging and release approval remain governed by the release workflow.
