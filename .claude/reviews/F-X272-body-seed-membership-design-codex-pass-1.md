# F-X272 body seed membership design review, pass 1

Result: zero defects, zero smells and zero nitpicks in this bounded design amendment. This is not implementation or provider approval.

## Scope and source

Reviewed only the normative constructor revision at `.claude/plans/F-X272-design.md:4973` and the appended Actual body seeds for inert Word membership section at6091. Plan SHA256 is `d9ec2d21a13755a6180088ad7ae7727a29d1d51ac001b59876c56404a5afe138`. Read CLAUDE.md and WORKFLOW.md. The plan belongs to Root and was not changed by this reviewer.

Read actual recovery source for the existing graph and visitor. Lower document.rs SHA256 is `adf97ba6fc582aa3b7c5d22127b87560f8c701beed97415253583804f326252b`, text.rs is `3c09961b0ef309a988febacf8881d6ea09a718c400351db15caf9bc7b0e2a5a3`, content_control.rs is `e871ce055662eeb02536cbcace5ba7609574ddb2c53de7f762edd0d7e7aa22b7`, and Native document.rs is `6f3e2a5b0dd7848ddf9fd8f62d2baee548f81fe0412c851549de641a619e1c6f`. These are read-only source observations, not an atomic implementation freeze.

## Contract checks

The single normative declaration at `.claude/plans/F-X272-design.md:4973` places current_bodies immediately after current_paragraphs. The amendment at6093 to6098 makes all actual callers and the external client adopt the same seven parameters and explicitly retires the six-parameter preparation as final proof. It adds no public type or carrier name. The existing CT_Body and shared borrowed lifetime suffice for this input.

The body seed closes a concrete existing gap. Recovery `crates/rdocx/src/document.rs:21467` leaves SdtContent::Run and RawXml out of the paragraph read visitor. Lower `crates/rdocx-oxml/src/document.rs:1194` traverses both current control content and original_resource_runs, whose stored original Run accessor is lower-private at `crates/rdocx-oxml/src/content_control.rs:2637`. A paragraph-free root control therefore cannot be completely enrolled through paragraph seeds or Native direct current-run callbacks. The actual borrowed main/glossary body seeds plus existing note paragraph seeds at plan6100 to6104 address that gap without changing public read selection.

The reused physical graph is concrete at lower `crates/rdocx-oxml/src/document.rs:1149`. It reaches paragraph revisions and Ruby runs, controls, tables, Field owners, revision paragraphs/runs and drawing text. Field resource_owner_children at `crates/rdocx-oxml/src/text.rs:7192` inspects stored owners rather than recursive display selection. The graph is already used for returned physical owners by commit_publication_owner at2948 to2971. Plan6105 to6110 correctly distinguishes this reusable child enumeration from the additional Text-leaf/raw-record inspection needed for membership. No claim that resource_children itself returns raw leaves is made.

Plan6107 requires owner identity deduplication before expansion. Overlapping Body, paragraph and Field seeds are explicitly allowed at6112. The existing constructor stores borrowed live identities rather than contents at lower `crates/rdocx-oxml/src/document.rs:5187`. The new design forbids parsing, lazy projection, normalized scalar creation, resource admission and selected-input funding during membership expansion at6109 to6110. Actual allocation of reference catalogs is not disguised as scalar materialization. Stored Field membership does not override writer/cache eligibility at6114.

The scope limit at plan6116 to6117 is material and honest. Root raw records outside the physical graph still need their actual separate source handoff when used. Body seeds are not an ordinary certificate or a claim to have inspected every XML record. Actual frozen Field relationships and producer-supplied TemplateWordLineage remain required at6115 and6118 to6120. No equal-byte discovery, post-preparation matching or clone-derived allowance is authorized.

The concrete controls at plan6122 cover the root-control gap, retained originals, stored nested Field Text, revision/Ruby, overlapping seeds, inert zero allowances, foreign candidates and unchanged Native read outputs. HLD03/HLD10 and the Rust0.17.0 inventory own the revised callable at6126. Existing public API, resource, shared-consumer and final riders remain due at6127 to6129. The amendment does not introduce a file, module, generic framework, new policy or review-cap reset.

## Remaining debt and prose check

The seven-parameter implementation, all real Native callers, borrowed membership lifetime controls, source and raw-owner coverage, actual clone/preparation lineage, exact external compile client and strict SemVer checks remain pending. No build, test, native operation or oracle was run. This review cannot qualify full C1 or a dependency provider.

Checked this review prose before writing. It contains no em dash, en dash or prose semicolon. Only this review file was written. Review pass ends here and returns control to Root.
