# F-X272, Root returned-reader functional support, pass 1

**Reviewed**: The bounded Root increment from `returned-reader-functional-batch/ROOT_TAKEOVER/source` to `ROOT_TAKEOVER/root-returned-reader-review-freeze/source`. Five paths are pinned, with four changed source files and unchanged Native field.rs. The textual delta is 2,876 additions and 627 deletions, including substantial formatting. This is independent support for Root final disposition of this increment, not another old external-model peer, a reset of any review cap or a whole F-X272 microscope.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in the bounded Root additions examined. Whole-provider acceptance is explicitly excluded.

## Defects

0 defects.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## Source identities and evidence

All 12 takeover snapshots authenticate against `ROOT_TAKEOVER/HANDOFF.json`. All five review snapshots authenticate against `root-returned-reader-review-freeze/MANIFEST.json`.

| Frozen path | SHA256 |
|---|---|
| crates/rdocx/src/document.rs | 4639ea49dab0c76891e0d6b127b91572805351efb4e999294f7c33e9217698b2 |
| crates/rdocx/src/template.rs | 1205589c9cd07d1e509c455cb2702d8c5a5f57e69dfd48fe599fed91a9ec4dc3 |
| crates/rdocx/src/field.rs | f4660edcc406f587c0cd9641c130d001a1e8b16a419cb24cbe4b16a745909f56 |
| crates/rdocx-oxml/src/document.rs | 53bcc06a158951f9bb7fb3a46310bf961e450aab79a7b7bdd9ef356db74e26b3 |
| crates/rdocx-oxml/src/text.rs | 772df33c5dd4738a1ba1591c5ed3f4e859beb0f39412c3c48e0e7735ae26807b |

The two archived checks-2 test binaries match their recorded hashes. Raw logs identify five lower returned-read passes, 18 Native scoped-template passes and 18 versioned-template passes. These are three named cohorts, not a whole suite or proof of every pending handoff. The tests cover the exact preceding production source. Comparison to the final freeze shows only the prepared-scalar tuple alias and final Option borrow/move cleanup afterward. Lower source is byte-identical between those tests and this freeze. The test results are not relabeled as an execution of different final Native whole-file hashes.

The checks-2 strict attempt is retained as exit 101. Its successor is separately represented by the final manifest's strict_exit 0 and archived strict log, which records completed rdocx/rdocx-oxml checking in 24.57 seconds. The recorded command uses all-targets and warnings denied. No all-features run is inferred from that command. No reviewer build, test, endpoint, native application or security probe was run.

## Functional assessment

The adopted derived-Field bridge at `crates/rdocx-oxml/src/document.rs:6090` checks operation identity, sticky refusal, successful reader facts and the actual registered Field map. It forwards existing origins and uses the same clone/normalization implementation with optional frozen input. The `None` path does not enroll a derived model as frozen ingress. Native preparation at `crates/rdocx/src/document.rs:26468` retains preflight bypass selection and the prior semantic-equality installation gate. The unchanged equality case does not install a different private model merely to retain a receipt. Related and private branches are not implied complete by this main-body connection.

The corrected block selection at `crates/rdocx/src/document.rs:10092` requires an actual scoped Source before reusing a cached record. It obtains that Source from the prior exact output handle or the frozen ingress record. Otherwise the existing substantive writer runs. A returned-reader source selects its actual paired writer. Ordinary None-context reuse is retained. Scoped whole reconstruction without a companion declines to the authoritative factory once, while separate qualified kept-model reuse remains. This addresses the observed drawing-reparse input loss without adding a qualifying second parse, credit by value or a new refusal policy.

The main drawing and nested-relationship branches carry selected source edits into the actual scoped reader. `apply_template_source_edits` at `crates/rdocx/src/document.rs:6725` validates ordered nonoverlapping edits and composes copied intervals with generated replacements. The next writer consumes the actual returned-model companion. Related-part planning keeps unchanged package storage when the edit plans make no byte change and reattaches the produced Source to that actual stored record. It does not select a new ingress role from later output equality.

Raw textbox preparation now uses the actual scoped part producer and existing original-literal walker. Selected scalar dependencies cross the rewrite tuple, while new replacement bytes stay generated. Publication uses the same actual Source through namespace closure, part splicing and the existing reader. The ordinary byte paths remain separate callers of their existing behavior. This closes the particular Root handoffs examined, not all current/imported/related story contexts.

`TemplateReplay` at `crates/rdocx/src/template.rs:1370` carries bytes beside their producer Source. Item slices, child splices, required-cell insertion and root composition keep actual ranges and generated framing distinct. Source-bearing final records and separately required legacy bytes use the existing paid copy operations. The structural-versus-scalar predicate is retained, so child-only paths do not acquire a new unconditional formatter call.

Prepared scalars retain borrowed JSON dependencies and shared callback events. `ProgramExecution::checked_clone` at `crates/rdocx/src/template.rs:228` reserves actual scalar String copies recursively before cloning and preserves the same callback delivery cells. The borrowed scoped-format tuple does not pay a nonexistent scalar clone. The ordinary owned tuple pays before its real clone. Literal non-emitting evaluation avoids constructing a discarded String. This is not acceptance of the still-pending complete validation-only expression path.

## Discriminating controls inspected

The entity replay control uses a real mapped Text and then the charged consumer, rather than byte equality from a generated-only source. The drawing control requires an actual changed drawing id, the paired reader result, next writer and a successful charged read of A&B. The cached-reader control reaches the real factory before returned-model emission. These controls distinguish output-byte parity from retained Word-origin authority.

The derived Field controls include parsed simple and nested cached forms, bypass and normalization, unread or foreign Sources and unregistered Fields. They exercise preparation followed by a real writer and charged Text consumer, rather than merely inspecting a receipt. The note controls distinguish selected later-note output from whole-family output, footnotes from endnotes, and selected append from whole-new-part drawing preparation. The ordinary selected-note behavior is not silently replaced by whole-part namespace preparation.

## Explicit boundaries

Importer and allocated-fragment Source closure, retained-main special branches, complete related/private-control/Revision/Ruby contexts, full non-emitting expression materialization, and remaining producer pre-growth obligations are not discharged by this report. The current importer work in the moving recovery tree is outside these immutable files. Same-operation model/Source pairing remains the explicit caller responsibility in the adopted contract. This pass does not expand that contract into a new public mutation or provenance policy.

Final source integration, external callable inventory, SemVer, complete local and integrated gates, hash/native/oracle riders, provider readiness and feature completion remain pending. The report neither approves another public method nor changes an allowance, quota or acceptance boundary.

## Not found

No concrete additional correctness, panic, namespace/schema-order, preservation or structural-rule defect was found in the bounded Root changes examined. No independent smell or nitpick was found. Root's earlier failed diagnostic and regression attempts remain distinct from the actual later passing cohorts.
