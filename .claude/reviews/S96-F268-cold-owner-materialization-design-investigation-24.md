# S96, cold owner materialization, investigation24

## Result and boundary

A small allocation candidate exists entirely in `word_piece_owner_addresses`. Direct body paragraph lookup can borrow its existing ContentLocation instead of cloning a temporary lookup key. A separate owned-key move can remove the map's initial clone, subject to actual transient capacity evidence. Neither requires changing finite qualification, accepted run paths, authority eagerness or any Source reader. These are preparation candidates only. Nothing is adopted or implemented here, and neither has an elapsed benefit established.

Read CLAUDE.md, the canonical workflow, current F268 plan and progress, Root29, Support11 and investigation22. Branch is sprint/s96 at cec87f9501eb13d0d5d60572472b227483262620. Existing dirty work is preserved. The inspected binding document.rs has SHA-256 ee722f7dbaa8060f9659f132a6f3cf8f6f1a10e17322ce67f3a7c342cc9398aa and Native document.rs has 9cbb239777eba5b9311838c3715d6fa8f9baa78713a880e31b74b1dd77f96e60. Both equal the actual entries in f268-guard-scheduling-implementation-20261108-1/FINAL-inputs.json. This authenticates the two inspected producer files against frozen21, not a new check of every repository input.

The profile comes from the earlier reviewed18/19/20 symbol artifact, not a new21 profile. Its cold profile SHA-256 is db07750986ab002dada9c3ce0d77d26712aa683202e7b3f07ee82f5a7d986519. Investigation22 authenticated that artifact and all five profiles. The cold OBSERVATION.json records90 actions, all semantics passing and two reopened cases. The Native and binding producers inspected here are unchanged by21. Its OXML scheduling change is separate and remains reviewed on its own evidence.

## Actual sampled producer

Under investigation22's active-ancestry union method, cold insertion has208 binding-active samples,143 in first-authority/owner materialization,53 in opening plain facts and42 in finite qualification. The latter categories nest within activation and cannot be added as separate costs. Allocator ancestry65 is also overlapping, not a count of allocations or milliseconds. Raw source: `.claude/scratch/S96-final-evidence/f268-five-remaining-symbol-diagnostic-20261108-1/cold-insert-attempt-1/profile.sample.txt:55`, `:58`, `:75` and `:192`.

The actual graph reaches word_piece_owner_addresses, ContentLocation hash-map growth and hashing at that same profile's `:98` and `:100`. It also reaches plain_paragraph_owners and AcceptedRunPath copying at `:138` and `:143`. append_word_owners_with_states appears at `:193`. Sampling does not resolve which inlined location clone contributes how much. The candidate below follows the actual source ownership, not an attribution of all these samples to one allocation.

## Candidate A, borrow the direct lookup key

At `crates/rdocx-py/src/document.rs:5272`, the one-index direct route produces Some(location.clone()), then immediately borrows that owned temporary for body_slots.get at `:5281` and discards it. For this exact arm, use body_slots.get(&location).copied() directly. Keep the existing Body guard, two-index ContentControl key construction, unsupported-path None and all result pushes unchanged.

The old and proposed lookup use the same concrete ContentLocation value and its same derived Eq/Hash. ContentLocation contains StoryId, item kind, index path and is_end at `crates/rdocx/src/document.rs:1304`. StoryId includes kind, String part name, owner index and fingerprint at `:1248`. Borrowing the original retains every component, including fingerprint and is_end. It must not substitute a newly constructed body location, a projection ordinal or a fingerprint-insensitive comparison.

The clone currently duplicates the owned part-name String and index-path Vec. Removing it removes that source-level copying operation for every one-index Body paragraph that reaches this branch. Actual optimized allocator calls and copied bytes remain to be observed. The lookup occurs before the mandatory retained Paragraph and Run location copies at `crates/rdocx-py/src/document.rs:5285` and `:5292`, which remain unchanged. There is no new reference stored in an owner or map, no changed retained buffer capacity and no lifetime beyond the get call.

This is the smallest independently selectable candidate. It preserves the existing order of all map insertions, reads and output construction.

## Candidate B, move the map key already owned by items

The first item loop clones each qualifying Body Paragraph or ContentControl location into body_slots at `crates/rdocx-py/src/document.rs:5188`. The next loop consumes those same items and immediately discards every non-Table location at `:5199`. body_slots is not read until paragraph processing at `:5271`.

A bounded alternative removes the first loop and, in the existing consuming loop's non-Table branch, applies its identical Body/kind/Some(slot) predicate and inserts the owned location directly. All Table processing stays in its original order and uses its original code. Other non-Table locations are still discarded. No preliminary partition, new buffer, helper, trait or borrowed-key map is needed.

The qualifying keys and slots, qualifying insertion order and last-write behavior for duplicate keys remain exact. Paragraph and ContentControl cannot also be Table because item_kind is one enum discriminant. Every table index, cell parent, paragraph ordinal, run index, section and hyperlink address is consequently produced in the same result order by the unchanged loops at `crates/rdocx-py/src/document.rs:5203`, `:5232`, `:5255`, `:5292`, `:5302` and `:5311`. The map is local and does not escape. Discarded read_facts are concrete None/String/Vec values, not user callbacks or Field readers, as defined at `:4445`.

This removes a second location clone per qualifying map insertion. It changes allocation/drop interleaving relative to table construction. It introduces no typed fallible operation, callback, Source access or public mutation in this infallible producer. An allocation abort is not a new recoverable error policy. The downstream authority transaction remains untouched.

Capacity is a qualification requirement for B. The original String/Vec may have spare capacity while their old clones may be length-sized. Moving the original into the local map can retain more spare capacity until this function returns, and duplicate-key replacement changes which original buffer remains live. Therefore B has no universal peak-memory claim. Record actual original/clone key capacities and local live-map payload before selecting it under existingQ22. Candidate A does not require B and can proceed without this representation change.

## Deliberately unchanged producers and blockers

The finite helper first checks paragraph properties and the entire paragraph_is_plain graph at `crates/rdocx/src/document.rs:8834`. It then rejects positions or non-ASCII-alphanumeric/non-space Text content before considering run properties at `:8837`. Property comparison and qualification-cache update occur at `:8844` and `:8880`. A fused run loop could change whether qualified_properties was updated before a later run fails. Comparing only a subset of properties could miss defaulted or unmodelled fields. Removing allowed-property construction without re-proving every field is not a simple equivalent repair. No such change is selected. The entire helper and its caller order stay byte-identical.

The18 materializer already preallocates successful finite text after checked sizing at `crates/rdocx/src/document.rs:35466`, retains its conservative original-reader fallback and calls the ordinary accepted_run_paths at `:35489`. Complete alignment, staged proof and unsupported-resource fallback remain at `:35491`, `:35506` and `:35539`. Neither candidate broadens that domain, caches eligibility or creates lazy authority.

plain_paragraph_owners still copies text and run paths at `crates/rdocx-py/src/document.rs:5443`. Some callers already own their values at `:5374` and `:6721`, but the main cold assembly borrows the opening tuple and separately uses its model at `:5647` and `:5667`. Changing one helper signature alone does not eliminate that cold copy. Consuming/restructuring the opening batch needs its own alignment/fallback/model-lifetime proof and is not selected here.

append_word_owners_with_states already extends a reserved destination using an owned locations iterator at `crates/rdocx-py/src/document.rs:8358`. scopes returns a static slice, not a per-owner Vec, at `:4950`. Its state clones retain concrete Arc generations, defined at `crates/oxml-py-support/src/lib.rs:200`. Dropping scopes, coalescing generations or replacing owners with ordinals is not justified by its sampled presence. No append repair is selected.

## Required evidence before any future adoption

Compare complete ordered address values, parent/body slots, projection ordinals and read facts, not map iteration order or text alone. Cover zero/one/many items, duplicate qualifying keys, direct paragraphs, ContentControl-contained paragraphs, unsupported path lengths, tables with nested cells, related stories, sections and hyperlinks. Preserve exact ContentLocation identity components. If B is selected, separately observe its capacity/lifetime difference and actual clone removal. Neither a source clone count nor samples establish elapsed acceptance.

A temporary source-bound diagnostic can count actual old clone execution at `crates/rdocx-py/src/document.rs:5196` and `:5273`, separately from candidate owned insertions and borrowed lookups. Record incoming and retained String/Vec lengths/capacities, duplicate replacements and local buffer release. Run the same immutable original cold insert/remove100/1000 inputs and required warm controls, preserving their original setup. Assert full values, Source observations and ordered address equality before copy-work assertions. Authenticate the old/new producer bytes and installed payloads. Instrumentation must be absent from the later ordinary artifact. This is a proposed diagnostic, not evidence already collected or permission to rerun timing.

Preserve all actual Source/admission/capture/attribute/event/resource work and counters. Keep private Field and malformed-input controls and exact first recoverable error. ensure_word_authority must still validate complete facts, reserve storage, prepare indexes and initialize scopes in the existing order at `crates/rdocx-py/src/document.rs:12416`, `:12418`, `:12435`, `:12438` and `:12440`, then publish only at `:12452`. Compare original immutable inputs, complete results/ZIPs/every member, owner IDs, callbacks, held/fresh handles, refusal atomicity and saved/reopened semantics. Preserve current budgets and original timing protocol.

## Findings and handback

Zero new product defects, zero smells and zero nitpicks are established. Candidate A is a concrete minimal next-sprint allocation investigation. B is separately selectable only with explicit transient capacity evidence. No finite-helper rewrite or other producer change is selected.

This pass performed read-only source/evidence inspection and hashing, and wrote only this report. No product, plan, test or progress edit, Cargo/build/test, runtime profile, elapsed qualification or Git/network mutation occurred. Root21 is not altered or delayed by this preparation. The current ordinary75/79 result and cold insertion3.019208ms failure remain retained, with no noise waiver or favorable retry. Final memory/equivalence/calibration, Office, installed typing, hosted/archive and integrated-union obligations remain pending. Control returns to Root.
