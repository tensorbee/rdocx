# F-X283, Bound whole-cell setters and Python open-replace-save cost

**Status**: approved
**Sprint**: S96
**Size**: L
**Depends on**: none

## Problem

Issue 445 reports body-dependent Python cell setters and unnecessary binding
inventory work for open, replace and save without live handles. PR 443 does not
discharge those outcomes. Registration in HLD14 retains both exact workloads
and the separate cell-paragraph setter observation.

Current direct body cell edits are still whole-document work. `PyCell::edit` at
`crates/rdocx-py/src/table.rs:2080` clones the document, resolves a
`TableLocation` (`crates/rdocx-py/src/table.rs:59`, `:126`) and then publishes
through `adopt_property_edit` (`crates/rdocx-py/src/document.rs:11268`), whose
`WordPieces::assemble` (`:5404`) compares every body block before and after the
edit. `PyCell::set_text` at `crates/rdocx-py/src/table.rs:2514` calls the native
`try_set_cell_text` at `crates/rdocx/src/document.rs:35199`, which clones,
reopens the whole package and then runs the full `prepare_word_revision`
inventory at `crates/rdocx-py/src/document.rs:10928`. `PyParagraph::set_text` at
`crates/rdocx-py/src/paragraph.rs:983` resolves its publication route through
`ParagraphLocation::native` (`:197`), and the cell branch reaches
`cell_paragraph_story_location` at `crates/rdocx/src/document.rs:35472`, which
clones and serializes the whole document. The prepared route already exists at
`crates/rdocx-py/src/paragraph.rs:434` and `crates/rdocx-py/src/document.rs:10209`.

This canonical draft consolidates the registered acceptance contract and Root's
source investigation. The operative simplification below is a separate Root
repair after actual design peer3 and Root-delegated final1. Grok1, DeepSeek2
and Codex3 remain the three used design peers, with none available and no reset.
The implementation review cap has not started. Status remains draft, pending
fresh Root final2 after adoption. No implementation is authorized by this text.

## Operative approval, 2026-11-02

Root approves this design after actual design peers Grok1, DeepSeek2 and Codex3, followed by final1 and clean final2. The final2 input is the authenticated simplification revision74a36d99a3c835cee3d99da9db058ae4da7d3c5ccdecc1f0351baa89bc57e385. This approval supersedes historical draft and retained-dormant-cache proposals below. It does not rewrite their review results, reset a counter or claim performance. Implementation peers remain zero.

The operative dormant mechanism retains no validation facts or reusable pieces between calls. Exact eager validation and its transient costs remain in open and positive replacement, and first authority activation always validates current native state again. All13criteria, exact workload paragraphs, public bridge/initializer contracts and final gates remain unchanged. No material product question remains.

F-X283 is approved but pending and unclaimed until the clean F-X282 integration Base is available. Its facade/binding/HLD resources run after F-X282 and before F-X281. This is exclusive resource scheduling, with no invented provider edge. A restricted F-X272 OOXML-only source wave can run alongside it if actual file ownership stays disjoint. Actual Base revalidation, implementation, independent review, scoped evidence and all final acceptance remain due. Publication still requires separate approval at the reviewed SHA.

## Spec reference

- `docs/hld/03-architecture.md`, staged edits and package identity.
- `docs/hld/10-bindings-spec.md`, Word checked XML owner authority and live handle routes.
- `docs/hld/12-testing-strategy.md`, performance strategy.
- `docs/hld/14-development-backlog.md:8862`, registered F-X283 contract.

Issue 445 body SHA-256 is
`60e20b2e1ca33a3ff0829d10dea13e973d61278f6c6f769a90503c28e32750c1`.
Registered PR 443 head is `7b3dcd0e4a3f89c4ecc7b88718c68a9a58159ad8`.
These are acceptance-source identities, not accepted implementation or CI proof.

## Approach

Retain the existing public Python signatures, errors and native field-specific
conversions. For a cell whose owner is a cell of a direct body table, publish
one local table candidate and one local commit. Related, nested and controlled
owners keep the existing checked fallback when they do not qualify. The
fallback preserves supported behavior rather than introducing a new refusal.

### Native candidate and commit route

The binding cannot reach `Cell.inner` (`crates/rdocx/src/table.rs:2166`).
Retain the two local-cell hidden facade bridges below and the consuming replacement bridge specified under dormant publication. Root authorizes those three additive Word0.17.0 items for design review, not implementation acceptance. Do not add
an invented `CT_Tbl::validate_resource_limits` or a third public query method.

```rust
#[doc(hidden)]
pub fn direct_body_table_cell_candidate(
    &self,
    body_slot: usize,
    row: usize,
    cell: usize,
    edit: impl FnOnce(&mut Cell<'_>) -> Result<()>,
) -> Result<Option<CT_Tbl>>

#[doc(hidden)]
pub fn commit_direct_body_table(&mut self, body_slot: usize, table: CT_Tbl)
```

`body_slot` indexes `self.document.body.content`, never the facade table ordinal
or the owner-vector index. The binding derives it from `direct_body_cell_slot`
(`crates/rdocx-py/src/document.rs:10347`). Invalid public handles first take their
existing validation/error path. The native candidate uses the existing table
lookup error family for a missing or wrong-kind slot, and preserves the calling
operation's existing missing-cell message. These are internal bridges, not a
new route for accepting invalid public indices.

The candidate returns `Ok(None)` for an unqualified source or selected cell,
without calling `edit`. It clones only the selected table, finds the cell, and
runs the concrete native closure over `Cell { inner: candidate_cell }`. Native
text/property validation remains unchanged. It moves the resulting table into
`CT_Body { content: vec![BodyContent::Table(Box::new(table))], sect_pr: None }`,
calls the existing `CT_Body::validate_resource_limits`
(`crates/rdocx-oxml/src/document.rs:5195`), and extracts that table for
`Ok(Some(table))`. `Document::body_item_owners` at `:32107` already uses this
singleton-body validation. No whole-document clone or serialization occurs.
An edit or validation refusal is returned before any live mutation.

Commit has the exclusive-borrow precondition that the same body slot is still
a table, the source proof below is still applicable, and no intervening native
mutation occurred. After the binding revision commit succeeds, it replaces
`**existing = table`, clears derived layout and resets the unsavable memo to
`OnceLock::from(None)` only under that source proof. This is nonfallible.
Ordinary `invalidate_layout` keeps its existing invalidation behavior for every
other publisher. A failed precondition is an internal programmer error, not an
input-dependent path supplied by a public caller.

SemVer impact is three additive hidden methods on published `rdocx`, including the consuming replacement method below.
`rdocx-py` is a separate crate and existing hidden facade bridges justify this
placement. The closure has concrete text and property consumers. No trait,
crate, source file or generic forwarding abstraction is introduced.

### Source proof and bounded unsavable qualification

`plain_local_edit_blocked == false` alone is insufficient after a flush.
`unsavable_body_edit()` is also not a bounded per-setter qualification because
`invalidate_layout` clears its memo (`crates/rdocx/src/document.rs:24947`) and
recomputation scans the source part (`:35188`, `:35199`, `:35215`).

Use a private source certificate in the existing facade file. It retains an
`Arc<Vec<u8>>` for the exact main-part bytes, the root and body namespace
metadata used by `plain_local_edit_facts`, and the established absence of
nested modeled namespace owners. The existing package stores shared part bytes
in `OpcPackage.parts` (`crates/oxml-opc/src/package.rs:69`). Creation/open may
record this certificate only when the actual source classification proves no
unsafe root/body prefix and an empty nested-owner list. Fresh compatible
package creation establishes the same fact from its constructed source. This
is actual source classification outside setter timing, not a guessed default.

The certificate also stores the actual selected package-part key and main-part
name. Obtain its retained Arc from the existing public `OpcPackage.parts` map
when classifying source, matching the actual selected get_part slice by pointer
and length if key resolution was indirect. That lookup is outside setter timing.
Do not require the new F-X282 get_part_shared method as a provider. A setter
uses the recorded key for one map lookup, checks the current main-part name,
and compares Arc allocation identity plus root/body namespace metadata. It
never repeats a part-map search, deterministic-key fallback scan or byte compare.

Refresh this certificate from actual already-computed source classification at
`from_package_reusing`, `set_story_source_xml` for the main story and successful
`flush_document_to_package` publication. Compatible construction establishes it
from the actual generated main bytes. A minimal package with no main bytes is
unqualified. Clone-for-staging shares it. A source publisher unable to establish
the complete fact invalidates it, or its changed allocation makes it inapplicable.
No general publisher gets to relabel a guessed source as owner-free. Include
Document.new plus many ordinary paragraph additions and table creation as a
positive gate, so proof custody does not accidentally disable the issue fixture
before its timed setters. Test after source replacement, namespace change and
save/reopen as well as direct opened input.

Candidate qualification compares the current part slice by pointer and length
with the retained allocation, and compares the current root/body namespace
metadata with the certificate. Retaining the Arc prevents address reuse.
A different source allocation, missing source, changed namespace metadata,
nonempty/unknown owner list or incomplete certificate returns `Ok(None)`.
Never fall back to byte equality or a source rescan inside the local setter.
General typed edits need not erase this source-only fact: applicability is
checked against the current source allocation and namespace metadata each time.
A changed/flushed source needs a new actual classification outside this path or
takes the existing checked fallback. Clone-for-staging copies the certificate
without copying source bytes. This certificate is private and is not a new
public cache-query API.

Under this certificate, the existing unsavable computation returns `None`:
its root/body prefix test has the same inputs, and its retained nested-owner
branches have no owners. The admitted concrete edits cannot introduce root/body
namespace metadata or change the package's main-part bytes. It is therefore
sound for this narrowly qualified commit to retain the proved `None` while
clearing layout. Do not apply this shortcut to compatible nested declarations,
unknown source classification or general native publishers. Controls must show
both initial source proof and its rejection after source replacement/namespace
change. A cached `None` without the certificate is not qualification.

### Binding local cell edit

Keep one private entry point in `crates/rdocx-py/src/document.rs`:

```rust
fn apply_local_cell_edit(
    &mut self,
    py: Python<'_>,
    body_slot: usize,
    table_ordinal: usize,
    row: usize,
    cell: usize,
    candidate: rdocx_oxml::CT_Tbl,
    operation: &'static str,
    publication: Option<WordPublication>,
) -> PyResult<bool>
```

`true` means committed, `false` means unqualified with no published mutation.
Native or binding errors propagate without publication. `publication` is None
for properties, Some(CellText(actual table location, row, cell)) for Cell.text,
and Some(ParagraphText(actual routed paragraph location)) for Paragraph.text.
Use the existing private WordPublication enum, not an operation-name heuristic
or a fourth public bridge. Before any publication, verify the supplied mode's
selected owner matches the already qualified table/cell/direct paragraph.
A mismatch or unsupported local mode takes checked fallback before mutation.
The two ordinals are different values. The current table
owner gives the body slot. The handle's current indexed table projection gives
the facade table ordinal. `indexes.cell_paragraphs` is keyed by table ordinal
(`:5639`), while native replacement and the piece slot use body slot. Missing
or inconsistent projections take fallback before mutation. The private
`PyCell::edit` accepts an `Fn` closure rather than `FnOnce` and borrows it for
the candidate attempt, so a later binding qualification failure can run the
same concrete edit on the existing checked fallback. Its five existing
property closures at `table.rs:2643`, `:2657`, `:2684`, `:2707` and `:2728`
borrow or copy their converted arguments and support `Fn`. The public native
candidate retains its `FnOnce` signature. Text setters retain their input text
and reconstruct their existing fallback call. No caller must repeat a consumed
closure or execute an edit on the live model while deciding fallback.

Preparation builds a concrete packet containing the table candidate, selected
owner replacements/newborns, removed index entries, prepared route/parent
buckets, read projections, one replacement piece and an optional prepared
revision. All fallible work and reservations occur here:

1. Require an aligned cached body piece at `body_slot`, matching source context,
   unchanged body-child count and a complete selected-table projection. The cached
   piece records its physical-table count once from prepared immutable facts,
   including nested tables rather than only the facade table projection. Local
   eligibility reads that cached count and requires exactly one physical table. A nested table anywhere in
   this selected body table takes checked fallback, including a nested neighbor
   of the otherwise clean selected cell. The local ordinal key does not prove
   containing-table identity for that topology. The
   selected cell contains no nested table or content control anywhere in that
   cell. If qualification or mapping fails, use the old checked route.
2. Build `word_block_owners` from the candidate only. `Err` propagates before
   mutation and `Ok(None)` selects fallback. Keep the new selected-table facts
   and one prepared `Arc<WordBlockPiece>` containing the candidate block.
3. Place standalone facts with the existing selected table body's item base and
   its first cell-story base using `placed_route`/`placed_story`. Verify that
   table/body item, table, cell-story and text-box counts are unchanged. A change
   that would shift another block's offsets takes fallback.
4. Find selected table descendants without scanning `self.owners`. Add private
   `cells_by_parent` and `runs_by_parent` indexes alongside the existing
   paragraph/table parent indexes. `tables_by_parent` contains tables, not
   cells. Traverse only the selected table's indexed subtree. Every constructor,
   full reindex and local publication maintains these two new indexes.
5. Join selected-cell paragraphs by direct ordinal, not raw story-item position.
   Property edits preserve owner/content/properties/grid states and relocate
   their routes. Cell.text preserves the selected cell owner and uses the exact
   existing CellText successor/retirement policy for its content and descendants,
   including retained later paragraphs. Paragraph.text uses the distinct existing
   ParagraphText policy: preserve its paragraph owner/properties and unrelated
   cell paragraph/run scopes, while changed content/replaced descendants receive
   their existing successors. Never infer CellText retirement from a boolean
   shared by both setters, and never reuse property-only states for changed
   paragraph content. Match each supplied publication to its real selected owner.
6. Prepare exact index deltas. Remove each affected old generation and route
   membership, install its relocated/replacement owner at a stable vector index
   where possible, and append newborn owners at reserved indices. Removed
   owners lose all current generation/route/parent memberships. Do not erase
   vector elements and shift unrelated owner indices. Qualification requires
   every existing owner route in the table to have a corresponding new owner
   under the ordinal join. Births are allowed, but disappearing owner records
   take the checked fallback before revision preparation. Thus this path needs
   no inactive-slot representation and never leaves stale retired records in
   `self.owners`. Replaced state generations still leave the generation index
   and receive their successors. This covers the exact empty-cell-to-one-run
   case without claiming a local fast path for arbitrary run removal. Prepare the affected
   `cells_by_parent`, `paragraphs_by_parent`, `runs_by_parent`, `tables_by_parent`
   buckets and complete selected-table `cell_paragraphs` entries. Add newborns
   to `by_generation` as well as `by_route`. Do not call `index_owners_again`.
7. Update selected-table and paragraph read facts, including insertion of a first
   paragraph into an empty cell. Body paragraph/table projections do not move.
   All matching and reservations are local to this table and the affected hash
   buckets. No full owners/indexes/pieces clone is permitted.
8. For content, prepare the exact supplied CellText or ParagraphText retirement
   set plus snapshot successor. None keeps property mode without a revision.
   The caller `retired` list excludes the document scope:
   `prepare_scopes` adds that scope automatically. Property mode does not
   advance the counter. No scope is retired during preparation.
9. Validate piece alignment and prepare the replacement Arc, without calling
   `WordPieces::splice` on the live cache. A missing, incomplete or misaligned
   cache selects fallback. Replacement has equal length, so no Vec splice,
   full-vector clone or cache clear is needed.

Commit text's prepared revision first. `commit_scopes` is the last fallible
operation and rejects before changing scopes. Then swap the native table,
install the prepared owner/read/index deltas and replace exactly
the already-qualified active `self.pieces.as_mut().expect("active pieces").blocks[body_slot]`. Property mode performs the same infallible
swaps without a revision commit. Do not call `install_inner`, which takes a
whole Document and would clear pieces without a pending cache. Do not set
`pending_pieces` for this in-place publication. A preparation/commit refusal
leaves bytes, owners, scopes, cache contents and held observations unchanged.
Capacity reservations before refusal may remain allocated but publish no facts.

### First `w:tcPr`, text semantics and checked fallback

Both absent and present-empty `tcPr` emit their first real properties before
cell story items. Relocate direct paragraph routes by ordinal, including held
first and later paragraphs, and update all their run routes/index buckets.
The entire selected cell must exclude nested tables and controls so their
routes cannot silently shift. Do not add a public properties-presence query.

`Cell::try_set_text` edits the first direct paragraph or inserts one, preserving
range anchors and replacing run content (`crates/rdocx/src/table.rs:2183`).
Notes/revisions/pictures in replaced runs follow that exact existing behavior.
The checked fallback retains comment reconciliation and serialize/reopen
(`crates/rdocx/src/document.rs:35099`). Anchor survival alone proves no package
or namespace equivalence. Compare local and fallback outputs for comments,
bookmarks, permissions, note/revision content and unmodified opaque properties
and siblings, including namespace aliases and shadowed serializer prefixes.
Prove saved-byte/part agreement where promised, not merely reopened text.
Unsupported qualification takes the existing checked path, never a new refusal.

### Cell paragraph publication

`PyParagraph::set_text` (`crates/rdocx-py/src/paragraph.rs:983`) resolves its
publication route through `location.native(&document.inner)`, which reaches
`cell_paragraph_story_location` and serializes the document. The setter changes
to `document.paragraph_route(&location, None)` (`crates/rdocx-py/src/document.rs:10209`),
the same prepared route `bind_current` already uses at `paragraph.rs:434`.

For a resolved direct body cell paragraph, the setter uses
`apply_local_cell_edit` with a closure over the cloned cell:
`cell.paragraph_mut(direct_ordinal)` then `set_text`. `direct_ordinal` is
counted from `indexes.cell_paragraphs[(table_ordinal, row, cell)]` as the number of
entries before the selected one whose `index_path().len() == 1`. The selected
paragraph must itself have `index_path().len() == 1`, so it is a direct cell
paragraph and `paragraphs_mut` (`crates/rdocx-oxml/src/table.rs:2274`) selects
the same physical paragraph. A controlled paragraph falls back to the existing
`edit_story_paragraph` route. Improving only the getter does not satisfy
A283-5, and debug physical-route assertions deliberately call the native full route. Their
work is not a release bound. Use an authenticated release binding for the actual
setter work/timing evidence, with debug diagnostic work reported separately.

### Dormant authority without retained validation caches

Retain the EXACT existing eager fallible construction work in `from_document`:
`WordPieces::default().assemble` and its existing full owner-inventory fallback,
native open/package checks and the explicit fallible related-part loop. Do not
replace the fallback by assemble's `.ok()` branches or silently suppress any
existing error. The same validation runs before every positive dormant candidate
publication. Open and positive candidate calls discard their resulting transient
addresses, read facts, projections and pieces before returning. They defer live
`RevisionState` allocation and authority indexing, not mandatory validation.

Factor the existing validation sequence into one private helper in the existing
binding file, with these actual consumers: open, positive dormant publication and
first activation. Its concrete result may be the existing types in one tuple:

```rust
fn validate_word_facts(
    inner: &rdocx::Document,
    py: Python<'_>,
) -> PyResult<(Vec<WordOwnerAddress>, WordIndexProjections, WordPieces)>
```

The helper preserves the current assemble/inventory-or-fallible-fallback sequence
and existing projection construction. It is not a second parser or new public
API. No `ValidatedWordFacts` field, address/projection cache or native-model epoch
is introduced on `PyDocument`. A private transient carrier is unnecessary when
this concrete tuple suffices. Do not retain that tuple across dormant calls.

Snapshot and removal scopes become `Option<RevisionState>` and the existing
binding `pieces` field becomes `Option<WordPieces>`. Dormant state has both scopes
absent, empty live owners/indexes, `pieces = None` and `pending_pieces = None`.
There are no reusable dormant addresses or projections hidden in another field.
Active state has both scopes present, complete live owners/indexes and
`pieces = Some(...)`, with pending pieces following the existing active commit
protocol. An incomplete assembled piece cache may remain inside that active
Some and selects the same checked fallback as today. No placeholder epoch clone
or independent boolean is used as authority. Existing active piece consumers
explicitly access Some only after their existing authority guard. Local cell
qualification still takes checked fallback for missing or incomplete pieces.

First activation ALWAYS reruns `validate_word_facts` on the current native model,
even immediately after open or a successful dormant edit. It then prepares and
installs live authority as specified below. It does not trust an earlier
validation result. This intentionally repeats immutable validation on first
handle access, whose existing setup location is outside Issue445 case1 setter
timing. It does not move any existing validation out of open or replacement.

All eager validation and transient allocation/deallocation remain inside the
actual open or positive replacement call being timed. Issue445 asks for total
Python time at most1.2times the same-build CLI. Its lazy-owner suggestion is not
a separate zero-facts criterion. Whether deferring live allocation/indexing while
discarding transient facts is sufficient is decided by the unchanged actual
timing gate. No success is promised. If it misses, improve implementation without
skipping validation, shifting work across measured boundaries, dropping the
transient deallocation cost or inflating CLI work.

Q-F283-01 is settled by preserving current checks. A bounded historical-artifact
probe gives a candidate witness: native open accepts a body hyperlink whose
declared `r:id` has no relationship, while owner snapshots refuse it. Baseline
and internal-anchor controls pass. The manifest is
`/private/tmp/rdocx-f283-open-link-probe/manifest.json`, SHA256
`2eecf3a0c9b8ef54d2529aee6ec4d7cb3d92c1545838b079c32dda666c00425c`.
This is not current-source qualification. Reproduce the missing-link and valid
controls against freshly identified implementation-Base native and Python
artifacts, retaining the same open failure and error family. Prior negative
probes never proved native and binding validation equivalent.

### Initial scopes at the current dormant serial

Explicitly amend the existing `RevisionCounter::initialize_scopes` construction
contract in unpublished `oxml-py-support`, whose Cargo.toml has publish=false.
Remove only its `current != 0` rejection. Retain the already-initialized and
unique-current-document-epoch guards, checked generation arithmetic and complete
allocation before changing `next_generation` or `initialized`. Do not reset
serial/origin, replace or retire the current epoch, or make activation a commit.
This is a named shared-support behavior amendment, not an assertion that the
existing method already permits it, and not a fourth published facade bridge.

Dormant successful positive publication calls existing
`prepare_scopes(operation, &[], &[])` and `commit_scopes` once. No authority
handles exist. Zero, mismatch and refusal commit nothing. First activation
initializes all owner/snapshot/removal scopes at that current serial, using
successive unused generations after `next_generation`. An outstanding prepared
transaction retains the current document epoch, so initialization refuses while
it lives. Drop it before activation. Do not invent an activation-then-stale
preparation witness that the uniqueness guard prevents.

Activation always computes fresh current validation facts and pieces through
`validate_word_facts`, irrespective of the previous dormant operation.
Prepare all source facts, owner-vector capacity, projections, scope list,
route buckets, parent buckets and generation-index capacity before initialization.
Split private address/index preparation from generation insertion instead of
calling the current allocating `WordOwnerIndexes::prepare` afterward.
`initialize_scopes` is the last recoverably fallible step. Then fill the reserved
owner vector and maps using moved facts and Arc clones, inserting one unique
primitive generation key per owner within reserved capacity. No parse,
serialization, reserve or error return follows success. Publish both scopes,
owners, indexes and `pieces = Some(prepared_pieces)` together. No dormant
facts exist to clear. Rust allocator abort
behavior is unchanged and is not described as a recoverable guarantee.

Tests cover serial0, one and several positive dormant edits, unchanged serial
and origin on activation, distinct owner generations, zero/refusal, second
initialization, a held current epoch and a live preparation blocking activation,
then successful initialization after dropping that preparation. Existing
PowerPoint zero-serial initialization behavior remains unchanged.

### Exact activation ingress and borrow inventory

Activate under an exclusive PyDocument borrow at these ten current Document
factories before reading owners or capturing authority:

| Factory | Current shape and required transition |
|---|---|
| `paragraphs` and `tables` | Acquire a temporary borrow_mut, activate, drop it before constructing the collection with Py::new. |
| `section` | Change borrow to borrow_mut, retain its existing numeric bounds check before activation, then capture. |
| `pictures` | Change `&self` to `&mut self`, activate before current capture. |
| `stories` | Change `&self` to `&mut self`, activate before current owner reads. |
| `header_footer_variants` | Same mutable activation before capture. |
| `hyperlinks` | Same mutable activation before current owner access. |
| `revision_snapshots` | Same mutable activation before capture. |
| `comments` | Change slf.borrow to borrow_mut and build snapshots under that exclusive borrow. |
| `story_items` | Same mutable activation and snapshot construction. |

The combined first row represents two factories, for ten total. Existing
checked descendants remain immutable: Paragraph/Run/Table/Row/Cell bind_current,
section furniture, note paragraph/table factories and ParagraphStoryItem's
paragraph. Their parent handle or snapshot proves active authority. Do not
reacquire mutable activation while an immutable checked descendant is borrowed.
Keep `owned_story_snapshot` behind these guarded ingresses. Plain section values
and settings/core/app/custom proxies carry no checked Word authority and need
no activation just for value/proxy access.

Change `prepare_word_revision` to `&mut self` and ensure current authority before
reading old owners. This covers numeric-selector mutations and prepared returns.
`plan_body_splice` reads indexes earlier, so change it to accept `&mut self` and
Python, activating before those reads. Its six production callers are add table,
remove content, remove range, insert paragraph, insert table and append paragraph.
Each already holds a mutable document borrow. Prepared captures operate on
prepared authority and do not activate again. `adopt_property_edit` is reached
through checked Table/Row/Cell handles and asserts active authority. Other
counted/scoped mutations use the existing activated publication route only when
a positive result reaches publication. The dormant replacement route below is
the sole explicitly separate positive publisher.

Constructible public records require an additional early-owner boundary.
A caller can construct an unbound `Story` and `Hyperlink` without using any
factory. In `remove_hyperlink`, retain current native record resolution and
private candidate removal before activation, preserving those refusal points.
Then ensure authority BEFORE the existing `self.owners` lookup that finds the
selected physical hyperlink, not only in the later publication helper. Owned
records already prove active authority. This is an early-index exception like
plan_body_splice, not an eleventh current factory or a blanket mutable borrow
inside immutable resolvers.

The adjacent `set_hyperlink_url` remains a direct native property publisher.
Resolve and validate its record and run its existing native edit. It does not
activate authority or advance the binding serial. Dormant facts and reusable
pieces do not exist, so neither success nor refusal needs cache invalidation.
First later activation always validates the current native state. Active behavior
and held-handle lifetime remain unchanged. Do not assume that this direct native
publisher passes through prepare_word_revision.

Add a fresh-open, internal-anchor hyperlink fixture addressed only by plain
public Story/Hyperlink constructors, then remove it before any checked factory
access. Match existing successful text/style/package effects and invalid or
ambiguous record errors. Separately retarget an unbound link before first
activation and verify current URLs/relationship ids, saved XML and later
factory observations. These controls reject the concrete F9 missing-owner
failure and any reintroduction of retained dormant cache reuse after a direct
property publisher.

These are callable/factory pins, not assertions that historical line numbers
remain current. Implementation audits every current capture/owner-index ingress
and any new equivalent publisher against this inventory. Tests exercise first
section, story, comments and collection access, numeric mutations and checked
descendants, including borrow-error and preparation-refusal controls.

### Consuming dormant replacement and exact publication

Add this third published, doc-hidden Word0.17.0 facade method in the existing
Document implementation, with no new source file or public model type:

```rust
#[doc(hidden)]
pub fn try_replace_all_expected_owned(
    self,
    pairs: &[(&str, &str, Option<usize>)],
) -> Result<std::result::Result<Option<(Self, Vec<usize>)>, ReplacementCountMismatch>>
```

It consumes one already-private candidate. Capture its existing staged origin
using `take_staged_commit_origin`, iterate pairs in exactly the current ordered
batch semantics, and call existing `replace_batch` once per pair. Preserve first
mismatch, placeholder/count fields, source/relationship/resource errors and
signature/cache handling. On positive counts finish through the existing
`finish_staged_candidate` mechanism. All-zero returns `Ok(Ok(None))` and no
candidate is published. Empty input returns no candidate. This does not assert
that replace_batch itself performs a full reopen. It reuses its actual existing
flush/raw/typed behavior, with the consuming finish matching current staged
publication. Keep the existing mutating native APIs' zero behavior unchanged.
Do not blindly delegate them to a None result that changes their contract.

For a dormant Python call, clone native once and pass that candidate to the
consuming method. No second native clone is permitted. The no-expect singular
`try_replace_text` retains the exact existing upfront replacement-text XML
character validation and mapped error before its batch operation. The expect
and ordered-batch routes preserve their own current validation/error ordering,
not a newly imposed universal policy. The same pairs and counts must agree with
the existing CLI entry point, which remains unchanged.

`None` means every returned count is zero, so the binding returns0 or a tuple of
zeros of the original pair count. It leaves native state and serial unchanged,
with no dormant facts or pieces to publish. A mismatch/error also discards the
private candidate. Positive
results include all ordered counts, including zeros among positive entries.
Positive counts that later undo their own byte changes still commit once, as
the existing binding does. Do not use byte equality to erase that revision.

Before any positive dormant live publication, run the SAME existing fallible
immutable candidate inventory and pieces/fallback preparation, starting from
`WordPieces::default()` through `validate_word_facts`. No old pieces or dormant
facts are reused. This retains mutate-time errors as well as open-time errors.
Prepare the integer or PyTuple result and the empty-scope revision before commit,
then discard transient validation addresses, projections and pieces within that
same call. Commit the counter last. Install the native candidate with infallible
swaps while keeping live owners/indexes empty, both optional authority scopes
absent and both `pieces` and `pending_pieces` None. Do not put discarded candidate
pieces into `pending_pieces`. No prepare_word_revision or live owner/index
allocation runs in this dormant branch. A failure publishes no candidate or
counter change. Active authority retains its current checked retirement and
prepared-piece publication route.

The existing `install_inner` must preserve the optional-piece invariant. Active
publication consumes matching pending pieces into Some, or installs Some of the
existing empty WordPieces fallback when no pending cache is supplied. Dormant
publication must not turn that empty fallback into a retained dormant cache:
its pieces and pending pieces stay None. The local active cell path still makes
its one prepared replacement directly and does not call install_inner.

Both callers after `expected_replacements`, namely expect-bearing
`try_replace_text` and `replace_all`, must dispatch on dormant versus active
publication. A helper flag alone does not suppress their current later owner
inventories. The no-expect singular route also bypasses the redundant outer
counted_mutation clone only while dormant. All active behavior remains checked.

### Direct native publishers, serialization and subsequent activation

No dormant address/projection/piece cache survives any call. Consequently a
native publisher cannot leave such a cache stale, and no new native epoch API or
exhaustive invalidation mechanism is needed. Do not weaken the mandatory
validation on open or positive dormant replacement to obtain this property.

Explicit no-handle direct native publishers are `set_hyperlink_url`,
`replace_image(relationship_id, data)`,
`replace_image_for_story(story, relationship_id, data)` and
`set_picture_sizes_for_relationship(relationship_id, width, height)`. The story
record may be caller-constructed and unbound. The image wrappers at current
`crates/rdocx-py/src/document.rs:13277`, `:13289` and `:13429` do not enter
prepare_word_revision. Native image replacement and relationship-wide resize
stage privately, call prepare_and_reopen_staged and install only on success at
current `crates/rdocx/src/document.rs:31210` and `:36771`. Preserve their existing
input/error order, returned count, native atomic refusal and active held-handle
behavior. Dormant calls do not activate or advance the binding serial. Both
success and refusal leave the binding's dormant pieces/facts absent. First
later authority access recomputes current native facts, rather than assuming
these reopens left owners unchanged. This is a design repair for stable F9,
not a claim of an executed stale-cache regression.

Native save/to_bytes may install normalized owners. Retain each serializer's
existing behavior and error ordering, including native byte construction,
PyBytes allocation and external stream write. A stream or filesystem failure
must not be claimed to undo native normalization that may already have happened.
There are no dormant facts/pieces to invalidate before or after any of these
steps. Success, native failure, PyBytes allocation failure and external write
failure cannot expose a retained dormant cache. Saving itself neither activates
live authority nor advances its binding serial. Active authority retains its
existing serializer behavior and caches, without a new blanket invalidation.

First subsequent authority access ALWAYS reruns current native immutable
validation/piece construction, then initializes authority at the unchanged
serial. Repeated serialization and direct image/link publications followed by
first access must agree with native/fresh/reopened owners, including field-result
normalization, source namespaces, relationship targets and picture dimensions.
Later positive dormant replacement also validates its current candidate from
empty pieces, without resurrecting facts from before a serializer or direct
publisher. Zero/refused replacement invents no activation or revision. All
validation and transient deallocation costs remain inside actual timed calls,
and the issue's1.2threshold is unchanged.

## Consumer map

| Rule | Entry points and outputs | Provider and caller policy | Lifecycle stages compared |
|---|---|---|---|
| Local cell content | Python `Cell.text`, native `Cell::try_set_text` | `direct_body_table_cell_candidate` plus `apply_local_cell_edit` with Some(CellText), unchanged checked text semantics | Held/fresh reads, native model, saved XML, reopen |
| Local properties | `Cell.vertical_alignment`, `shading`, `width`, `set_border`, `set_margins` | Same candidate route with None publication, existing native conversions and validation | First property, repeated edit, held route, saved XML |
| Cell paragraph publication | `Paragraph.text` setter | `paragraph_route` plus the local cell route with Some(ParagraphText) for direct cell paragraphs, existing `edit_story_paragraph` otherwise | Held paragraph/run handles, fresh/native/reopened reads |
| Dormant replacement | `Document.open`, `try_replace_text` with both expect modes, `replace_all`, `save`, `to_bytes` | Consuming checked batch, exact eager validation with transient results discarded, empty-scope serial commit and no retained dormant facts/pieces | Open, zero/nonzero/refused replace, save, first handle |
| Active replacement | Same methods after a handle exists | Existing `prepare_word_revision` retirement and candidate installation | Held handles before and after, fresh handles, reopen |
| CLI comparison | `rdocx replace` with one pattern/value pair | `try_replace_all_expected` and `try_replace_text` share `replace_batch` and staged commit semantics | Same input and successful counted output |

## Behavior contracts and evidence

| Surface | Exact behavior | Primary source | Discriminator |
|---|---|---|---|
| `Cell.text` | Checked text, first direct paragraph only, existing anchors retained, run content replaced | `crates/rdocx/src/table.rs:2183` | Multiple paragraphs and anchored text versus complex fallback |
| `Cell.width` | Existing infallible `i64` EMU extraction and truncation toward zero by 635, negative widths stored | `crates/rdocx/src/table.rs:2306`, `crates/oxml-core/src/length.rs:72` | 634, 635, 636, -634 and -635 EMU plus both i32-twip wrap boundaries, native/getter/reopen agreement |
| Margins | Existing checked nonnegative lengths, atomic refusal | `crates/rdocx/src/table.rs:2370` | Negative later argument leaves all margins unchanged |
| First properties | `None` and `Some(empty)` both emit the first `w:tcPr` on assignment | `crates/rdocx-oxml/src/table.rs:2035`, `:2128` | Held first and later paragraph handles survive the index shift |
| Local publication | No visible partial mutation on refusal | HLD10 checked Word authority | Injected preparation failure then byte and handle equality, then later success |
| Dormant authority | Exact eager fallible validation retained and transient results discarded, no dormant facts/pieces, first activation always rebuilds current native facts and live authority | Issue 445 case 2 and current binding routes | No-handle workflow versus prior-handle workflow |
| Candidate route | No live mutation before every fallible native and binding step succeeds | New hidden methods above | Injected `commit_scopes` refusal leaves bytes and handles identical |

## Dependency readiness

There are no formal providers. F-X283 is resource-serial after F-X282 because
they share facade, binding and HLD files. It does not rely on an unapproved
F-X282 bridge and uses no F282 helper. Clean Base, approved complete design and
normal claim are still required. No claim or implementation is authorized by
this draft.

## Acceptance criteria

| ID | Input or setup | Expected observation | Consumers | Local test | Final evidence |
|---|---|---|---|---|---|
| A283-1 | Issue case 1 `Cell.text` | Distinct PASS under the exact timing contract below | Installed Python | Text correctness and actual traversal counts | Linux release wheel case 1 |
| A283-2 | Case 1 `vertical_alignment` | Distinct PASS under the same timing contract | Installed Python | First/repeated properties and traversal counts | Linux release wheel case 1 |
| A283-3 | Case 1 `shading` | Distinct PASS under the same timing contract | Installed Python | Fill value and traversal counts | Linux release wheel case 1 |
| A283-4 | Issue case 2 | 1000 replacements and Python total at most 1.2 times CLI | Same-build wheel and CLI | Dormant/active equivalence and counted output | Exact Linux release pair |
| A283-5 | Cell paragraph text assignment | Setter publication avoids unrelated body traversal, no invented time threshold | Python `Paragraph.text` setter | Held/fresh/native/reopened text | Installed regression |
| A283-6 | `Cell.width` | Existing conversion and negative behavior preserved | Python and native | Boundary values above | Installed regression |
| A283-7 | `Cell.set_border` | Existing edge/style/color/size validation and other-edge preservation | Python and native | One valid edge plus invalid atomic assignment | Installed regression |
| A283-8 | `Cell.set_margins` | Existing units and nonnegative checks, no partial update | Python and native | Four distinct margins plus invalid last argument | Installed regression |
| A283-9 | Held cell/paragraph/run owners | Existing identity and retirement rules, correct shifted routes | Held and fresh Python handles | Empty/absent/existing tcPr, adjacent cells | Installed regression |
| A283-10 | Refused edit or preparation | XML, package, revisions and held observations unchanged | Native and Python | Failure followed by successful edit | Installed regression |
| A283-11 | Comments, notes and revisions | Existing checked operation semantics, explicit XML and part observations | Local path and fallback | Anchors plus note/revision-bearing first paragraph | Final schema and native riders |
| A283-12 | Unknown XML and properties | Preserve unmodified opaque bytes and schema child order | Native, Python serialization | tcPr first, unknown siblings, reopened package | Schema, hash and native riders |
| A283-13 | Dormant then active authority | Mandatory open checks, lazy activation, correct active invalidation and cache state | Open/replace/save and live handles | Zero/nonzero/refused edits before and after activation | Installed regression and case 2 |

Case 1 is exactly bodies 0, 500 and 2000, followed by one 20 by 5 table.
Acquire the 100 cell handles once before the setter loops and reuse that same
list. Time `Cell.text` first, then `Cell.vertical_alignment`, then `Cell.shading`,
matching the issue script order. Each setter runs once per cell, sequentially,
with no invented warmup. Each of the three 2000-body costs must be below 1.0 ms
per call and at most `1.5 * empty_body_cost + 0.05 ms`.

Case 2 is exactly 4000 paragraphs with bold prefix, normal middle and italic
end. Every fourth middle contains water, giving 1000 replacements. Python
timing starts inside the running process before `Document.open` and includes
replace and save. CLI timing includes subprocess start. Three paired trials
select the lexicographic minimum Python tuple `(total, opened, replaced, saved)`
and the minimum CLI command time. Never combine independent stage minima.
Final qualification uses an actual Linux x86_64 release wheel and matching CLI.

## Test matrix

| Case | Combined states | Expected agreement | Wrong implementation rejected |
|---|---|---|---|
| C1 | Empty, 500 and 2000 body paragraphs, same table | Successful correct edits with bounded unrelated work | Native-only optimization or hidden whole-body binding scan |
| C2 | `None`, `Some(empty)`, nonempty tcPr, two held paragraphs | Correct route shift once, no stale owner | Presence-only first-property predicate |
| C3 | Text edit, retained later paragraph, held runs | Correct text and causal invalidation | Rebuilding all cell paragraphs or retaining removed run facts |
| C4 | Direct body, nested table, controlled and related owner | Same supported checked semantics | Treating a containing table as a direct body slot |
| C5 | Anchors, note reference, tracked content, opaque siblings | Existing operation's exact XML and package outcome | Fictional note/revision preservation or skipped reconciliation |
| C6 | Dormant zero/nonzero/refused mutation, then first handle | No unnecessary activation, no retained dormant facts/pieces, current facts always rebuilt at first handle | Skipped install, deferred mandatory validation or resurrected cached piece |
| C7 | Handle acquired before replacement | Correct scope retirement and fresh values | Dormant fast path used with live authority |
| C8 | Invalid color/border/margin and preparation failure | No publication, later valid edit succeeds | Live edit before fallible index or revision update |
| C9 | `commit_scopes` refusal injected after native candidate success | Bytes, owners and handles identical, later edit succeeds | Native swap before revision commit |
| C10 | Property edit then text edit on the same cell | Property keeps handles, text preserves cell owner and retires the existing CellText descendant/content scopes | One retirement policy applied to both |
| C11 | Zero, one and several dormant commits, then activation, live preparation and held epoch | Same serial/origin at activation, distinct owner generations, initialization refusal leaves counter unchanged | Resetting serial, invented mutation or activating through an outstanding epoch clone |
| C12 | First section/story/comments/collection factory and numeric mutation, then checked descendants | Exact guarded mutable ingress without nested borrow failure | Blanket activation in immutable descendant binders or an unguarded direct factory |
| C13 | Missing external hyperlink relationship and valid/internal-anchor controls | Current native and binding behavior retained at open and candidate publication | Assuming native open subsumes the binding's fallible inventory |
| C14 | Dormant save/to_bytes, fields and namespaces, later first handle, repeated serialization, PyBytes or stream failure | No dormant facts/pieces survive any outcome, first access rebuilds current native/reopened owners, no new binding serial for save | Retaining pre-normalization facts, inventing native rollback after external failure or allocating authority solely to save |
| C15 | Fresh dormant document plus caller-constructed unbound Story/Hyperlink, remove and URL-retarget controls, invalid/ambiguous records | Removal activates before its physical owner lookup, retarget preserves serial with no dormant facts/pieces, first later authority always rebuilds, existing errors retained | Assuming every supported link record came through a checked factory or retaining dormant facts across native publication |

| C16 | Fresh dormant document, replace_image, replace_image_for_story using an unbound Story, and relationship-wide picture resize, each success/refusal then first handle, normalization-sensitive fields and namespaces | Native operation/error/count and active lifetimes unchanged, no dormant facts/pieces retained, first activation uses current native/fresh/reopened owners and image targets/dimensions | Treating direct native reopens as prepare_word_revision publishers or reusing pre-publication validation facts |

C17 extends A283-9 with Cell.text versus Paragraph.text versus property publication,
held selected paragraph/run and later paragraph/run handles, and cell/table property
handles, across absent/empty/present tcPr. Compare each exact existing checked
publisher, current values and saved/reopened facts. A blanket CellText retirement
for Paragraph.text must fail this discriminator. Public signatures stay unchanged.

C18 extends A283-9 with a clean selected cell beside enclosing and nested
row0/cell0 neighbors. Hold their paragraphs, runs, Font and paragraph-property
views with distinct values through property, CellText and ParagraphText edits.
The existing checked path must preserve physical identity and exact retirement.
Native/reopened bytes alone must not pass exchanged held authority.

## Rejected alternatives

- Whole-document cloning and full owner rebuilding retain the reported scaling defect.
- Mutate then repair indexes violates atomic publication.
- A public properties-presence method answers the wrong serialization question.
- Returning a cloned cell or table alone does not prove binding publication atomicity.
- A new trait or generic solely to forward the edit closure increases places to look.
- A native-only timing result does not prove the Python issue.
- Donor timings, debug counters and reduced workloads do not replace final acceptance.
- Retaining dormant validation facts/pieces creates unnecessary source-custody and invalidation obligations across direct native publishers. Transient validation plus unconditional first-activation rebuild removes that cache risk.

## Test plan

**Test gate**: regression.

Local RED/GREEN cases must cover each A-ID using existing test entrypoints.
Record actual old-source failures separately from compile failures. Use scoped
checks while iterating. Instrument actual traversed work, not a constant entry
counter. Pin the release artifact and build before timing. No timing or test
has passed for this draft.

Final integrated checks include installed Python 3.9 and 3.12, strict typing
and stubs, exact Linux performance workloads, schema and applicable native
controls, hash 49 unchanged, and the retained F-X268 union riders: 79 flat calls, 252-call catalogue,
original Issue407 gate, five linear calls, counted work, memory and the
applicable PowerPoint248 rider. These are retained union obligations, not
new Issue445 timing criteria. Run
the shared full gate once over the integrated union. All remain pending.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`
- `docs/hld/15-build-and-toolchain.md`

## Risk routing

Performance budget: authenticated release artifact, exact sampling and work
controls. PyO3 bindings: installed runtime, typing/stubs and final WASM check.
Cached read projections and shared semantics: held/fresh/native/reopened
comparisons after fact-changing edits, including unchanged-scope publication.
XML creation and relocation: tcPr child order, preserved unknown bytes and
schema checks. Compatibility and units: retain field-specific existing
conversions. Contributed acceptance: preserve the pinned issue workload and all
thirteen A-IDs. Public hidden bridges: all three additive Word0.17.0 methods enter SemVer, archive and publication dry-run review. The existing unpublished support initializer amendment has zero-serial PowerPoint compatibility controls.

## Hash harness

Expected unchanged. Hash 49 and saved-byte controls remain final union gates.
No baseline movement is authorized.

## Implementation checklist

- [x] Add `direct_body_table_cell_candidate` and `commit_direct_body_table` with the stated signatures.
- [x] Add the private `apply_local_cell_edit` preparation and commit ordering.
- [x] Add private `cells_by_parent` and `runs_by_parent` indexes and maintain all full/local index publishers.
- [x] Establish exact-source namespace certificates and source-change rejection controls.
- [x] Implement the specified initializer amendment, exact ingress guards and consuming replacement publisher with their discriminating controls.
- [x] Route `Cell.text`, the property setters and `Paragraph.text` in a direct body cell through the local path.
- [x] Make `Paragraph.text` resolve its publication route through `paragraph_route`.
- [x] Retain exact eager fallible validation at open and positive dormant publication, discard transient results within each call, retain no dormant facts/pieces and implement each enumerated ingress.
- [x] Reproduce the historical missing-link witness against fresh native/Python artifacts, retaining current open errors and valid controls.
- [x] Implement all thirteen A-IDs with discriminating local evidence, unconditional first-activation rebuild, no retained dormant caches, all named direct publishers and serializer/error outcomes.
- [x] Update exactly the HLD impact list before implementation freeze.
- [x] Independent implementation microscope, scoped gate and structured handoff.
- [ ] Final integrated and release artifact qualification.


## Contract closure self-review

- Proposed callable route: local candidate, binding preparation and nonfallible
  native commit are specified. This is not source compilation evidence. `Cell.text`,
  the five property setters and `Paragraph.text` in a direct body cell all
  reach `direct_body_table_cell_candidate` and `commit_direct_body_table`.
- Supported kinds: `BodyContent::Table` direct body slots qualify. Nested
  tables, content controls and related stories are excluded and fall back. The
  exclusion is explicit, not an umbrella statement.
- Signatures: the three hidden methods and the private binding entry point are
  stated once and referenced, with no divergent copies. `Cell.width` remains
  infallible and negative values remain stored. Non-finite values are not `i64`.
- Imported states: absent and empty `tcPr` both qualify. Opaque `extra_xml`
  boundaries and unknown siblings are preserved because the candidate is a
  clone of the parsed model.
- Source pins: every claim above is a current-source pin at
  `5820539cd277b417ae060e15010eb73bd1d00b01` with the preserved F272 dirty work.
  This is a design read, never a clean implementation Base.
- Performance: the timed operation is one public Python setter call on a reused
  handle list. Setup, handle acquisition and the first authority activation are
  outside the timed loop. Work-bound controls count visited owners and body
  children, not a constant entry counter.
- Cached read projection: source facts are the selected cell's paragraphs,
  runs and table row counts. The preparation boundary is
  `apply_local_cell_edit`. Publishers are the local cell path and the existing
  active-authority publication. Held, fresh, native and save/reopen
  observations map to C2, C3, C7 and C10.
- Mechanisms: Q-F283-01 preserves exact validation, Q-F283-02 names the three additive methods, and M283-D1through3 now have explicit source amendments and publishers. Status stays draft pending the separate fresh Root final2 after the simplification repair. Three actual design peers remain used, not reset. No unperformed implementation or timing pass is claimed.

## Open questions

| ID | Affected F-IDs and A-IDs | Decision and recommendation | Status and answer/source |
|---|---|---|---|
| Q-F283-01 | A283-13, dormant authority | Retain exact eager fallible validation and discard transient immutable facts/pieces while deferring live authority. First activation always rebuilds from current native state. No open error is moved to first handle. Historical missing-link witness is evidence for the distinction, with fresh native/Python reproduction pending. | Closed by Root compatibility-preserving mechanism choice. The unchanged1.2timing gate remains deciding. |
| Q-F283-02 | A283-1 to A283-4, A283-13 | The two local-cell methods plus consuming `try_replace_all_expected_owned` are three additive doc-hidden published Word0.17.0 methods. Amend the existing unpublished support initializer without adding a fourth facade method. | Technical surface authorized by Root for design review, with strict SemVer/archive inventory and final design approval still required. |


## Root author adjudication of design peer 2

The actual peer pass remains used. F1 to F8, F10 to F12, F14, F16 and F17
identify real callable, bounded-work, atomicity or proof gaps. This revision
specifies singleton-body validation, distinct ordinals, selected-table child
indexes and deltas, one prepared piece replacement after revision commit,
source-bound unsavable proof, explicit fallbacks and separate text/property
state policy. These are proposed mechanisms, not implemented or validated.

F9's borrow problem is real, but its 34 raw captures include prepared/internal
sites and are not an activation count. M283-D2 preserves the exact remaining
inventory work. F13 is downgraded to a wording clarification: the prior plan
never explicitly included the document scope in `retired`, so no duplicate
call was established. This revision explicitly excludes it. F15's provenance
is corrected: F283 registration at HLD14:8876 does not name 252/407. They remain
required through F268's final union rider at HLD14:8681 and are now named here.

S1 is addressed by the later consuming bridge mechanism below. No implementation is claimed. S2 requires release-bound evidence distinct from debug
physical-route diagnostics. S3/S4 take checked fallback for unavailable caches
or routes. S5 preserves caller-specific existing errors, including the actual
property wording `cell index out of range`. S6 adds signed i32-twip wrap
controls. N1 uses double dereference, and N2 retains existing invalid-slot error
families. No review record is rewritten and no finding is recorded as tested.

The above author checkpoint was historical. The subsequent Root mechanism revision below settles those design choices with explicit amendments, retained validation and actual ingress/publisher inventories. Final design review and actual implementation evidence remain required.
Selected-table disappearing owner records explicitly select checked fallback,
so this proposal does not require an inactive-slot representation. No source implementation is authorized by
this draft and no acceptance row is removed.


## Root mechanism closure revision, 2026-11-02

Historical mechanism checkpoint, superseded by the operative simplification below. Root retained all thirteen acceptance rows and original exact workloads. M283-D1through3 were concrete implementation proposals, not implemented code. The then-available design peer3 has since run, and all three design peers remain used. Three implementation peers remain separately available. Status remains draft and F-X283 stays pending/unclaimed.

The missing-hyperlink probe used authenticated historical rlibs, not current source, and provides a candidate error-timing witness rather than a current-source pass. Its fixture SHA256 is32c7a5877b96fe4a714bb2762cda441d92e3ad02eda698010974b92e6b9e0aec. Actual Issue445 body at the pinned hash leaves owner strategy to the implementer and retains the1.2total-time criterion. No immutable validation work is removed from timed calls and the CLI denominator is unchanged.

The historical checkpoint proposed conservative dormant serializer cache invalidation because native save/to_bytes may normalize owners. The operative simplification removes that retained dormant cache entirely, replacing conditional invalidation/rebuild with unconditional first-activation validation. No stale-cache execution is claimed. Fresh discriminating tests remain mandatory. Source pins in earlier sections remain their explicitly historical producers. Current implementation Base must be revalidated after the F-X282 resource predecessor hands back. No source, threshold, gate, baseline or release status changes in this design-only revision.


## F9 targeted repair after actual design peer3

Codex is the actual third design peer after Grok1 and DeepSeek2. Its report finds one high defect, stable F9, and no separate smell or nit. All three design peers are now used. Root's targeted author repair above adds the unbound hyperlink removal boundary and the adjacent direct URL publisher policy. It does not claim an executed failure against unimplemented code. The previously generic ingress inventory was incomplete despite its ten named factories. No fourth peer is authorized or counter reset. Root final design review remains due before approval, and all actual implementation/performance evidence is still pending.


## Root simplification after final design review1

Root-delegated final1 retained stable F9 for three direct image publishers that
reopen native state without passing through prepare_word_revision. This separate
author repair removes all retained dormant address/projection facts and reusable
pieces rather than adding another invalidation inventory. Exact eager fallible
validation still runs at open and positive dormant publication, with transient
allocation and deallocation inside the actual timed calls. First authority
activation always recomputes the current native facts and pieces. Local active
setter pieces, the three published hidden bridges and the initializer amendment
remain unchanged. Direct URL/image publishers and serializers preserve existing
errors, native state and binding serial semantics without any dormant cache to
invalidate. No implementation, actual performance result or final acceptance is
claimed. All thirteen A-IDs and original workloads/gates remain. Three design
peers remain used, with no fourth peer or reset. Status is draft, and fresh
Root final2 follows actual adoption.

## Root implementation contract clarification, paragraph retirement

The implementation author identified a contradiction missed by the design peers
and Root finals. A shared retire boolean could not preserve the distinct existing
CellText and ParagraphText handle policies. Root authorizes the private helper's
Option<WordPublication> parameter and exact selected-owner checks above. This
repairs the mechanism for existing A283-9, without changing public signatures,
the three bridge inventory, thresholds, provider edges or final gates. Design and
implementation counters do not reset. Actual old/current handle observations and
independent implementation review remain required.

## Scoped instrumentation applicability

Root authorizes a debug-only underscored _word_piece_alignments diagnostic in
the existing binding module for the named F-X283 structural tests. It counts
actual aligned_block comparisons and keeps _word_piece_work owner-read semantics
unchanged. Release __all__, public interfaces and the three native facade bridges
remain unchanged. Earlier Base and intermediate artifacts counted alignment in
_word_piece_work. Their exact hashes and instrumentation revision remain in
evidence, not relabeled as the final counter revision. Combined comparison/read
counts are debug structural observations, never Linux release timing evidence.

## Implementation peer1 D1 remediation

Peer1 source-traced an unproven table-wide join between nested and enclosing
cell paragraphs. Actual installed execution on frozen author artifact d9884b17
fails all four new operation cases with exchanged held paragraph text while
saved bytes remain correct. The actual Base artifact passes the same four
cases. The repair keeps the existing checked publisher for any selected-table
facts containing more than the single root table. It does not change the
common relocation key, public API, original workloads, criteria or counters.
This is author remediation, not an independent finding disposition. Peer2
decides D1 closure. Implementation peers used1/3 remain used.

C18 boundary controls also execute two same-depth controls, field-bearing
physical paragraphs and textbox paragraphs under one physical table. Base,
frozen author and repaired artifacts each pass all12 controls. No additional
RED is claimed. Same-cell controlled and field-bearing physical paragraphs
share a cell story but have distinct per-story projection ordinals in
word_piece_owner_addresses. The local key includes that ordinal. Textbox
owners select the existing nonzero text_boxes checked fallback. These controls
do not broaden the local topology or change supported input behavior.

The current cfg(test) commit-fault binding-object diagnostic is executed on the
repaired source at Root's explicit request. It is not an installed-wheel
injection and no new fault hook is exported. Installed ordinary refusal and
later-success observations remain separate. Peer2 and Root assess evidence
coverage without treating earlier or compiled-only diagnostics as current.


## Peer2 D1 complete membership amendment

The same physical-authority invariant remains D1. Native physical cell inventory
counts controlled and empty cells even when binding direct coordinates omit them.
Each immutable piece now caches complete direct-table membership: one physical
root table, every native physical cell represented once in contiguous source order,
qualified table/part identity, unique valid direct coordinates and every paragraph
story belonging to that cell range. Local publication requires this certificate
before revision preparation and requires it again for candidate facts. Ambiguous
membership uses the existing checked publisher. All three constructors, rebuilds,
splices and local piece replacement retain this immutable-fact custody. Setter
qualification remains O(1). All13 A-IDs, numeric limits and counters remain unchanged.

C19 extends A283-9 with controlled-row and controlled-cell neighbors before/after
the selected direct cell, both text publishers and property controls. Assert
physical story/XML routes, held/fresh views, the actual checked retirement policy
and follow-up edits/save/reopen. Actual corrected Base16PASS and frozen8FAIL/8PASS
precede production edits. An initial fixture incorrectly expected controlled held
paragraphs to survive checked text publication. Its separate observations are
retained as fixture history, not a project regression. Implementation peers2/3
remain used. Current C9 diagnostic and final affected evidence follow separately.


Peer2 remediation evidence: repaired installed49PASS, including C1916PASS,
original affected21PASS and retained paragraph-ancestry12PASS. Actual native
membership diagnostic1PASS and current C9 binding-object diagnostic1PASS have
complete input and executable hashes. Installed fault injection stays unperformed.
Scoped check, relevant all-feature Clippy, format and affected prose pass on the
frozen source. Source/artifact evidence is under D1-membership. No release timing,
provider, completion or clean-review claim. Peer3 and Root final remain required.


## Local integration preparation

Checkpoint: sprint-final-acceptance. The three actual implementation peers are
consumed. Root final .claude/reviews/F-X283-codex-final-1.md records0 defects and
0 smells on the frozen product source. /verify --scoped F-X283 binds applicable
local evidence in prepare/VERIFY_SCOPED.json under the existing tmp evidence root.
The structured handoff carries the reviewed source commit as Head and is committed
separately afterward. No crates change follows that source commit. This mechanical
metadata update changes no product, test or HLD intent. The canonical private
Option<WordPublication> policy at a995352b is retained exactly. All13 acceptance
rows and numeric workloads remain. Design peers3/3 and implementation peers3/3
remain exhausted. No new peer phase or provider acceptance is inferred.

Plan stays approved and final qualification stays unticked. Original exact Linux
release case1/case2 timings, installed3.9/3.12 typing/stubs, native/schema, hash49,
F26879/252/407/five-linear/count/memory/PowerPoint248, WASM, full integrated union,
docs, archives/publication dry-run, hosted CI and reviewed-SHA publication approval
remain due. Installed fault injection remains unperformed and separate from the
actual current C9 binding-object diagnostic. No delivery completion is claimed.


## Approved final benchmark execution route

The exact A283-1 through A283-4 workloads and thresholds remain unchanged.
The reviewed product implementation is integrated at
514a076e592517ef128e555f52cef1085fa702de. The resumed worker preserves original
Base 101c34c4a6d5fa62d9ba0e0ff745fd44c9aeaf13 and all prior evidence. No
product dependency, public API, review counter or performance threshold changes.

Add scripts/benchmark_f283.py from the reviewed runner, SHA-256
9d944b03b199e16e58e4d8b404f319c06b5722e9bf7933b9f382fe0ce6d07acc.
Root asked under AGENTS.md's new-file rule and the user instructed continuation.
The runner preserves the pinned workloads, checks input geometry and requires
authenticated final source, release artifacts, build configuration and interpreter
identities. Its 19 synthetic checks are tool evidence, not product timing.

The operative route after Codex final review2 is an optional benchmark job
alongside the normal final CI invocation in the registered ci.yml. Preserve all
existing normal and owner-only selectors. Keep CI gate unconditional with
always(). Add the benchmark job to its needs and check its result: requested
benchmarks must succeed, while an unrequested benchmark job must be skipped.
The final ordinary manual CI invocation can request both normal validation and
the exact performance workloads once. Never emit a successful required context
merely because a benchmark-only route skipped normal validation.

There is no benchmark-only mode and no new normal-SemVer authority exclusion.
Restore the existing controller exactly. The new job retains its owner/source
admission, read-only permissions, pinned second-phase release artifacts and
conflicting-owner-mode refusal. Missing producer or interpreter proof, oversized
transfer, failed timing or a missing workload remains failed or pending. No
build, signing, publication or tag action belongs to the benchmark job.

Update the existing workflow tests for optional benchmark selection and the
normal gate's requested-success versus unrequested-skip behavior. Preserve the
existing owner, aggregate, signing and release checks. Local scope is the
runner, deciding gate and affected workflow-policy controls. Full source-qualified
CI, exact final Linux timings and all original union riders remain pending.

HLD15 describes the implemented optional job and normal gate contract before
freeze. Prior HLD03, HLD10 and HLD12 product updates remain intact with no new
lease in this phase. Risk-routing's performance, new-file, release-scripting,
shared-consumer and evidence-adoption obligations remain. No source push or
dispatch is authorized by this revision alone. Use the reviewed source-availability
checkpoint after its prerequisites pass. This technical revision supersedes the
benchmark-only proposal rejected in .claude/reviews/F-X283-codex-final-2.md.


## A283-14, one-story dispatch preparation trial

Root assigns only this bounded implementation to GLM-5.3 Flash through OpenCode,
requesting high effort. All other implementation stays with Codex. Original
claim Base101c34 and prior reviewed product integration514a remain. This is an
internal callable preparation step for the existing final benchmark route.

Add --prepare-dispatch to the existing benchmark script. It takes --transfer,
--transfer-sha256, --source-ref, --workflow-sha and --output. Output is an absolute
new JSON file for a GitHub workflow dispatch request, never a network operation.
Preparation works with the standard library on the current Mac and never imports
rdocx, calls authenticate, starts benchmarks, loads a CLI or contacts any service.

T1. Read the transfer as UTF-8 bytes, authenticate its exact supplied lowercase
SHA-256, require schema f283-root-transfer-v1 and preserve its original text
exactly in inputs.f283_approval_bundle. Do not normalize or rewrite it.

T2. Before output, validate the already existing CI transfer contract: positive
integer producer/artifact IDs received as numbers or decimal strings, exact
proof_json SHA and schema, 40-character
lowercase source_commit, accepted build-mode attestation, and SHA-256 identities
of source_inputs_json, wheel_build_receipt_json and cli_build_receipt_json against
the nested proof. These are received approvals, never approvals made by this tool.
Reject bool IDs, incomplete records, malformed JSON and inconsistent identities.
Do not read Linux artifact paths or infer actual build acceptance from file names.

T3. Require a nonempty refs/heads/ branch source-ref and lowercase 40-character
workflow-sha. Build ref from that branch name and all eight existing CI inputs:
owner_compatibility=false, source_sha from the authenticated proof, source_ref,
workflow_sha, runner_label=ubuntu-24.04, config_sha from the authenticated transfer,
f283_benchmark=true and f283_approval_bundle with the exact original text. Use
JSON booleans for the two typed boolean inputs. Do not add workflow input keys.

T4. Serialize the complete request as ASCII JSON. Admit at most65,535bytes,
including ref, inputs, formatting and final newline, before creating output.
This intentionally conservative complete-request rule fits the documented input
payload bound. Test exactly at and just above the bound using synthetic data.
No truncation, dropped source identities or shortened approval data is permitted.

T5. Refuse overwrite and leave no requested output on validation failure. Reuse
the existing atomic writer where its byte serialization matches the size check.
Print only a concise output path, byte count and digest, not approval contents.

T6. Reject incompatible execution/self-test/case options. Original benchmark
execution, --self-test and all pinned builder/timer AST behavior stay unchanged.
Run the existing19synthetic checks and affected existing workflow tests, adding
finite positive/negative command tests to the existing test entrypoint only.
No product test, build, timing, package install, GitHub dispatch or new test file.

T7. Update HLD15's optional benchmark preparation intent and the existing progress
record. Do not touch other HLD, CI workflow, controller, trackers, manifests,
baselines or product files. No new module, dependency, trait or generic framework.

T8. Self-review once, preserve actual commands/exits, and stop with a concise
handoff and pending evidence. Do not commit, push, tag or publish. Root reviews,
requests up to three focused repair cycles if needed, then takes all work back.
The previous runner9d944b03 predates this extension. Its workload semantics stay
pinned, while the new runner identity awaits Root review and final qualification.

## Hosted benchmark interpreter observation acquisition

Root approves this ordinary execution-preparation clarification. The existing
manual rdocx Linux GNU CLI rehearsal selects the benchmark job's identical
pinned CPython3.12.13 setup after all CLI tests, then logs platform, executable
SHA256, version and actual source root/commit. Existing builds and tests keep
their original interpreter and configuration. The retained job log supplies an
observation for the separately reviewed transfer. It emits no approval or
acceptance boolean. The later benchmark still checks its actual runtime digest.
No new job, public contract, threshold, authority or review-cycle reset is added.
HLD15 records the acquisition method. Hosted execution and actual transfer
approval remain pending, with all prior final gates unchanged.


## H4 reviewed source availability renewal, 2026-11-06

Root prepares a successor to actual failed H3 CI42306259639. Bounded repairs
pass24unique focused Word controls by applicable21+3runs, strict all-targets/
all-features Clippy, formatting and two default-feature cache controls. Actual
new Native outputs pass50SDK cases, and four actual sole-ReviewerB Word
Accept All/Reject All saved results pass SDK and semantic checks. Six passive
Word copies and all50originals remain unchanged, with all15prior generated
windows retained. Root repair audit and affected archive support5 have zero
defects and smells. F271A7 separately prepares the exact same fresh public
consumer before its unchanged300second run, with15passing harness self-tests.

Actual H3 hosted PowerPoint passed1041tests with zero skips. Its unchanged
local release artifacts have1039applicable cases plus two debug-inapplicable
cases per runtime, with actual hosted coverage of both counters. Earlier full
commands remain failed, and source-qualified composite applicability is
retained. Word's old bytecode-containing wheel authentication failures remain
failed. Ignored bytecode is quarantined and fresh clean release packaging,
both runtime suites and typing remain pending.

The one affected Word archive was measured and independently authenticated.
Final native/performance/source-diagnostic/full riders, current normal CI,
verified publication dry-run, current CLI hash/pixel checks and build-only
26payload rehearsal remain pending. Source availability permits only the
reviewed clean sprint/s96 ref push and normal CI routing. No ledger completion,
dependency unblocking, main merge, sprint/release tag or publication follows.
Actual upcoming SHA, push and CI run identity will be appended after observation.


## H5 reviewed source availability renewal, 2026-11-07

The successor to b26837a9 repairs the actual16H4 integration failures. The
current guarded native-host batch passes17cases, including the unchanged
PDF/SVG logical-order regression and a stored-field control. Bounded independent
review and Root final pass2 have zero findings. Full workspace Clippy and
Rustdoc include both PyO3 binding crates and actually pass with warnings denied.
The two affected Word archives have actual authenticated measurements and
post-adoption packages within the existing64byte compression tolerance, with
correct per-page Nov7dates. Other package records remain unchanged.

The canonical orchestrator renews only the previously consumed source-
availability checkpoint. Actual clean sprint ref, source SHA, workflow bytes,
remote ancestry, push and normal CI request/run identities are recorded after
observation in the existing S96run state and progress. Both optional CI flags
remain false. Only the reviewed sprint ref may be pushed without force.

H4CI42315170211 remains failed at Test/MSRV, despite actual public conformance,
Python, hash, Docs and six SemVer producer/signer successes. Current source
normal CI, fresh Word release wheel, remaining full/risk/native/performance
acceptance, verified publish dry run and build-only26payload rehearsal remain
due. No completion ledger, dependency unblock, main merge, tag, reporter
comment, yank or publication is authorized by this source checkpoint.


## H6 reviewed source availability renewal, 2026-11-07

The successor to a5765aa0 repairs the actual70H5 integration failures. The
source-stable host batch passes73regressions, including all70original cases,
plus18Native controls and18lower semantic and normalization controls. Root
final repair passes3and4 and the retained bounded independent reviews record their actual
scope and source. Full workspace Clippy and Rustdoc actually pass, with warnings
denied for Clippy. The explicit F272 amendment adds one doc-hidden lower current
bookmark projection accessor in the prepared minor. HLD04 states its current
versus preserved-source distinction. Actual four-package archive measurements and post-adoption checks pass. A later
rdocx-only cache repair has a genuine measured successor and a pending
post-adoption check before push.

The canonical orchestrator renews only the previously consumed source-
availability checkpoint. Actual clean sprint ref, source SHA, workflow bytes,
remote ancestry, local evidence and normal CI request/run identities are
recorded after observation in the existing S96run state and progress. Both
optional CI flags remain false. Only the reviewed sprint ref may be pushed
without force. Historical consumed integration a13b1292 remains a historical
fact, not current source qualification or a fabricated new worker handoff.

H5CI42322974698 remains failed at Test, MSRV and CI gate. Its successful jobs
are not a green current-source CI run. The initial H6policy run retained7actual
environment errors. The four affected policy methods actually pass under the required normal
process and download access. An initial hash failure exposed three redundant
nested namespace declarations after authored style publication. The private
cache now retains an actual emitted-model witness separately from its genuine
parsed source proof. All16focused style controls pass, and all49hashes match
the unchanged baseline. Full workspace Clippy and Rustdoc pass again on that
actual source. The exact original failed receipts remain. Fresh-document cache
parsing is actual counted work, not a model clone labeled as a parse.

Current-source normal CI, fresh installed Word artifact and full functional
qualification, remaining complete/risk/native/performance acceptance, exact
source-diagnostic applicability, verified publish dry run and the subsequent
build-only26payload rehearsal remain due. The permanent named native exception
remains unproved. No completion ledger, dependency unblock, main merge, tag,
reporter comment, yank or publication is authorized by this source checkpoint.


## H7 CLI fixture source availability renewal

The H6 normal run42335559891fails Test and MSRV on the same17CLI integration
cases after the Word library/integration/regression suites pass. Its later
Test goldens and final compatibility decider are skipped, not passed.

The successor changes only CLI tests and affected archive metadata. Production
code, version carriers, public interfaces, parser/renderer/CLI behavior, all
thresholds and the hash baseline are unchanged. Actual current macOS CLI
integration passes116cases, with the one existing Linux-only case retained for
hosted qualification. Source-qualified XML positives, preserved invalid
negatives, exact diagnostics/JSON and corrected valid PNG retain their intended
assertions. Root bounded repair pass6 reviews these exact changes.

The consumed source-availability checkpoint is renewed for the reviewed clean
canonical sprint ref and normal CI only. Record the actual SHA/tree, origin,
ancestry, actor, request/run and outcomes after observation. H6failed CI is not
current green evidence. Current product artifacts retain their actual4cf
producer identities and require explicit nonproduct applicability, never a
fabricated rebuild. Other21archive-source inputs remain unchanged, with the
updated CLI archive measured and post-adoption checked. All final native,
performance, hosted, verification/review and close debt remains pending. No
completion ledger, dependency unblock, main merge, tag, registry publication,
yank or reporter comment is authorized by this checkpoint.


## H8 integrated repair source availability renewal

The H7 run42341613212 passes the117CLI cases and its Word library, integration
and regression suites. Its MSRV job fails15later layout cases. Historical
producer, signer, Python and fidelity successes keep their exact H7 identity.
The failed whole job and any unperformed final decider remain unqualified.

The current repair restores stored Field display hyphenation policy, neutral
pre-laid body wrapping compatibility, instruction events before stored Field
cache display, and checked Engine preparation of deferred revision Fields.
Actual producer paint and multilingual metadata replace stale synthetic
fixture assumptions. The complete current layout library passes419cases with
zero failures and one unchanged ignored case. All49hashes match the unchanged
baseline, and full workspace Clippy passes with warnings denied.

The exact original Issue437 counter workload also exposed unnecessary story
refreshes in both run-font and run-collection view capture. The repaired
current-owner capture records zero body scans in all11batches. Both exact
new-document and reopened lifecycle regressions pass. Instrumented clocks do
not establish the ordinary release ratios. Earlier partial repairs and failed
count observations remain preserved. Public signatures, version carriers,
resource limits, numeric acceptance thresholds and the baseline are unchanged.

The consumed source-availability checkpoint is renewed only for reviewed clean
canonical sprint source and complete normal CI. Bind actual source, tree,
origin, ancestry, actor, request, run and outcomes after observation. H7
signed subjects and its decider cannot be relabeled at the successor source.
Current-source installed artifacts, full runtime and typing, verified changed
package dry runs, remaining native, quiet performance and memory acceptance,
build-only26payloads, full verification, sprint review and close remain due
where not actually performed. The named native exception remains unproved.
No feature completion, dependency unblock, main merge, tag, publication, yank
or reporter comment is authorized by this checkpoint.

## H9 property namespace source availability renewal

The H8 normal run42346571659 has an actual MSRV failure after419layout cases
pass. Two older assertions conflate a namespace declaration with its expanded
name, and a later generic path refusal with the earlier checked invalid
measurement refusal. Strengthened namespace controls also exposed genuine
property publication and inherited foreign binding capture defects.

The successor preserves Word identities at generated property roots while
retaining foreign opaque records at their actual reader. The same tests cover
empty, default, nonempty, positioned and direct property writers, local
namespace shadowing, invalid measurement refusal and unchanged caller state.
The private source-qualified lower library passes1289cases. Actual integrated-source and hash receipts remain bound to their recorded
producer snapshots. Independent support pass2 closes S1 with zero defects
and zero smells. The final Root microscope and affected source/metadata checks
remain explicit prerequisites.
The empty table-property parser repair retains a qualified self-closing
property owner at its schema position. The PowerPoint number-format test
separately proves checked authoring refusal and imported empty-code fallback.
Both retain their actual prior failures and successor results.

The consumed source availability checkpoint is renewed only for reviewed clean
canonical sprint source and complete normal CI. Record the actual source,
tree, origin, ancestry, actor, request, run and outcomes after observation.
H8 remains failed and its successful categories retain their exact source.
No earlier native subject, signature, decider, wheel or diagnostic artifact
is relabeled at the successor. Existing per-feature pending arrays and states
remain unchanged, with additional affected input debt recorded separately.

Current release and diagnostic artifacts, installed functional and typing
checks, affected package checks, remaining native, quiet performance and
memory riders, build-only26payloads, full verification, sprint review and
close remain due where not actually performed. The named native exception
remains unproved. No feature completion, dependency unblock, main merge,
tag, publication, yank or reporter comment is authorized by this checkpoint.
