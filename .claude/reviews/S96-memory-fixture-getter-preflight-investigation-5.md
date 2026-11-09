# S96 memory fixture getter preflight, investigation 5

## Disposition and actual interrupted observation

The optional debug preparation was interrupted, not rejected by document
admission. Its retained transport exits130 after Root sent SIGINT to verified
own PID30775. The receipt explicitly records no resource refusal and no full
semantic/RSS acceptance at
`.claude/scratch/S96-final-evidence/memory-final-source-preparation-20261108-1/admission-debug-24-1/INTERRUPTED.json:2`.
The observation identifies KeyboardInterrupt at `OBSERVATION.json:9` in that
same directory. Zero product defects, zero smells and zero nitpicks are
established by this source/evidence investigation.

The script calls Document.open before inspect. Its traceback enters inspect
after the paragraph-count assertion and reaches the representative p.runs
expression at
`.claude/scratch/S96-final-evidence/memory-final-source-preparation-20261108-1/admission_debug_24.py:13`
and `:15`. Therefore Document.open returned and the20000count assertion passed.
The script assigns record['opened'] only after inspect returns at `:19`.
Its absence from the partial JSON does not mean open failed. Sample index and
individual native getter phases were not logged. The remaining formatting,
save and reopen work was not completed or qualified.

This Python traceback is not a native stack profile or an isolated elapsed
measurement. It cannot establish which internal frame consumed time, which
sample index was active, a leak, current release cost or an admission failure.
No cause is invented from the process interruption.

## Source applicability

The actual extension is ordinary debug21+24A Native
`7600718ec16a350e2d02fb228c98cfabfce9363d2763d9a3663c52781658087b`,
as identified in the observation. Its successful producer records the ordinary
debug Maturin command and unchanged source in
`.claude/scratch/S96-final-evidence/f268-direct-owner-key-implementation-20261108-1/ordinary-debug/build-result.json`.
Read-only hash comparisons against that producer's source-inputs manifest match
the inspected Python paragraph, run, formatting and document files, support
lib.rs, and Native paragraph/document files. A matching frozen Python document
copy also exists under that packet's `new-diagnostic/before/` tree. The citations
below describe those inspected producer-matching paths, not later Root changes.

No product import/runtime, compiler, Cargo, test, timing, network, mutation or
preflight execution occurred in this investigation. Only this review was
written. Root retains sole source/Cargo ownership during the namespace and
prospective performance work.

## Actual collection and current-authority path

1. `doc.paragraphs` is not a bare container read. The getter at
   `crates/rdocx-py/src/document.rs:17511` calls ensure_word_authority, then
   creates PyParagraphCollection. That initializer at `:12413` either returns
   for an existing snapshot or validates actual Word facts, allocates checked
   owner scopes/indexes, initializes revision states and installs current
   owners/pieces. validate_word_facts at `:8392` tries WordPieces assembly,
   falls back to actual full owner locations if unavailable, and retains the
   debug full-rebuild comparison when piece inventory succeeds. The actual
   preflight already completed this getter before its count assertion.
2. Paragraph collection length at `crates/rdocx-py/src/paragraph.rs:1201` uses
   cached_body_paragraph_count when available and otherwise the native facade
   count. The cached count comes from current body-paragraph indexes at
   `crates/rdocx-py/src/document.rs:10702`. Native count at
   `crates/rdocx/src/document.rs:39072` counts the real facade paragraph iterator.
3. Selection at paragraph.rs`:1190` captures a path and binds the paragraph
   through bind_current at `:416`. It resolves the current physical owner via
   paragraph_route, then captures checked owner scopes. A route miss retains
   the existing snapshot fallback. Under debug assertions, an indexed route
   is also compared with the actual native route at
   `crates/rdocx-py/src/document.rs:10483` and `:4806`. Native paragraph routing
   at `crates/rdocx/src/document.rs:40119` selects the real facade paragraph,
   validates stories/source and only then takes the existing direct-body route.
   Those checks can reach substantially more work than a container lookup.
   Their cost in this interrupted attempt was not measured.
4. The p.runs getter at paragraph.rs`:991` calls PyParagraph::validate at `:493`,
   captures a current Content view for an owned paragraph, and allocates a
   PyRunCollection holding checked paths/owner states. For the fixture's direct
   body paragraph handle, validate keeps the Body route after scope validation.
   It does not take the Related-owner refresh branch. capture_current_view at
   Python document.rs`:10560` uses indexed_owner's generation lookup at `:10438`,
   then the existing word_owner_capture at `:4863`. That captures document,
   removal, owner/content and actual parent scopes. The getter does not call
   ensure_word_authority again or enumerate runs in this direct-body branch.
5. If `len(runs)` is reached, its separate path is
   `crates/rdocx-py/src/run.rs:639` to len at `:605`, collection validation at
   `:583`, then ParagraphLocation::read at paragraph.rs`:216`. Direct Body reads
   call native paragraph(index), whose facade iterator uses nth at Native
   document.rs`:37759`. ParagraphRef::run_count at Native paragraph.rs`:2948`
   computes the actual accepted_run_paths length. Thus collection length is a
   checked accepted-view read, not necessarily Vec::len. Related/cell routes
   retain their current owner refresh/native story read fallbacks. The traceback
   does not prove this length operation completed.

ContentPath::validate_scopes at `crates/oxml-py-support/src/lib.rs:640` checks
document identity and captured states' first invalidation causes. capture_scopes
at `:403` validates the actual scope set and stores cloned state handles.
These are required ownership checks, not permission to use an unchecked numeric
index merely because this input is regular.

## Additional work in the original formatting loop

`for run in p.runs` creates another run collection. Iterator next at
`crates/rdocx-py/src/run.rs:698` validates/counts and obtains each real run.
RunCollection::item at `:614` resolves the accepted run path and binds current
owner authority. PyRun::bind_current at `:59` again uses physical paragraph
routing, with its debug indexed-route comparison or snapshot fallback.
Font access at `:510` validates the run and captures a Content view. The font
name getter and run.bold projection subsequently validate and read actual
selected run properties. This is more than inspecting eighteen cached strings.

Those paths and debug reference checks could contribute work in a completed
debug smoke. They are source-level possibilities, not proven hotspots here:
the interrupted attempt did not finish or profile the formatting loop. Do not
attribute the interruption specifically to those later operations or assume
release behavior has the same cost.

## Actual mandatory Python setter path

A future companion must exercise the genuine public operation rather than
constructing a native mutation that bypasses binding checks. Its operand path
is doc.paragraphs[index], paragraph.runs[run_index], run.font, then assigning
font.name. Font.set_name at `crates/rdocx-py/src/formatting.rs:615` calls apply
at `:434`, which validates the current font/run path at `:397` and invokes
apply_run_update at `:453`. That uses ParagraphLocation::edit at
`crates/rdocx-py/src/paragraph.rs:250`, the real paragraph edit_run, FontUpdate::Name
and Run::set_font_value. Native body editing reaches paragraph_mut at Native
document.rs`:39192`. No whole-package prepare/reopen is added per font assignment.

Ordinary getter/handle acquisition required by this operand remains part of
the declared first-edit phase. An extra font getter, run enumeration, count or
semantic scan before that phase would change activation and must not be inserted
into a deciding RSS worker merely because it appeared in this optional smoke.

## Minimal future non-deciding preflight

Use a new output directory and authenticate the exact unchanged fixture SHA
`a1fbdfb7a40dfe63d9a4f273b8bb6c1000d5bc43541a1c04a6e9f4e78924e204`,
interpreter, artifact/profile and source. Preserve the original20000-by-six
fixture and all Source/resource limits. Choose an explicit operational stop
boundary before launching. An interrupted child remains interrupted, and only
Root's verified own PID is stopped before quiet timing readiness. No broad
process termination or reduced fixture substitutes for success.

Two isolated disposable children are sufficient to separate the questions:

- **Getter localization child:** emit flushed begin/end records for open,
  paragraph-collection activation, paragraph count, and then each existing
  representative index0/10000/19999. Give paragraph selection, text assertion,
  p.runs getter and len(runs) separate Python statements and separate records.
  Only after those complete, separately record run iteration, each run item and
  the existing font/bold checks. Do not put all operations on one source line or
  claim a completed sample until its checks return. This is the existing sampled
  smoke with progress visibility, not an RSS or performance benchmark.
- **Setter companion child:** start from a fresh document with no preceding
  inspection. Record open and each mandatory operand step separately, obtain
  one declared direct-body run/font handle, and perform the actual first changed
  font assignment specified by the eventual qualified worker. Its value must
  differ from the fixture's Arial value. Do not read font.name or enumerate runs
  to prewarm that child before the assignment. After the setter returns, verify
  selected font, unchanged text/sibling formatting and recorded owner behavior.
  Then separately record public save, fresh reopen and post-reopen checks on
  that disposable result. A setter changes main XML, so original whole-package
  byte equality is not an appropriate success predicate for this child.

Any optional phase clocks are explicitly operational, non-deciding observations
with the actual start/end/exit retained. Their values neither qualify elapsed
performance nor identify a native hotspot without profiling. Both children may
use separate authenticated debug/release configurations when needed, but results
and assertion/reference differences remain distinct. Changing a profile does
not convert the existing interrupted debug attempt into a pass.

Destroy both children, all documents/handles and saved buffers before any
deciding RSS execution. A deciding worker starts in a new process, reopens the
same untouched input and performs only its predeclared mandatory sequence.
Neither a live preflight document nor its warmed caches, owner indexes or
inspection results may be supplied to that worker. Keep comprehensive semantic,
omission equivalence and activated-history companions separate from RSS phases.

## Pending facts and unchanged requirements

The exact sample index and native operation active at interruption remain
unknown. There is no phase-resolved successful smoke, completed sampled-format
result, saved/reopened outcome or RSS result from this attempt. Actual future
first-setter and full sampled/save/reopen completion remain pending. A genuine
resource refusal, if encountered, retains its exact first cause and does not
authorize a larger allowance, Source credit or smaller convenient fixture.

Memory investigations2/3/4 still require current-source Normal/omission
equivalence, all mandatory identifiers and provenance, candidate/key/loan
construction suppression, Native/Python ordinary-private calibration, three
fresh alternating pairs, raw current/peak units and equal-terminal N/4N expiry
proof. This proposal changes no endpoint, memory ceiling, timing threshold,
source budget, product/plan, release status or publication authority. End the
bounded pass and return control to Root.
