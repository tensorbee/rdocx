# F-X272, Root import Source functional support, pass 2

**Reviewed**: Only the D1 successor from `ROOT_TAKEOVER/root-main-note-import-batch/source` to `ROOT_TAKEOVER/root-main-note-strict-repair/source`. Two Native files change. The exact diff contains six additions and five deletions.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. D1 is closed for this bounded successor.

## Defects

0 defects. The original D1 strict failure remains historical evidence in pass 1.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## D1 disposition

`crates/rdocx/src/document.rs:1392` now stores the large concrete Source variant as `Box<ScopedWordSource>`. The producer at `crates/rdocx/src/template.rs:1828` boxes the existing paid repeat result. The importer at `crates/rdocx/src/document.rs:47613` immediately moves that same Source out before continuing the existing branch. The second construction change is in the existing note regression. There is no XML clone, trait object, new parser, changed provenance selection or new allowance in this repair.

`crates/rdocx/src/document.rs:28197` now moves the existing Arc directly into emitted_part, removing the redundant conversion without changing ownership or publication behavior.

Both repaired source hashes authenticate against `MANIFEST.json`:

| Path | SHA256 |
|---|---|
| crates/rdocx/src/document.rs | 0a7a16ebf1dca21a6726dc39d78c813f587fb538bf63abb075c5ed9c33392af3 |
| crates/rdocx/src/template.rs | 9fac16c4f2ff5099b00ee868661298910c582ea4efe398c15816ce605f71909e |

The manifest records strict_exit 0. Its archived strict log records completed checking in 17.65 seconds, with no diagnostics. This is the targeted successor to the strict all-targets, warnings-denied command recorded in pass 1. No all-features or runtime execution is inferred from that log.

## Evidence applicability and boundaries

The original 21 scoped, 18 versioned and 10 fragment passing cohorts predate this concrete Box and Arc cleanup. They remain attached to their original frozen source and are not relabeled as runtime execution of these two successor hashes. No reviewer build or test was run.

The functional review and every exclusion in pass 1 remain unchanged. Changed or rebuilt revision returned contexts, inline-control returned maps, complete non-emitting expression evaluation, remaining related-comment/glossary work, final joins and whole C1/provider acceptance remain pending. Final integration, callable inventory, SemVer and integrated hash/native/oracle riders are not discharged here.

## Not found

The exact repair introduces no additional correctness, ownership, panic, XML preservation, namespace or structural-rule finding. This pass neither repeats the broader source review nor approves completion beyond closing D1.
