# F-X275, final union CI repair, Root affected pass 1

**Verdict**: 0 defects, 0 smells, 0 nitpicks. Accept the bounded local repair for later source-qualified hosted execution. No independent peer-cap reset, native/API qualification or final sprint pass follows.

## Actual failures and repair

`.github/workflows/ci.yml:1012`
`scripts/sprint_workflow.py:2435`

Actual H run42293933089macOS diagnostics identify small Homebrew wrapper scripts as inherited bare rustc/cargo/rustdoc. The existing identity refusal is correct. The two added producer lines obtain the installed1.97.1compiler path from rustup and prepend its concrete directory before reuse/fresh execution. Existing host, version, override, lexical and content-hash checks remain unchanged. No arbitrary wrapper allowlist, ignored mismatch or fake native proof was added. New real macOS execution remains mandatory. The earlier reuse inventory failure is preserved separately from fresh setup failure.

`crates/rdocx/Cargo.toml:37`
`Cargo.lock:1372`

The real supply-chain gate reports RUSTSEC-2026-0433against0.13.5. Its primary RustSec source identifies0.13.6as patched. The exact pin and registry checksum update only this package in Cargo.lock. Runtime parser options and disabled wikilinks remain unchanged. No advisory ignore or exposure claim is introduced. Current HLD, plan constraint and unpublished release prose now state the actual version and both relevant advisories. Published changelog history remains exact.

`.github/workflows/ci.yml:280`

The existing Test job retains exactly21named synthetic output files, the seven deterministic samples in DOCX/PDF/PNG forms, under a pinned existing upload action. These outputs have a real current consumer: final baseline reviewers inspecting the actual hosted generation and rasterizer. No arbitrary workspace, credentials or user files are uploaded. Missing diagnostic files are ignored, but earlier failed tests remain failed and the artifact supplies no acceptance flag. No job, event route, permission, product payload or release authority was added.

## Actual local checks

The affected all-target Word check passes with0.13.6. Eight actual all-feature Markdown regressions pass. Seven existing focused native-tool, fallback, deciding-gate and release-note policies pass, plus both carrier/deciding policies after diagnostic retention. The existing cargo deny check reports advisories, bans, licenses and sources all okay, actual exit0. Original CI failures and a local DNS-failed update attempt remain separate. The network-enabled exact lock update actually succeeds. Source, commands and observations are retained in LOCAL_F275_REPAIR.json and raw logs under /private/tmp/rdocx-ci-triage-42293933089-20261106.

No functional baseline is accepted by this pass. Actual16hash deltas still need final-source output inspection and labelled attribution, including the separate PDF ActualText and automatic-colour effects. The ZIP compatibility repair is independently reviewed under F280b. Archive records affected by source/dependency changes, complete source-qualified normal CI, all three native producers/signers and the full integrated gate remain pending.

## Defects

0.

## Smells

0.

## Nitpicks

0.

## Not found

No identity, source-binding, diagnostic-privacy, dependency-policy or scope defect established in this bounded change. Only this report was written during review. The audit ends before commit and further gate execution.
