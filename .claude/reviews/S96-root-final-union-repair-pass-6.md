# S96 bounded CLI expectation repair, Root pass6

Read-only review phase. Only this review record is written in this pass.
This is a source-availability correction, not final sprint acceptance.

## Defects

0.

## Smells

0.

## Nitpicks

0.

## Actual failure and repair

H6Test and MSRV fail the same17CLI integration cases. Actual MSRV Word
library1009passes, integration559passes and regression2531passes precede that
failure. Those successful observations keep their exact source/configuration.
The original failed overall CI and skipped later goldens/decider are retained.

The correction changes only CLI tests and their archive measurement metadata.
Production Rust, public interfaces, schema rules, parser, renderer and CLI
behavior are unchanged. The strict global success helper remains strict.
Named render calls allow exactly the approved Calibri to Carlito diagnostic.
Forced output checks allow only empty stderr or that exact single diagnostic,
with the existing classifier still rejecting unrelated warnings.

Exact JSON expectations include the approved revision text and crossing reply
range facts. Authored footnote leading spaces are retained in plain, structured
and diff output. Invalid positive controls now have required revision identity,
real note definitions/relationships and controls before the final sectPr.
Actual source-qualified experiments validate82positive fragments and retain
10invalid negatives. Existing visibility, warning-count, source-preservation and
geometry assertions are unchanged. The malformed negative controls still refuse,
now requiring the actual earliest closed-root/CDATA diagnostic.

The old1pixel PNG failed its zlib checksum and chunk CRC. The replacement is a
valid1pixel RGB image in source with the same dimensions. Its image decode
warning is eliminated by fixing the asset, not by allowing image diagnostics.

## Evidence

The first current host CLI batch passes115and fails the later stale CDATA
assertion. Actual isolated CLI output establishes text outside root element.
That exact negative assertion is corrected and its full test passes. The
subsequent complete current macOS CLI integration command passes116of116,
with source stable. The original Linux-only case is unchanged and remains a
fresh hosted check. Actual compiled CLI bytes stay36e583df throughout the
repair, independently confirming no program change.

The affected CLI source archive is actually126182compressed bytes,
593817normalized member bytes and8members. Its README and script row agree.
Post-adoption repackaging passes exact normalized bytes/count and the unchanged
64byte compressed tolerance. Other21archive rows and speed rows are unchanged.
The initial network-restricted package preparation failure is retained, followed
by actual offline acquisition success. This is not a verified publication run.

All49hash entries already match on unchanged production/source generator inputs.
No baseline change is introduced. Current wheel/native/installed and other
product evidence retain their actual4cf producer identities, with later
nonproduct applicability assessed explicitly. Fresh deciding CI, package
verification, full final riders, review and close remain due. No main, tag,
registry publication, reporter comment or native-pin substitution is authorized.

## Reviewed diff pins

- `crates/rdocx-cli/README.md`: `c79cc9e50daa4d3460eb756051054dc2cbc5e4b0be0d8cd71bc03adb3da0010b`.
- `crates/rdocx-cli/tests/integration.rs`: `733ed084ea75b5506aab4dc7244779b5770b3ff2dcf2c16a93219e3e68223a3b`.
- `scripts/readme_doctests.py`: `1f4cddec3a7c97117fa87786a21293fe5f20bc9b8f18589f9a983cec11e24d9f`.
