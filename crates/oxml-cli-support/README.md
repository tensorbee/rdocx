# oxml-cli-support

Give OOXML command-line tools consistent range parsing, output naming,
publication, and JSON contracts.

## Capabilities

- Positive one-based inclusive ranges with sorted, deduplicated results.
- Output extension replacement and collision checks, including refusal of an
  output that is the input file under any spelling of its path, or that is not
  a regular file.
- Adjacent temporary-file staging with cleanup and rollback after errors.
- Atomic replacement of existing files when the caller allows it. A replacing
  set keeps the files it has already swapped in when a later one fails, since
  the files they replaced cannot be restored.
- Versioned JSON object envelopes shared by Word and Presentation CLIs.
- Replacement maps, a JSON array of ordered pairs with optional expected
  counts, read the same way by both CLIs.
- The local wall-clock time and civil date arithmetic for date arguments.
- Bounded expansion and staged multi-output publication without replacement
  keep large or failed requests from leaving partial command results.

## Measured footprint and speed

The archive row is regenerated from the package that carries this README.

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: oxml-cli-support | 11,281 compressed bytes, 40,193 member bytes, 6 members | 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --no-verify` | Tracked `oxml-cli-support` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-09 |

## Use it when

Use this crate when building a repository DOCX or PPTX CLI that must follow the
same output-path and structured-output conventions. Application code should
use `rdocx` or `rpptx` instead.

## Relationship

This format-neutral crate is consumed by `rdocx-cli` and `rpptx-cli` and does
not depend on either document model. It does not provide argument parsing,
document I/O, or a user-facing CLI. Filesystem rollback is best effort.

## Example

```rust,no_run
let slides = oxml_cli_support::parse_range("2,4-6")?;
assert_eq!(slides, vec![2, 4, 5, 6]);
# Ok::<(), oxml_cli_support::Error>(())
```
