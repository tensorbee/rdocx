# rpptx-cli

`rpptx-cli` makes complete PPTX workflows available to shell scripts. It
inspects and extracts deck content, validates package invariants, compares or
replaces text, and produces deterministic fixed output.

## Capabilities

- Human-readable inspection, and JSON inspection with per-shape kind,
  placeholder, geometry, rotation, autofit, and text detail.
- Slide-order text extraction and recursive outlines, as plain text or schema-1
  JSON, with speaker notes on request.
- PDF, PNG, JPEG, and multi-page TIFF conversion.
- Selected-slide rendering, thumbnails, text diff, replacement, and validation.
- Modern comment thread listing, addition, replies, resolution, and removal.
- Scriptable output covers slide order, notes, comments, recursive group
  content, relationship-backed media, and deterministic rendering diagnostics.


## Package signatures

```sh
rpptx sign input.pptx --key private-key.pem --cert certificate.pem -o signed.pptx --json
rpptx verify-signatures signed.pptx --json
```

The key must be PKCS#8 RSA and the certificate X.509, as DER or one unencrypted
PEM block each, at most1MiB per input. Signing refuses existing destinations,
including aliases, and publishes only a valid complete candidate. Verification
returns schema1 JSON with status and signatures. Status is valid, invalid,
partial or unsigned. Exit0 requires at least one valid complete signature and
no invalid or incomplete report. Runtime and integrity failures exit1, and
usage errors exit2. Certificate-chain trust is not evaluated. The native signer
refuses already signed inputs and packages with external relationships.
Both commands return exit1 on report output errors, including a closed stdout.
Signing publishes before reporting, so this error can leave a valid new output.

## Measured footprint and speed

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rpptx-cli | 67,538 compressed bytes, 310,592 member bytes, 8 members | 0.14.3 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo publish --workspace --dry-run --allow-dirty --locked` | Tracked `rpptx-cli` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-23 |

## Use it when

Use the CLI for shell automation. Use `rpptx` when the same operations belong inside a Rust application.

## Relationship

It uses `oxml-cli-support` for shared command conventions and the real `rpptx` facade for document behavior.

## Example

```sh
cargo install rpptx-cli --version '^0.14.3'
rpptx inspect deck.pptx --json
rpptx text deck.pptx --json
rpptx outline deck.pptx --notes
rpptx comment add deck.pptx --slide 2 --author Reviewer \
  --text "Check the figures" --date 2026-09-25T10:00:00Z -o reviewed.pptx
rpptx comment list reviewed.pptx --json
rpptx convert deck.pptx --to pdf -o deck.pdf
rpptx thumbnail deck.pptx -o thumbnail.png
```

`inspect --json` keeps its existing keys and adds `shape_details` beside each
slide's shape count. Each shape reports its z-order index, id, name, kind,
placeholder type and index, direct position and size in EMU, rotation in
degrees, direct autofit mode, paragraphs, table size, and children. A
placeholder that inherits its geometry from its layout reports `null` geometry,
and a placeholder without an explicit type reports a `null` type.

Both inspection forms report every core property the deck sets. Beside title,
creator, subject, description, keywords, last modified by, created and
modified, they report category, content status, identifier, language, last
printed, revision and version, which `inspect --json` adds to its `metadata`
object. Human-readable inspection prints `(none)` when the deck sets no core
property.

`text --json` reports each slide's one-based number, slide id, paragraphs, and
speaker notes, or `null` notes when the slide has no notes part. Each paragraph
has a typed zero-based `path` of shape, table row, table cell, and paragraph
positions, the id of the shape that owns it, its level, its visible text, and
its regular runs. Run indexes are the facade's run positions, so fields and line
breaks appear only in the paragraph text, where a line break is U+000B. Run
`formatting` is `null` when a run has no direct properties. Otherwise it records
nullable direct bold, italic, underline, font, point size, and sRGB colour
values.

`outline --json` reports each slide's title, its outline items with their
levels, and its speaker notes. Notes are the plain text of the notes body, with
paragraphs and line breaks both written as newlines. `--notes` prints one
`Notes:` line for each non-empty notes line after a slide's plain text or
outline. JSON output always carries the notes.

`comment` works on modern PowerPoint comments. Legacy comments stay preserved
but are not listed. `list` shows each comment and reply as `open`, `resolved`,
or `closed`, and its JSON carries the `resolved` flag beside the raw `status`
token. `add --slide` takes a one-based slide number. `add` and `reply` require
an RFC 3339 `--date`, reuse an existing author with the same name, and add the
author otherwise. New ids are sequential GUIDs, so the output depends on
neither a clock nor a random source. `resolve` takes a thread id. `remove`
takes a thread id, which also removes its replies, or a reply id. Every
mutation requires `-o/--output`, refuses an existing output, publishes only a
complete presentation, and supports a schema-1 record through `--json`.

`convert`, `render`, and `thumbnail` refuse an output file that already exists
unless `--force` is given. Even with `--force` they refuse their own input file
under any spelling of its path, and an output that is not a regular file, such
as a directory, a symbolic link, a FIFO, or a device like `/dev/null`. A run
checks every file it would write before it writes the first one, and publishes
each file only once it is complete, so a failed run leaves no truncated output.

The output extension of `replace` and of every comment mutation selects the
package class the output declares. A `.potx` template written to `deck.pptx`
becomes a presentation, and a deck that carries a VBA project cannot change to
a macro-free extension.

Run `rpptx --help` or `rpptx <command> --help` for the complete command surface.

## One-shot editing and fit reports

```sh
rpptx replace deck.pptx --map pairs.json --json -o filled.pptx
rpptx slide add deck.pptx --layout 7 --at 1 --json -o added.pptx
rpptx slide duplicate deck.pptx 2 -o copied.pptx
rpptx slide remove deck.pptx 2 -o removed.pptx
rpptx slide move deck.pptx 3 --to 1 -o reordered.pptx
rpptx slide hide deck.pptx 2 -o hidden.pptx
rpptx slide show hidden.pptx 2 -o shown.pptx
rpptx notes set deck.pptx 1 --from-file notes.txt -o noted.pptx
rpptx fit deck.pptx --json
rpptx meta get deck.pptx --json
rpptx meta set deck.pptx --title "Quarterly report" --author Ada -o titled.pptx
```

Replacement maps are ordered JSON arrays of objects with `placeholder`,
`value` and optional `expect`. Unknown fields and invalid types refuse.
An absent or null expectation permits any count, including zero. Each pair
searches the preceding pair's result, so repeated pairs and cascades are
meaningful. An empty array succeeds unchanged. Empty placeholders refuse.
Single `-p/-v` and `--map` are mutually exclusive, and map mode refuses a
command-line `--expect`. Single replacement retains its refusal on zero matches
unless `--expect 0` is supplied. Slides and speaker notes share the native
literal scope. All pairs run in one checked native candidate. A later count
mismatch or malformed input creates no output and leaves the source unchanged.

Replacement JSON has schema1, action, mode, ordered pairs with index,
placeholder, value, count and nullable expected, total_count and output.
Every mutation prints JSON only after publication. New editing commands require
an output, refuse an existing destination and protect input aliases, symlinks
and hard links. They use the native extension-based package class and VBA
contracts. Recoverable staging failures remove this invocation's new files.
Filesystem cleanup remains best effort rather than crash-atomic publication.

Slide and numeric layout positions are one-based. A decimal layout selects its
checked numeric index. Other layout values require an exact unique name.
`--at` is the final insertion position, including append, and defaults to
append. Creation and placement form one shared native transaction. Duplicate
inserts directly after its source. Move uses the final destination position,
and moving to the same position succeeds unchanged. Hide and show also retain
no-op behavior. Mutation JSON records action, selected slide, resulting slide
count, output and changed where applicable.

Notes require exactly one of `--text` and `--from-file`. Complete UTF-8 input
is read before editing. LF, CRLF, carriage returns, leading/trailing spaces,
Unicode and empty text are retained without trimming or an added newline.
Existing notes body formatting and opaque payload stay preserved. JSON reports
the selected slide and Unicode scalar character_count.

`fit` is read-only and consumes the native all-owner `fit_report`, including
its inherited contexts, admitted table cells and diagnostics. It uses bundled
deterministic fonts or the strict caller loader through `--font-dir`.
Frames carry owner part, child path, cell, shape identity/name, context, mode,
status, current stored values, overflow, needed total scale/reduction and the
provider's quantized and persisted trials. Dimensions/heights are points,
old/new bounds are EMU, and needed scale/reduction use hundred-thousandths.
No permitted fitting trial means nullable needed_scale with
`reason=unfit-at-floor`. Empty frames remain accounted for. Unsupported
mandatory owners refuse with diagnostics instead of disappearing. Exit0 means
all admitted frames were assessed without overflow, exit1 means overflow,
and exit2 means a checked error. The command does not modify autofit or save.

Core metadata get includes every native core field. Set accepts title, author,
subject, keywords, description and category, preserving all unspecified fields
and opaque imported metadata. An empty supplied value removes that field.
Duplicate flags and an empty set request refuse before publication.

The same common operations are exposed in Python as
`Presentation.replace_all`, `SlideCollection.add_slide_at` and
`Presentation.fit_report`. Ordered counts are tuples, count mismatches retain
index/placeholder/expected/found, and failed or actual no-op batches preserve
held content handles. FitReport, FitFrameReport and FitTrial are frozen
snapshots rather than mutable owner selectors.

## Licensed font embedding

Use caller-owned font files only after confirming their embedding license.
Embedding requires explicit authorization and exact license identity. Commands
write a new output and refuse input/font aliases or existing destinations.

```bash
rpptx font embed input.pptx --name "My Face" --font ./MyFace-Regular.ttf \
  --license-authorized --license-identity "SIL Open Font License 1.1" \
  --output embedded.pptx
rpptx font list embedded.pptx --json
rpptx font settings embedded.pptx --embed-true-type-fonts false --output disabled.pptx
rpptx font remove embedded.pptx --name "My Face" --kind regular --output removed.pptx
```

The four CLI kinds are regular, bold, italic and bold-italic. Settings accept
true, false or none, with none removing the owned value. Listing reports actual
native part, format and byte-length facts in the schema1 JSON envelope.
PowerPoint preserves compressed/unknown imported font data and reports absent
decoded bytes. save-subset-fonts is the other presence-sensitive setting.
