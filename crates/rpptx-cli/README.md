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
- One-shot edits: batch replacement from a JSON map, slide add, duplicate,
  remove, move, hide and show, speaker notes, and core properties.
- A `fit` check that lists every overflowing text frame with the font scale it
  needs and exits 1 when one overflows.
- Scriptable output covers slide order, notes, comments, recursive group
  content, relationship-backed media, and deterministic rendering diagnostics.

## Measured footprint and speed

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rpptx-cli | 48,380 compressed bytes, 216,922 member bytes, 8 members | 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --no-verify` | Tracked `rpptx-cli` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-09 |

## Use it when

Use the CLI for shell automation. Use `rpptx` when the same operations belong inside a Rust application.

## Relationship

It uses `oxml-cli-support` for shared command conventions and the real `rpptx` facade for document behavior.

## Example

```sh
cargo install rpptx-cli --version '^0.14.0'
rpptx inspect deck.pptx --json
rpptx text deck.pptx --json
rpptx outline deck.pptx --notes
rpptx comment add deck.pptx --slide 2 --author Reviewer \
  --text "Check the figures" --date 2026-09-25T10:00:00Z -o reviewed.pptx
rpptx comment list reviewed.pptx --json
rpptx convert deck.pptx --to pdf -o deck.pdf
rpptx thumbnail deck.pptx -o thumbnail.png
rpptx replace deck.pptx --map pairs.json -o filled.pptx --json
rpptx slide add deck.pptx --layout "Title and Content" --at 2 -o added.pptx
rpptx slide duplicate deck.pptx 3 -o copied.pptx
rpptx slide move deck.pptx 4 --to 1 -o moved.pptx
rpptx slide hide deck.pptx 5 -o hidden.pptx
rpptx notes set deck.pptx 2 --from-file notes.txt -o noted.pptx
rpptx fit deck.pptx --json
rpptx meta set deck.pptx --title "Quarterly review" --author Ada -o titled.pptx
```

Every editing command takes `-o`, refuses an output that already exists, the
input included, and writes nothing when the edit is refused. Slide numbers are
one-based. With `--json`, each prints a schema-1 record with its `action`, the
one-based `slide` it produced and the `output`. A slide edit also reports
`changed`, false when the deck stays as it was, such as a slide moved to its
own position or a hidden slide hidden again.

`replace --map PAIRS_JSON` reads a JSON array of
`{"placeholder", "value", "expect"}` objects and applies the pairs in order, so
a later pair sees the text an earlier one wrote. A pair that gives `expect`
must replace exactly that count, and one without must replace at least one
occurrence, or nothing is written and the error names the zero-based pair.
`--json` reports the count of each pair and the total.

`slide add --layout` takes a layout name, matched exactly and then without
case, or its one-based number in master order, and `--at` the one-based
position of the new slide. An unknown layout lists the available ones.
`slide duplicate N` inserts the copy right after slide N. `slide remove N`
removes the slide with its notes and comments. `slide move N --to M` makes
slide N the M-th slide. `slide hide N` and `slide show N` set whether the slide
show skips it.

`notes set N --text TEXT` or `--from-file PATH` replaces the speaker notes of
slide N, one paragraph per line, and creates its notes slide when it has none.

`fit` lays out the text of every slide shape as rendering does and lists each
frame whose text overflows it, with its slide, shape id, name, autofit mode,
the font scale it renders at, and `needed_font_scale`: the largest scale, in
steps of 2.5% down to 25%, at which the text fits, as PowerPoint's shrink text
on overflow computes it, or `null` when even 25% overflows. Measures are
rounded to four decimals. It exits 1 when a frame overflows and 2 on an error,
so it can gate a script. Tables and SmartArt are not checked. It wraps
`Presentation::text_fit_report`.

`meta get` prints the core properties, with the creator as `author`, and
`meta set` writes the title, author, subject, keywords, description and
category it is given, keeping the others.

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
