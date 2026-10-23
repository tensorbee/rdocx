# rdocx-cli

`rdocx-cli` turns DOCX files into inspectable and scriptable shell artifacts.
It extracts content, reports structure, validates packages, applies changes,
and produces fixed or flow output without an Office host.

## Capabilities

- Human-readable or JSON structure and metadata inspection.
- Accepted-view plain text or schema-1 rich extraction with typed nested paths,
  followed by the text boxes, headers, footers, notes, and comments.
- Deterministic point-space body layout fragments for shell automation.
- PDF, HTML, Markdown, PNG, JPEG, and multi-page TIFF conversion.
- Page-range rendering, guarded literal replacement, diffing, and validation
  verdicts that check every related part and every style id.
- Comment thread inspection and mutation with explicit body run ranges or an
  anchor text, and optional RFC 3339 dates.
- Tracked revision inspection, filtered resolution, table-of-contents rebuilds,
  and document comparison at run, word, or character granularity with ignore
  options.
- Atomic form filling accepts repeatable `--tag` assignments or strict flat
  `--values` JSON over unique SDT tags/aliases and legacy names.
- Package-preserving edits cover the final S74 paragraph, run, typography,
  table, section, settings, field, form, equation, and drawing surface.


## Package signatures

```sh
rdocx sign input.docx --key private-key.pem --cert certificate.pem -o signed.docx --json
rdocx verify-signatures signed.docx --json
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
| Crates.io archive: rdocx-cli | 111,061 compressed bytes, 516,119 member bytes, 8 members | 0.16.3 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo publish --workspace --dry-run --allow-dirty --locked` | Tracked `rdocx-cli` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-22 |

## Use it when

Use this crate for shell automation. Use the
[`rdocx`](https://docs.rs/rdocx) library when these operations need to run
inside a Rust application.

## Relationship

The binary delegates document behavior to `rdocx` and shares path and JSON
conventions with `rpptx-cli` through `oxml-cli-support`.

## Example

```sh
cargo install rdocx-cli --version '^0.16.3'

rdocx inspect report.docx
rdocx text report.docx
rdocx text report.docx --json
rdocx layout report.docx --json
rdocx replace template.docx -p TOKEN -v ready --expect 1 -o report.docx
rdocx convert report.docx --to pdf -o report.pdf
rdocx convert redline.docx --to pdf --revision-view tracked -o redline.pdf
rdocx render redline.docx --revision-view tracked --revision-style legal
rdocx diff before.docx after.docx
rdocx diff before.docx after.docx --exit-code --json
rdocx validate report.docx
rdocx render report.docx --page 0 -o rendered
rdocx comment list report.docx --json
rdocx comment add report.docx --start-paragraph 0 --start-run 0 \
  --end-paragraph 0 --end-run 1 --author Reviewer --text 'Check this' \
  --date 2026-09-13T12:00:00Z -o commented.docx
rdocx comment add report.docx --anchor 'target words' --occurrence 1 \
  --author Reviewer --text 'Check this' -o commented.docx
rdocx revision accept reviewed.docx --author Reviewer -o accepted.docx
rdocx compare original.docx edited.docx --author Reviewer \
  --timestamp 2026-09-13T12:00:00Z -o redline.docx
rdocx compare original.docx edited.docx --author Reviewer \
  --timestamp 2026-09-13T12:00:00Z --granularity word --ignore-comments \
  --ignore-story header -o redline.docx
rdocx toc rebuild report.docx -o refreshed.docx
```

Comment `add` ranges use zero-based body paragraph and run boundaries. The
start is inclusive and the end is exclusive. Run boundaries count the runs that
`text --json` lists, including the runs inside inline content controls and
tracked insertions. A range that cannot be anchored exactly, such as one that
crosses the edge of an inline content control, is refused. In place of the
four range flags, `--anchor TEXT` comments on the zero-based `--occurrence`
(default 0) of a literal, case-sensitive text of the main story, through body
paragraphs, tables and block content controls, and splits the runs at both
ends of the match. A text that does not occur, an occurrence past the last
match, and a match that cannot be anchored exactly exit unsuccessfully without
creating the output. Comment replies,
resolution, and removal select a decimal comment id. Comment `add` and `reply`
write an optional `--date` RFC 3339 timestamp as the comment date. An invalid
timestamp exits unsuccessfully without creating the output, and without
`--date` the comment stays undated.

`compare` replaces each changed run whole by default (`--granularity run`),
like the Rust and Python APIs. `--granularity word` marks only the changed
words, and `--granularity character` marks only the changed characters.
`--ignore-formatting`, `--ignore-whitespace`, `--ignore-fields`,
and `--ignore-comments` keep the original side of those differences.
`--ignore-comments` keeps the original's comments and anchors and drops the
comments of the edited file. The repeatable `--ignore-story KIND` keeps one
story of the original, where `KIND` is a Python `Story.kind` name: `body`,
`header`, `footer`, `comment`, `text_box`, `footnote`, or `endnote`. Ignoring
the `comment` story excludes only the comments part, so a pair whose comments
differ needs `--ignore-comments`. An unknown granularity or story name is a
usage error, a repeated story fails, and neither creates the output. The
`--json` record states the options that ran.

Revision `list` reports every supported story and names the story of each
revision. Revision `accept` and `reject` operate across every supported story
and accept at most one selector: `--id`,
`--author`, or the paired `--start-date` and `--end-date` RFC 3339 bounds.
Omitting a selector resolves all modeled revisions. Every mutation, comparison,
and TOC rebuild requires `-o/--output`, publishes only a complete validated
DOCX, and supports a schema-1 record through `--json`. `compare` reports how
many revisions it created in each story while retaining the main-body count.
A `.docx`, `.docm`, `.dotx`, or `.dotm` output extension selects the package class the output
declares, so a template edited into `report.docx` is written as a document. An
input that carries a VBA project cannot change to a macro-free extension.

`text` prints each paragraph with the same accepted-view text as `text --json`:
tracked insertions and move destinations are included, and tracked deletions
and move sources are left out. Plain `text` covers body paragraphs and the
paragraphs inside content controls, table cells, and nested tables in document
order.

After the body, `text` prints every other story: text boxes, headers, footers,
footnotes, endnotes, and comments, in the order of `Document::stories`. Each
package part starts with a line that names the kind of its first story and the
part, such as `--- header (/word/header1.xml) ---`. Each direct paragraph or
block content control of its stories follows on one line, table cells
included, with the accepted-view text of Python `StoryItem.text`. A part
without any text, such as an empty header variant, is left out. A text box
that Word writes twice, DrawingML in `mc:Choice` and a VML copy in
`mc:Fallback`, is never read from the copy. When a story part cannot be read,
such as a truncated header or a header relationship to a missing part, `text`
and the Markdown and HTML conversions keep the body, print one warning on
stderr that names the part when it can, list no other story, and exit
successfully. `text --json` then states `main` scope. `validate` reports that
part as an error.

`text --json` reports accepted-view paragraphs in source order. Each paragraph
has a zero-based `body_index`, a typed zero-based path within that body item,
its direct style and numbering, and accepted-view runs. Its `text` also holds
the text inside smart tags, inline custom XML elements, and simple fields,
which `runs` does not list. Run `formatting` is
`null` when no direct run properties exist. Otherwise it records nullable
direct bold, italic, strike, underline, font, point size, colour, highlight,
language, and character style values. The record has `all-supported-stories`
scope, and its `stories` array lists the other stories as plain `text` does,
one entry per story with its Python `Story.kind` name as `kind`, its
`part_name`, its `owner_index`, and `items`. Each item is a direct paragraph or
block content control with its story `index_path`, its `kind`, and its `text`.

`convert --to md` and `convert --to html` append the same stories after the
body, one section per part under a bold label. They leave comments out,
because comments annotate a document rather than belong to it.

`convert` to PDF or images and `render` use the accepted revision view by
default. `--revision-view tracked` shows both sides of tracked changes, and
`--revision-style legal` colours them as legal comparison tools do: insertions
blue with a double underline, deletions red and struck through, moved text
green. `--revision-style` needs `--revision-view tracked`.
Unknown view names are usage errors. HTML and Markdown conversion refuse
the tracked view before creating output.

`validate` exits unsuccessfully when a relationship of the main document points
at a missing part, when a part has no declared content type, when an XML part
that the main document relates to is not well formed, or when the main
document, a header, a footer, the notes, or the comments name a paragraph,
character, or table style id that no style defines. Word silently falls back
to the default style for such an id. A style id inside a tracked property
change is not checked, because it records the formatting before the change,
and neither is one inside `mc:Fallback`, which Word does not read, or an empty
id.
Empty paragraphs, heading level gaps, and missing metadata are warnings.

`layout --json` uses bundled deterministic fonts. It lists every direct body
item, including preserved items that have no fragments. Each laid-out fragment
uses points from the top-left page origin and records one-based physical and
displayed page numbers. A body item that crosses a page boundary has one
fragment on each occupied page.

`diff` compares the accepted-view paragraph text of every story: body
paragraphs, the paragraphs of the body's table cells and nested tables, text
boxes, the headers and footers of each section, footnotes, endnotes, and
comments. Each story is compared as one sequence through a shortest edit
script, found by Myers' linear-space algorithm, so a few edits in a long
document stay fast and memory stays proportional to the story length. Between
two matched paragraphs, removed and added paragraphs pair in order as changed
paragraphs. A changed paragraph prints a `-` line and a `+` line, an added one
only `+`, and a removed one only `-`. The summary line reads
`N paragraph(s) changed, A added, R removed.`, which replaces the
`N paragraph(s) differ.` line of earlier releases.

Each line locates its paragraph between brackets. A body paragraph keeps its
one-based position among the body paragraphs, such as `[2]`. A table cell
paragraph of the body is located as `[table 1, row 1, cell 2, paragraph 1]`,
where the table counts the tables placed directly in the body and the rest is
the `text --json` path made one-based. A header or footer is named by its type
and the first section that references it, as
`[header default, section 1, paragraph 1]` or
`[footer first, section 2, paragraph 1]`, and the two files pair their headers
and footers by that name. Inserting a new first section with its own header
therefore shows the old header as changed and itself as added under section 2.
Notes and comments count in their part, as `[footnote 1, paragraph 1]` or
`[comment 2, paragraph 1]`, and a text box of the body as
`[text box 1, paragraph 1]`. A table cell inside a header, a footer, or a
notes or comments part is labelled `table cell K` in that part, without a row
and cell path, as `[header default, section 1, table cell 3, paragraph 1]`.
Outside the body, `paragraph N` counts the paragraphs of its story and a block
content control reads as `content control N`, counted apart. A story that
cannot be read is printed as `(not compared: ...)` instead of being counted as
equal.

`diff --json` writes a schema-1 record with the counts, each difference with
its story kind, locations, and texts, and the stories not compared.
`--exit-code` exits with 1 when the compared text differs and 2 on an
error or an unreadable story, as `diff` and `cmp` do. An unreadable story
remains in `not_compared` even when another story differs. Without
`--exit-code`, `diff` exits with 0 after printing a partial comparison.

`replace --expect N` publishes only when the run-aware replacement count is
exactly `N`. A mismatch exits unsuccessfully without creating or replacing the
requested output.

The count covers the text a reader sees in the body, tables,
content controls, headers, footers, footnotes, endnotes, and tracked insertions,
in the text boxes of the body, headers, and footers, in the labels of the
charts in the body, and inside smart tags, inline custom XML elements, and the
cached results of simple fields. Deleted text is not counted, except in a text
box inside a deleted run, which is replaced like any other text box. A text box
inside a note and the separators of the notes parts are not counted.

`convert` and `render` refuse an output file that already exists unless
`--force` is given. Even with `--force` they refuse their own input file under
any spelling of its path, and an output that is not a regular file, such as a
directory, a symbolic link, a FIFO, or a device like `/dev/null`. A run checks
every file it would write before it writes the first one, and publishes each
file only once it is complete, so a failed run leaves no truncated output.

`validate` exits unsuccessfully on a structural error: a relationship to a
missing part, a part without a content type, or a prefix that `mc:Ignorable`
or `mc:MustUnderstand` lists without a namespace declaration. Empty
paragraphs, heading level gaps, and missing metadata are warnings only.


Run `rdocx --help` or `rdocx <command> --help` for the complete option set.

## Ordered one-shot edits

`replace INPUT -p OLD -v NEW -o OUTPUT` retains the single-pair command.
`replace INPUT --map PAIRS.json -o OUTPUT --json` reads a UTF-8 JSON array:

```json
[
  {"placeholder": "A", "value": "B", "expect": 2},
  {"placeholder": "B", "value": "C", "expect": null}
]
```

Pairs run in array order, including repeated placeholders. Each later pair
searches the earlier result. An omitted or null `expect` is unrestricted.
An empty array succeeds with empty counts and publishes the unchanged document
to the required output. Unknown keys, wrong types, invalid replacement text,
negative or fractional counts and late expected-count mismatches refuse the
whole edit. Duplicate object keys follow serde_json's last occurrence wins
policy. `--map` cannot be combined with `-p`, `-v` or command-line `--expect`.

`--regex` uses native Rust regex syntax and capture substitutions such as `$1`.
Patterns match replaceable stretches of adjacent ordinary runs within one
paragraph. Wrapper, field-instruction, deleted-revision and paragraph boundaries
remain boundaries. Supported body, table, control, header, footer, note and
textbox stories participate. Chart labels are excluded. Zero-width matches are
skipped. Every pattern and replacement is validated before publishing one
candidate. JSON success carries schema1, operation, mode, ordered pairs with
index, placeholder, value, count and nullable expected count, total_count and
output. Errors go to stderr with a nonzero exit and no success JSON.

```bash
rdocx fields update input.docx --now 2026-10-10T08:30:00 -o fields.docx --json
rdocx fill input.docx --tag client=Ada --tag status=ready -o filled.docx --json
rdocx meta get input.docx --json
rdocx meta set input.docx --title Report --custom count=int:4 -o metadata.docx --json
rdocx images extract input.docx images --json
```

`fields update` atomically combines ordinary and deterministic layout cache
updates through the native facade. Omitted `--now` stays unresolved, preserving
stored DATE/TIME caches and retention diagnostics. Explicit dates use checked
Gregorian wall-clock values without time-zone conversion. The record reports
ordinary and per-kind page-cache counts, complete ordered evaluation facts,
combined diagnostics, total_updated and saved_cache_scope. The selected input
basename and canonical full path supply FILENAME context. Dynamic header/footer
PAGE and NUMPAGES caches retain their provider policy. Unsupported switches do
not become successful updates.

`fill --tag NAME=VALUE` selects tags only, never aliases. Values retain every
character after the first `=`. Duplicate tags, a missing tag, invalid text or
an unsupported bound owner refuse the complete batch. Each matching main-body
control and its supported custom XML binding changes together. Counts follow
the supplied tag order.

`meta set` changes supplied core fields and preserves unspecified metadata.
An empty core string removes that value. Custom syntax is `NAME=TYPE:VALUE`:
`str`, signed32 `int`, finite64 `float`, `bool` with true/false, `filetime` with
a checked Gregorian RFC3339 value and explicit timezone, or `empty` with no
payload. Strings retain additional `:` and `=` characters. Existing property
PID, fmtid, position and string subtype remain preserved. Duplicate custom
names and set/remove collisions refuse. `--remove-custom NAME` is a no-op when
absent. Reads include immutable typed property identity and preserved raw data.

`inspect --json` retains existing schema1 keys and adds word, Unicode scalar
character and deterministic page counts, their accepted-main-body scope and
layout diagnostics. Spaces contribute to character_count. Pictures come from
the existing physical occurrence provider, including supported related story
owners. Occurrences retain story identity, index path, ordinal, relationship,
EMU dimensions, description, title and decorative state.

`images extract` writes each unique internal media part once, in first
occurrence order. Generated flat names such as `image-000001.png` use a bounded
content-type mapping, with `bin` for an unknown image type. Shared media has one
file and multiple occurrence records. Distinct parts remain distinct even when
bytes are equal. Embedded bytes are preserved exactly. External linked images
are listed with status external, URL metadata and file null, without fetching.
Their native media_part_name is null and bytes/content_type are empty.
Unsupported physical ownership and missing internal media refuse through the
checked provider. Linked occurrence resizing retains its existing refusal.

Extraction rejects symlink directories and existing file collisions and has no
`--force`. It preloads and stages the complete output set before publishing.
Recoverable publication failure removes this invocation's files and a newly
created directory only if empty. Existing unrelated files remain. This is
recoverable rollback, not crash-atomic multi-file replacement. Empty extraction
creates no directory. JSON is printed only after publication.

Every mutation requires an output and refuses existing outputs, input identity,
symlink aliases and hard links. Native package-class, VBA and signature checks
remain authoritative. No success report precedes publication.

## Licensed font embedding

Use caller-owned font files only after confirming their embedding license.
Embedding requires explicit authorization and exact license identity. Commands
write a new output and refuse input/font aliases or existing destinations.

```bash
rdocx font embed input.docx --name "My Face" --font ./MyFace-Regular.ttf \
  --license-authorized --license-identity "SIL Open Font License 1.1" \
  --output embedded.docx
rdocx font list embedded.docx --json
rdocx font settings embedded.docx --embed-true-type-fonts false --output disabled.docx
rdocx font remove embedded.docx --name "My Face" --kind regular --output removed.docx
```

The four CLI kinds are regular, bold, italic and bold-italic. Settings accept
true, false or none, with none removing the owned value. Listing reports actual
native part, format and byte-length facts in the schema1 JSON envelope.
Word also accepts --font-key and --subsetted, and its settings include
--embed-system-fonts and --save-subset-fonts.
