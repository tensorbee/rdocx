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
- Package-preserving edits cover the final S74 paragraph, run, typography,
  table, section, settings, field, form, equation, and drawing surface.

## Measured footprint and speed

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rdocx-cli | 73,938 compressed bytes, 329,766 member bytes, 8 members | 0.16.0 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --no-verify` | Tracked `rdocx-cli` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-09 |

## Use it when

Use this crate for shell automation. Use the
[`rdocx`](https://docs.rs/rdocx) library when these operations need to run
inside a Rust application.

## Relationship

The binary delegates document behavior to `rdocx` and shares path and JSON
conventions with `rpptx-cli` through `oxml-cli-support`.

## Example

```sh
cargo install rdocx-cli --version '^0.16.0'

rdocx inspect report.docx
rdocx text report.docx
rdocx text report.docx --json
rdocx layout report.docx --json
rdocx replace template.docx -p TOKEN -v ready --expect 1 -o report.docx
rdocx convert report.docx --to pdf -o report.pdf
rdocx convert redline.docx --to pdf --revision-view tracked -o redline.pdf
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
default. `--revision-view tracked` shows both sides of tracked changes.
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
Empty paragraphs, heading level gaps, and missing metadata are warnings. So
are a section whose `w:pgSz` lacks a width or a height, which leaves each
consumer to guess the page, and an even-page header or footer while
`w:evenAndOddHeaders` is off, which Word and Google Docs ignore.

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
