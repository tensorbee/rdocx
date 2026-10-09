# rpptx-py

Native rpptx for Python

`rpptx` gives Python applications a native PPTX workflow for creating,
opening, editing, and saving PowerPoint-compatible presentations. It also
supports comments, speaker notes, and deterministic rendering. It works
directly with OOXML packages and does not require Microsoft PowerPoint,
LibreOffice, a conversion service, or a Java or .NET runtime.

The package runs the Rust `rpptx` presentation engine locally, exposes a typed
Python API, and ships a default template for from-scratch deck creation. It
renders slides and speaker notes without a remote service, supports review
through notes and modern comment threads, and preserves safe package content
outside the focused Python editing surface.

## Installation

Install the current release from PyPI:

```sh
python -m pip install rpptx
```

`rpptx` requires CPython 3.9 or newer. Its platform wheels use the Python 3.9
stable ABI, so one wheel supports multiple compatible Python versions. Source
installation requires a Rust toolchain and maturin.

## Quick start

```python
from rpptx import Presentation
from rpptx.util import Inches

presentation = Presentation()
slide = presentation.slides.add_slide(presentation.slide_layouts[6])
text_box = slide.shapes.add_textbox(
    Inches(1), Inches(1), Inches(8), Inches(1)
)
text_box.text = "Quarterly review"
presentation.save("review.pptx")

pdf = presentation.to_pdf()
with open("review.pdf", "wb") as output:
    output.write(pdf)
```

## Capabilities

- File and byte-based PPTX input and output.
- Slide layouts, slides, placeholders, shapes, text frames, paragraphs, runs,
  pictures, preset shapes, and tables.
- Deterministic PDF and PNG output for slides and speaker notes.
- Formatting-preserving text replacement across slides and speaker notes,
  on one slide or in one text frame, with an optional expected count that
  leaves the deck unchanged on a mismatch.
- Package and PresentationML validation that returns the issues
  `rpptx validate` reports as typed values.
- Speaker-note text plus modern comment authors, threads, replies, thread
  resolution and removal, and ordered comment movement.
- Master, layout, placeholder, theme, shape, chart, media, and relationship
  state remains inside the native presentation engine during package edits.
- Read speaker-note text and inspect or mutate modern comment threads.
- Python collections with negative indexes, slices, iteration, and explicit
  stale-handle errors after structural changes.
- Table cell merge and split, cell fills, margins, and borders, and row
  heights.
- Table rows and columns added with `table.rows.add_row()` and
  `table.columns.add_column()` and removed with `remove`, keeping merged cells
  and the frame size in step.
- Picture crop, read and written as in python-pptx.
- Shape z-order through `slide.shapes.move(from_, to)`.
- Group members added through `group.shapes`, with the group refit to its
  members as in python-pptx.
- Run hyperlinks, read, added, retargeted, and removed as in python-pptx.
- Shape click actions, a hyperlink or a slide jump, through `shape.click_action`
  as in python-pptx.
- `shape.is_placeholder` and `shape.placeholder_format` with `PP_PLACEHOLDER`
  as in python-pptx.
- Slide numbers, footer and date on every slide in one call,
  `prs.set_header_footer(slide_number=True, footer="ACME", hide_on_title=True)`,
  per slide through `slide.header_footer`, and `paragraph.add_field("slidenum")`.
- `prs.slide_master.theme` colours and fonts, and `slide.transition` with type,
  direction, duration, advance timing and `apply_to_all()`.

## Use it when

Use `rpptx` when a Python service, desktop application, or automation task
needs to create, inspect, update, comment on, or render presentations locally.
It is suited to generated decks, template workflows, review, and previews.

## Relationship

The Python API is intentionally focused. It covers the documented `rpptx`
surface and the reviewed python-pptx Getting Started workflows, but it is not a
drop-in replacement for every python-pptx method or private lxml object.

## Example

Open and inspect an existing presentation:

```python
from rpptx import Presentation

presentation = Presentation("deck.pptx")
print(len(presentation.slides))
png = presentation.render_slide_to_png(0)
for slide in presentation.slides:
    print(slide.notes_text or "")
```

## Type checking

The distribution includes hand-written extension stubs and a `py.typed`
marker. Editors and type checkers can inspect concrete presentation, slide,
shape, text, table, notes, and comment types.

```sh
python -m pip install mypy
python -m mypy your_application.py
```

The release gate validates the installed package with strict mypy and
`stubtest` in addition to its runtime suite.

## Project links

- [Source repository](https://github.com/tensorbee/rdocx)
- [Issue tracker](https://github.com/tensorbee/rdocx/issues)
- [Changelog](https://github.com/tensorbee/rdocx/blob/main/CHANGELOG.md)
- [Binding specification](https://github.com/tensorbee/rdocx/blob/main/docs/hld/10-bindings-spec.md)

## License

Licensed under either the
[MIT License](https://github.com/tensorbee/rdocx/blob/main/LICENSE) or the
[Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0), at your
option.
