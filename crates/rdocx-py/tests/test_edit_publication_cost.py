"""Which whole-document passes a mutating call makes when it publishes, and that the direct routes agree with the
checked ones.

A whole-document pass is work in proportion to the document whatever the edit: a staging clone of the document, a
walk of a story to resolve a route, a comparison of two staged saves, a publication that matches the whole owner
inventory, a rebuild of that inventory, a reopen of the package. The debug build of the extension counts them
(`_word_publication_work`, `_word_whole_body_work`, `_reopen_work`) outside the guards that check a direct route, so
the counts are those of the call. No timing is read.

FLAT names the calls that make no such pass and do the same counted work in a body four times larger. A cell edit
also does the same counted work in a table twenty times larger and beside a cell of a hundred thousand characters.
The counted work includes the bytes of the blocks whose owners the checked reader reads again
(`_word_piece_bytes`), so a cell edit that read its whole table would show. GROWING names the others with the passes
each one still makes. The flat test is a strict expected failure for a GROWING call, and its passes are asserted
exactly, so a change that makes a call flat or removes one of its passes fails until the entry
moves or shrinks, and a change that adds a pass fails as well.

`_word_force_checked_route` makes every publication resolve its route by a story walk and compare staged saves, and
every properties edit stage the document: the checked whole-document routes the direct ones replace. Each call must
give the same result, the same saved parts and the same held handles on both.
"""
import datetime
import re
from io import BytesIO
from types import SimpleNamespace
from zipfile import ZipFile

import pytest

import rdocx
from rdocx import _rdocx

pytestmark = pytest.mark.skipif(
    not hasattr(_rdocx, "_word_publication_work"),
    reason="the publication counters and the route switch exist in debug builds only",
)

BODY = 8
STAMP = datetime.datetime(2026, 10, 1, 12, tzinfo=datetime.timezone.utc)


def build(paragraphs, rows=2, columns=2, last_cell=""):
    """A body of `paragraphs` paragraphs, then a table whose first cell holds "gamma" and whose last cell holds
    `last_cell`, and a paragraph of the runs "alpha " and "beta"."""
    document = rdocx.Document()
    for index in range(paragraphs):
        document.add_paragraph(f"Body paragraph {index}")
    table = document.add_table(rows, columns)
    table.cell(0, 0).text = "gamma"
    if last_cell:
        table.cell(rows - 1, columns - 1).text = last_cell
    document.add_paragraph("alpha ").add_run("beta")
    return document


def targets(document):
    paragraph = document.paragraphs[-1]
    return SimpleNamespace(document=document, paragraph=paragraph, run=paragraph.runs[1], table=document.tables[0])


def setter(owner, name, value):
    """Set `name` on the object `owner` gives, taken before the call is counted."""
    return lambda t: (lambda target=owner(t): setattr(target, name, value))


def method(owner, name, *arguments):
    return lambda t: (lambda target=owner(t): getattr(target, name)(*arguments))


def core(t):
    return t.document.core_properties


def app(t):
    return t.document.app_properties


def custom(t):
    return t.document.custom_properties


def with_client(t):
    properties = t.document.custom_properties
    properties.set("Client", "ACME")
    return properties


def cell(t):
    return t.table.cell(0, 0)


def cell_with_properties(t):
    """A cell that already writes its properties, so a property edit moves none of its items."""
    t.table.cell(0, 0).width = rdocx.Inches(1)
    return t.document.tables[0].cell(0, 0)


def cell_with_table(t):
    """A cell that holds a table, which the selected-table candidate of a cell edit declines."""
    t.table.cell(0, 0).add_table(1, 1)
    return t.document.tables[0].cell(0, 0)


def bookmarked(t):
    start, end = rdocx.RunPosition(body_index=0, run_index=0), rdocx.RunPosition(body_index=0, run_index=1)
    t.document.add_bookmark("mark", rdocx.RunRange(start=start, end=end))
    return t.document.paragraphs[-1]


def with_chart(t):
    data = rdocx.CategoryChartData()
    data.categories = ["a", "b"]
    data.add_series("series", (1.0, 2.0))
    chart = t.document.add_chart(rdocx.XL_CHART_TYPE.COLUMN_CLUSTERED, data, rdocx.Inches(3), rdocx.Inches(2))
    return lambda: t.document.set_chart_legend(chart, visible=True, position="b")


CALLS = {
    "AppProperties.application": setter(app, "application", "An application"),
    "AppProperties.application_version": setter(app, "application_version", "16.0000"),
    "AppProperties.company": setter(app, "company", "A company"),
    "AppProperties.manager": setter(app, "manager", "A manager"),
    "AppProperties.presentation_format": setter(app, "presentation_format", "A format"),
    "AppProperties.template": setter(app, "template", "A template"),
    "CustomProperties.set": method(custom, "set", "Client", "ACME"),
    "CustomProperties.__setitem__": method(custom, "__setitem__", "Client", "ACME"),
    "CustomProperties.delete": method(with_client, "delete", "Client"),
    "CustomProperties.__delitem__": method(with_client, "__delitem__", "Client"),
    "Font.bold": setter(lambda t: t.run.font, "bold", True),
    "Paragraph.alignment": setter(lambda t: t.paragraph, "alignment", rdocx.WD_ALIGN_PARAGRAPH.CENTER),
    "ParagraphFormat.left_indent": setter(lambda t: t.paragraph.paragraph_format, "left_indent", rdocx.Inches(1)),
    "Cell.text": setter(cell, "text", "12.50"),
    "Cell.shading": setter(cell, "shading", "D9E2F3"),
    "Cell.vertical_alignment": setter(cell, "vertical_alignment", rdocx.WD_CELL_VERTICAL_ALIGNMENT.CENTER),
    "Cell.width": setter(cell, "width", rdocx.Inches(1)),
    "Cell.set_border": lambda t: (lambda target=cell(t): target.set_border("top", "single", size=4, color="000000")),
    "Cell.set_margins": lambda t: (lambda target=cell(t): target.set_margins(top=10, right=10, bottom=10, left=10)),
    "Paragraph.text of a cell paragraph": setter(lambda t: t.table.cell(0, 0).paragraphs[0], "text", "12.50"),
    "Paragraph.add_run": method(lambda t: t.paragraph, "add_run", " gamma"),
    "Paragraph.add_hyperlink": method(lambda t: t.paragraph, "add_hyperlink", "link", "https://example.com/a"),
    "Paragraph.add_internal_hyperlink": method(bookmarked, "add_internal_hyperlink", "link", "mark"),
    "Paragraph.add_equation": lambda t: (lambda target=t.paragraph: target.add_equation("x^2", format="latex")),
    "Cell.text of a cell that holds a table": setter(cell_with_table, "text", "12.50"),
    "Document.set_chart_legend": with_chart,
    "Paragraph.text": setter(lambda t: t.paragraph, "text", "changed"),
    "Run.text": setter(lambda t: t.run, "text", "changed"),
    "Run.add_break": method(lambda t: t.run, "add_break"),
    "Run.add_tab": method(lambda t: t.run, "add_tab"),
    "Table.style": setter(lambda t: t.table, "style", "Table Grid"),
    "Table.add_row": method(lambda t: t.table, "add_row"),
    "Table.remove_row": method(lambda t: t.table, "remove_row", 1),
    "Settings.track_revisions": setter(lambda t: t.document.settings, "track_revisions", True),
    "Document.add_paragraph": method(lambda t: t.document, "add_paragraph", "new"),
    "Document.add_table": method(lambda t: t.document, "add_table", 1, 1),
}
for name in ("author", "category", "comments", "content_status", "identifier", "keywords", "language",
             "last_modified_by", "subject", "title", "version"):
    CALLS[f"CoreProperties.{name}"] = setter(core, name, f"A {name}")
for name in ("created", "modified", "last_printed"):
    CALLS[f"CoreProperties.{name}"] = setter(core, name, STAMP)
CALLS["CoreProperties.revision"] = setter(core, "revision", 7)

# Cell edits that read the facts of their cell alone, whatever the body and the table.
CELL_CALLS = ("Cell.text", "Cell.shading", "Cell.vertical_alignment", "Cell.width", "Cell.set_border",
              "Cell.set_margins", "Paragraph.text of a cell paragraph")

# Calls that make no whole-document pass and do the same counted work whatever the body.
FLAT = (
    "AppProperties.application", "AppProperties.application_version", "AppProperties.company",
    "AppProperties.manager", "AppProperties.presentation_format", "AppProperties.template",
    "CoreProperties.author", "CoreProperties.category", "CoreProperties.comments", "CoreProperties.content_status",
    "CoreProperties.created", "CoreProperties.identifier", "CoreProperties.keywords", "CoreProperties.language",
    "CoreProperties.last_modified_by", "CoreProperties.last_printed", "CoreProperties.modified",
    "CoreProperties.revision", "CoreProperties.subject", "CoreProperties.title", "CoreProperties.version",
    "CustomProperties.__delitem__", "CustomProperties.__setitem__", "CustomProperties.delete",
    "CustomProperties.set",
    "Font.bold", "Paragraph.alignment", "ParagraphFormat.left_indent",
) + CELL_CALLS

# Calls whose work grows with the body, with the whole-document passes each one makes. A walk in an entry resolves
# the story address its native edit takes, not the route of its publication. An empty entry makes no pass but visits
# every owner record.
GROWING = {
    "Paragraph.add_run": {"clones": 1, "publications": 1},
    "Paragraph.add_hyperlink": {"clones": 2, "publications": 1},
    "Paragraph.add_internal_hyperlink": {"clones": 6, "walks": 1, "publications": 1, "rebuilds": 1, "reopens": 1},
    "Paragraph.add_equation": {"clones": 1, "publications": 1},
    "Cell.text of a cell that holds a table": {"clones": 4, "publications": 1, "rebuilds": 1, "reopens": 1},
    "Document.set_chart_legend": {"clones": 5, "compares": 1, "publications": 1, "rebuilds": 1, "reopens": 1},
    "Paragraph.text": {"clones": 1, "publications": 1},
    "Run.text": {"clones": 1, "publications": 1},
    "Run.add_break": {"clones": 1},
    "Run.add_tab": {"clones": 1},
    "Table.style": {"clones": 1},
    "Table.add_row": {"clones": 3, "publications": 1, "rebuilds": 1, "reopens": 1},
    "Table.remove_row": {"clones": 4, "publications": 1, "rebuilds": 1, "reopens": 1},
    "Settings.track_revisions": {"clones": 1},
    "Document.add_paragraph": {},
    "Document.add_table": {},
}


def work():
    """The whole-document passes and the other counted work since the last read. Reading resets every counter."""
    clones, walks, compares, publications = _rdocx._word_publication_work()
    rebuilds, source_reads = _rdocx._word_whole_body_work()
    passes = {"clones": clones, "walks": walks, "compares": compares, "publications": publications,
              "rebuilds": rebuilds, "reopens": _rdocx._reopen_work()[0]}
    rest = {"source_reads": source_reads, "piece_reads": _rdocx._word_piece_work(),
            "piece_bytes": _rdocx._word_piece_bytes(), "alignments": _rdocx._word_piece_alignments(),
            "owner_visits": _rdocx._word_splice_work()[0], "staging": list(_rdocx._staging_work())}
    return {name: count for name, count in passes.items() if count}, rest


def measure(name, paragraphs, rows=2, columns=2, last_cell=""):
    act = CALLS[name](targets(build(paragraphs, rows, columns, last_cell)))
    work()
    act()
    return work()


@pytest.mark.parametrize("name", [
    pytest.param(name, marks=pytest.mark.xfail(strict=True, reason="its work still grows with the body"))
    if name in GROWING else name
    for name in sorted(CALLS)
])
def test_call_makes_no_whole_document_pass_and_the_same_work_in_a_larger_body(name):
    small, large = measure(name, BODY), measure(name, 4 * BODY)
    assert small[0] == {} and large[0] == {}
    assert small[1] == large[1]


@pytest.mark.parametrize("name", CELL_CALLS)
def test_cell_call_makes_no_whole_document_pass_and_the_same_work_in_a_larger_table(name):
    small, large = measure(name, BODY, 2, 5), measure(name, BODY, 40, 5)
    assert small[0] == {} and large[0] == {}
    assert small[1] == large[1]
    assert small[1]["piece_bytes"] > 0


@pytest.mark.parametrize("name", CELL_CALLS)
def test_cell_call_does_the_same_work_beside_a_large_cell(name):
    small, large = measure(name, BODY, 2, 5, "x"), measure(name, BODY, 2, 5, "x" * 100_000)
    assert small[0] == {} and large[0] == {}
    assert small[1] == large[1]


@pytest.mark.parametrize("name", sorted(GROWING))
def test_growing_call_makes_only_its_declared_whole_document_passes(name):
    for paragraphs in (BODY, 4 * BODY):
        assert measure(name, paragraphs)[0] == GROWING[name]


def test_every_call_is_flat_or_growing():
    assert not set(FLAT) & set(GROWING)
    assert sorted(set(FLAT) | set(GROWING)) == sorted(CALLS)


def test_run_text_that_changes_nothing_publishes_nothing_and_saves_nothing():
    document = build(BODY)
    run = document.paragraphs[-1].runs[1]
    held = document.paragraphs[-1].runs[0]
    work()
    run.text = "beta"
    assert work()[0] == {"clones": 1}
    assert (run.text, held.text) == ("beta", "alpha ")


def test_run_text_of_a_paragraph_that_is_not_plain_keeps_the_staged_save_comparison():
    document = build(BODY)
    paragraph = document.paragraphs[-1]
    paragraph.add_hyperlink("link", "https://example.com/a")
    run = document.paragraphs[-1].runs[1]
    work()
    run.text = "changed"
    passes = work()[0]
    assert passes["compares"] == 1 and passes["publications"] == 1
    assert document.paragraphs[-1].text == "alpha changedlink"


def test_run_text_publishes_beside_content_that_no_save_accepts_and_save_reports_it():
    def document():
        result = rdocx.Document()
        result.add_paragraph("first")
        result.add_paragraph("alpha ").add_run("beta")
        result.paragraphs[0].text = "a\x00b"
        return result

    for text, expected in (("changed", "alpha changed"), ("beta", "alpha beta")):
        direct = document()
        direct.paragraphs[1].runs[1].text = text
        assert direct.paragraphs[1].text == expected
        with pytest.raises(rdocx.PackageError, match="U\\+0000"):
            direct.save(BytesIO())
        checked = document()
        run = checked.paragraphs[1].runs[1]
        previous = _rdocx._word_force_checked_route(True)
        try:
            with pytest.raises(rdocx.PackageError, match="U\\+0000"):
                run.text = text
        finally:
            _rdocx._word_force_checked_route(previous)
        assert checked.paragraphs[1].text == "alpha beta"


TEXT_BOX = "<w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p></w:txbxContent>"
TEXT_BOX_BODY = (
    "<w:body><w:p><w:r><w:t>anchor </w:t></w:r><w:r><mc:AlternateContent "
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:v="urn:schemas-microsoft-com:vml">'
    '<mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>'
    '<wp:docPr id="1" name="box"/><a:graphic>'
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
    f"<wps:wsp><wps:txbx>{TEXT_BOX}</wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>"
    f'</mc:Choice><mc:Fallback><w:pict><v:shape id="fallback"><v:textbox>{TEXT_BOX}</v:textbox></v:shape>'
    "</w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p><w:sectPr/></w:body>"
)


def with_body(body):
    """A document opened from a package whose main part holds `body`."""
    result = BytesIO()
    with ZipFile(BytesIO(rdocx.Document().to_bytes())) as package, ZipFile(result, "w") as output:
        for member in package.infolist():
            data = package.read(member.filename)
            if member.filename == "word/document.xml":
                data = re.sub(rb"<w:body>.*</w:body>", body.encode(), data, flags=re.S)
            output.writestr(member, data)
    return rdocx.Document.from_bytes(result.getvalue())


def text_box_format():
    """The document and the format of the paragraph in the text box of TEXT_BOX_BODY."""
    document = with_body(TEXT_BOX_BODY)
    paragraph = next(item.paragraph for item in document.story_items
                     if item.kind == "paragraph" and item.story.kind == "text_box")
    return document, paragraph.paragraph_format


def test_text_box_paragraph_format_walks_for_its_edit_address_alone():
    saved = {}
    for forced, walks in ((False, 2), (True, 3)):
        document, paragraph_format = text_box_format()
        work()
        previous = _rdocx._word_force_checked_route(forced)
        try:
            paragraph_format.left_indent = rdocx.Inches(1)
        finally:
            _rdocx._word_force_checked_route(previous)
        passes = work()[0]
        assert (passes["walks"], passes["compares"], passes["publications"]) == (walks, 1, 1)
        saved[forced] = parts(document)
    assert saved[False] == saved[True]


def cell_xml(content, properties=""):
    return f"<w:tc>{properties}{content}</w:tc>"


def text(value):
    return f"<w:p><w:r><w:t>{value}</w:t></w:r></w:p>"


def table_xml(columns, *rows):
    grid = "".join('<w:gridCol w:w="1000"/>' for _ in range(columns))
    return (f"<w:body><w:tbl><w:tblGrid>{grid}</w:tblGrid>" + "".join(f"<w:tr>{row}</w:tr>" for row in rows)
            + "</w:tbl><w:p/><w:sectPr/></w:body>")


# Opened tables whose cells a lone cell might read differently from the whole table, with the cells edited.
CELL_TABLES = {
    "merged cells": (table_xml(
        3,
        cell_xml(text("a"), '<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>')
        + cell_xml(text("b"), '<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>'),
        cell_xml("<w:p/>", "<w:tcPr><w:vMerge/></w:tcPr>") + cell_xml(text("c")) + cell_xml(text("d")),
    ), ((0, 1), (1, 0))),
    "a bookmark across two cells": (table_xml(
        2,
        cell_xml('<w:p><w:bookmarkStart w:id="1" w:name="span"/><w:r><w:t>left</w:t></w:r></w:p>')
        + cell_xml('<w:p><w:r><w:t>right</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>'),
    ), ((0, 0), (0, 1))),
    "a field in one cell": (table_xml(
        2,
        cell_xml(text("left"))
        + cell_xml('<w:p><w:fldSimple w:instr=" SEQ Table "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>'),
    ), ((0, 0), (0, 1))),
    "row attributes and properties, and markers between cells and rows": (
        '<w:body><w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>'
        '<w:tr w:rsidR="00AB12CD"><w:trPr><w:cantSplit/></w:trPr>' + cell_xml(text("a"))
        + '<w:bookmarkStart w:id="2" w:name="rows"/>' + cell_xml(text("b")) + "</w:tr>"
        + '<w:bookmarkEnd w:id="2"/><w:tr>' + cell_xml(text("c")) + cell_xml(text("d")) + "</w:tr>"
        "</w:tbl><w:p/><w:sectPr/></w:body>",
        ((0, 0), (0, 1), (1, 0)),
    ),
    "an empty cell property element": (table_xml(
        2, cell_xml(text("left"), "<w:tcPr/>") + cell_xml(text("right")),
    ), ((0, 0),)),
}

CELL_EDITS = {
    "Cell.shading": lambda cell: setattr(cell, "shading", "D9E2F3"),
    "Cell.text": lambda cell: setattr(cell, "text", "12.50"),
    "Paragraph.text of a cell paragraph": lambda cell: setattr(cell.paragraphs[0], "text", "12.50"),
}


@pytest.mark.parametrize("table", sorted(CELL_TABLES))
@pytest.mark.parametrize("edit", sorted(CELL_EDITS))
def test_cell_edit_of_an_opened_table_gives_the_result_handles_and_saved_parts_of_the_checked_route(table, edit):
    body, coordinates = CELL_TABLES[table]
    for row, column in coordinates:
        results = {}
        for forced in (False, True):
            document = with_body(body)
            held = {}
            for r, cells in enumerate(document.tables[0].rows):
                for c, current in enumerate(cells.cells):
                    held[f"cell {r} {c}"] = current
                    held[f"cell {r} {c} paragraph"] = current.paragraphs[0]
                    held.update({f"cell {r} {c} run": run for run in current.paragraphs[0].runs[:1]})
            target = document.tables[0].rows[row].cells[column]
            work()
            previous = _rdocx._word_force_checked_route(forced)
            try:
                CELL_EDITS[edit](target)
            finally:
                _rdocx._word_force_checked_route(previous)
            # Each shape stays on the local route, which makes no whole-document pass.
            assert (work()[0] == {}) is not forced
            results[forced] = ({name: probe(handle) for name, handle in held.items()}, parts(document))
        assert results[False] == results[True], (row, column)


def parts(document):
    stream = BytesIO()
    document.save(stream)
    with ZipFile(BytesIO(stream.getvalue())) as package:
        return {name: package.read(name) for name in package.namelist()}


def probe(handle):
    """What a held handle reads after the call, or the error that says it is stale."""
    try:
        for name in ("text", "bold", "left_indent", "page_width"):
            if hasattr(type(handle), name):
                return "read", repr(getattr(handle, name))
        return "read", len(handle.rows)
    except Exception as error:  # noqa: BLE001 - the error is the result compared
        return type(error).__name__, str(error)


def held_handles(document):
    first, last = document.paragraphs[0], document.paragraphs[-1]
    table = document.tables[0]
    return {
        "another paragraph": first, "its run": first.runs[0], "the paragraph": last, "its first run": last.runs[0],
        "its second run": last.runs[1], "its font": last.runs[1].font, "its format": last.paragraph_format,
        "the table": table, "a cell": table.cell(0, 0), "a cell paragraph": table.cell(1, 1).paragraphs[0],
        "the first cell's paragraph": table.cell(0, 0).paragraphs[0],
        "its cell run": table.cell(0, 0).paragraphs[0].runs[0], "a cell of the same row": table.cell(0, 1),
        "the paragraph of that cell": table.cell(0, 1).paragraphs[0],
        "the section": document.section(0),
    }


def outcome(prepare, forced):
    document = build(BODY)
    held = held_handles(document)
    act = prepare(targets(document))
    previous = _rdocx._word_force_checked_route(forced)
    try:
        try:
            value = act()
            result = ("returned", type(value).__name__, getattr(value, "text", value if value is None else None))
        except Exception as error:  # noqa: BLE001 - the error is the result compared
            result = (type(error).__name__, str(error))
    finally:
        _rdocx._word_force_checked_route(previous)
    return result, {name: probe(handle) for name, handle in held.items()}, parts(document)


DIFFERENTIAL = dict(CALLS)
DIFFERENTIAL.update({
    "Run.text unchanged": setter(lambda t: t.run, "text", "beta"),
    "Run.text with a character no save accepts": setter(lambda t: t.run, "text", "a\x00b"),
    "Run.text with a tab and a line break": setter(lambda t: t.run, "text", "a\tb\nc"),
    "Run.text emptied": setter(lambda t: t.run, "text", ""),
    "AppProperties.application_version refused": setter(app, "application_version", "sixteen"),
    "CustomProperties.delete of a missing name": method(custom, "__delitem__", "Missing"),
    "Cell.shading of a cell that writes its properties": setter(cell_with_properties, "shading", "D9E2F3"),
    "Cell.text with a tab and a line break": setter(cell, "text", "a\tb\nc"),
    "Cell.text with a character no save accepts": setter(cell, "text", "a\x00b"),
    "Cell.set_margins refused": lambda t: (
        lambda target=cell(t): target.set_margins(top=-1, right=0, bottom=0, left=0)),
})


@pytest.mark.parametrize("name", sorted(DIFFERENTIAL))
def test_direct_route_gives_the_result_handles_and_saved_parts_of_the_checked_route(name):
    direct, checked = outcome(DIFFERENTIAL[name], False), outcome(DIFFERENTIAL[name], True)
    assert direct[0] == checked[0]
    assert direct[1] == checked[1]
    assert direct[2] == checked[2]


@pytest.mark.parametrize("name", [
    "Run.text with a character no save accepts", "AppProperties.application_version refused",
    "CustomProperties.delete of a missing name", "Cell.text with a character no save accepts",
    "Cell.set_margins refused",
])
def test_refused_edit_changes_nothing(name):
    untouched = build(BODY)
    result, held, saved = outcome(DIFFERENTIAL[name], False)
    assert result[0] != "returned"
    assert saved == parts(untouched)
    assert all(kind == "read" for kind, _ in held.values())


def test_forced_checked_route_makes_the_passes_the_direct_routes_replace():
    forced = {}
    previous = _rdocx._word_force_checked_route(True)
    try:
        for name in ("CoreProperties.title", "Paragraph.add_run", "Run.text", "Table.add_row", "Cell.shading",
                     "Cell.text"):
            forced[name] = measure(name, BODY)[0]
    finally:
        _rdocx._word_force_checked_route(previous)
    assert forced["CoreProperties.title"] == {"clones": 1}
    assert forced["Paragraph.add_run"]["walks"] == 2
    assert (forced["Run.text"]["walks"], forced["Run.text"]["compares"]) == (1, 1)
    assert forced["Table.add_row"]["walks"] == 1
    assert forced["Cell.shading"] == {"clones": 1}
    assert (forced["Cell.text"]["walks"], forced["Cell.text"]["publications"]) == (1, 1)
