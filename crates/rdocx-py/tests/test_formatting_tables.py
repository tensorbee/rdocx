from rdocx.enum.dml import MSO_COLOR_TYPE
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from rdocx import RGBColor


def _document_xml(document_bytes):
    with ZipFile(BytesIO(document_bytes)) as package:
        return package.read("word/document.xml").decode()


def _replace_document_xml(document_bytes, old, new):
    source_bytes = BytesIO(document_bytes)
    output_bytes = BytesIO()
    with ZipFile(source_bytes) as source, ZipFile(
        output_bytes, "w", compression=ZIP_DEFLATED
    ) as output:
        for member in source.infolist():
            contents = source.read(member.filename)
            if member.filename == "word/document.xml":
                assert old in contents
                contents = contents.replace(old, new)
            output.writestr(member, contents)
    return output_bytes.getvalue()


def test_unset_run_bold_is_none():
    from rdocx import Document

    document = Document()
    run = document.add_paragraph("").add_run("plain")

    assert run.font.bold is None


def test_run_bool_tristate_round_trips():
    from rdocx import Document

    document = Document()
    font = document.add_paragraph("").add_run("formatted").font
    font.bold = False
    font.italic = True
    font.strike = False

    reopened = Document.from_bytes(document.to_bytes())
    font = reopened.paragraphs[0].runs[0].font
    assert font.bold is False
    assert font.italic is True
    assert font.strike is False


def test_none_clears_direct_formatting():
    from rdocx import Document, Inches

    document = Document()
    font = document.add_paragraph("").add_run("formatted").font
    font.bold = True
    font.italic = False
    font.underline = True
    font.bold = None
    font.italic = None
    font.underline = None
    paragraph_format = document.paragraphs[0].paragraph_format
    paragraph_format.keep_with_next = True
    paragraph_format.keep_together = True
    paragraph_format.page_break_before = True
    paragraph_format.widow_control = True
    paragraph_format.first_line_indent = Inches(-0.25)
    paragraph_format.keep_with_next = None
    paragraph_format.keep_together = None
    paragraph_format.page_break_before = None
    paragraph_format.widow_control = None
    paragraph_format.first_line_indent = None

    reopened = Document.from_bytes(document.to_bytes())
    font = reopened.paragraphs[0].runs[0].font
    paragraph_format = reopened.paragraphs[0].paragraph_format
    assert font.bold is None
    assert font.italic is None
    assert font.underline is None
    assert paragraph_format.keep_with_next is None
    assert paragraph_format.keep_together is None
    assert paragraph_format.page_break_before is None
    assert paragraph_format.widow_control is None
    assert paragraph_format.first_line_indent is None


def test_font_and_paragraph_format_values_round_trip():
    from rdocx import (
        Document,
        Inches,
        Pt,
        RGBColor,
        WD_ALIGN_PARAGRAPH,
        WD_UNDERLINE,
    )

    document = Document()
    paragraph = document.add_paragraph("")
    run = paragraph.add_run("formatted")
    run.font.name = "Aptos"
    run.font.size = Pt(12)
    run.font.color = RGBColor(0x12, 0x34, 0x56)
    run.font.underline = WD_UNDERLINE.DOT_DASH
    paragraph = document.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.left_indent = Inches(0.5)
    paragraph.paragraph_format.right_indent = Inches(0.25)
    paragraph.paragraph_format.first_line_indent = Inches(-0.25)
    paragraph.paragraph_format.line_spacing = Pt(18)
    paragraph.paragraph_format.keep_with_next = False
    paragraph.paragraph_format.keep_together = True
    paragraph.paragraph_format.page_break_before = False
    paragraph.paragraph_format.widow_control = True

    reopened = Document.from_bytes(document.to_bytes())
    paragraph = reopened.paragraphs[0]
    font = paragraph.runs[0].font
    paragraph_format = paragraph.paragraph_format
    assert font.name == "Aptos"
    assert font.size == Pt(12)
    assert font.color.rgb == RGBColor(0x12, 0x34, 0x56)
    assert font.underline == WD_UNDERLINE.DOT_DASH
    assert paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT
    assert paragraph_format.alignment == WD_ALIGN_PARAGRAPH.RIGHT
    assert paragraph_format.space_before == Pt(3)
    assert paragraph_format.space_after == Pt(6)
    assert paragraph_format.left_indent == Inches(0.5)
    assert paragraph_format.right_indent == Inches(0.25)
    assert paragraph_format.first_line_indent == Inches(-0.25)
    assert paragraph_format.line_spacing == Pt(18)
    assert paragraph_format.keep_with_next is False
    assert paragraph_format.keep_together is True
    assert paragraph_format.page_break_before is False
    assert paragraph_format.widow_control is True


def test_all_approved_underline_codes_round_trip():
    from rdocx import Document, WD_UNDERLINE

    for style in (WD_UNDERLINE.DOT_DASH, WD_UNDERLINE.DOT_DOT_DASH):
        document = Document()
        font = document.add_paragraph("").add_run("underlined").font
        font.underline = style
        reopened = Document.from_bytes(document.to_bytes())
        assert reopened.paragraphs[0].runs[0].font.underline == style


def test_format_subhandles_become_stale_after_structure_change():
    from rdocx import Document, StaleElementError

    document = Document()
    paragraph = document.add_paragraph("")
    font = paragraph.add_run("held").font
    paragraph_format = document.paragraphs[0].paragraph_format

    spacing = paragraph_format.space_after
    document.add_paragraph("retires the earlier run")

    with pytest.raises(StaleElementError):
        _ = font.bold
    assert paragraph_format.space_after == spacing
    assert paragraph.text == "held"


def test_nested_paragraph_stale_error_names_complete_recovery_path():
    from rdocx import Document, StaleElementError

    document = Document()
    cell = document.add_table(rows=1, cols=1).rows[0].cells[0]
    nested = cell.add_paragraph("held")
    retired_run = nested.runs[0]

    document.add_paragraph("retires the earlier nested run")
    assert nested.text == "held"

    with pytest.raises(
        StaleElementError,
        match=(
            r"doc\.tables\[0\]\.rows\[0\]\.cells\[0\]"
            r"\.paragraphs\[1\]"
        ),
    ):
        _ = retired_run.text


def test_cell_text_replacement_invalidates_nested_run_and_font():
    from rdocx import Document, StaleElementError

    document = Document()
    document.add_table(rows=1, cols=1)
    document.tables[0].rows[0].cells[0].text = "before"
    run = document.tables[0].rows[0].cells[0].paragraphs[0].runs[0]
    font = run.font

    document.tables[0].rows[0].cells[0].text = "after"

    with pytest.raises(StaleElementError, match=r"revision 2.*revision 3"):
        _ = run.text
    with pytest.raises(StaleElementError, match=r"revision 2.*revision 3"):
        _ = font.bold


def test_table_handles_write_through_and_reopen():
    from rdocx import (
        Document,
        Inches,
        Pt,
        WD_CELL_VERTICAL_ALIGNMENT,
        WD_TABLE_ALIGNMENT,
    )

    document = Document()
    table = document.add_table(rows=2, cols=2)
    table.style = "TableGrid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.width = Inches(4)
    document.tables[0].rows[0].cells[0].text = "alpha"
    document.tables[0].rows[1].cells[1].text = "omega"
    document.tables[0].rows[1].cells[1].vertical_alignment = (
        WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    )
    document.tables[0].rows[1].cells[1].width = Inches(2)
    document.tables[0].rows[0].cells[0].add_paragraph("nested")
    nested = document.tables[0].rows[0].cells[0].paragraphs[-1]
    nested.paragraph_format.space_after = Pt(6)
    nested.paragraph_format.left_indent = Inches(0.25)

    reopened = Document.from_bytes(document.to_bytes())
    table = reopened.tables[0]
    assert len(reopened.tables) == 1
    assert len(list(reopened.tables)) == 1
    assert len(table.rows) == 2
    assert [len(row.cells) for row in table.rows] == [2, 2]
    assert len(table.rows[0].cells) == 2
    assert table.style == "TableGrid"
    assert table.alignment == WD_TABLE_ALIGNMENT.CENTER
    assert table.width == Inches(4)
    assert table.rows[0].cells[0].text == "alpha\nnested"
    assert table.rows[1].cells[1].text == "omega"
    assert (
        table.rows[1].cells[1].vertical_alignment
        == WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    )
    assert table.rows[1].cells[1].width == Inches(2)
    nested = table.rows[0].cells[0].paragraphs[-1]
    assert [paragraph.text for paragraph in table.rows[0].cells[0].paragraphs] == [
        "alpha",
        "nested",
    ]
    assert nested.paragraph_format.space_after == Pt(6)
    assert nested.paragraph_format.left_indent == Inches(0.25)


def test_unrepresentable_table_justify_reads_as_none_after_reopen():
    from rdocx import Document, WD_TABLE_ALIGNMENT

    document = Document()
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    with pytest.raises(ValueError, match="unsupported table alignment"):
        table.alignment = 3

    justified = _replace_document_xml(
        document.to_bytes(),
        b'<w:jc w:val="center"/>',
        b'<w:jc w:val="both"/>',
    )
    reopened = Document.from_bytes(justified)

    assert reopened.tables[0].alignment is None


def test_automatic_font_color_reads_as_none_after_reopen():
    from rdocx import Document, RGBColor

    document = Document()
    font = document.add_paragraph("").add_run("automatic").font
    font.color = RGBColor(0x12, 0x34, 0x56)
    automatic = _replace_document_xml(
        document.to_bytes(),
        b'<w:color w:val="123456"/>',
        b'<w:color w:val="auto"/>',
    )
    reopened = Document.from_bytes(automatic)

    assert reopened.paragraphs[0].runs[0].font.color.rgb is None


def test_table_rows_are_cloned_with_their_formatting_and_removed():
    from rdocx import (
        Document,
        Inches,
        RdocxError,
        StaleElementError,
        WD_CELL_VERTICAL_ALIGNMENT,
    )

    document = Document()
    document.add_table(rows=2, cols=2)
    document.tables[0].rows[0].cells[0].text = "header"
    document.tables[0].rows[1].cells[0].text = "entry"
    document.tables[0].rows[1].cells[0].paragraphs[0].runs[0].font.bold = True
    template = document.tables[0].rows[1].cells[1]
    template.width = Inches(2)
    template.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM

    table = document.tables[0]
    held = table.rows[0]
    copied = table.clone_row(-1)
    with pytest.raises(StaleElementError):
        _ = held.cells
    copied.cells[1].text = "new entry"

    reopened = Document.from_bytes(document.to_bytes())
    rows = reopened.tables[0].rows
    assert [row.cells[0].text for row in rows] == ["header", "entry", "entry"]
    assert rows[2].cells[0].paragraphs[0].runs[0].font.bold is True
    assert rows[2].cells[1].text == "new entry"
    assert rows[2].cells[1].width == Inches(2)
    assert rows[2].cells[1].vertical_alignment == WD_CELL_VERTICAL_ALIGNMENT.BOTTOM

    reopened.tables[0].clone_row(0, at=0)
    reopened.tables[0].remove_row(1)
    reopened.tables[0].remove_row(-1)
    assert [row.cells[0].text for row in reopened.tables[0].rows] == [
        "header",
        "entry",
    ]
    with pytest.raises(IndexError):
        reopened.tables[0].remove_row(2)
    with pytest.raises(IndexError):
        reopened.tables[0].clone_row(0, at=3)
    reopened.tables[0].remove_row(0)
    with pytest.raises(RdocxError, match="at least one row"):
        reopened.tables[0].remove_row(0)
    assert [row.cells[0].text for row in reopened.tables[0].rows] == ["entry"]


def test_table_borders_margins_and_grid_widths_round_trip():
    from rdocx import Document, Inches, Pt

    document = Document()
    table = document.add_table(rows=2, cols=3)
    assert table.border("top") is None
    assert table.cell_margins is None
    table.set_borders("single", size=4, color="000000")
    table.set_border("insideV", "dashed", size=8, color="FF0000")
    table.set_cell_margins(top=Pt(1), right=Pt(2), bottom=Pt(3), left=Pt(4))
    table.grid_widths = [Inches(1), Inches(2), Inches(3)]
    table.set_column_width(-1, Inches(1.5))

    xml = _document_xml(document.to_bytes())
    assert '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>' in xml
    assert '<w:insideV w:val="dashed" w:sz="8" w:space="0" w:color="FF0000"/>' in xml
    assert "<w:tblCellMar>" in xml
    assert '<w:top w:w="20" w:type="dxa"/>' in xml
    assert '<w:right w:w="40" w:type="dxa"/>' in xml
    assert '<w:gridCol w:w="1440"/>' in xml
    assert '<w:gridCol w:w="2160"/>' in xml

    table = Document.from_bytes(document.to_bytes()).tables[0]
    assert table.border("top") == ("single", 4, RGBColor.from_string("000000"))
    assert table.border("insideH") == ("single", 4, RGBColor.from_string("000000"))
    assert table.border("insideV") == ("dashed", 8, RGBColor.from_string("FF0000"))
    assert table.cell_margins == (Pt(1), Pt(2), Pt(3), Pt(4))
    assert table.grid_widths == (Inches(1), Inches(2), Inches(1.5))
    assert table.width == Inches(4.5)
    assert [cell.width for cell in table.rows[1].cells] == [
        Inches(1),
        Inches(2),
        Inches(1.5),
    ]


def test_row_height_split_and_header_round_trip():
    from rdocx import Document, Pt, WD_ROW_HEIGHT_RULE

    document = Document()
    table = document.add_table(rows=2, cols=1)
    row = table.rows[0]
    assert row.height is None
    assert row.height_rule is None
    assert row.cant_split is None
    assert row.is_header is None
    row.height = Pt(20)
    assert row.height_rule == WD_ROW_HEIGHT_RULE.AT_LEAST
    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    row.height = Pt(30)
    row.cant_split = True
    row.is_header = True
    table.rows[1].cant_split = False

    xml = _document_xml(document.to_bytes())
    assert '<w:trHeight w:val="600" w:hRule="exact"/>' in xml
    assert "<w:cantSplit/>" in xml
    assert "<w:tblHeader/>" in xml

    rows = Document.from_bytes(document.to_bytes()).tables[0].rows
    assert rows[0].height == Pt(30)
    assert rows[0].height_rule == WD_ROW_HEIGHT_RULE.EXACTLY
    assert rows[0].cant_split is True
    assert rows[0].is_header is True
    assert rows[1].cant_split is False
    assert rows[1].is_header is None

    rows[0].is_header = None
    rows[0].cant_split = None
    assert rows[0].is_header is None
    assert rows[0].cant_split is None


def test_cell_shading_borders_and_margins_round_trip():
    from rdocx import Document, Pt

    document = Document()
    cell = document.add_table(rows=1, cols=2).cell(0, 1)
    assert cell.shading is None
    assert cell.border("bottom") is None
    assert cell.margins is None
    cell.shading = "D9D9D9"
    cell.set_border("bottom", "double", size=6, color="auto")
    cell.set_margins(top=0, right=Pt(5), bottom=0, left=Pt(5))

    xml = _document_xml(document.to_bytes())
    assert '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' in xml
    assert '<w:bottom w:val="double" w:sz="6" w:space="0" w:color="auto"/>' in xml

    cell = Document.from_bytes(document.to_bytes()).tables[0].cell(0, 1)
    assert cell.shading == RGBColor.from_string("D9D9D9")
    assert cell.border("bottom") == ("double", 6, "auto")
    assert cell.border("top") is None
    assert cell.margins == (0, Pt(5), 0, Pt(5))


def test_invalid_table_formatting_changes_nothing_and_keeps_handles_live():
    from rdocx import Document, RdocxError, WD_ROW_HEIGHT_RULE

    document = Document()
    table = document.add_table(rows=1, cols=2)
    row = table.rows[0]
    cell = row.cells[0]
    before = document.to_bytes()

    with pytest.raises(ValueError, match="border style"):
        table.set_borders("bogus", size=4, color="000000")
    with pytest.raises(ValueError, match="border edge"):
        cell.set_border("middle", "single", size=4, color="000000")
    with pytest.raises(RdocxError, match="border width"):
        table.set_border("top", "single", size=97, color="000000")
    with pytest.raises(ValueError, match=r"Cell\.set_border\(color\).*six ASCII hexadecimal digits"):
        cell.set_border("top", "single", size=4, color="red")
    with pytest.raises(RdocxError, match="cannot be negative"):
        table.set_cell_margins(top=-635, right=0, bottom=0, left=0)
    with pytest.raises(RdocxError, match="cannot be negative"):
        cell.set_margins(top=0, right=0, bottom=0, left=-635)
    with pytest.raises(ValueError, match=r"Cell\.shading.*six ASCII hexadecimal digits"):
        cell.shading = "yellow"
    with pytest.raises(RdocxError, match="exactly 2 positive column widths"):
        table.grid_widths = [914400]
    with pytest.raises(RdocxError, match="must be positive"):
        table.grid_widths = [914400, 0]
    with pytest.raises(IndexError):
        table.set_column_width(2, 914400)
    with pytest.raises(ValueError, match="nonnegative"):
        table.set_column_width(0, -635)
    with pytest.raises(RdocxError, match="cannot be negative"):
        row.height = -635
    with pytest.raises(ValueError, match="set height first"):
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    with pytest.raises(ValueError, match="unsupported row height rule"):
        row.height_rule = 0

    assert document.to_bytes() == before
    table.set_borders("single", size=4, color="auto")
    row.cant_split = True
    cell.shading = "FFFFFF"
    assert (row.cant_split, cell.shading) == (True, RGBColor.from_string("FFFFFF"))


def test_insert_table_at_a_body_index_returns_the_new_table():
    from rdocx import Document, StaleElementError

    seed = Document()
    seed.add_table(rows=1, cols=1).cell(0, 0).text = "in control"
    seed.add_paragraph("after")
    document = Document.from_bytes(
        _replace_document_xml(
            _replace_document_xml(seed.to_bytes(), b"<w:tbl>", b"<w:sdt><w:sdtContent><w:tbl>"),
            b"</w:tbl>",
            b"</w:tbl></w:sdtContent></w:sdt>",
        )
    )
    held = document.tables[0]
    before = document.to_bytes()
    with pytest.raises(IndexError, match="content index out of range"):
        document.insert_table(3, 1, 1)
    assert document.to_bytes() == before

    inserted = document.insert_table(1, 2, 2)
    with pytest.raises(StaleElementError):
        _ = held.rows
    assert document.find_content_index(inserted) == 1
    assert len(inserted.rows) == 2
    inserted.cell(1, 1).text = "inserted"

    reopened = Document.from_bytes(document.to_bytes())
    body = [
        (item.kind, item.text)
        for item in reopened.story_items
        if item.direct_body_index is not None
    ]
    assert body == [("content_control", None), ("table", None), ("paragraph", "after")]
    assert [table.cell(0, 0).text for table in reopened.tables] == ["in control", ""]
    assert reopened.tables[1].cell(1, 1).text == "inserted"


def test_cells_merge_through_the_checked_table_operations():
    from rdocx import Document, RdocxError, StaleElementError

    document = Document()
    table = document.add_table(rows=3, cols=3)
    held = table.cell(0, 2)
    table.set_cell_grid_span(0, 0, 2)
    with pytest.raises(StaleElementError):
        _ = held.text

    table = document.tables[0]
    assert [len(row.cells) for row in table.rows] == [2, 3, 3]
    assert [cell.grid_span for cell in table.rows[0].cells] == [2, 1]
    last = table.cell(0, -1)
    table.set_cell_vertical_merge(0, -1, "restart")
    table.set_cell_vertical_merge(1, -1, "continue")
    assert last.vertical_merge == "restart"
    assert [row.cells[-1].vertical_merge for row in table.rows] == [
        "restart",
        "continue",
        None,
    ]

    xml = _document_xml(document.to_bytes())
    assert '<w:gridSpan w:val="2"/>' in xml
    assert '<w:vMerge w:val="restart"/>' in xml
    assert "<w:vMerge/>" in xml

    table.cell(1, 1).text = "full"
    table = document.tables[0]
    before = document.to_bytes()
    with pytest.raises(RdocxError, match="nonempty cell"):
        table.set_cell_grid_span(1, 0, 2)
    with pytest.raises(RdocxError, match="no matching cell above"):
        table.set_cell_vertical_merge(1, 0, "continue")
    with pytest.raises(RdocxError, match="exceeds the row grid"):
        table.set_cell_grid_span(0, 0, 4)
    with pytest.raises(ValueError, match="vertical merge"):
        table.set_cell_vertical_merge(1, 0, "sideways")
    with pytest.raises(IndexError):
        table.set_cell_grid_span(3, 0, 2)
    assert document.to_bytes() == before

    table.set_cell_vertical_merge(1, -1, None)
    table.set_cell_grid_span(0, 0, None)
    reopened = Document.from_bytes(document.to_bytes()).tables[0]
    assert [len(row.cells) for row in reopened.rows] == [3, 3, 3]
    assert reopened.cell(1, -1).vertical_merge is None


def test_table_acceptance_workflow_writes_the_native_body():
    # table_acceptance_workflow_writes_the_body_the_python_binding_pins in
    # crates/rdocx/tests/integration_test.rs pins the same body for the native
    # calls, so CI checks that these calls write what the native ones write.
    from rdocx import Document, WD_ROW_HEIGHT_RULE

    document = Document()
    document.add_paragraph("before")
    document.add_paragraph("after")
    table = document.insert_table(1, 3, 3)
    table.set_cell_grid_span(0, 0, 2)
    table = document.tables[0]
    table.set_cell_vertical_merge(1, 2, "restart")
    table.set_cell_vertical_merge(2, 2, "continue")
    table.set_borders("single", size=4, color="000000")
    table.set_border("insideV", "dashed", size=8, color="FF0000")
    table.set_cell_margins(top=0, right=63500, bottom=12700, left=127000)
    table.grid_widths = [1371600, 1828800, 1828800]
    table.set_column_width(-1, 914400)
    cell = table.cell(1, 0)
    cell.shading = "D9D9D9"
    cell.set_margins(top=12700, right=0, bottom=25400, left=6350)
    cell.set_border("bottom", "double", size=6, color="auto")
    row = table.rows[0]
    row.height = 254000
    row.cant_split = True
    row.is_header = True
    table.rows[1].height = 381000
    table.rows[1].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    xml = _document_xml(document.to_bytes())
    body = xml[xml.index("<w:body>") : xml.index("</w:body>") + len("</w:body>")]
    assert "".join(line.strip() for line in body.splitlines()) == (
        '<w:body><w:p><w:r><w:t>before</w:t></w:r></w:p><w:tbl><w:tblPr>'
        '<w:tblW w:w="6480" w:type="dxa"/><w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="dashed" w:sz="8" w:space="0" w:color="FF0000"/>'
        '</w:tblBorders><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
        '<w:left w:w="200" w:type="dxa"/><w:bottom w:w="20" w:type="dxa"/>'
        '<w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="2160"/><w:gridCol w:w="2880"/>'
        '<w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/>'
        '<w:trHeight w:val="400" w:hRule="atLeast"/><w:tblHeader/></w:trPr>'
        '<w:tc><w:tcPr><w:tcW w:w="5040" w:type="dxa"/>'
        '<w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc><w:tc><w:tcPr>'
        '<w:tcW w:w="1440" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>'
        '<w:tr><w:trPr><w:trHeight w:val="600" w:hRule="exact"/></w:trPr>'
        '<w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/><w:tcBorders>'
        '<w:bottom w:val="double" w:sz="6" w:space="0" w:color="auto"/>'
        '</w:tcBorders>'
        '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/><w:tcMar>'
        '<w:top w:w="20" w:type="dxa"/><w:left w:w="10" w:type="dxa"/>'
        '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
        '</w:tcMar></w:tcPr><w:p/></w:tc><w:tc><w:tcPr>'
        '<w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p/></w:tc><w:tc>'
        '<w:tcPr><w:tcW w:w="1440" w:type="dxa"/>'
        '<w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc></w:tr><w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr><w:p/>'
        '</w:tc><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>'
        '<w:p/></w:tc><w:tc><w:tcPr><w:tcW w:w="1440" w:type="dxa"/>'
        '<w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p><w:r>'
        '<w:t>after</w:t></w:r></w:p><w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:gutter="0" w:header="720" w:footer="720"/></w:sectPr></w:body>'
    )


def test_python_paragraph_and_run_formatting_matches_native_facades():
    from rdocx import Document

    document = Document()
    paragraph = document.add_paragraph("Title")
    assert paragraph.style is None
    assert paragraph.numbering is None
    paragraph.style = "Heading1"
    paragraph.numbering = (1, 2)
    assert paragraph.text == "Title"

    reopened = Document.from_bytes(document.to_bytes())
    assert reopened.paragraphs[0].style == "Heading1"
    assert reopened.paragraphs[0].numbering == (1, 2)

    with pytest.raises(ValueError, match="numbering level"):
        paragraph.numbering = (1, 9)
    paragraph.style = None
    paragraph.numbering = None
    reopened = Document.from_bytes(document.to_bytes())
    assert reopened.paragraphs[0].style is None
    assert reopened.paragraphs[0].numbering is None

    document = Document()
    run = document.add_paragraph("").add_run("marked")
    assert run.style_id is None
    assert run.font.highlight is None
    assert run.font.shading is None
    run.style_id = "Strong"
    run.font.highlight = "yellow"
    run.font.shading = "FFFF00"

    reopened = Document.from_bytes(document.to_bytes()).paragraphs[0].runs[0]
    assert reopened.style_id == "Strong"
    assert reopened.font.highlight == "yellow"
    assert reopened.font.shading == RGBColor.from_string("FFFF00")

    run.font.shading = "AUTO"
    reopened = Document.from_bytes(document.to_bytes()).paragraphs[0].runs[0]
    assert reopened.font.shading == "auto"

    with pytest.raises(ValueError, match="highlight"):
        run.font.highlight = "FFFF00"
    with pytest.raises(ValueError, match="hexadecimal"):
        run.font.shading = "yellow"
    run.style_id = None
    run.font.highlight = None
    run.font.shading = None
    reopened = Document.from_bytes(document.to_bytes()).paragraphs[0].runs[0]
    assert reopened.style_id is None
    assert reopened.font.highlight is None
    assert reopened.font.shading is None


def test_paragraph_style_accepts_a_defined_id_or_name_and_rejects_the_rest():
    from rdocx import Document

    source = BytesIO(Document().to_bytes())
    output = BytesIO()
    with ZipFile(source) as archive, ZipFile(
        output, "w", compression=ZIP_DEFLATED
    ) as rewritten:
        for member in archive.infolist():
            contents = archive.read(member.filename)
            if member.filename == "word/styles.xml":
                assert b"</w:styles>" in contents
                contents = contents.replace(
                    b"</w:styles>",
                    b'<w:style w:type="character" w:styleId="Strong">'
                    b'<w:name w:val="Strong"/></w:style>'
                    b'<w:style w:type="character" w:styleId="Emphasis">'
                    b'<w:name w:val="Emphasis Char"/></w:style>'
                    b'<w:style w:type="paragraph" w:styleId="EmphPara">'
                    b'<w:name w:val="Emphasis"/></w:style></w:styles>',
                )
            rewritten.writestr(member, contents)
    document = Document.from_bytes(output.getvalue())
    document.add_paragraph("Title")
    paragraph = document.paragraphs[0]

    for value, expected in (
        ("Heading 1", "Heading1"),
        ("Normal", "Normal"),
        ("heading 1", "Heading1"),
        ("Emphasis", "EmphPara"),
        ("Heading1", "Heading1"),
    ):
        paragraph.style = value
        assert paragraph.style == expected

    before = document.to_bytes()
    for value, error, message in (
        ("NoSuchStyle", KeyError, "no style with ID or name 'NoSuchStyle'"),
        ("Heading 10", KeyError, "no style"),
        ("Strong", ValueError, "character style"),
        ("emphasis char", ValueError, "character style"),
    ):
        with pytest.raises(error, match=message):
            paragraph.style = value
    assert document.to_bytes() == before
    assert paragraph.style == "Heading1"
    assert Document.from_bytes(before).paragraphs[0].style == "Heading1"


def _part(document, name):
    with ZipFile(BytesIO(document.to_bytes())) as archive:
        return archive.read(name).decode()


def test_new_document_defines_the_common_word_styles_and_opened_ones_keep_theirs():
    from rdocx import Document

    document = Document()
    styles = {style.style_id: style for style in document.styles}
    expected = {f"Heading{level}" for level in range(1, 10)} | {
        "Normal",
        "Title",
        "Subtitle",
        "NoSpacing",
        "Quote",
        "ListParagraph",
        "Caption",
        "TableGrid",
    }
    assert set(styles) == expected
    assert styles["Heading2"].name == "heading 2"
    assert styles["Heading2"].based_on == "Normal"
    assert styles["TableGrid"].style_type == "table"
    assert [style.is_default for style in document.styles].count(True) == 1

    document.add_paragraph("Chapter")
    document.paragraphs[0].style = "Heading 2"
    assert document.paragraphs[0].style == "Heading2"
    document.paragraphs[0].style = "Caption"
    assert document.paragraphs[0].style == "Caption"

    trimmed = BytesIO()
    with ZipFile(BytesIO(document.to_bytes())) as source, ZipFile(
        trimmed, "w", compression=ZIP_DEFLATED
    ) as output:
        for member in source.infolist():
            contents = source.read(member.filename)
            if member.filename == "word/styles.xml":
                contents = (
                    b'<w:styles xmlns:w="http://schemas.openxmlformats.org/'
                    b'wordprocessingml/2006/main"><w:style w:type="paragraph" '
                    b'w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
                    b"</w:style></w:styles>"
                )
            output.writestr(member, contents)
    opened = Document.from_bytes(trimmed.getvalue())
    assert [style.style_id for style in opened.styles] == ["Normal"]


def test_add_style_derives_the_word_id_and_writes_its_formatting(tmp_path):
    from rdocx import Document, Inches, Pt, RGBColor, Style

    document = Document()
    document.add_paragraph("body")
    held = document.paragraphs[0]

    note = document.add_style("Note", style_type="paragraph", based_on="Normal")
    assert note == Style(
        style_id="Note",
        name="Note",
        based_on="Normal",
        style_type="paragraph",
        linked_style=None,
        next_style=None,
        priority=None,
        auto_redefine=None,
        hidden=None,
        semi_hidden=None,
        unhide_when_used=None,
        quick_format=None,
        locked=None,
        is_default=False,
    )
    boxed = document.add_style(
        "Boxed Note",
        based_on="note",
        next_style="Normal",
        font_name="Arial",
        font_size=Pt(12),
        bold=True,
        italic=False,
        color=RGBColor(0x11, 0x22, 0x33),
        space_before=Pt(6),
        space_after=Pt(3),
        left_indent=Inches(0.5),
        right_indent=Inches(0.25),
        first_line_indent=-Inches(0.25),
    )
    assert (boxed.style_id, boxed.based_on, boxed.next_style) == (
        "BoxedNote",
        "Note",
        "Normal",
    )
    assert document.add_style("heading x", style_id="Custom").style_id == "Custom"
    assert document.add_style("Callout", "character", bold=True).style_type == "character"
    assert document.add_style("Plain Grid", "table", based_on="Table Grid").based_on == (
        "TableGrid"
    )
    assert held.text == "body"
    held.style = "Boxed Note"
    assert held.style == "BoxedNote"

    path = tmp_path / "styles.docx"
    document.save(path)
    reopened = {style.style_id: style for style in Document(path).styles}
    assert reopened["BoxedNote"] == boxed
    xml = _part(document, "word/styles.xml")
    start = xml.index('w:styleId="BoxedNote"')
    boxed_xml = xml[start : xml.index("</w:style>", start)]
    for fragment in (
        '<w:spacing w:before="120" w:after="60"/>',
        '<w:ind w:left="720" w:right="360" w:hanging="360"/>',
        '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>',
        '<w:i w:val="false"/>',
        '<w:color w:val="112233"/>',
        '<w:sz w:val="24"/>',
    ):
        assert fragment in boxed_xml, fragment

    docx = pytest.importorskip("docx")
    oracle = docx.Document(str(path)).styles["Boxed Note"]
    assert oracle.style_id == "BoxedNote"
    assert oracle.base_style.name == "Note"
    assert (oracle.font.name, oracle.font.size, oracle.font.bold) == ("Arial", Pt(12), True)
    assert oracle.font.italic is False
    assert str(oracle.font.color.rgb) == "112233"
    assert oracle.paragraph_format.space_before == Pt(6)
    assert oracle.paragraph_format.left_indent == Inches(0.5)
    assert oracle.paragraph_format.first_line_indent == -Inches(0.25)


def test_add_style_derives_word_ids_that_survive_save_and_reopen(tmp_path):
    from rdocx import Document

    expected = {
        "Q&A": "QA",
        "Note (draft)": "Notedraft",
        "My-Style.v2": "My-Stylev2",
        "\u00dcberschrift Eigen": "berschriftEigen",
        'A "q" <x>': "Aqx",
        "\u898b\u51fa\u3057 \u30ab\u30b9\u30bf\u30e0": "a",
        "\u5225\u306e\u30b9\u30bf\u30a4\u30eb": "a0",
    }
    document = Document()
    for name, style_id in expected.items():
        assert document.add_style(name, based_on="Normal").style_id == style_id
        document.add_paragraph(name).style = style_id

    path = tmp_path / "ids.docx"
    document.save(path)
    reopened = Document(path)
    assert set(expected.values()) <= {style.style_id for style in reopened.styles}
    assert [paragraph.style for paragraph in reopened.paragraphs] == list(
        expected.values()
    )
    docx = pytest.importorskip("docx")
    oracle = {style.style_id for style in docx.Document(str(path)).styles}
    assert set(expected.values()) <= oracle


def test_add_style_rejects_bad_input_without_changing_the_document():
    from rdocx import Document

    document = Document()
    document.add_style("Callout", "character")
    before = document.to_bytes()
    for arguments, keywords, error, message in (
        (("Heading 2",), {}, ValueError, "ID 'Heading2' already exists"),
        (("normal",), {"style_id": "Other"}, ValueError, "named 'normal' already exists"),
        (("Note",), {"style_id": " "}, ValueError, "cannot be blank"),
        (("Note", "numbering"), {}, ValueError, "style_type must be"),
        (("Note",), {"based_on": "Missing"}, KeyError, "no style with ID or name 'Missing'"),
        (("Note",), {"based_on": "Callout"}, ValueError, "is a character style, not a paragraph"),
        (("Note",), {"next_style": "Table Grid"}, ValueError, "is a table style, not a paragraph"),
        (("Note", "character"), {"next_style": "Normal"}, ValueError, "only a paragraph style"),
        (("Note", "character"), {"space_after": 12700}, ValueError, "no paragraph formatting"),
    ):
        with pytest.raises(error, match=message):
            document.add_style(*arguments, **keywords)
    assert document.to_bytes() == before


def test_a_style_with_a_numbering_level_renders_its_paragraphs_numbered():
    from rdocx import Document, ListLevel, Pt

    def recipe(style_numbering, direct_numbering):
        document = Document()
        definition = document.add_numbering_definition(
            [ListLevel(format="decimal", text="%1."), ListLevel(format="lowerLetter")]
        )
        instance = document.add_numbering_instance(definition)
        document.add_style(
            "Step",
            based_on="Normal",
            font_name="Arial",
            font_size=Pt(14),
            bold=True,
            space_before=Pt(12),
            space_after=Pt(6),
        )
        if style_numbering:
            document.link_style_to_numbering("Step", instance, 0)
        for text in ("Mix", "Bake"):
            paragraph = document.add_paragraph(text)
            paragraph.style = "Step"
            if direct_numbering:
                paragraph.numbering = (instance, 0)
        return document

    document = recipe(style_numbering=True, direct_numbering=False)
    assert document.paragraphs[0].numbering is None
    styles_xml = _part(document, "word/styles.xml")
    step = styles_xml[styles_xml.index('w:styleId="Step"') :]
    assert '<w:numId w:val="1"/>' in step[: step.index("</w:style>")]
    numbering_xml = _part(document, "word/numbering.xml")
    assert '<w:pStyle w:val="Step"/>' in numbering_xml
    assert '<w:numFmt w:val="lowerLetter"/>' in numbering_xml
    assert '<w:lvlText w:val="%2."/>' in numbering_xml

    rendered = document.render_page_to_png(0, 72)
    assert rendered == recipe(False, True).render_page_to_png(0, 72)
    assert rendered != recipe(False, False).render_page_to_png(0, 72)
    reopened = Document.from_bytes(document.to_bytes())
    assert reopened.render_page_to_png(0, 72) == rendered


def test_list_levels_and_numbering_calls_check_their_input():
    import rdocx
    from rdocx import Document, Inches, ListLevel, RdocxError

    level = ListLevel(format="bullet", text="-", start=3, left_indent=Inches(1))
    assert (level.format, level.text, level.start) == ("bullet", "-", 3)
    assert (level.left_indent, level.hanging_indent) == (Inches(1), None)
    assert level == ListLevel(format="bullet", text="-", start=3, left_indent=Inches(1))
    assert ListLevel().format == "decimal"
    with pytest.raises(ValueError, match="'Decimal' is not a Word numbering format"):
        ListLevel(format="Decimal")
    with pytest.raises(ValueError, match="hanging_indent"):
        ListLevel(hanging_indent=-1)
    with pytest.raises(AttributeError):
        level.start = 1  # type: ignore[misc]

    document = Document()
    definition = document.add_numbering_definition(
        [level, ListLevel(format="upperRoman", hanging_indent=Inches(0.5))]
    )
    numbering_xml = _part(document, "word/numbering.xml")
    for fragment in (
        '<w:start w:val="3"/>',
        '<w:lvlText w:val="-"/>',
        '<w:ind w:left="1440" w:hanging="360"/>',
        '<w:ind w:left="1440" w:hanging="720"/>',
    ):
        assert fragment in numbering_xml, fragment
    instance = document.add_numbering_instance(definition)
    assert instance != document.add_numbering_instance(definition)

    before = document.to_bytes()
    for call in (
        lambda: document.add_numbering_definition([]),
        lambda: document.add_numbering_definition([ListLevel(text="%2.")]),
        lambda: document.add_numbering_instance(definition + 100),
        lambda: document.link_style_to_numbering("Normal", instance + 100, 0),
        lambda: document.link_style_to_numbering("Normal", instance, 5),
    ):
        with pytest.raises(RdocxError):
            call()
    with pytest.raises(ValueError, match="is a table style, not a paragraph"):
        document.link_style_to_numbering("TableGrid", instance, 0)
    assert document.to_bytes() == before
    assert rdocx.ListLevel is ListLevel


def test_styles_can_be_removed_and_made_the_default():
    from rdocx import Document, RdocxError

    document = Document()
    document.add_style("Note", based_on="Normal")
    document.add_style("Aside", based_on="Note")
    document.add_paragraph("aside").style = "Aside"

    assert document.remove_style("No Such Style") is False
    before = document.to_bytes()
    with pytest.raises(RdocxError, match="referenced by style 'Aside'"):
        document.remove_style("Note")
    with pytest.raises(RdocxError, match="referenced by document content"):
        document.remove_style("Aside")
    assert document.to_bytes() == before

    document.set_default_style("note")
    defaults = [style.style_id for style in document.styles if style.is_default]
    assert defaults == ["Note"]
    with pytest.raises(KeyError, match="no style with ID or name 'Missing'"):
        document.set_default_style("Missing")

    assert document.remove_style("Subtitle") is True
    assert "Subtitle" not in {style.style_id for style in document.styles}


def test_word_highlight_keywords_round_trip_and_clear():
    from rdocx import Document, RGBColor

    document = Document()
    document.add_paragraph("").add_run("marked").font.color = RGBColor(0x12, 0x34, 0x56)
    keyword = _replace_document_xml(
        document.to_bytes(),
        b'<w:color w:val="123456"/>',
        b'<w:color w:val="123456"/><w:highlight w:val="darkBlue"/>',
    )
    reopened = Document.from_bytes(keyword)
    font = reopened.paragraphs[0].runs[0].font
    assert font.highlight == "darkBlue"

    font.highlight = "yellow"
    reopened = Document.from_bytes(reopened.to_bytes())
    assert reopened.paragraphs[0].runs[0].font.highlight == "yellow"
    reopened.paragraphs[0].runs[0].font.highlight = None
    assert Document.from_bytes(reopened.to_bytes()).paragraphs[0].runs[0].font.highlight is None


def test_table_indent_accepts_negative_zero_positive_and_none():
    from rdocx import Document, Inches

    document = Document()
    table = document.add_table(rows=1, cols=1)
    assert table.indent is None
    for value in (Inches(-0.25), 0, Inches(0.25), None):
        table.indent = value
        assert table.indent == value


def test_invalid_table_indent_leaves_source_and_handles_unchanged():
    from rdocx import Document, Inches
    document = Document()
    table = document.add_table(rows=1, cols=1)
    table.indent = Inches(-0.25)
    cell = table.cell(0, 0)
    before = document.to_bytes()
    from rdocx import RdocxError
    for value in ("0.25", 0.25, (2**31) * 635, (-(2**31) - 1) * 635):
        with pytest.raises((TypeError, RdocxError)):
            table.indent = value
        assert document.to_bytes() == before
        assert cell.text == ""
    reopened = Document.from_bytes(document.to_bytes())
    assert reopened.tables[0].indent == Inches(-0.25)
    table.indent = -636
    assert table.indent == -635
    assert Document.from_bytes(document.to_bytes()).tables[0].indent == -635
    table.indent = None
    assert b"tblInd" not in _document_xml(document.to_bytes()).encode()


def _f_x197_word_colour_surface(document, surface):
    if surface == "style-create":
        return lambda value: document.add_style("Colour Control", color=value), None
    if surface == "style-update":
        document.add_style("Colour Control")
        return lambda value: document.set_style("Colour Control", color=value), None
    if surface.startswith("font"):
        font = document.add_paragraph("").add_run("colour").font
        if surface == "font-direct":
            return lambda value: setattr(font, "color", value), lambda: font.color.rgb
        if surface == "font-rgb":
            return lambda value: setattr(font.color, "rgb", value), lambda: font.color.rgb
        return lambda value: setattr(font, "shading", value), lambda: font.shading
    table = document.add_table(1, 1)
    if surface == "table-borders":
        return lambda value: table.set_borders("single", size=4, color=value), lambda: (table.border("top") or (None, None, None))[2]
    if surface == "table-border":
        return lambda value: table.set_border("top", "single", size=4, color=value), lambda: (table.border("top") or (None, None, None))[2]
    cell = table.cell(0, 0)
    if surface == "cell-border":
        return lambda value: cell.set_border("left", "single", size=4, color=value), lambda: (cell.border("left") or (None, None, None))[2]
    assert surface == "cell-shading"
    return lambda value: setattr(cell, "shading", value), lambda: cell.shading


_F_X197_WORD_COLOUR_SURFACES = ["style-create", "style-update", "font-direct", "font-rgb", "font-shading", "table-borders", "table-border", "cell-border", "cell-shading"]


@pytest.mark.parametrize("surface", _F_X197_WORD_COLOUR_SURFACES)
@pytest.mark.parametrize("form", ["rgb", "triple", "hex", "hash"])
def test_every_word_colour_setter_accepts_rgb_and_optional_hash_hex(surface, form):
    from rdocx import Document, RGBColor
    from xml.etree import ElementTree as ET
    document = Document()
    expected = RGBColor(0xA1, 0xB2, 0xC3)
    assert RGBColor.from_string("#a1b2c3") == expected
    assert RGBColor.from_string("a1b2c3") == expected
    value = {"rgb": expected, "triple": tuple(expected), "hex": "a1b2c3", "hash": "#a1b2c3"}[form]
    setter, getter = _f_x197_word_colour_surface(document, surface)
    setter(value)
    if getter is not None:
        assert getter() == expected
        assert isinstance(getter(), RGBColor)
    saved = document.to_bytes()
    if surface != "style-create":
        setter(value)
        assert document.to_bytes() == saved
    reopened = Document.from_bytes(saved)
    with ZipFile(BytesIO(reopened.to_bytes())) as package:
        member = "word/styles.xml" if surface.startswith("style") else "word/document.xml"
        tree = ET.fromstring(package.read(member))
        namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        if surface.startswith("style"):
            tree = next(node for node in tree.iter(namespace + "style") if node.get(namespace + "styleId") == "ColourControl")
        assert any(node.get(namespace + name) == "A1B2C3" for node in tree.iter() for name in ("val", "fill", "color"))


def test_style_colour_inputs_are_validated_before_style_mutation():
    from rdocx import Document
    document = Document()
    document.add_style("Colour Control", color="#123456")
    before = document.to_bytes()
    with pytest.raises(ValueError, match="hex"):
        document.set_style("Colour Control", color="##123456")
    assert document.to_bytes() == before


def test_word_font_colour_format_supports_rgb_theme_and_direct_assignment():
    from rdocx import Document, RGBColor
    document = Document()
    font = document.add_paragraph("").add_run("colour").font
    colour = font.color
    assert colour.rgb is None
    assert colour.type is None
    font.color = "#123456"
    assert colour.rgb == RGBColor(0x12, 0x34, 0x56)
    colour.theme_color = "accent1"
    assert colour.theme_color == "accent1"
    font.color = None
    assert colour.rgb is None


def test_word_scalar_rgb_getters_return_rgbcolor_and_keep_auto():
    from rdocx import Document, RGBColor
    document = Document()
    cell = document.add_table(1, 1).cell(0, 0)
    cell.shading = "123456"
    assert cell.shading == RGBColor(0x12, 0x34, 0x56)
    cell.shading = "auto"
    assert cell.shading == "auto"


@pytest.mark.parametrize("surface", _F_X197_WORD_COLOUR_SURFACES)
@pytest.mark.parametrize("value,error_type", [(object(), TypeError), (7, TypeError), ((1, 2), TypeError), ((1.5, 2, 3), TypeError), ((10**100, 2, 3), ValueError), ((256, 2, 3), ValueError), ("##123456", ValueError), ("１２３４５６", ValueError), ("12345\ud800", ValueError), (" 123456", ValueError)])
def test_invalid_colour_inputs_name_the_parameter_and_leave_source_unchanged(surface, value, error_type):
    from rdocx import Document
    document = Document()
    setter, getter = _f_x197_word_colour_surface(document, surface)
    before = document.to_bytes()
    with pytest.raises(error_type) as error:
        setter(value)
    parameter = {"style-create": "Document.add_style(color)", "style-update": "Document.set_style(color)", "font-direct": "Font.color", "font-rgb": "Font.color.rgb", "font-shading": "Font.shading", "table-borders": "Table.set_borders(color)", "table-border": "Table.set_border(color)", "cell-border": "Cell.set_border(color)", "cell-shading": "Cell.shading"}[surface]
    assert parameter in str(error.value)
    assert "RGBColor" in str(error.value) or "hex" in str(error.value) or "channels" in str(error.value)
    if error_type is TypeError:
        assert type(value).__name__ in str(error.value)
    assert document.to_bytes() == before
    if getter is not None:
        assert getter() is None


def test_colour_format_handles_follow_structural_invalidation():
    from rdocx import Document, StaleElementError
    document = Document()
    colour = document.add_paragraph("").add_run("colour").font.color
    assert colour.rgb is None
    document.add_paragraph("invalidate")
    with pytest.raises(StaleElementError):
        _ = colour.rgb


def _f_x197_word_pdf_pixels(pdf, tmp_path, name):
    import shutil
    import subprocess
    from PIL import Image
    executable = shutil.which("pdftoppm")
    assert executable is not None
    version = subprocess.run([executable, "-v"], capture_output=True, text=True, timeout=60)
    assert version.returncode == 0 and "26.01.0" in version.stdout + version.stderr
    source = tmp_path / (name + ".pdf")
    source.write_bytes(pdf)
    assert pdf.startswith(b"%PDF")
    prefix = tmp_path / name
    result = subprocess.run([executable, "-f", "1", "-singlefile", "-r", "72", "-png", str(source), str(prefix)], capture_output=True, text=True, timeout=60)
    assert result.returncode == 0, result.stdout + result.stderr
    return Image.open(prefix.with_suffix(".png")).convert("RGB")


def test_colour_inputs_and_nested_font_formats_reach_actual_rendering(tmp_path):
    from rdocx import Document, Pt
    from xml.etree import ElementTree as ET
    namespace = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    def render(value, theme=None):
        document = Document()
        font = document.add_paragraph("").add_run("COLOUR CONTROL").font
        font.name = "Liberation Sans"
        font.size = Pt(28)
        font.color.rgb = value
        if theme is not None:
            font.color.theme_color = theme
        output = BytesIO()
        with ZipFile(BytesIO(document.to_bytes())) as source, ZipFile(output, "w") as target:
            for member in source.infolist():
                data = source.read(member.filename)
                if member.filename == "word/theme/theme1.xml":
                    tree = ET.fromstring(data)
                    for token, rgb in (("accent1", "FF0000"), ("accent2", "0000FF")):
                        colour = tree.find(".//" + namespace + token + "/" + namespace + "srgbClr")
                        assert colour is not None
                        colour.set("val", rgb)
                    data = ET.tostring(tree)
                target.writestr(member, data)
        reopened = Document.from_bytes(output.getvalue())
        if theme is not None:
            colour = reopened.paragraphs[0].runs[0].font.color
            assert colour.theme_color == theme
            assert colour.rgb == RGBColor(0, 255, 0)
        return _f_x197_word_pdf_pixels(reopened.to_pdfa_deterministic(), tmp_path, (theme or value).replace("#", "hash"))
    red = render("#FF0000")
    blue = render("#0000FF")
    assert red.tobytes() == render("ff0000").tobytes()
    assert red.tobytes() != blue.tobytes()
    assert any(r > g + 40 and r > b + 40 for r, g, b in red.getdata())
    assert any(b > g + 40 and b > r + 40 for r, g, b in blue.getdata())
    assert render("00FF00", "accent1").tobytes() == red.tobytes()
    assert render("00FF00", "accent2").tobytes() == blue.tobytes()


def test_word_theme_colour_view_retains_cached_rgb_modifiers_and_opaque_source():
    from rdocx import Document, RGBColor, MSO_THEME_COLOR
    from xml.etree import ElementTree as ET
    document = Document()
    document.add_paragraph("").add_run("source").font.color = "123456"
    output = BytesIO()
    with ZipFile(BytesIO(document.to_bytes())) as source, ZipFile(output, "w") as target:
        for member in source.infolist():
            data = source.read(member.filename)
            if member.filename == "word/document.xml":
                assert data.count(b'w:val="123456"') == 1
                data = data.replace(b'w:val="123456"', b'w:val="123456" w:themeColor="accent1" w:themeTint="80" w:themeShade="40"')
                data = data.replace(b"</w:rPr>", b'<w:unknownColourMetadata value="retain"/></w:rPr>', 1)
            target.writestr(member, data)
    document = Document.from_bytes(output.getvalue())
    colour = document.paragraphs[0].runs[0].font.color
    before = document.to_bytes()
    assert colour.type is MSO_COLOR_TYPE.THEME and colour.theme_color == "accent1"
    assert colour.rgb == RGBColor(0x12, 0x34, 0x56)
    assert document.to_bytes() == before
    colour.theme_color = MSO_THEME_COLOR.ACCENT_2
    reopened = Document.from_bytes(document.to_bytes())
    assert reopened.paragraphs[0].runs[0].font.color.theme_color == "accent2"
    with ZipFile(BytesIO(reopened.to_bytes())) as package:
        xml = package.read("word/document.xml")
        assert b'<w:unknownColourMetadata value="retain"/>' in xml
        root = ET.fromstring(xml)
        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        node = root.find(".//" + ns + "color")
        assert {key: node.get(ns + key) for key in ("val", "themeColor", "themeTint", "themeShade")} == {"val": "123456", "themeColor": "accent2", "themeTint": "80", "themeShade": "40"}
    before = document.to_bytes()
    for invalid in ("unknown", 0, -2, 17, True):
        with pytest.raises((TypeError, ValueError)):
            colour.theme_color = invalid
        assert document.to_bytes() == before
    colour.theme_color = None
    assert colour.type is MSO_COLOR_TYPE.RGB and colour.rgb == RGBColor(0x12, 0x34, 0x56)


@pytest.mark.parametrize("surface", ["font-shading", "table-borders", "table-border", "cell-border", "cell-shading"])
def test_word_shading_and_border_scalar_auto_remains_distinct_from_rgb(surface):
    from rdocx import Document
    document = Document()
    setter, getter = _f_x197_word_colour_surface(document, surface)
    setter("auto")
    assert getter() == "auto"
    assert '="auto"' in _document_xml(Document.from_bytes(document.to_bytes()).to_bytes())


@pytest.mark.parametrize("code,token", list(enumerate(["dark1", "light1", "dark2", "light2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink", "followedHyperlink", "text1", "background1", "text2", "background2"], start=1)))
def test_word_supported_theme_enum_values_round_trip_without_invented_rgb(code, token):
    from rdocx import Document, MSO_THEME_COLOR
    document = Document()
    colour = document.add_paragraph("").add_run("theme").font.color
    colour.theme_color = MSO_THEME_COLOR(code)
    assert colour.rgb is None and colour.type is MSO_COLOR_TYPE.THEME
    assert colour.theme_color == token
    assert Document.from_bytes(document.to_bytes()).paragraphs[0].runs[0].font.color.theme_color == token


def test_word_unset_colour_view_and_none_clear_preserve_absent_source():
    from rdocx import Document, ColorFormat
    document = Document()
    font = document.add_paragraph("").add_run("unset").font
    before = document.to_bytes()
    assert isinstance(font.color, ColorFormat)
    assert (font.color.rgb, font.color.type, font.color.theme_color) == (None, None, None)
    font.color.rgb = None
    font.color.theme_color = None
    font.color = None
    assert document.to_bytes() == before


def test_word_cell_font_colour_view_uses_accepted_run_paths_and_rejects_structural_staleness():
    from rdocx import Document, RGBColor, StaleElementError
    document = Document()
    cell = document.add_table(1, 1).cell(0, 0)
    cell.text = "cell colour"
    cell = document.tables[0].cell(0, 0)
    colour = cell.paragraphs[0].runs[0].font.color
    colour.rgb = "#123456"
    assert colour.rgb == RGBColor(0x12, 0x34, 0x56)
    assert Document.from_bytes(document.to_bytes()).tables[0].cell(0, 0).paragraphs[0].runs[0].font.color.rgb == colour.rgb
    document.add_paragraph("invalidate")
    before = document.to_bytes()
    with pytest.raises(StaleElementError):
        _ = colour.rgb
    with pytest.raises(StaleElementError):
        colour.rgb = "ABCDEF"
    assert document.to_bytes() == before


def test_fx200_document_defaults_preserve_normal_named_styles_and_live_paragraphs():
    from rdocx import Document
    from rdocx.shared import Pt

    document = Document()
    paragraph = document.add_paragraph("unchanged")
    from xml.etree import ElementTree as ET
    def named_style_xml(data):
        with ZipFile(BytesIO(data)) as package:
            root = ET.fromstring(package.read("word/styles.xml"))
        return tuple(ET.tostring(child) for child in root if child.tag.endswith("}style"))
    original_styles = named_style_xml(document.to_bytes())
    document.default_font_name = "Gelasio"
    document.default_font_size = Pt(13)
    assert document.default_font_name == "Gelasio"
    assert document.default_font_size == Pt(13)
    assert paragraph.text == "unchanged"
    assert named_style_xml(document.to_bytes()) == original_styles
    reopened = Document.from_bytes(document.to_bytes())
    assert reopened.default_font_name == "Gelasio"
    assert reopened.default_font_size == Pt(13)
    reopened.default_font_name = None
    reopened.default_font_size = None
    assert reopened.default_font_name is None and reopened.default_font_size is None
    assert reopened.paragraphs[0].text == "unchanged"


def test_fx200_defaults_absent_read_clear_and_bad_size_preserve_package_bytes():
    from rdocx import Document
    import pytest

    from xml.etree import ElementTree as ET
    source = Document().to_bytes()
    result = BytesIO()
    with ZipFile(BytesIO(source)) as package, ZipFile(result, "w") as target:
        for member in package.infolist():
            data = package.read(member.filename)
            if member.filename == "word/styles.xml":
                root = ET.fromstring(data)
                for child in list(root):
                    if child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docDefaults":
                        root.remove(child)
                data = ET.tostring(root)
            target.writestr(member, data)
    document = Document.from_bytes(result.getvalue())
    before = document.to_bytes()
    assert document.default_font_name is None and document.default_font_size is None
    document.default_font_name = None
    document.default_font_size = None
    assert document.to_bytes() == before
    for invalid in (-1, (2**32) * 6350):
        with pytest.raises(ValueError, match="default_font_size"):
            document.default_font_size = invalid
        assert document.to_bytes() == before


@pytest.mark.parametrize("cell_owned", [False, True])
def test_fx200_tabs_owned_snapshots_stable_insert_delete_and_clear(cell_owned):
    from rdocx import Document, Inches, WD_TAB_ALIGNMENT, WD_TAB_LEADER

    document = Document()
    paragraph = document.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0] if cell_owned else document.add_paragraph("tabs")
    tabs = paragraph.paragraph_format.tab_stops
    before = document.to_bytes()
    assert len(tabs) == 0 and list(tabs) == []
    assert document.to_bytes() == before
    retained = tabs.add_tab_stop(Inches(2), WD_TAB_ALIGNMENT.END, WD_TAB_LEADER.MIDDLE_DOT)
    tabs.add_tab_stop(Inches(-1), WD_TAB_ALIGNMENT.START)
    before_duplicate = document.to_bytes()
    with pytest.raises(ValueError, match="position"):
        tabs.add_tab_stop(Inches(2), WD_TAB_ALIGNMENT.LIST, WD_TAB_LEADER.HEAVY)
    assert document.to_bytes() == before_duplicate
    assert retained.position == Inches(2) and paragraph.text == ("" if cell_owned else "tabs")
    tabs.add_tab_stop(Inches(3), WD_TAB_ALIGNMENT.LIST, WD_TAB_LEADER.HEAVY)
    assert [tab.position for tab in tabs] == [Inches(-1), Inches(2), Inches(3)]
    assert tabs[-1].alignment == WD_TAB_ALIGNMENT.LIST
    del tabs[-1]
    assert retained.position == Inches(2) and retained.alignment == WD_TAB_ALIGNMENT.END
    assert retained.leader == WD_TAB_LEADER.MIDDLE_DOT
    with pytest.raises(AttributeError):
        retained.position = Inches(3)
    before = document.to_bytes()
    for invalid in (-(2**31 + 1) * 635, 2**31 * 635):
        with pytest.raises(ValueError, match="signed twip"):
            tabs.add_tab_stop(invalid)
        assert document.to_bytes() == before
    reopened = Document.from_bytes(before)
    paragraph = reopened.tables[0].cell(0, 0).paragraphs[0] if cell_owned else reopened.paragraphs[0]
    restored = paragraph.paragraph_format.tab_stops
    assert [(tab.position, tab.alignment, tab.leader) for tab in restored] == [
        (Inches(-1), WD_TAB_ALIGNMENT.START, WD_TAB_LEADER.SPACES),
        (Inches(2), WD_TAB_ALIGNMENT.END, WD_TAB_LEADER.MIDDLE_DOT),
    ]
    tabs.clear_all()
    assert len(tabs) == 0
    assert retained.position == Inches(2)
    after = document.to_bytes()
    tabs.clear_all()
    assert document.to_bytes() == after


@pytest.mark.parametrize("alignment", [0, 1, 2, 3, 4, 6, 7, 102, 103, 104])
@pytest.mark.parametrize("leader", [None, 0, 1, 2, 3, 4, 5])
def test_fx200_tabs_exact_vocabulary_round_trip(alignment, leader):
    from rdocx import Document, Pt
    document = Document()
    tabs = document.add_paragraph("typed").paragraph_format.tab_stops
    result = tabs.add_tab_stop(Pt(18), alignment, leader)
    reopened = Document.from_bytes(document.to_bytes()).paragraphs[0].paragraph_format.tab_stops[0]
    assert reopened.position == result.position == Pt(18)
    assert reopened.alignment == result.alignment == alignment
    assert reopened.leader == result.leader == (0 if leader is None else leader)


@pytest.mark.parametrize("cell_owned", [False, True])
def test_fx200_paragraph_border_shading_outline_rtl_roundtrip_and_atomic_errors(cell_owned):
    from rdocx import Document, RGBColor
    document = Document()
    paragraph = document.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0] if cell_owned else document.add_paragraph("owner")
    formatting = paragraph.paragraph_format
    before = document.to_bytes()
    assert formatting.border("top") is None and formatting.shading is None
    assert formatting.outline_level is None and formatting.rtl is None
    assert document.to_bytes() == before
    for edge in ("top", "bottom", "left", "right", "between", "bar"):
        formatting.set_border(edge, "double", size=12, color=RGBColor(18, 52, 86))
        assert formatting.border(edge) == ("double", 12, RGBColor(18, 52, 86))
    formatting.shading = "auto"
    formatting.outline_level = 9
    formatting.rtl = False
    before = document.to_bytes()
    for bad in (-1, 10):
        with pytest.raises(ValueError, match="outline_level"):
            formatting.outline_level = bad
        assert document.to_bytes() == before
    with pytest.raises(ValueError):
        formatting.set_border("top", "single", size=12, color="red")
    assert document.to_bytes() == before
    reopened = Document.from_bytes(before)
    paragraph = reopened.tables[0].cell(0, 0).paragraphs[0] if cell_owned else reopened.paragraphs[0]
    restored = paragraph.paragraph_format
    assert restored.border("bar") == ("double", 12, RGBColor(18, 52, 86))
    assert restored.shading == "auto" and restored.outline_level == 9 and restored.rtl is False
    formatting.clear_border("top")
    assert formatting.border("top") is None and formatting.border("bar") is not None
    formatting.clear_borders()
    formatting.shading = None
    formatting.outline_level = None
    formatting.rtl = None
    assert formatting.border("bar") is None and formatting.shading is None
    assert formatting.outline_level is None and formatting.rtl is None

def test_fx200_style_keyword_snapshots_round_trip_and_explicit_empty_clears(tmp_path):
    from rdocx import Document, Pt, Inches
    from rdocx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from xml.etree import ElementTree as ET
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    document = Document()
    document.add_style("Complete", font_name="Gelasio", bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5,
        underline=WD_UNDERLINE.DOUBLE, keep_with_next=True,
        keep_together=False, page_break_before=False,
        borders={"bottom": ("single", 8, "123456")}, shading=RGBColor(1, 2, 3),
        tab_stops=[(Inches(1), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)])
    paragraph = document.add_paragraph("styled output")
    paragraph.style = "Complete"
    path = tmp_path / "style-snapshot.docx"
    document.save(path)
    reopened = Document(path)
    def style_node(document):
        with ZipFile(BytesIO(document.to_bytes())) as package:
            root = ET.fromstring(package.read("word/styles.xml"))
        return next(node for node in root.findall(w + "style") if node.get(w + "styleId") == "Complete")
    node = style_node(reopened)
    ppr, rpr = node.find(w + "pPr"), node.find(w + "rPr")
    assert ppr.find(w + "jc").get(w + "val") == "center"
    assert ppr.find(w + "spacing").get(w + "line") == "360"
    assert ppr.find(w + "spacing").get(w + "lineRule") == "auto"
    assert ppr.find(w + "keepNext") is not None
    assert ppr.find(w + "keepLines").get(w + "val") in ("0", "false")
    assert ppr.find(w + "pageBreakBefore").get(w + "val") in ("0", "false")
    assert ppr.find(w + "pBdr").find(w + "bottom").get(w + "color") == "123456"
    assert ppr.find(w + "shd").get(w + "fill") == "010203"
    assert ppr.find(w + "tabs").find(w + "tab").get(w + "pos") == "1440"
    assert rpr.find(w + "u").get(w + "val") == "double"
    reopened.set_style("Complete", borders={}, tab_stops=[], keep_with_next=False, line_spacing=Pt(18))
    reopened.save(path)
    final = Document(path)
    node = style_node(final)
    ppr, rpr = node.find(w + "pPr"), node.find(w + "rPr")
    selected_borders = ppr.findall(w + "pBdr")[-1]
    assert all(selected_borders.find(w + edge) is None
               for edge in ("top", "bottom", "left", "right", "between", "bar"))
    selected_tabs = ppr.findall(w + "tabs")[-1]
    assert selected_tabs.findall(w + "tab") == []
    assert ppr.find(w + "jc").get(w + "val") == "center"
    assert ppr.find(w + "spacing").get(w + "line") == "360"
    assert ppr.find(w + "spacing").get(w + "lineRule") == "exact"
    assert ppr.find(w + "shd").get(w + "fill") == "010203"
    assert rpr.find(w + "u").get(w + "val") == "double"
    assert final.paragraphs[-1].text == "styled output"
    assert final.paragraphs[-1].style == "Complete"
    saved = final.to_bytes()
    restored = Document.from_bytes(saved)
    restored_ppr = style_node(restored).find(w + "pPr")
    selected_borders = restored_ppr.findall(w + "pBdr")[-1]
    assert all(selected_borders.find(w + edge) is None
               for edge in ("top", "bottom", "left", "right", "between", "bar"))
    selected_tabs = restored_ppr.findall(w + "tabs")[-1]
    assert selected_tabs.findall(w + "tab") == []
    assert restored.to_bytes() == saved
    restored.set_style("Complete", borders={}, tab_stops=[])
    assert restored.to_bytes() == saved


def test_fx200_style_keyword_validation_precedes_every_source_commit():
    from rdocx import Document, Inches
    from rdocx.enum.text import WD_TAB_ALIGNMENT
    document = Document()
    document.add_style("Existing", bold=True)
    paragraph = document.add_paragraph("still live")
    for keyword in [dict(alignment=999), dict(underline=999), dict(line_spacing=float("nan")),
        dict(borders={"unknown": ("single", 8, "123456")}), dict(shading="invalid"),
        dict(tab_stops=[(Inches(1), WD_TAB_ALIGNMENT.LEFT, 999)])]:
        before = document.to_bytes()
        with pytest.raises(ValueError): document.set_style("Existing", italic=True, **keyword)
        assert document.to_bytes() == before and paragraph.text == "still live"
        with pytest.raises(ValueError): document.add_style("MustNotExist", italic=True, **keyword)
        assert document.to_bytes() == before
        assert all(style.name != "MustNotExist" for style in document.styles)
    for keyword in [dict(alignment=0), dict(keep_together=False), dict(borders={}), dict(tab_stops=[])]:
        before = document.to_bytes()
        with pytest.raises(ValueError, match="character style"):
            document.add_style("CharacterRefusal", style_type="character", **keyword)
        assert document.to_bytes() == before


def test_fx200_style_update_preserves_imported_unrelated_xml_and_theme_clear_intent():
    from rdocx import Document
    from xml.etree import ElementTree as ET
    document = Document()
    document.add_style("Imported", bold=True)
    source = document.to_bytes()
    output = BytesIO()
    with ZipFile(BytesIO(source)) as package, ZipFile(output, "w") as target:
        for member in package.infolist():
            data = package.read(member.filename)
            if member.filename == "word/styles.xml":
                text = data.decode()
                start = text.index('w:styleId="Imported"')
                end = text.index('</w:style>', start)
                part = text[start:end]
                assert "<w:rPr>" in part
                part = part.replace("<w:rPr>", '<w:rPr xmlns:f="urn:foreign" f:keep="&#65;"><w:rFonts w:asciiTheme="minorHAnsi" w:hint="cs"/><f:opaque token="same"/>', 1)
                text = text[:start] + part + text[end:]
                data = text.encode()
            target.writestr(member, data)
    document = Document.from_bytes(output.getvalue())
    document.set_style("Imported", font_name="Gelasio", color="123456", keep_together=False)
    with ZipFile(BytesIO(document.to_bytes())) as package:
        saved = package.read("word/styles.xml")
    assert b'<f:opaque token="same"/>' in saved
    assert b'f:keep="&#65;"' in saved
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    node = next(node for node in ET.fromstring(saved).findall(w+"style") if node.get(w+"styleId")=="Imported")
    fonts = node.find(w+"rPr").find(w+"rFonts")
    assert fonts.get(w+"ascii")=="Gelasio" and fonts.get(w+"hint")=="cs"
    assert fonts.get(w+"asciiTheme") is None
    reopened=Document.from_bytes(document.to_bytes())
    assert next(style for style in reopened.styles if style.style_id=="Imported").name=="Imported"


@pytest.mark.parametrize("cell_owned",[False,True])
@pytest.mark.parametrize("name,value",[
    ("superscript",True),("subscript",True),("all_caps",False),("small_caps",True),
    ("double_strike",True),("hidden",False),("rtl",True),("language","en-US"),
    ("language_east_asia","ja-JP"),("language_bidi","ar-SA"),
    ("name_east_asia","Carlito"),("name_complex_script","Carlito"),("character_spacing",-635),
])
def test_fx200_font_added_direct_properties_round_trip_and_clear(cell_owned,name,value):
    from rdocx import Document
    document=Document()
    if cell_owned: paragraph=document.add_table(1,1).cell(0,0).paragraphs[0]
    else: paragraph=document.add_paragraph("")
    run=paragraph.add_run("direct")
    font=run.font
    assert getattr(font,name) is None
    before=document.to_bytes()
    setattr(font,name,None)
    assert document.to_bytes()==before
    setattr(font,name,value)
    assert getattr(font,name)==value
    reopened=Document.from_bytes(document.to_bytes())
    restored=(reopened.tables[0].cell(0,0).paragraphs[0] if cell_owned else reopened.paragraphs[0]).runs[0].font
    assert getattr(restored,name)==value
    setattr(restored,name,None)
    assert getattr(restored,name) is None
    final=Document.from_bytes(reopened.to_bytes())
    final_run=(final.tables[0].cell(0,0).paragraphs[0] if cell_owned else final.paragraphs[0]).runs[0]
    assert getattr(final_run.font,name) is None and final_run.text=="direct"


@pytest.mark.parametrize("cell_owned",[False,True])
def test_fx200_font_spacing_checked_range_preserves_bytes_and_live_handles(cell_owned):
    from rdocx import Document
    document=Document()
    paragraph=document.add_table(1,1).cell(0,0).paragraphs[0] if cell_owned else document.add_paragraph("")
    run=paragraph.add_run("range")
    font=run.font
    before=document.to_bytes()
    for invalid in [(2**31)*635,(-2**31-1)*635]:
        with pytest.raises(ValueError,match="twip range"): font.character_spacing=invalid
        assert document.to_bytes()==before and run.text=="range" and font.character_spacing is None


def test_fx200_python_style_saved_output_paints_exact_expected_colours_with_deterministic_native_cli(tmp_path):
    from rdocx import Document, Pt
    from PIL import Image
    from pathlib import Path
    import os,subprocess
    # The existing CLI render path uses layout_deterministic_with_options.
    # It consumes Python-authored saved OOXML. This is a companion native paint
    # gate, rather than adding a new public Python deterministic render API.
    root=Path(__file__).resolve().parents[3]
    # Coordinator builds and authenticates this exact current-prefix binary
    # before admitting this native companion gate. No implicit Cargo fallback.
    cli_path=os.environ.get("RDOCX_CLI")
    assert cli_path, "native companion gate requires coordinator-authenticated RDOCX_CLI"
    assert Path(cli_path).is_file(), "RDOCX_CLI must identify the admitted native binary"
    cli=[cli_path]
    document=Document()
    document.add_style("Paint",font_name="Carlito",font_size=Pt(24),color="FF0000",shading="00FF00")
    paragraph=document.add_paragraph("MMMMMMMM")
    paragraph.style="Paint"
    source=tmp_path/"python-style-paint.docx"
    rendered=[]
    for index,(foreground,background) in enumerate([("FF0000","00FF00"),("0000FF","FFFF00")]):
        if index: document.set_style("Paint",color=foreground,shading=background)
        document.save(source)
        saved=source.read_bytes()
        output=tmp_path/f"native-{index}"
        result=subprocess.run(cli+["render",str(source),"--output-dir",str(output),"--dpi","144","--page","0"],cwd=root,capture_output=True,text=True,timeout=300)
        assert result.returncode==0,(result.stdout,result.stderr)
        assert source.read_bytes()==saved
        images=list(output.glob("*.png"))
        assert len(images)==1
        with Image.open(images[0]) as image: pixels=list(image.convert("RGB").getdata())
        expected_fg=tuple(bytes.fromhex(foreground));expected_bg=tuple(bytes.fromhex(background))
        assert sum(pixel==expected_fg for pixel in pixels)>20
        assert sum(pixel==expected_bg for pixel in pixels)>200
        rendered.append(images[0].read_bytes())
        assert Document.from_bytes(saved).paragraphs[0].text=="MMMMMMMM"
    assert rendered[0]!=rendered[1]



def test_fx200_checklist_presets_create_real_bullet_definitions_and_reference_them_on_reopen():
    from rdocx import Document, ListLevel
    from xml.etree import ElementTree as ET
    document=Document()
    expected=[]
    for checked in [False,True]:
        preset=ListLevel.checklist(checked=checked)
        marker="☑" if checked else "☐"
        assert preset.format=="bullet" and preset.text==marker
        definition=document.add_numbering_definition([preset])
        instance=document.add_numbering_instance(definition)
        paragraph=document.add_paragraph("checked" if checked else "unchecked")
        paragraph.numbering=(instance,0)
        expected.append((definition,instance,marker))
    saved=document.to_bytes()
    with ZipFile(BytesIO(saved)) as package:
        numbering=ET.fromstring(package.read("word/numbering.xml"))
    w="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for definition,instance,marker in expected:
        abstract=next(node for node in numbering.findall(w+"abstractNum") if node.get(w+"abstractNumId")==str(definition))
        level=abstract.find(w+"lvl")
        assert level.find(w+"numFmt").get(w+"val")=="bullet"
        assert level.find(w+"lvlText").get(w+"val")==marker
        num=next(node for node in numbering.findall(w+"num") if node.get(w+"numId")==str(instance))
        assert num.find(w+"abstractNumId").get(w+"val")==str(definition)
    reopened=Document.from_bytes(saved)
    assert [paragraph.text for paragraph in reopened.paragraphs]==["unchecked","checked"]
    assert [paragraph.numbering for paragraph in reopened.paragraphs]==[(instance,0) for _,instance,_ in expected]
    assert document.to_bytes()==saved
    # Actual U2610/U2611 deterministic glyph paint remains the explicit joint
    # F200/F210 gate against selected3968-byte provider. No square substitution
    # or missing glyph is accepted by this semantic-only control.


@pytest.mark.parametrize("clear", [False, True])
@pytest.mark.parametrize("started", [False, True])
def test_fx200_imported_style_border_root_and_duplicate_clear_survive_save_reopen(clear, started):
    from rdocx import Document
    from xml.etree import ElementTree as ET
    document=Document()
    document.add_style("ImportedBorder", keep_with_next=True)
    original=document.to_bytes()
    output=BytesIO()
    w="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    border='''<q:pBdr f:first="&#65;"><q:top q:val="single"/><f:opaque first="same"/></q:pBdr><f:pBdr f:keep="foreign"/><q:pBdr f:last="&#66;"><q:bottom q:val="single" f:edge="&#67;"/><!--border--><f:opaque>  keep  </f:opaque></q:pBdr>'''
    if started:
        border=border.replace('<q:bottom q:val="single" f:edge="&#67;"/>', '<q:bottom q:val="single" f:edge="&#67;"><f:inside>  edge  </f:inside><!--inside--></q:bottom>')
    with ZipFile(BytesIO(original)) as package, ZipFile(output,"w") as target:
        for member in package.infolist():
            data=package.read(member.filename)
            if member.filename=="word/styles.xml":
                text=data.decode()
                start=text.index('w:styleId="ImportedBorder"')
                end=text.index('</w:style>',start)
                part=text[start:end]
                assert '<w:pPr>' in part
                part=part.replace('<w:pPr>', '<w:pPr xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:f="urn:foreign" f:outer="&#68;">'+border,1)
                data=(text[:start]+part+text[end:]).encode()
            target.writestr(member,data)
    document=Document.from_bytes(output.getvalue())
    document.set_style("ImportedBorder", borders={} if clear else {"bottom":("double",8,"123456")})
    saved=document.to_bytes()
    with ZipFile(BytesIO(saved)) as package: raw=package.read("word/styles.xml")
    for preserved in [b'f:first="&#65;"', b'f:last="&#66;"', b'f:outer="&#68;"', b'<f:pBdr f:keep="foreign"/>', b'<!--border-->', b'<f:opaque>  keep  </f:opaque>']:
        assert preserved in raw
    if clear:
        assert b'f:edge="&#67;"' not in raw
        assert b'<f:inside>  edge  </f:inside>' not in raw and b'<!--inside-->' not in raw
    else:
        assert b'f:edge="&#67;"' in raw
        if started:
            assert b'<f:inside>  edge  </f:inside>' in raw and b'<!--inside-->' in raw
    style=next(node for node in ET.fromstring(raw).findall(w+"style") if node.get(w+"styleId")=="ImportedBorder")
    properties=style.findall(w+"pPr")[-1]
    selected=properties.findall(w+"pBdr")[-1]
    assert properties.find(w+"keepNext") is not None
    if clear:
        assert all(selected.find(w+edge) is None for edge in ["top","bottom","left","right","between","bar"])
    else:
        edge=selected.find(w+"bottom")
        assert edge.get(w+"val")=="double" and edge.get(w+"sz")=="8" and edge.get(w+"color")=="123456"
    reopened=Document.from_bytes(saved)
    assert reopened.to_bytes()==saved
    with ZipFile(BytesIO(reopened.to_bytes())) as package: final=ET.fromstring(package.read("word/styles.xml"))
    final_style=next(node for node in final.findall(w+"style") if node.get(w+"styleId")=="ImportedBorder")
    final_border=final_style.findall(w+"pPr")[-1].findall(w+"pBdr")[-1]
    assert ET.tostring(final_border)==ET.tostring(selected)


@pytest.mark.parametrize("physical,logical",[("start","left"),("end","right")])
def test_fx200_imported_started_logical_border_setters_keep_one_owner_payload_and_shadowed_peers(physical,logical):
    from rdocx import Document
    from xml.etree import ElementTree as ET
    document=Document()
    document.add_style("Logical",keep_with_next=True)
    original=document.to_bytes()
    opaque='<f:opaque xmlns:q="urn:shadow"><q:bottom q:val="double"></q:bottom><w:bottom w:val="double"></w:bottom></f:opaque>'
    border=f'<q:pBdr f:keep="&#65;">{opaque}<q:{physical} q:val="single" f:edge="&#66;"><f:payload>  unique  </f:payload><!--unique--></q:{physical}></q:pBdr>'
    output=BytesIO()
    with ZipFile(BytesIO(original)) as package,ZipFile(output,"w") as target:
        for member in package.infolist():
            data=package.read(member.filename)
            if member.filename=="word/styles.xml":
                text=data.decode();start=text.index('w:styleId="Logical"');end=text.index('</w:style>',start);part=text[start:end]
                assert '<w:pPr>' in part
                part=part.replace('<w:pPr>','<w:pPr xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:f="urn:foreign">'+border,1)
                data=(text[:start]+part+text[end:]).encode()
            target.writestr(member,data)
    document=Document.from_bytes(output.getvalue())
    w="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for style,size,color in [("double",8,"123456"),("single",12,"654321")]:
        document.set_style("Logical",borders={logical:(style,size,color)})
        saved=document.to_bytes()
        with ZipFile(BytesIO(saved)) as package:raw=package.read("word/styles.xml")
        assert opaque.encode() in raw
        assert raw.count(f'<q:{physical} '.encode())==1 and raw.count(f'</q:{physical}>'.encode())==1
        assert raw.count(b'<f:payload>  unique  </f:payload>')==1 and raw.count(b'<!--unique-->')==1
        node=next(node for node in ET.fromstring(raw).findall(w+"style") if node.get(w+"styleId")=="Logical")
        selected=node.find(w+"pPr").findall(w+"pBdr")[-1]
        assert len(selected.findall(w+physical))==1 and selected.find(w+"bottom") is None
        edge=selected.find(w+physical)
        assert (edge.get(w+"val"),edge.get(w+"sz"),edge.get(w+"color"))==(style,str(size),color)
        document=Document.from_bytes(saved)
        assert document.to_bytes()==saved
        document.set_style("Logical",borders={logical:(style,size,color)})
        assert document.to_bytes()==saved


@pytest.mark.parametrize("cell_owned", [False, True])
@pytest.mark.parametrize("initial", [None, "baseline", "superscript", "subscript", "inherited-super", "inherited-sub"])
@pytest.mark.parametrize("operation", ["superscript", "subscript"])
@pytest.mark.parametrize("value", [None, False, True])
def test_fx200_installed_super_sub_complete_transition_matrix(cell_owned, initial, operation, value):
    from rdocx import Document, StaleElementError
    from xml.etree import ElementTree as ET

    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    document = Document()
    paragraph = document.add_table(1, 1).cell(0, 0).paragraphs[0] if cell_owned else document.add_paragraph("")
    paragraph.add_run("matrix")
    data = document.to_bytes()
    inherited = initial in ("inherited-super", "inherited-sub")
    direct = None if inherited else initial
    if initial is not None:
        properties = ('<w:rStyle w:val="InheritedScript"/>' if inherited else '<w:vertAlign w:val="' + initial + '"/>')
        data = _replace_document_xml(data, b'<w:r><w:t>matrix</w:t></w:r>',
            ('<w:r><w:rPr>' + properties + '</w:rPr><w:t>matrix</w:t></w:r>').encode())
    if inherited:
        output = BytesIO()
        inherited_mode = "superscript" if initial == "inherited-super" else "subscript"
        with ZipFile(BytesIO(data)) as source, ZipFile(output, "w") as target:
            for member in source.infolist():
                payload = source.read(member.filename)
                if member.filename == "word/styles.xml":
                    assert payload.count(b"</w:styles>") == 1
                    style = ('<w:style w:type="character" w:styleId="InheritedScript"><w:name w:val="InheritedScript"/>'
                             '<w:rPr><w:vertAlign w:val="' + inherited_mode + '"/></w:rPr></w:style>').encode()
                    payload = payload.replace(b"</w:styles>", style + b"</w:styles>")
                target.writestr(member, payload)
        data = output.getvalue()
    document = Document.from_bytes(data)
    paragraph = document.tables[0].cell(0, 0).paragraphs[0] if cell_owned else document.paragraphs[0]
    run = paragraph.runs[0]
    font = run.font
    live = run.font

    def package_run(bytes_):
        with ZipFile(BytesIO(bytes_)) as archive:
            xml = ET.fromstring(archive.read("word/document.xml"))
            styles = ET.fromstring(archive.read("word/styles.xml"))
        nodes = [node for node in xml.iter(w + "r") if ''.join(t.text or '' for t in node.findall(w + "t")) == "matrix"]
        assert len(nodes) == 1
        return nodes[0], styles

    def assert_direct(owner_font, mode):
        assert owner_font.superscript is (None if mode is None else mode == "superscript")
        assert owner_font.subscript is (None if mode is None else mode == "subscript")

    before = document.to_bytes()
    original_node, original_styles = package_run(before)
    assert_direct(font, direct)
    if initial is None:
        assert original_node.find(w + "rPr") is None
    if inherited:
        inherited_style = next(node for node in original_styles.findall(w + "style") if node.get(w + "styleId") == "InheritedScript")
        assert inherited_style.find(w + "rPr").find(w + "vertAlign").get(w + "val") == inherited_mode
        assert run.style_id == "InheritedScript"
        assert original_node.find(w + "rPr").find(w + "vertAlign") is None
    setattr(font, operation, value)
    expected = operation if value is True else (None if value is None or direct == operation else direct)
    assert_direct(font, expected)
    assert_direct(live, expected)
    assert run.text == "matrix" and paragraph.text == "matrix"
    saved = document.to_bytes()
    node, styles = package_run(saved)
    properties = node.find(w + "rPr")
    alignments = [] if properties is None else properties.findall(w + "vertAlign")
    assert len(alignments) == (0 if expected is None else 1)
    if expected is not None:
        assert alignments[0].get(w + "val") == expected
    if inherited:
        assert properties.find(w + "rStyle").get(w + "val") == "InheritedScript"
        assert ET.tostring(styles) == ET.tostring(original_styles)
    if initial is None and value is not True:
        assert properties is None, "absent False/None must not materialize run properties"
    if expected == direct:
        assert saved == before, "opposite False, baseline False and absent/null no-ops preserve bytes"
    reopened = Document.from_bytes(saved)
    restored = (reopened.tables[0].cell(0, 0).paragraphs[0] if cell_owned else reopened.paragraphs[0]).runs[0]
    assert_direct(restored.font, expected)
    assert restored.text == "matrix"
    assert reopened.to_bytes() == saved
    again = Document.from_bytes(reopened.to_bytes())
    final = (again.tables[0].cell(0, 0).paragraphs[0] if cell_owned else again.paragraphs[0]).runs[0]
    assert_direct(final.font, expected)
    assert final.text == "matrix" and again.to_bytes() == saved
    document.add_paragraph("topology change")
    with pytest.raises(StaleElementError):
        _ = live.superscript


def test_fx200_checked_list_helpers_start_overrides_and_fresh_returned_handles():
    from rdocx import Document, ListLevel, StaleElementError, RdocxError
    from xml.etree import ElementTree as ET
    from io import BytesIO
    from zipfile import ZipFile
    import pytest
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    document = Document()
    original = document.add_paragraph("before")
    for name in ["add_bullet_list_item", "add_numbered_list_item"]:
        before = document.to_bytes()
        for level in [-1, 9]:
            with pytest.raises(ValueError): getattr(document, name)("bad", level)
            assert document.to_bytes() == before
        assert original.text == "before"
        for level in range(9):
            paragraph = getattr(document, name)("item", level)
            assert paragraph.text == "item"
            if level == 0:
                with pytest.raises(StaleElementError, match="Document.add_"+name[4:]):
                    _ = original.text
            paragraph.add_run(" added")
        original = document.paragraphs[0]
    definition = document.add_numbering_definition([ListLevel(format="decimal", text="%1.")])
    plain = document.add_numbering_instance(definition)
    restarted = document.add_numbering_instance(definition, start=7)
    zero = document.add_numbering_instance(definition, start=0)
    before = document.to_bytes()
    for kwargs in [dict(level=1), dict(start=-1), dict(start=2**32), dict(start=1, level=9)]:
        with pytest.raises(ValueError): document.add_numbering_instance(definition, **kwargs)
        assert document.to_bytes() == before
    with pytest.raises(RdocxError, match="definition"):
        document.add_numbering_instance(2**32-1)
    assert document.to_bytes() == before
    with ZipFile(BytesIO(before)) as package:
        numbering = ET.fromstring(package.read("word/numbering.xml"))
    by_id = {int(node.get(w+"numId")): node for node in numbering.findall(w+"num")}
    assert by_id[plain].find(w+"lvlOverride") is None
    for instance, start in [(restarted, "7"), (zero, "0")]:
        owner = by_id[instance]
        assert owner.find(w+"abstractNumId").get(w+"val") == str(definition)
        override = owner.find(w+"lvlOverride")
        assert override.get(w+"ilvl") == "0" and override.find(w+"startOverride").get(w+"val") == start
    reopened = Document.from_bytes(before)
    assert len(reopened.paragraphs) == 19
    for _ in range(2):
        reopened = Document.from_bytes(reopened.to_bytes())
        assert len(reopened.paragraphs) == 19


@pytest.mark.parametrize("operation", ["add_row", "insert_column", "delete_column", "split"])
def test_fx202_issue305_table_surface_on_installed_base(operation):
    from rdocx import Document

    document = Document()
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "retained"
    table = document.tables[0]
    if operation == "add_row":
        row = table.add_row()
        assert len(row.cells) == 2
        assert [cell.text for cell in row.cells] == ["", ""]
        assert len(document.tables[0].rows) == 3
    elif operation == "insert_column":
        table.insert_column(0, before=True)
        assert len(document.tables[0].rows[0].cells) == 3
        assert document.tables[0].cell(0, 1).text == "retained"
    elif operation == "delete_column":
        table.delete_column(1)
        assert len(document.tables[0].rows[0].cells) == 1
        assert document.tables[0].cell(0, 0).text == "retained"
    else:
        before = document.to_bytes()
        table.cell(0, 0).split()
        assert document.to_bytes() == before
        assert table.cell(0, 0).text == "retained"
    assert Document.from_bytes(document.to_bytes()).tables[0].cell(0, 0).text in (
        "retained", ""
    )


@pytest.mark.parametrize("property_name", [
    "autofit", "first_row", "last_row", "first_column", "last_column", "banded_rows", "banded_columns"
])
def test_fx202_table_look_and_layout_on_installed_base(property_name):
    from rdocx import Document

    document = Document()
    table = document.add_table(rows=1, cols=1)
    setattr(table, property_name, True)
    assert getattr(table, property_name) is True
    assert getattr(Document.from_bytes(document.to_bytes()).tables[0], property_name) is True
    setattr(table, property_name, False)
    assert getattr(table, property_name) is False


def test_fx202_table_style_rejects_undefined_without_publication():
    from rdocx import Document

    document = Document()
    table = document.add_table(rows=1, cols=1)
    before = document.to_bytes()
    with pytest.raises(KeyError):
        table.style = "X202UndefinedTableStyle"
    assert document.to_bytes() == before


def test_fx202_completed_nested_table_provider_remains_available():
    from rdocx import Document

    document = Document()
    table = document.add_table(rows=1, cols=1)
    nested = table.cell(0, 0).add_table(rows=1, cols=1)
    nested.cell(0, 0).text = "nested retained"
    reopened = Document.from_bytes(document.to_bytes())
    assert reopened.tables[0].cell(0, 0).tables[0].cell(0, 0).text == "nested retained"


def test_fx202_direct_table_after_nested_owner_does_not_edit_the_nested_table():
    from rdocx import Document
    document = Document()
    nested = document.add_table(rows=1, cols=1).cell(0, 0).add_table(rows=1, cols=1)
    nested.cell(0, 0).text = "nested retained"
    document.add_table(rows=1, cols=1)
    assert len(document.tables) == 2
    document.tables[1].cell(0, 0).text = "second direct retained"
    assert document.tables[0].cell(0, 0).tables[0].cell(0, 0).text == "nested retained"
    assert document.tables[1].cell(0, 0).text == "second direct retained"
    reopened = Document.from_bytes(document.to_bytes())
    assert reopened.tables[0].cell(0, 0).tables[0].cell(0, 0).text == "nested retained"
    assert reopened.tables[1].cell(0, 0).text == "second direct retained"


def test_fx202_table_look_setter_preserves_default_banding():
    import rdocx
    document = rdocx.Document()
    table = document.add_table(1, 1)
    assert table.banded_rows is True
    assert table.banded_columns is True
    table.first_row = True
    assert table.first_row is True
    assert table.banded_rows is True
    assert table.banded_columns is True


def _fresh_run_text(document, table=0, row=0, cell=0, run=0):
    return document.tables[table].cell(row, cell).paragraphs[0].runs[run].text


def _one_pixel_png():
    import struct
    import zlib

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    header = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    pixels = zlib.compress(bytes((0, 1, 2, 3)))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", pixels) + chunk(b"IEND", b"")


def test_issue388_cell_property_edits_preserve_live_owners():
    """Expected on old source: a fresh cell run lookup raises ValueError
    'element has no current checked physical owner'. A nested held run raises
    RdocxError because its story item is no longer a paragraph. Reopen of the
    edited package still contains the cell text.
    """
    from rdocx import Document

    document = Document()
    table = document.add_table(rows=1, cols=2)
    cell = table.cell(0, 0)
    cell.text = "kept"
    held = cell.paragraphs[0].runs[0]
    cell.shading = "D9D9D9"
    cell.width = 914400
    table.set_column_width(0, 914400)
    assert held.text == "kept"
    assert cell.paragraphs[0].runs[0].text == "kept"
    assert _fresh_run_text(document) == "kept"
    held.text = "later"
    reopened = Document.from_bytes(document.to_bytes())
    assert reopened.tables[0].cell(0, 0).paragraphs[0].runs[0].text == "later"
    assert reopened.tables[0].cell(0, 0).shading == RGBColor.from_string("D9D9D9")

    nested_document = Document()
    nested = nested_document.add_table(rows=1, cols=1).cell(0, 0).add_table(rows=1, cols=1)
    nested.cell(0, 0).text = "nested"
    nested_held = nested.cell(0, 0).paragraphs[0].runs[0]
    nested.cell(0, 0).shading = "ABCDEF"
    nested.cell(0, 0).width = 457200
    assert nested_held.text == "nested"
    fresh_nested = nested_document.tables[0].cell(0, 0).tables[0].cell(0, 0)
    assert fresh_nested.paragraphs[0].runs[0].text == "nested"

    empty = Document()
    empty_cell = empty.add_table(rows=1, cols=1).cell(0, 0)
    assert empty_cell.paragraphs[0].text == ""
    assert len(empty_cell.paragraphs[0].runs) == 0
    empty_cell.shading = "010203"
    assert empty_cell.paragraphs[0].text == ""
    assert len(empty_cell.paragraphs[0].runs) == 0
    assert Document.from_bytes(empty.to_bytes()).tables[0].cell(0, 0).paragraphs[0].text == ""

    many = Document()
    many_cell = many.add_table(rows=1, cols=1).cell(0, 0)
    many_cell.text = "one"
    many_cell.paragraphs[0].add_run(" two")
    many_cell.shading = "ABCDEF"
    assert [_fresh_run_text(many, run=index) for index in (0, 1)] == ["one", " two"]
    with pytest.raises(IndexError, match="run index out of range"):
        many.tables[0].cell(0, 0).paragraphs[0].runs[2].text


def test_issue388_property_edit_owner_boundaries():
    """Expected on old source: vertical alignment, borders, margins, grid
    widths, a non-removing grid span and vertical merge leave a fresh run
    lookup without a physical owner. Replacing the cell text still retires
    the held run. A grid span that removes a cell still retires that run.
    """
    from rdocx import Document, RdocxError, StaleElementError, WD_CELL_VERTICAL_ALIGNMENT

    document = Document()
    table = document.add_table(rows=2, cols=2)
    cell = table.cell(0, 0)
    cell.text = "kept"
    held = cell.paragraphs[0].runs[0]
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    cell.set_border("bottom", "single", size=4, color="000000")
    cell.set_margins(top=0, right=635, bottom=0, left=635)
    table.grid_widths = [914400, 914400]
    table.set_cell_grid_span(0, 0, 1)
    table.set_cell_vertical_merge(1, 0, "restart")
    assert held.text == "kept"
    assert cell.paragraphs[0].runs[0].text == "kept"
    assert _fresh_run_text(document) == "kept"
    assert document.tables[0].cell(1, 0).paragraphs[0].text == ""

    replaced = Document()
    replaced_cell = replaced.add_table(rows=1, cols=1).cell(0, 0)
    replaced_cell.text = "before"
    replaced_run = replaced_cell.paragraphs[0].runs[0]
    before = replaced.to_bytes()
    replaced_cell.text = "after"
    assert replaced.to_bytes() != before
    with pytest.raises(StaleElementError, match="Cell.text"):
        _ = replaced_run.text
    assert replaced.tables[0].cell(0, 0).paragraphs[0].runs[0].text == "after"
    assert Document.from_bytes(replaced.to_bytes()).tables[0].cell(0, 0).text == "after"

    spanned = Document()
    spanned_table = spanned.add_table(rows=1, cols=2)
    spanned_table.cell(0, 0).text = "origin"
    removed_paragraph = spanned_table.cell(0, 1).paragraphs[0]
    spanned_table.set_cell_grid_span(0, 0, 2)
    with pytest.raises(StaleElementError, match="Table.set_cell_grid_span"):
        _ = removed_paragraph.text
    assert _fresh_run_text(spanned) == "origin"

    refused = Document()
    refused_table = refused.add_table(rows=1, cols=2)
    refused_table.cell(0, 0).text = "origin"
    refused_table.cell(0, 1).text = "kept"
    kept_run = refused_table.cell(0, 1).paragraphs[0].runs[0]
    before_refusal = refused.to_bytes()
    with pytest.raises(RdocxError, match="nonempty"):
        refused_table.set_cell_grid_span(0, 0, 2)
    assert refused.to_bytes() == before_refusal
    assert kept_run.text == "kept"


def test_issue389_range_removal_is_one_owned_transaction():
    """Additive API. Old source has no remove_content_range, so this fails
    while resolving the method. That absence is not the owner-regression RED.
    The assertions below are the contract once the method exists.
    """
    from rdocx import Document, RdocxError

    document = Document()
    for text in ("zero", "one", "two", "three"):
        document.add_paragraph(text)
    assert document.remove_content_range(1, 1) == 0
    assert document.remove_content_range(4, 4) == 0
    assert [paragraph.text for paragraph in document.paragraphs] == [
        "zero",
        "one",
        "two",
        "three",
    ]
    assert document.remove_content_range(1, 2) == 1
    assert [paragraph.text for paragraph in document.paragraphs] == ["zero", "two", "three"]
    assert document.remove_content_range(1, 3) == 2
    assert [paragraph.text for paragraph in document.paragraphs] == ["zero"]
    reopened = Document.from_bytes(document.to_bytes())
    assert [paragraph.text for paragraph in reopened.paragraphs] == ["zero"]

    blocked = Document()
    blocked.add_paragraph("before")
    table = blocked.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "inside"
    nested = table.cell(0, 0).add_table(rows=1, cols=1)
    nested.cell(0, 0).text = "nested"
    blocked.add_paragraph("after")
    assert blocked.remove_content_range(2, 3) == 1
    assert blocked.tables[0].cell(0, 0).paragraphs[0].text == "inside"
    assert blocked.tables[0].cell(0, 0).tables[0].cell(0, 0).paragraphs[0].text == "nested"
    assert blocked.remove_content_range(1, 2) == 1
    assert len(blocked.tables) == 0
    assert [paragraph.text for paragraph in blocked.paragraphs] == ["before"]

    invalid = Document()
    invalid.add_paragraph("stay")
    invalid.add_paragraph("put")
    before = invalid.to_bytes()
    with pytest.raises((ValueError, IndexError, RdocxError)):
        invalid.remove_content_range(2, 1)
    with pytest.raises((ValueError, IndexError, RdocxError)):
        invalid.remove_content_range(0, 3)
    assert invalid.to_bytes() == before
    assert [paragraph.text for paragraph in invalid.paragraphs] == ["stay", "put"]


def test_issue389_edit_cost_preserves_checked_mutation_contract():
    """Control. A refused XML edit publishes nothing. Removing one of two
    body drawings that share one image part leaves that part in the package.
    """
    from rdocx import Document, XmlError

    document = Document()
    paragraph = document.add_paragraph("stay")
    before = document.to_bytes()
    with pytest.raises(XmlError):
        paragraph.replace_xml("<not-xml")
    assert document.to_bytes() == before
    assert paragraph.text == "stay"

    pictured = Document()
    pictured.add_paragraph("keep")
    png = _one_pixel_png()
    pictured.add_picture(png, "shared.png", 914400, 914400)
    xml = _document_xml(pictured.to_bytes())
    embed = xml.find('r:embed="rId1"')
    assert embed > 0
    start = xml.rfind("<w:p", 0, embed)
    end = xml.find("</w:p>", embed) + len("</w:p>")
    shared = _replace_document_xml(pictured.to_bytes(), b"</w:body>", (xml[start:end] + "</w:body>").encode())
    opened = Document.from_bytes(shared)
    assert opened.remove_content(1) is True
    package = ZipFile(BytesIO(opened.to_bytes()))
    assert "word/media/image1.png" in package.namelist()
    assert package.read("word/media/image1.png") == png
    assert 'r:embed="rId1"' in _document_xml(opened.to_bytes())


def test_issue389_local_edits_do_not_scale_with_unrelated_package():
    """Replace and remove the same paragraph in a 50-paragraph document and
    an 800-paragraph document. The edited paragraph changes and a later
    paragraph keeps its text. The visit bound is a library test. The release
    ratio is final-sprint evidence.
    """
    from rdocx import Document

    for count in (50, 800):
        document = Document()
        for index in range(count):
            document.add_paragraph(f"p{index}")
        paragraph = document.paragraphs[0]
        xml = paragraph.xml
        changed = xml.replace(">p0<", ">q0<", 1)
        assert changed != xml
        paragraph.replace_xml(changed)
        assert paragraph.text == "q0"
        assert document.remove_content(1) is True
        assert document.paragraphs[0].text == "q0"
        assert document.paragraphs[count - 2].text == f"p{count - 1}"


def _fx283_piece_work():
    # Debug-only structural diagnostics. Release runs retain the semantic cases.
    from rdocx import _rdocx

    reads = getattr(_rdocx, "_word_piece_work", None)
    comparisons = getattr(_rdocx, "_word_piece_alignments", None)
    if reads is None:
        return None
    return reads() + (comparisons() if comparisons is not None else 0)


def _fx283_body_document(body):
    from rdocx import Document

    seed = Document().to_bytes()
    old = _document_xml(seed).encode()
    new = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{body}<w:sectPr/></w:body></w:document>'
    ).encode()
    return Document.from_bytes(_replace_document_xml(seed, old, new))


@pytest.mark.parametrize("properties", ["", "<w:tcPr/>", '<w:tcPr><w:tcW w:w="900" w:type="dxa"/></w:tcPr>'])
def test_fx283_first_properties_relocate_held_routes_and_text_retires_descendants(properties):
    from rdocx import Document, StaleElementError, WD_CELL_VERTICAL_ALIGNMENT
    _word_piece_work = _fx283_piece_work

    document = _fx283_body_document(
        '<w:tbl><w:tr><w:tc>' + properties +
        '<w:p><w:r><w:t>first</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>later</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t>neighbor</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    )
    cell, neighbor = document.tables[0].cell(0, 0), document.tables[0].cell(0, 1)
    first, later = cell.paragraphs[0], cell.paragraphs[1]
    run = first.runs[0]
    _word_piece_work()
    cell.shading = "123456"
    work = _word_piece_work()
    assert work is None or work <= 3
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    assert (first.text, later.text, run.text) == ("first", "later", "first")
    _word_piece_work()
    cell.text = "after"
    work = _word_piece_work()
    assert work is None or work <= 3
    assert cell.text == "after\nlater" and neighbor.text == "neighbor"
    for held in (first, later, run):
        with pytest.raises(StaleElementError):
            _ = held.text
    paragraph = cell.paragraphs[1]
    _word_piece_work()
    paragraph.text = "changed later"
    work = _word_piece_work()
    assert work is None or work <= 3
    reopened = Document.from_bytes(document.to_bytes()).tables[0].cell(0, 0)
    assert reopened.text == cell.text == "after\nchanged later"


@pytest.mark.parametrize("count", [0, 40, 120])
def test_fx283_reused_cell_setters_do_not_visit_unrelated_body_pieces(count):
    from rdocx import Document, WD_CELL_VERTICAL_ALIGNMENT
    _word_piece_work = _fx283_piece_work

    document = Document()
    for _ in range(count):
        document.add_paragraph("unrelated")
    table = document.add_table(rows=2, cols=2)
    cells = [table.cell(row, cell) for row in range(2) for cell in range(2)]
    for name, value in [("text", "value"), ("vertical_alignment", WD_CELL_VERTICAL_ALIGNMENT.CENTER), ("shading", "123456")]:
        for cell in cells:
            _word_piece_work()
            setattr(cell, name, value)
            work = _word_piece_work()
            assert work is None or work <= 3
    assert [cell.text for cell in cells] == ["value"] * 4


@pytest.mark.parametrize("emu", [634, 635, 636, -634, -635, 2147483648 * 635, -2147483649 * 635])
def test_fx283_cell_width_keeps_signed_twip_truncation_and_wrap(emu):
    from rdocx import Document

    document = Document()
    cell = document.add_table(rows=1, cols=1).cell(0, 0)
    cell.width = emu
    twips = abs(emu) // 635 * (-1 if emu < 0 else 1)
    expected = ((twips + 2**31) % 2**32 - 2**31) * 635
    assert cell.width == expected
    assert Document.from_bytes(document.to_bytes()).tables[0].cell(0, 0).width == expected


def test_fx283_dormant_replacement_keeps_errors_counts_and_later_activation():
    from rdocx import Document, RdocxError, ReplacementCountError, StaleElementError

    source = _fx283_body_document('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>water</w:t></w:r></w:p>').to_bytes()
    document = Document.from_bytes(source)
    assert document.try_replace_text("missing", "none") == 0
    with pytest.raises(ReplacementCountError):
        document.try_replace_text("water", "wine", expect=2)
    with pytest.raises(RdocxError, match="replacement text.*U\\+0001"):
        document.try_replace_text("missing", "\x01")
    assert document.to_bytes() == source
    assert document.replace_all([("missing", "none", 0), ("water", "wine", 1), ("wine", "water", 1)]) == (0, 1, 1)
    assert document.try_replace_text("water", "wine") == 1
    saved = document.to_bytes()
    held = document.paragraphs[0]
    assert held.text == Document.from_bytes(saved).paragraphs[0].text == "wine"
    assert held.runs[0].font.bold is True
    assert document.try_replace_text("wine", "water", expect=1) == 1
    with pytest.raises(StaleElementError):
        _ = held.text


def test_fx283_missing_external_link_still_refuses_open():
    from rdocx import RdocxError

    with pytest.raises(RdocxError):
        _fx283_body_document('<w:p><w:hyperlink r:id="missing"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>')
    control = _fx283_body_document('<w:p><w:hyperlink w:anchor="inside"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>')
    assert control.paragraphs[0].text == "link"


def test_fx283_first_authority_access_rebuilds_current_facts_after_dormant_replace():
    from rdocx import Document
    _word_piece_work = _fx283_piece_work

    document = _fx283_body_document('<w:p><w:r><w:t>water</w:t></w:r></w:p>')
    assert document.try_replace_text("water", "wine") == 1
    saved = document.to_bytes()
    _word_piece_work()
    paragraph = document.paragraphs[0]
    work = _word_piece_work()
    assert work is None or work > 0, "first checked factory must validate current native facts afresh"
    assert paragraph.text == Document.from_bytes(saved).paragraphs[0].text == "wine"


@pytest.mark.parametrize("publisher", ["replace_image", "replace_image_for_story", "resize"])
def test_fx283_direct_images_and_failed_stream_save_rebuild_at_first_handle(publisher):
    import base64
    from rdocx import Document, RdocxError, Story
    _word_piece_work = _fx283_piece_work

    png = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4z8AAAAIAAeIhvDMAAAAASUVORK5CYII=")
    seed = Document()
    seed.add_picture(png, "before.png", width=1270, height=1905)
    rel = seed.pictures()[0].relationship_id
    field = b'<w:p xmlns:k="urn:opaque-scope"><w:fldSimple w:instr="MERGEFIELD name"><w:r><w:t>cache</w:t></w:r></w:fldSimple></w:p>'
    source = _replace_document_xml(seed.to_bytes(), b"<w:body>", b"<w:body>" + field)
    document = Document.from_bytes(source)
    with pytest.raises(RdocxError):
        document.replace_image("missing", png)
    if publisher == "replace_image":
        document.replace_image(rel, png)
    elif publisher == "replace_image_for_story":
        document.replace_image_for_story(Story(kind="body", part_name="/word/document.xml", owner_index=0), rel, png)
    else:
        assert document.set_picture_sizes_for_relationship(rel, 2540, 3175) == 1
    saved = document.to_bytes()

    class ShortStream:
        def write(self, value):
            return 0

    with pytest.raises(OSError, match="short or invalid"):
        document.save(ShortStream())
    _word_piece_work()
    actual = document.pictures()[0]
    work = _word_piece_work()
    assert work is None or work > 0
    expected = Document.from_bytes(saved).pictures()[0]
    assert (actual.width, actual.height, actual.bytes) == (expected.width, expected.height, expected.bytes)
    assert document.paragraphs[0].text == Document.from_bytes(saved).paragraphs[0].text


@pytest.mark.parametrize("properties", ["", "<w:tcPr/>", '<w:tcPr><w:tcW w:w="900" w:type="dxa"/></w:tcPr>'])
@pytest.mark.parametrize("fallback", [False, True])
def test_fx283_paragraph_text_keeps_its_distinct_existing_owner_policy(properties, fallback):
    from rdocx import Document, StaleElementError

    controlled = '<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t/></w:r></w:sdtContent></w:sdt>' if fallback else ""
    document = _fx283_body_document(
        '<w:tbl><w:tr><w:tc>' + properties + '<w:p>'
        '<w:r><w:t>first</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>later</w:t></w:r>' + controlled + '</w:p></w:tc></w:tr></w:tbl>'
    )
    table = document.tables[0]
    cell = table.cell(0, 0)
    selected, later = cell.paragraphs[0], cell.paragraphs[1]
    selected_format, later_format = selected.paragraph_format, later.paragraph_format
    selected_run, later_run = selected.runs[0], later.runs[0]
    cell.shading = "123456"
    assert selected.text == selected_run.text == "first"
    assert later.text == later_run.text == "later"
    selected.text = "changed paragraph"
    assert selected.text == "changed paragraph"
    assert later.text == "later"
    assert selected_format.space_after is None and later_format.space_after is None
    for run in (selected_run, later_run):
        with pytest.raises(StaleElementError):
            _ = run.text
    assert cell.text == "changed paragraph\nlater"
    assert Document.from_bytes(document.to_bytes()).tables[0].cell(0, 0).text == cell.text
    cell.text = "changed cell"
    assert cell.text == "changed cell\nlater" and len(table.rows) == 1
    for paragraph in (selected, later):
        with pytest.raises(StaleElementError):
            _ = paragraph.text
    for properties_view in (selected_format, later_format):
        with pytest.raises(StaleElementError):
            _ = properties_view.space_after


@pytest.mark.parametrize("factory", ["paragraphs", "tables", "section", "pictures", "stories", "comments", "story_items", "header_footer_variants", "hyperlinks", "revisions"])
def test_fx283_first_factory_after_field_serialization_uses_current_facts(factory):
    from rdocx import Document

    document = _fx283_body_document(
        '<w:p xmlns:k="urn:opaque-scope"><w:r><w:t>water </w:t></w:r>'
        '<w:fldSimple w:instr="MERGEFIELD name"><w:r><w:t>cache</w:t></w:r></w:fldSimple></w:p>'
    )
    assert document.try_replace_text("water", "wine") == 1
    saved = document.to_bytes()
    if factory == "section":
        document.section(0)
    elif factory == "pictures":
        document.pictures()
    else:
        getattr(document, factory)
    expected = Document.from_bytes(saved)
    assert document.paragraphs[0].text == expected.paragraphs[0].text
    assert [item.text for item in document.story_items] == [item.text for item in expected.story_items]
    assert document.to_bytes() == saved


def test_fx283_plain_hyperlink_records_remove_and_retarget_before_first_factory():
    from rdocx import Document, Hyperlink, RdocxError, Story

    plain = Story(kind="body", part_name="/word/document.xml", owner_index=0)
    document = _fx283_body_document('<w:p><w:hyperlink w:anchor="inside"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>')
    record = Hyperlink(story=plain, index_path=(0,), text="link", url=None, anchor="inside", relationship_id=None)
    invalid = Hyperlink(story=plain, index_path=(9,), text="missing", url=None, anchor="inside", relationship_id=None)
    before = document.to_bytes()
    with pytest.raises(RdocxError):
        document.remove_hyperlink(invalid)
    assert document.to_bytes() == before
    document.remove_hyperlink(record)
    assert document.paragraphs[0].text == "link" and document.hyperlinks == ()
    seed = Document()
    seed.add_paragraph("prefix").add_hyperlink("external", "https://before.example")
    link = seed.hyperlinks[0]
    record = Hyperlink(story=plain, index_path=link.index_path, text=link.text, url=link.url,
                      anchor=link.anchor, relationship_id=link.relationship_id)
    document = Document.from_bytes(seed.to_bytes())
    document.set_hyperlink_url(record, "https://after.example")
    assert document.hyperlinks[0].url == "https://after.example"
    assert Document.from_bytes(document.to_bytes()).hyperlinks[0].url == "https://after.example"


@pytest.mark.parametrize("operation", ["shading", "alignment", "cell_text", "paragraph_text"])
def test_fx283_clean_cell_edit_does_not_exchange_enclosing_and_nested_neighbor_authority(operation):
    from rdocx import Document, StaleElementError, WD_CELL_VERTICAL_ALIGNMENT

    document = _fx283_body_document(
        '<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w="1111" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/><w:i w:val="0"/></w:rPr><w:t>outer</w:t></w:r></w:p>'
        '<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w="2222" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:spacing w:after="240"/></w:pPr><w:r><w:rPr><w:b w:val="0"/><w:i/></w:rPr><w:t>inner</w:t></w:r></w:p>'
        '</w:tc></w:tr></w:tbl></w:tc><w:tc><w:p><w:r><w:t>clean</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    )
    table = document.tables[0]
    outer_cell, clean = table.cell(0, 0), table.cell(0, 1)
    inner_cell = outer_cell.tables[0].cell(0, 0)
    outer, inner = outer_cell.paragraphs[0], inner_cell.paragraphs[0]
    outer_run, inner_run = outer.runs[0], inner.runs[0]
    outer_font, inner_font = outer_run.font, inner_run.font
    outer_format, inner_format = outer.paragraph_format, inner.paragraph_format
    assert (outer.text, inner.text) == ("outer", "inner")
    assert (outer_format.space_after, inner_format.space_after) == (120 * 635, 240 * 635)
    assert (outer_font.bold, inner_font.bold, outer_font.italic, inner_font.italic) == (True, False, False, True)
    if operation == "shading":
        clean.shading = "123456"
    elif operation == "alignment":
        clean.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    elif operation == "cell_text":
        clean.text = "changed"
    else:
        clean.paragraphs[0].text = "changed"
    # Serialization remains correct even when an authority join is wrong.
    reopened = Document.from_bytes(document.to_bytes())
    assert reopened.tables[0].cell(0, 0).paragraphs[0].text == "outer"
    assert reopened.tables[0].cell(0, 0).tables[0].cell(0, 0).paragraphs[0].text == "inner"
    assert (outer.text, inner.text) == ("outer", "inner")
    assert (outer_cell.width, inner_cell.width) == (1111 * 635, 2222 * 635)
    assert (outer_format.space_after, inner_format.space_after) == (120 * 635, 240 * 635)
    if operation in ("shading", "alignment"):
        assert (outer_run.text, inner_run.text) == ("outer", "inner")
        assert (outer_font.bold, inner_font.bold, outer_font.italic, inner_font.italic) == (True, False, False, True)
    else:
        # Ordinary runs retain the existing document-epoch retirement policy.
        for run in (outer_run, inner_run):
            with pytest.raises(StaleElementError):
                _ = run.text
        for font in (outer_font, inner_font):
            with pytest.raises(StaleElementError):
                _ = font.bold


@pytest.mark.parametrize("neighbor", ["controls", "fields", "textboxes"])
@pytest.mark.parametrize("operation", ["shading", "alignment", "cell_text", "paragraph_text"])
def test_fx283_one_table_control_field_and_textbox_neighbors_keep_existing_authority(neighbor, operation):
    from rdocx import StaleElementError, WD_CELL_VERTICAL_ALIGNMENT

    def paragraph(label, index):
        run = '<w:b/><w:i w:val="0"/>' if index == 0 else '<w:b w:val="0"/><w:i/>'
        field = '<w:fldSimple w:instr="MERGEFIELD name"><w:r><w:t>cache</w:t></w:r></w:fldSimple>' if neighbor == "fields" else ""
        return f'<w:p><w:pPr><w:spacing w:after="{120 * (index + 1)}"/></w:pPr><w:r><w:rPr>{run}</w:rPr><w:t>{label}</w:t></w:r>{field}</w:p>'

    first, second = paragraph("first", 0), paragraph("second", 1)
    if neighbor == "controls":
        content = ''.join('<w:sdt><w:sdtPr/><w:sdtContent>' + value + '</w:sdtContent></w:sdt>' for value in (first, second))
    elif neighbor == "textboxes":
        shapes = ''.join(f'<v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="box{index}" style="width:10pt;height:10pt"><v:textbox><w:txbxContent>{value}</w:txbxContent></v:textbox></v:shape>' for index, value in enumerate((first, second)))
        content = '<w:p><w:r><w:pict>' + shapes + '</w:pict></w:r></w:p>'
    else:
        content = first + second
    document = _fx283_body_document('<w:tbl><w:tr><w:tc>' + content + '</w:tc><w:tc><w:p><w:r><w:t>clean</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
    if neighbor == "textboxes":
        held = [item.paragraph for item in document.story_items if item.kind == "paragraph" and item.story.kind == "text_box"]
    else:
        held = list(document.tables[0].cell(0, 0).paragraphs)
    assert len(held) == 2
    runs = [paragraph.runs[0] for paragraph in held]
    fonts = [run.font for run in runs]
    formats = [paragraph.paragraph_format for paragraph in held]
    expected = ["firstcache", "secondcache"] if neighbor == "fields" else ["first", "second"]
    assert [paragraph.text for paragraph in held] == expected
    assert [view.space_after for view in formats] == [120 * 635, 240 * 635]
    assert [view.bold for view in fonts] == [True, False]
    clean = document.tables[0].cell(0, 1)
    if operation == "shading":
        clean.shading = "123456"
    elif operation == "alignment":
        clean.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    elif operation == "cell_text":
        clean.text = "changed"
    else:
        clean.paragraphs[0].text = "changed"
    text_edit = operation in ("cell_text", "paragraph_text")
    if neighbor == "textboxes" and text_edit:
        # The existing checked publisher conservatively retires textbox owners.
        for paragraph, view in zip(held, formats):
            with pytest.raises(StaleElementError):
                _ = paragraph.text
            with pytest.raises(StaleElementError):
                _ = view.space_after
    else:
        assert [paragraph.text for paragraph in held] == expected
        assert [view.space_after for view in formats] == [120 * 635, 240 * 635]
    if text_edit:
        for run, view in zip(runs, fonts):
            with pytest.raises(StaleElementError):
                _ = run.text
            with pytest.raises(StaleElementError):
                _ = view.bold
    else:
        assert [run.text for run in runs] == ["first", "second"]
        assert [view.bold for view in fonts] == [True, False]


@pytest.mark.parametrize("neighbor", ["controlled_row", "controlled_cell"])
@pytest.mark.parametrize("before", [True, False])
@pytest.mark.parametrize("operation", ["cell_text", "paragraph_text", "shading", "alignment"])
def test_fx283_control_owned_cell_membership_preserves_physical_authority(neighbor, before, operation):
    from rdocx import Document, StaleElementError, WD_CELL_VERTICAL_ALIGNMENT

    controlled = ('<w:tc><w:tcPr><w:tcW w:w="2222" w:type="dxa"/></w:tcPr>'
                  '<w:p><w:pPr><w:spacing w:after="240"/></w:pPr>'
                  '<w:r><w:rPr><w:b/></w:rPr><w:t>controlled</w:t></w:r></w:p></w:tc>')
    clean_xml = ('<w:tc><w:p><w:pPr><w:spacing w:after="120"/></w:pPr>'
                 '<w:r><w:rPr><w:i/></w:rPr><w:t>clean</w:t></w:r></w:p></w:tc>')
    if neighbor == "controlled_row":
        controlled = '<w:tr>' + controlled + '</w:tr>'
        clean_xml = '<w:tr>' + clean_xml + '</w:tr>'
    controlled = '<w:sdt><w:sdtPr/><w:sdtContent>' + controlled + '</w:sdtContent></w:sdt>'
    content = controlled + clean_xml if before else clean_xml + controlled
    if neighbor == "controlled_cell":
        content = '<w:tr>' + content + '</w:tr>'
    document = _fx283_body_document('<w:tbl>' + content + '</w:tbl>')
    table = document.tables[0]
    clean = table.cell(0, 0)
    selected = clean.paragraphs[0]
    held = next(item.paragraph for item in document.story_items
                if item.kind == "paragraph" and item.text == "controlled")
    run, selected_run = held.runs[0], selected.runs[0]
    view, selected_view = held.paragraph_format, selected.paragraph_format
    font = run.font
    assert clean.text == selected.text == "clean"
    assert held.text == run.text == "controlled"
    assert (view.space_after, selected_view.space_after) == (240 * 635, 120 * 635)
    if operation == "cell_text":
        clean.text = "changed"
    elif operation == "paragraph_text":
        selected.text = "changed"
    elif operation == "shading":
        clean.shading = "123456"
    else:
        clean.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    text_edit = operation in ("cell_text", "paragraph_text")
    expected = "changed" if text_edit else "clean"
    physical = [item for item in document.story_items if item.kind == "paragraph"]
    assert sorted(item.text for item in physical) == sorted(["controlled", expected])
    assert len({item.story.owner_index for item in physical}) == 2
    if text_edit:
        with pytest.raises(StaleElementError):
            _ = held.text
        with pytest.raises(StaleElementError):
            _ = view.space_after
        held = next(item.paragraph for item in physical if item.text == "controlled")
    else:
        assert held.text == "controlled" and "controlled" in held.xml
        assert view.space_after == 240 * 635
    assert clean.text == expected
    fresh = clean.paragraphs[0]
    assert fresh.text == expected and expected in fresh.xml
    assert fresh.paragraph_format.space_after == 120 * 635
    if operation == "cell_text":
        with pytest.raises(StaleElementError):
            _ = selected.text
        with pytest.raises(StaleElementError):
            _ = selected_view.space_after
    else:
        assert selected.text == expected and selected_view.space_after == 120 * 635
    if text_edit:
        for old in (run, selected_run):
            with pytest.raises(StaleElementError):
                _ = old.text
        with pytest.raises(StaleElementError):
            _ = font.bold
    else:
        assert run.text == "controlled" and font.bold is True
        assert selected_run.text == "clean"
    held.text = "neighbor changed"
    assert held.text == "neighbor changed"
    fresh.text = "follow up"
    assert clean.text == "follow up"
    assert sorted(item.text for item in document.story_items if item.kind == "paragraph") == ["follow up", "neighbor changed"]
    reopened = Document.from_bytes(document.to_bytes())
    assert sorted(item.text for item in reopened.story_items if item.kind == "paragraph") == ["follow up", "neighbor changed"]
