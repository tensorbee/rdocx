from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

import pytest


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
    assert font.color == RGBColor(0x12, 0x34, 0x56)
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

    document.add_paragraph("invalidates both")

    with pytest.raises(StaleElementError):
        _ = font.bold
    with pytest.raises(StaleElementError):
        _ = paragraph_format.space_after


def test_nested_paragraph_stale_error_names_complete_recovery_path():
    from rdocx import Document, StaleElementError

    document = Document()
    cell = document.add_table(rows=1, cols=1).rows[0].cells[0]
    nested = cell.add_paragraph("held")

    document.add_paragraph("invalidates nested paragraph")

    with pytest.raises(
        StaleElementError,
        match=(
            r"doc\.tables\[0\]\.rows\[0\]\.cells\[0\]"
            r"\.paragraphs\[1\]"
        ),
    ):
        _ = nested.text


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

    assert reopened.paragraphs[0].runs[0].font.color is None


def test_paragraph_style_and_numbering_round_trip_and_clear():
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


def test_run_style_and_highlight_round_trip_and_clear():
    from rdocx import Document

    document = Document()
    run = document.add_paragraph("").add_run("marked")
    assert run.style is None
    assert run.font.highlight is None
    run.style = "Strong"
    run.font.highlight = "FFFF00"

    reopened = Document.from_bytes(document.to_bytes()).paragraphs[0].runs[0]
    assert reopened.style == "Strong"
    assert reopened.font.highlight == "FFFF00"

    with pytest.raises(ValueError, match="hexadecimal"):
        run.font.highlight = "yellow"
    run.style = None
    run.font.highlight = None
    reopened = Document.from_bytes(document.to_bytes()).paragraphs[0].runs[0]
    assert reopened.style is None
    assert reopened.font.highlight is None


def test_word_highlight_keyword_reads_and_clears_through_font_highlight():
    from rdocx import Document, RGBColor

    document = Document()
    document.add_paragraph("").add_run("marked").font.color = RGBColor(0x12, 0x34, 0x56)
    keyword = _replace_document_xml(
        document.to_bytes(),
        b'<w:color w:val="123456"/>',
        b'<w:color w:val="123456"/><w:highlight w:val="yellow"/>',
    )
    reopened = Document.from_bytes(keyword)
    font = reopened.paragraphs[0].runs[0].font
    assert font.highlight == "yellow"

    font.highlight = None
    assert Document.from_bytes(reopened.to_bytes()).paragraphs[0].runs[0].font.highlight is None
