import io
import zipfile

import pytest


def _replace_document_body(document, body):
    source = io.BytesIO(document.to_bytes())
    result = io.BytesIO()
    with zipfile.ZipFile(source) as source_zip:
        with zipfile.ZipFile(result, "w") as result_zip:
            for info in source_zip.infolist():
                data = source_zip.read(info.filename)
                if info.filename == "word/document.xml":
                    start = data.index(b"<w:body>") + len(b"<w:body>")
                    end = data.index(b"</w:body>")
                    data = data[:start] + body.encode() + data[end:]
                result_zip.writestr(info, data)
    return type(document).from_bytes(result.getvalue())


def _document_xml(document):
    with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as archive:
        return archive.read("word/document.xml")


def _document_with_structure_snapshots(document):
    word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    body = f"""
        <w:p><w:hyperlink r:id="rIdScoped"><w:r><w:t>body link</w:t></w:r></w:hyperlink></w:p>
        <w:p><w:pPr><w:sectPr>
          <w:headerReference w:type="default" r:id="rIdHeader0"/>
          <w:footerReference w:type="default" r:id="rIdFooter0"/>
          <w:type w:val="nextPage"/>
          <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
          <w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="720" w:footer="720" w:gutter="120"/>
          <w:pgNumType w:start="3"/><w:cols w:num="2" w:space="360"/>
          <w:titlePg/>
        </w:sectPr></w:pPr><w:r><w:t>section zero</w:t></w:r></w:p>
        <w:p><w:pPr><w:sectPr>
          <w:type w:val="continuous"/>
          <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
          <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>
        </w:sectPr></w:pPr><w:r><w:t>section one</w:t></w:r></w:p>
        <w:p><w:r><w:t>section two</w:t></w:r></w:p>
        <w:sectPr>
          <w:headerReference w:type="default" r:id="rIdHeader2"/>
          <w:footerReference w:type="default" r:id="rIdFooter2"/>
          <w:type w:val="oddPage"/>
          <w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>
          <w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000"/>
        </w:sectPr>
    """
    header0 = f"""<w:hdr xmlns:w="{word}" xmlns:r="{rel}">
      <w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>header table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
      <w:p><w:hyperlink r:id="rIdScoped"><w:r><w:t>header zero link</w:t></w:r></w:hyperlink></w:p>
    </w:hdr>"""
    header2 = f"""<w:hdr xmlns:w="{word}" xmlns:r="{rel}">
      <w:p><w:hyperlink r:id="rIdScoped"><w:r><w:t>header two link</w:t></w:r></w:hyperlink></w:p>
    </w:hdr>"""
    footer0 = f"<w:ftr xmlns:w=\"{word}\"><w:p><w:r><w:t>footer zero</w:t></w:r></w:p></w:ftr>"
    footer2 = f"<w:ftr xmlns:w=\"{word}\"><w:p><w:r><w:t>footer two</w:t></w:r></w:p></w:ftr>"
    external_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdScoped" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{}" TargetMode="External"/>
      </Relationships>"""
    source = io.BytesIO(document.to_bytes())
    result = io.BytesIO()
    with zipfile.ZipFile(source) as source_zip:
        with zipfile.ZipFile(result, "w") as result_zip:
            for info in source_zip.infolist():
                data = source_zip.read(info.filename)
                if info.filename == "word/document.xml":
                    start = data.index(b"<w:body>") + len(b"<w:body>")
                    end = data.index(b"</w:body>")
                    data = data[:start] + body.encode() + data[end:]
                elif info.filename == "word/_rels/document.xml.rels":
                    additions = """
                      <Relationship Id="rIdScoped" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://body.example/" TargetMode="External"/>
                      <Relationship Id="rIdHeader0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header0.xml"/>
                      <Relationship Id="rIdFooter0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer0.xml"/>
                      <Relationship Id="rIdHeader2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
                      <Relationship Id="rIdFooter2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
                    """.encode()
                    data = data.replace(b"</Relationships>", additions + b"</Relationships>")
                elif info.filename == "[Content_Types].xml":
                    additions = """
                      <Override PartName="/word/header0.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                      <Override PartName="/word/footer0.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                      <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                      <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                    """.encode()
                    data = data.replace(b"</Types>", additions + b"</Types>")
                elif info.filename == "word/styles.xml":
                    style = b"""<w:style w:type="paragraph" w:styleId="SnapshotStyle"><w:name w:val="Snapshot Style"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="7"/><w:qFormat/></w:style>"""
                    data = data.replace(b"</w:styles>", style + b"</w:styles>")
                result_zip.writestr(info, data)
            result_zip.writestr("word/header0.xml", header0)
            result_zip.writestr("word/header2.xml", header2)
            result_zip.writestr("word/footer0.xml", footer0)
            result_zip.writestr("word/footer2.xml", footer2)
            result_zip.writestr(
                "word/_rels/header0.xml.rels",
                external_rels.format("https://header-zero.example/"),
            )
            result_zip.writestr(
                "word/_rels/header2.xml.rels",
                external_rels.format("https://header-two.example/"),
            )
    return type(document).from_bytes(result.getvalue())


def test_stale_paragraph_after_structural_removal_raises_named_error():
    import rdocx

    doc = rdocx.Document()
    for text in ("zero", "one", "two", "three", "four"):
        doc.add_paragraph(text)

    held = doc.paragraphs[3]
    assert doc.remove_content(1) is True

    with pytest.raises(
        rdocx.StaleElementError,
        match=(
            r"paragraph handle was created at document revision 5, but the document "
            r"is now at revision 6"
        ),
    ):
        _ = held.text


def test_lazy_collections_support_index_slice_and_iteration():
    import rdocx

    doc = rdocx.Document()
    for text in ("alpha", "beta", "gamma"):
        doc.add_paragraph(text)

    paragraphs = doc.paragraphs
    assert len(paragraphs) == 3
    assert paragraphs[-1].text == "gamma"
    assert [paragraph.text for paragraph in paragraphs[0:3:2]] == ["alpha", "gamma"]
    assert [paragraph.text for paragraph in paragraphs] == ["alpha", "beta", "gamma"]

    paragraph = doc.paragraphs[0]
    paragraph.add_run(" one")
    paragraph = doc.paragraphs[0]
    paragraph.add_run(" two")
    runs = doc.paragraphs[0].runs
    assert len(runs) == 3
    assert runs[-1].text == " two"
    assert [run.text for run in runs[0:3:2]] == ["alpha", " two"]
    assert [run.text for run in runs] == ["alpha", " one", " two"]


def test_failed_removal_does_not_stale_live_handles():
    import rdocx

    doc = rdocx.Document()
    doc.add_paragraph("live")
    held = doc.paragraphs[0]

    assert doc.remove_content(99) is False
    assert held.text == "live"


def test_core_text_mutations_survive_bytes_round_trip():
    import rdocx

    doc = rdocx.Document()
    paragraph = doc.add_paragraph("Hello")
    run = paragraph.add_run(" world")
    reopened = rdocx.Document.from_bytes(doc.to_bytes())

    assert reopened.paragraphs[0].text == "Hello world"
    reopened.paragraphs[0].runs[1].text = " Rust"
    reopened_again = rdocx.Document.from_bytes(reopened.to_bytes())
    assert reopened_again.paragraphs[0].text == "Hello Rust"


def test_constructor_accepts_an_optional_input_path(tmp_path):
    import rdocx

    assert len(rdocx.Document().paragraphs) == 0

    path = tmp_path / "input.docx"
    source = rdocx.Document()
    source.add_paragraph("opened by constructor")
    source.save(path)

    reopened = rdocx.Document(path)
    assert reopened.paragraphs[0].text == "opened by constructor"


def test_priority_word_operations_return_typed_snapshots_and_remain_atomic():
    import rdocx

    position = rdocx.RunPosition(body_index=0, run_index=0)
    range_ = rdocx.RunRange(
        start=position,
        end=rdocx.RunPosition(body_index=0, run_index=1),
    )
    with pytest.raises(AttributeError):
        position.run_index = 2

    commented = rdocx.Document()
    commented.add_paragraph("review this")
    held_before_comment = commented.paragraphs[0]
    comment_id = commented.add_comment(
        range_, author="Ada", text="Please revise", initials="AL"
    )
    with pytest.raises(
        rdocx.StaleElementError, match=r"revision 1, but the document is now at revision 2"
    ):
        _ = held_before_comment.text
    held_before_reply = commented.paragraphs[0]
    reply_id = commented.reply_to(comment_id, author="Grace", text="Done")
    with pytest.raises(
        rdocx.StaleElementError, match=r"revision 2, but the document is now at revision 3"
    ):
        _ = held_before_reply.text
    held_before_resolve = commented.paragraphs[0]
    assert commented.resolve_comment(comment_id, resolved=True) is True
    with pytest.raises(
        rdocx.StaleElementError, match=r"revision 3, but the document is now at revision 4"
    ):
        _ = held_before_resolve.text
    assert commented.comments == (
        rdocx.Comment(
            id=comment_id,
            author="Ada",
            initials="AL",
            date=None,
            text="Please revise",
            parent_id=None,
            resolved=True,
        ),
        rdocx.Comment(
            id=reply_id,
            author="Grace",
            initials=None,
            date=None,
            text="Done",
            parent_id=comment_id,
            resolved=False,
        ),
    )
    reopened_comments = rdocx.Document.from_bytes(commented.to_bytes())
    assert reopened_comments.comments == commented.comments

    live = reopened_comments.paragraphs[0]
    before_failure = reopened_comments.to_bytes()
    invalid = rdocx.RunRange(
        start=rdocx.RunPosition(body_index=99, run_index=0),
        end=rdocx.RunPosition(body_index=99, run_index=1),
    )
    with pytest.raises(rdocx.RdocxError):
        reopened_comments.add_comment(invalid, author="Ada", text="invalid")
    assert live.text == "review this"
    assert reopened_comments.to_bytes() == before_failure
    held_before_remove = reopened_comments.paragraphs[0]
    assert reopened_comments.remove_comment(reply_id) is True
    with pytest.raises(
        rdocx.StaleElementError, match=r"revision 0, but the document is now at revision 1"
    ):
        _ = held_before_remove.text
    live_after_noop = reopened_comments.paragraphs[0]
    assert reopened_comments.remove_comment(reply_id) is False
    assert live_after_noop.text == "review this"

    original = rdocx.Document()
    original.add_paragraph("before")
    edited = rdocx.Document()
    edited.add_paragraph("after")
    held_before_compare = original.paragraphs[0]
    diagnostics = original.compare(
        edited, author="Ada", timestamp="2026-09-14T09:00:00Z"
    )
    assert isinstance(diagnostics, tuple)
    assert all(isinstance(item, rdocx.ComparisonDiagnostic) for item in diagnostics)
    with pytest.raises(
        rdocx.StaleElementError, match=r"revision 1, but the document is now at revision 2"
    ):
        _ = held_before_compare.text
    reopened_redline = rdocx.Document.from_bytes(original.to_bytes())
    assert b"before" in _document_xml(reopened_redline)
    assert b"after" in _document_xml(reopened_redline)
    with pytest.raises(rdocx.RdocxError, match="existing modeled revisions"):
        reopened_redline.compare(
            edited, author="Ada", timestamp="2026-09-14T09:01:00Z"
        )

    failed_compare = rdocx.Document()
    failed_compare.add_paragraph("stable")
    live_after_failure = failed_compare.paragraphs[0]
    before_failure = failed_compare.to_bytes()
    with pytest.raises(rdocx.RdocxError):
        failed_compare.compare(edited, author="Ada", timestamp="not-a-timestamp")
    assert live_after_failure.text == "stable"
    assert failed_compare.to_bytes() == before_failure

    unchanged = rdocx.Document()
    unchanged.add_paragraph("same")
    identical = rdocx.Document.from_bytes(unchanged.to_bytes())
    held_before_identical_compare = unchanged.paragraphs[0]
    assert unchanged.compare(
        identical, author="Ada", timestamp="2026-09-14T09:02:00Z"
    ) == ()
    assert held_before_identical_compare.text == "same"

    laid_out = rdocx.Document()
    laid_out.add_paragraph("A deterministic layout fragment")
    fragments = laid_out.layout()
    assert fragments
    assert fragments[0].body_index == 0
    assert fragments[0].physical_page == 1
    assert fragments[0].displayed_page == 1
    assert isinstance(fragments[0].bounds, rdocx.BoundingBox)
    assert fragments[0].bounds.width > 0.0
    page = laid_out.layout_page(0)
    assert page == rdocx.LayoutPage(
        page_number=1,
        displayed_page_number=1,
        width=page.width,
        height=page.height,
    )
    assert laid_out.layout_page(1) is None

    toc_source = rdocx.Document()
    toc_source.add_paragraph("placeholder")
    toc = _replace_document_body(
        toc_source,
        """
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \\o "1-1"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>
        """,
    )
    held_before_toc = toc.paragraphs[-1]
    report = toc.rebuild_toc()
    assert report == rdocx.TocRebuildReport(
        entry_count=1, bookmark_count=1, diagnostic_count=0
    )
    with pytest.raises(
        rdocx.StaleElementError, match=r"revision 0, but the document is now at revision 1"
    ):
        _ = held_before_toc.text
    reopened_toc = rdocx.Document.from_bytes(toc.to_bytes())
    assert "Heading" in [paragraph.text for paragraph in reopened_toc.paragraphs]

    no_toc = rdocx.Document()
    no_toc.add_paragraph("no table of contents")
    live_after_noop = no_toc.paragraphs[0]
    assert no_toc.rebuild_toc() == rdocx.TocRebuildReport(
        entry_count=0, bookmark_count=0, diagnostic_count=0
    )
    assert live_after_noop.text == "no table of contents"


def test_update_page_fields_writes_layout_page_numbers():
    import rdocx

    source = rdocx.Document()
    source.add_paragraph("placeholder")
    document = _replace_document_body(
        source,
        """
        <w:p><w:r><w:t>Page one.</w:t></w:r></w:p>
        <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>
        """,
    )
    held_before_update = document.paragraphs[1]
    assert document.update_page_fields() == 2
    with pytest.raises(
        rdocx.StaleElementError, match=r"revision 0, but the document is now at revision 1"
    ):
        _ = held_before_update.text
    xml = _document_xml(document)
    assert b"<w:t>7</w:t>" not in xml
    assert b"<w:t>9</w:t>" not in xml
    assert xml.count(b"<w:t>2</w:t>") == 2

    no_page_fields = rdocx.Document()
    no_page_fields.add_paragraph("no page fields")
    live_after_noop = no_page_fields.paragraphs[0]
    assert no_page_fields.update_page_fields() == 0
    assert live_after_noop.text == "no page fields"


def test_word_structure_snapshots_preserve_order_ownership_and_types():
    import rdocx

    document = _document_with_structure_snapshots(rdocx.Document())
    reopened = rdocx.Document.from_bytes(document.to_bytes())

    assert len(reopened.sections) == 3
    assert reopened.sections[0] == rdocx.Section(
        ordinal=0,
        is_final=False,
        orientation=None,
        page_width=7_772_400,
        page_height=10_058_400,
        margin_top=914_400,
        margin_right=685_800,
        margin_bottom=914_400,
        margin_left=685_800,
        gutter=76_200,
        column_count=2,
        column_spacing=228_600,
        page_number_start=3,
        header_distance=457_200,
        footer_distance=457_200,
        different_first_page=True,
        break_type="nextPage",
    )
    assert [section.ordinal for section in reopened.sections] == [0, 1, 2]
    assert [section.is_final for section in reopened.sections] == [False, False, True]
    with pytest.raises(AttributeError):
        reopened.sections[0].orientation = "landscape"

    assert reopened.styles[-1] == rdocx.Style(
        style_id="SnapshotStyle",
        name="Snapshot Style",
        based_on="Normal",
        style_type="paragraph",
        linked_style=None,
        next_style="Normal",
        priority=7,
        auto_redefine=None,
        hidden=None,
        semi_hidden=None,
        unhide_when_used=None,
        quick_format=True,
        locked=None,
        is_default=False,
    )
    assert [style.style_id for style in reopened.styles][-1] == "SnapshotStyle"

    variants = reopened.header_footer_variants
    assert len(variants) == 18
    assert variants[0] == rdocx.HeaderFooterVariant(
        section_index=0,
        kind="header",
        variant="default",
        story=rdocx.Story(
            kind="header", part_name="/word/header0.xml", owner_index=0
        ),
        source_section=0,
        inherited=False,
    )
    assert variants[6].story == variants[0].story
    assert variants[6].source_section == 0
    assert variants[6].inherited is True
    assert variants[12].story == rdocx.Story(
        kind="header", part_name="/word/header2.xml", owner_index=0
    )
    assert variants[12].source_section == 2
    assert variants[12].inherited is False
    assert variants[1].story is None

    header_items = [
        item for item in reopened.story_items if item.story == variants[0].story
    ]
    assert [(item.kind, item.index_path, item.text) for item in header_items] == [
        ("table", (0,), None),
        ("paragraph", (1,), "header zero link"),
    ]
    assert all(isinstance(item, rdocx.StoryItem) for item in reopened.story_items)
    assert all(isinstance(story, rdocx.Story) for story in reopened.stories)

    assert [link.relationship_id for link in reopened.hyperlinks] == [
        "rIdScoped",
        "rIdScoped",
        "rIdScoped",
    ]
    assert [link.url for link in reopened.hyperlinks] == [
        "https://body.example/",
        "https://header-zero.example/",
        "https://header-two.example/",
    ]
    assert [link.text for link in reopened.hyperlinks] == [
        "body link",
        "header zero link",
        "header two link",
    ]
    assert [link.index_path for link in reopened.hyperlinks] == [(0,), (1,), (0,)]

    nested = _replace_document_body(
        rdocx.Document(),
        """
        <w:sdt><w:sdtContent><w:sdt><w:sdtContent><w:p>
          <w:hyperlink w:anchor="nested-target"><w:r><w:t>nested link</w:t></w:r></w:hyperlink>
        </w:p></w:sdtContent></w:sdt><w:p>
          <w:hyperlink w:anchor="after-target"><w:r><w:t>after link</w:t></w:r></w:hyperlink>
        </w:p></w:sdtContent></w:sdt>
        <w:sectPr/>
        """,
    )
    assert [
        (link.text, link.anchor, link.index_path) for link in nested.hyperlinks
    ] == [
        ("nested link", "nested-target", (1,)),
        ("after link", "after-target", (0,)),
    ]

    captured_items = reopened.story_items
    captured_links = reopened.hyperlinks
    reopened.add_paragraph("later body content")
    assert captured_items[-1].text != "later body content"
    assert len(reopened.story_items) == len(captured_items) + 1
    assert reopened.hyperlinks == captured_links
