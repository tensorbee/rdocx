import io
import posixpath
import re
import struct
import time
import zipfile
import zlib

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


def _replace_settings_xml(document, settings):
    source = io.BytesIO(document.to_bytes())
    result = io.BytesIO()
    with zipfile.ZipFile(source) as source_zip:
        with zipfile.ZipFile(result, "w") as result_zip:
            for info in source_zip.infolist():
                data = source_zip.read(info.filename)
                if info.filename == "word/settings.xml":
                    data = settings.encode()
                result_zip.writestr(info, data)
    return type(document).from_bytes(result.getvalue())


def _tracked_document():
    import rdocx

    original = rdocx.Document()
    original.add_paragraph("alpha")
    edited = rdocx.Document()
    edited.add_paragraph("alpha beta")
    document = rdocx.Document.from_bytes(original.to_bytes())
    document.compare(edited, "Ada", "2026-01-02T03:04:05Z")
    return document


def test_revisions_are_snapshots_and_resolution_reports_counts():
    import rdocx

    document = _tracked_document()
    revisions = document.revisions
    assert revisions
    assert {revision.author for revision in revisions} == {"Ada"}
    kinds = {revision.kind for revision in revisions}
    assert "insertion" in kinds
    assert kinds <= {
        "insertion",
        "deletion",
        "move_from",
        "move_to",
        "run_property_change",
        "paragraph_property_change",
        "table_property_change",
        "section_property_change",
    }
    assert {revision.timestamp for revision in revisions} == {"2026-01-02T03:04:05Z"}

    held = document.paragraphs[0]
    assert document.reject_revisions_by_author("Grace") == 0
    held.text
    with pytest.raises(rdocx.RdocxError):
        document.accept_revisions_in_date_range(start="yesterday", end="2026-01-03T00:00:00Z")
    held.text

    assert (
        document.accept_revisions_in_date_range(
            start="2026-01-01T00:00:00Z", end="2026-01-03T00:00:00Z"
        )
        > 0
    )
    with pytest.raises(rdocx.StaleElementError):
        held.text
    assert document.revisions == ()
    assert [paragraph.text for paragraph in document.paragraphs] == ["alpha beta"]


def test_reject_revision_id_then_reject_all_restores_the_original():
    document = _tracked_document()
    first = document.revisions[0]
    assert document.reject_revision_id(first.id) > 0
    assert all(revision.id != first.id for revision in document.revisions)
    document.reject_all()
    assert document.revisions == ()
    assert [paragraph.text for paragraph in document.paragraphs] == ["alpha"]


def test_counted_replacement_spans_runs_and_a_bad_regex_changes_nothing():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("Dear {{na").add_run("me}}, hello")
    held = document.paragraphs[0]
    assert document.try_replace_text("{{missing}}", "x") == 0
    held.text
    assert document.try_replace_text("{{name}}", "Ada") == 1
    with pytest.raises(rdocx.StaleElementError):
        held.text
    assert document.paragraphs[0].text == "Dear Ada, hello"

    before = document.to_bytes()
    with pytest.raises(rdocx.RdocxError, match="invalid regex"):
        document.replace_all_regex([("hello", "bye"), ("(", "x")])
    assert document.to_bytes() == before
    assert document.replace_all_regex([("h(el)lo", "bye")]) == 1
    assert document.paragraphs[0].text == "Dear Ada, bye"


def test_update_fields_takes_a_keyword_context_and_counts_updates():
    import datetime

    import rdocx

    document = _replace_document_body(
        rdocx.Document(),
        '<w:p><w:fldSimple w:instr=" FILENAME "><w:r><w:t>old.docx</w:t></w:r></w:fldSimple></w:p>'
        '<w:p><w:fldSimple w:instr=" MERGEFIELD Name "><w:r><w:t>Name</w:t></w:r></w:fldSimple></w:p>'
        '<w:p><w:fldSimple w:instr=" DATE \\@ &quot;yyyy-MM-dd&quot; "><w:r><w:t>2000-01-01</w:t></w:r></w:fldSimple></w:p>',
    )
    held = document.paragraphs[0]
    count = document.update_fields(
        now=datetime.datetime(2026, 9, 15, 10, 30),
        file_name="report.docx",
        merge_fields={"Name": "Ada"},
    )
    assert count == 3
    with pytest.raises(rdocx.StaleElementError):
        held.text
    xml = _document_xml(document)
    for cached in (b"report.docx", b"Ada", b"2026-09-15"):
        assert cached in xml
    assert b"old.docx" not in xml and b"2000-01-01" not in xml


def test_python_story_revision_field_and_xml_operations_are_typed_and_atomic():
    import rdocx

    document = _tracked_document()
    revisions = document.revisions
    assert revisions
    assert {revision.author for revision in revisions} == {"Ada"}
    assert document.accept_all() == len(revisions)

    document.set_header("Draft")
    document.set_footer("Page footer")
    header = next(story for story in document.stories if story.kind == "header")
    document.add_hyperlink_to_story(header, "home", "https://example.com/")
    header_item = next(
        item
        for item in document.story_items
        if item.story == header and item.kind == "paragraph"
    )
    assert isinstance(header_item.xml, bytes)
    document.set_story_text(header_item, "Final")

    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert "Final" in [item.text for item in reopened.story_items]
    assert [(link.text, link.url) for link in reopened.hyperlinks] == [
        ("home", "https://example.com/")
    ]

    with pytest.raises(rdocx.StaleElementError, match="story item handle"):
        document.set_story_text(header_item, "stale")

    linked = rdocx.Document()
    paragraph = linked.add_paragraph("See ")
    paragraph.add_hyperlink("old link", "https://example.com/old")
    linked_item = next(
        item
        for item in linked.story_items
        if item.story.kind == "body" and item.kind == "paragraph"
    )
    linked.set_story_text(linked_item, "New text without a link")
    assert linked.hyperlinks == ()
    assert b"<w:hyperlink" not in _document_xml(linked)

    commented = rdocx.Document()
    commented.add_paragraph("commented")
    commented.add_paragraph("destination")
    commented.add_comment(
        rdocx.RunRange(
            start=rdocx.RunPosition(body_index=0, run_index=0),
            end=rdocx.RunPosition(body_index=0, run_index=1),
        ),
        author="Ada",
        text="review",
    )
    commented.clone_content(commented.paragraphs[0], 0)
    cloned_xml = _document_xml(commented)
    assert cloned_xml.count(b"commentRangeStart") == 1
    assert cloned_xml.count(b"commentRangeEnd") == 1
    assert cloned_xml.count(b"commentReference") == 1

    lookup = _replace_document_body(
        rdocx.Document(),
        '<w:sdt><w:sdtContent><w:p><w:r><w:t>Background</w:t></w:r></w:p></w:sdtContent></w:sdt>'
        '<w:p><w:r><w:t>Background</w:t></w:r></w:p>',
    )
    assert lookup.find_content_index("Background") == 1
    assert lookup.find_content_indices("Background") == (1, 0)


def test_paragraph_views_include_accepted_nested_runs_in_source_order():
    import rdocx

    document = _replace_document_body(
        rdocx.Document(),
        '<w:p><w:r><w:t xml:space="preserve">Start </w:t></w:r>'
        '<w:ins w:id="1" w:author="Ada"><w:r><w:t xml:space="preserve">inserted </w:t></w:r></w:ins>'
        '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>'
        '<w:r><w:t xml:space="preserve">controlled </w:t></w:r>'
        '<w:ins w:id="3" w:author="Ada"><w:r><w:t xml:space="preserve">nested </w:t></w:r></w:ins>'
        '</w:sdtContent></w:sdt>'
        '<w:del w:id="2" w:author="Ada"><w:r><w:delText>deleted </w:delText></w:r></w:del>'
        '<w:moveFrom w:id="4" w:author="Ada"><w:r><w:t>moved away </w:t></w:r></w:moveFrom>'
        '<w:r><w:t>end</w:t></w:r></w:p>',
    )
    item = next(
        item
        for item in document.story_items
        if item.story.kind == "body" and item.kind == "paragraph"
    )
    paragraph = document.paragraphs[0]

    assert item.text == "Start inserted controlled nested end"
    assert paragraph.text == item.text
    assert [run.text for run in paragraph.runs] == [
        "Start ",
        "inserted ",
        "controlled ",
        "nested ",
        "end",
    ]

    inserted = paragraph.runs[1]
    controlled = paragraph.runs[2]
    nested = paragraph.runs[3]
    inserted.text = "changed "
    controlled.font.bold = True
    nested.text = "composed "
    assert inserted.text == "changed "
    assert controlled.font.bold is True
    assert nested.text == "composed "

    boundary = document.split_run(body_index=0, run_index=1, character_offset=4)
    assert boundary == 2
    with pytest.raises(rdocx.StaleElementError):
        _ = inserted.text
    assert [run.text for run in document.paragraphs[0].runs] == [
        "Start ",
        "chan",
        "ged ",
        "controlled ",
        "composed ",
        "end",
    ]

    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert reopened.paragraphs[0].text == "Start changed controlled composed end"
    assert reopened.paragraphs[0].runs[3].font.bold is True
    assert reopened.paragraphs[0].runs[4].text == "composed "


def test_python_story_inventory_scales_linearly():
    import rdocx

    def timed_snapshot(paragraph_count):
        body = []
        for index in range(paragraph_count):
            body.append(
                f'<w:p><w:r><w:t>paragraph {index}</w:t></w:r>'
                f'<w:hyperlink w:anchor="target{index}"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>'
            )
            if index % 10 == 0:
                cells = "".join(
                    f'<w:tc><w:p><w:r><w:t>cell {index} {cell}</w:t></w:r></w:p></w:tc>'
                    for cell in range(4)
                )
                body.append(f"<w:tbl><w:tr>{cells}</w:tr></w:tbl>")
        document = _replace_document_body(rdocx.Document(), "".join(body))
        started = time.perf_counter()
        items = document.story_items
        links = document.hyperlinks
        elapsed = time.perf_counter() - started
        return elapsed, len(items), len(links)

    small_elapsed, small_items, small_links = timed_snapshot(50)
    large_elapsed, large_items, large_links = timed_snapshot(100)

    assert large_items == small_items * 2
    assert large_links == small_links * 2
    assert large_elapsed <= small_elapsed * 3.0 + 0.02, (
        f"doubling the story inventory took {large_elapsed / small_elapsed:.2f} times longer"
    )


def test_update_fields_on_open_sets_clears_and_removes_the_setting():
    import rdocx

    document = rdocx.Document()
    assert document.update_fields_on_open is None
    for value in (True, False, None):
        document.update_fields_on_open = value
        assert document.update_fields_on_open is value
        document = rdocx.Document.from_bytes(document.to_bytes())
        assert document.update_fields_on_open is value

    word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    for settings in (
        f'<w:settings xmlns:w="{word}"><w:updateFields/><w:updateFields w:val="false"/></w:settings>',
        f'<w:settings xmlns:w="{word}"><w:updateFields w:val="invalid"/></w:settings>',
    ):
        document = _replace_settings_xml(rdocx.Document(), settings)
        before = document.to_bytes()
        assert document.update_fields_on_open is None
        with pytest.raises(rdocx.XmlError, match="ambiguous or malformed"):
            document.update_fields_on_open = True
        assert document.to_bytes() == before


def _one_pixel_png():
    def chunk(kind, data):
        crc = struct.pack(">I", zlib.crc32(kind + data))
        return struct.pack(">I", len(data)) + kind + data + crc

    header = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(b"\x00\xff\xff\xff"))
        + chunk(b"IEND", b"")
    )


def _relationship_target(document, part_name, relationship_id):
    owner = part_name.lstrip("/")
    directory, filename = posixpath.split(owner)
    relationships = posixpath.join(directory, "_rels", f"{filename}.rels")
    with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as archive:
        xml = archive.read(relationships)
        match = re.search(
            rb'<Relationship Id="'
            + re.escape(relationship_id.encode())
            + rb'" Type="[^"]+" Target="([^"]+)"',
            xml,
        )
        assert match is not None
        target = match.group(1).decode()
        return posixpath.normpath(posixpath.join(directory, target)).lstrip("/")


def test_replace_image_preserves_drawings_and_story_relationship_ownership():
    docx = pytest.importorskip("docx")
    import rdocx

    source = docx.Document()
    source.add_picture(io.BytesIO(_one_pixel_png()))
    source.sections[0].header.paragraphs[0].add_run().add_picture(
        io.BytesIO(_one_pixel_png())
    )
    source.sections[0].footer.paragraphs[0].add_run().add_picture(
        io.BytesIO(_one_pixel_png())
    )
    buffer = io.BytesIO()
    source.save(buffer)
    document = rdocx.Document.from_bytes(buffer.getvalue())
    held = document.paragraphs[0]
    stories = {story.kind: story for story in document.stories}
    relationship_id = re.search(
        rb'r:embed="([^"]+)"', _document_xml(document)
    ).group(1).decode()
    with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as archive:
        drawing_xml = {
            story.part_name: archive.read(story.part_name.lstrip("/"))
            for story in (stories["body"], stories["header"], stories["footer"])
        }
    with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as archive:
        header_id = re.search(
            rb'r:embed="([^"]+)"', archive.read(stories["header"].part_name.lstrip("/"))
        ).group(1).decode()
        footer_id = re.search(
            rb'r:embed="([^"]+)"', archive.read(stories["footer"].part_name.lstrip("/"))
        ).group(1).decode()

    jpeg = b"\xff\xd8\xff\xd9"
    document.replace_image(relationship_id, jpeg)
    document.replace_image_for_story(stories["header"], header_id, jpeg)
    document.replace_image_for_story(stories["footer"], footer_id, jpeg)
    assert document.image_data(relationship_id) == jpeg
    assert held.text == ""
    assert rdocx.Document.from_bytes(document.to_bytes()).image_data(relationship_id) == jpeg
    with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as archive:
        assert b'ContentType="image/jpeg"' in archive.read("[Content_Types].xml")
        for story, rel_id in [
            (stories["body"], relationship_id),
            (stories["header"], header_id),
            (stories["footer"], footer_id),
        ]:
            target = _relationship_target(document, story.part_name, rel_id)
            assert archive.read(target) == jpeg
            assert archive.read(story.part_name.lstrip("/")) == drawing_xml[story.part_name]

    with pytest.raises(rdocx.RdocxError):
        document.replace_image("rIdMissing", jpeg)
    assert document.image_data("rIdMissing") is None


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


def test_python_indexed_content_mutation_is_counted_and_atomic():
    import rdocx

    document = rdocx.Document()
    first = document.add_paragraph("alpha {{TOKEN")
    first.add_run("}}")
    document.add_table(1, 1)
    document.add_paragraph("omega")

    first = document.paragraphs[0]
    table = document.tables[0]
    assert document.find_content_index(first) == 0
    assert document.find_content_index(table) == 1

    held = document.paragraphs[0]
    inserted = document.insert_paragraph(1, "middle")
    assert inserted.text == "middle"
    with pytest.raises(rdocx.StaleElementError):
        _ = held.text

    fragment = document.pop_content(document.find_content_index(inserted))
    assert fragment.kind == "paragraph"
    with pytest.raises(rdocx.StaleElementError):
        _ = inserted.text

    document.insert_content(2, fragment)
    middle = document.paragraphs[1]
    document.clone_content(middle, 0)
    with pytest.raises(rdocx.StaleElementError):
        _ = middle.text

    table = document.tables[0]
    document.move_content(table, 5)
    with pytest.raises(rdocx.StaleElementError):
        _ = table.style

    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert [paragraph.text for paragraph in reopened.paragraphs] == [
        "middle",
        "alpha {{TOKEN}}",
        "middle",
        "omega",
    ]
    assert reopened.find_content_index(reopened.tables[0]) == 4

    held = reopened.paragraphs[1]
    assert reopened.try_replace_text("{{TOKEN}}", "done") == 1
    with pytest.raises(rdocx.StaleElementError):
        _ = held.text

    held = reopened.paragraphs[1]
    before = reopened.to_bytes()
    with pytest.raises(rdocx.RdocxError):
        reopened.replace_all_regex([(r"[", "broken")])
    assert reopened.to_bytes() == before
    assert held.text == "alpha done"

    assert reopened.replace_all_regex([(r"middle", "center")]) == 2
    with pytest.raises(rdocx.StaleElementError):
        _ = held.text
    held = reopened.paragraphs[1]
    assert reopened.try_replace_text("not present", "ignored") == 0
    assert held.text == "alpha done"

    before = reopened.to_bytes()
    held = reopened.paragraphs[0]
    with pytest.raises(IndexError):
        reopened.pop_content(99)
    with pytest.raises(IndexError):
        reopened.move_content(reopened.tables[0], 99)
    foreign_document = rdocx.Document()
    foreign = foreign_document.add_paragraph("foreign")
    with pytest.raises(ValueError, match="different document"):
        reopened.clone_content(foreign, 0)
    assert reopened.to_bytes() == before
    assert held.text == "center"

    fragment_document = rdocx.Document()
    fragment_document.add_paragraph("reusable")
    reusable = fragment_document.pop_content(0)
    assert reusable.kind == "paragraph"
    with pytest.raises(TypeError):
        rdocx.ContentFragment()
    fragment_document.insert_content(0, reusable)
    fragment_document.insert_content(1, reusable)
    assert [paragraph.text for paragraph in fragment_document.paragraphs] == [
        "reusable",
        "reusable",
    ]

    nested = _replace_document_body(
        rdocx.Document(),
        "<w:sdt><w:sdtContent><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:sdtContent></w:sdt>",
    )
    nested_paragraph = nested.paragraphs[0]
    with pytest.raises(ValueError, match="not a direct body child"):
        nested.find_content_index(nested_paragraph)
    assert nested_paragraph.text == "nested"


def test_clone_content_scales_linearly_and_names_invalid_arguments():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("source")
    document.add_paragraph("destination")

    with pytest.raises(
        TypeError, match="source must be a Paragraph or Table handle"
    ):
        document.clone_content(0, 1)
    with pytest.raises(
        TypeError, match="destination must be a direct body index integer"
    ):
        document.clone_content(document.paragraphs[0], document.paragraphs[1])

    def clone_elapsed(paragraph_count, destination):
        body = "".join(
            f"<w:p><w:r><w:t>Paragraph {index} with realistic text.</w:t></w:r></w:p>"
            for index in range(paragraph_count)
        )
        candidate = _replace_document_body(rdocx.Document(), body)
        source = candidate.paragraphs[paragraph_count // 2]
        expected = source.text
        started = time.perf_counter()
        destination_index = (
            paragraph_count if destination == "end" else paragraph_count // 2 + 1
        )
        candidate.clone_content(source, destination_index)
        elapsed = time.perf_counter() - started
        assert candidate.paragraphs[destination_index].text == expected
        return elapsed

    for destination in ("middle", "end"):
        small = clone_elapsed(50, destination)
        large = clone_elapsed(100, destination)
        assert large <= small * 4.0 + 0.05, (destination, small, large)


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
        range_,
        author="Ada",
        text="Please revise",
        initials="AL",
        date="2026-09-16T10:15:30Z",
    )
    with pytest.raises(
        rdocx.StaleElementError, match=r"revision 1, but the document is now at revision 2"
    ):
        _ = held_before_comment.text
    held_before_reply = commented.paragraphs[0]
    reply_id = commented.reply_to(
        comment_id,
        author="Grace",
        text="Done",
        date="2026-09-16T11:00:00+01:00",
    )
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
            date="2026-09-16T10:15:30Z",
            text="Please revise",
            parent_id=None,
            resolved=True,
        ),
        rdocx.Comment(
            id=reply_id,
            author="Grace",
            initials=None,
            date="2026-09-16T11:00:00+01:00",
            text="Done",
            parent_id=comment_id,
            resolved=False,
        ),
    )
    reopened_comments = rdocx.Document.from_bytes(commented.to_bytes())
    assert reopened_comments.comments == commented.comments

    before_invalid_date = reopened_comments.to_bytes()
    with pytest.raises(rdocx.RdocxError, match="invalid RFC 3339 comment timestamp"):
        reopened_comments.add_comment(
            range_, author="Ada", text="invalid", date="2026-02-30"
        )
    assert reopened_comments.to_bytes() == before_invalid_date

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
        entry_count=1, bookmark_count=1, diagnostics=()
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
        entry_count=0, bookmark_count=0, diagnostics=()
    )
    assert live_after_noop.text == "no table of contents"


def test_word_default_toc_switch_rebuilds_and_reports_ordered_diagnostics():
    import rdocx

    source = rdocx.Document()
    source.add_paragraph("placeholder")
    document = _replace_document_body(
        source,
        """
        <w:p><w:fldSimple w:instr="TOC"><w:r><w:t>simple cache</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \\x</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:t>complex cache</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        """,
    )

    report = document.rebuild_toc()
    assert report.diagnostics == (
        "simple table of contents fields are not rebuilt, stored display retained",
        "field TOC uses unsupported switch \\x, stored display retained",
    )
    assert report.diagnostic_count == len(report.diagnostics) == 2
    with pytest.raises(AttributeError):
        report.diagnostics = ()


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


def test_update_layout_backed_fields_returns_owned_report():
    import rdocx

    source = rdocx.Document()
    source.add_paragraph("placeholder")
    document = _replace_document_body(
        source,
        """
        <w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>stale page</w:t></w:r></w:fldSimple><w:fldSimple w:instr="NUMPAGES"><w:r><w:t>stale count</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:fldSimple w:instr="PAGEREF destination"><w:r><w:t>stale target</w:t></w:r></w:fldSimple></w:p>
        <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:bookmarkStart w:id="7" w:name="destination"/><w:r><w:t>Target</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>
        """,
    )

    report = document.update_layout_backed_fields()
    assert report == rdocx.LayoutBackedFieldUpdateReport(
        page_fields=1,
        num_pages_fields=1,
        page_reference_fields=1,
        diagnostics=report.diagnostics,
    )
    assert report.updated_count == 3
    assert report.diagnostic_count == len(report.diagnostics)
    with pytest.raises(AttributeError):
        report.page_fields = 0
    xml = _document_xml(document)
    assert b"stale page" not in xml
    assert b"stale count" not in xml
    assert b"stale target" not in xml


def test_word_structure_snapshots_preserve_order_ownership_and_types():
    import rdocx

    document = _document_with_structure_snapshots(rdocx.Document())
    reopened = rdocx.Document.from_bytes(document.to_bytes())

    assert len(reopened.sections) == 3
    # The untouched body is saved as written, so the explicit portrait
    # orientation survives the round trip.
    assert reopened.sections[0] == rdocx.Section(
        ordinal=0,
        is_final=False,
        orientation="portrait",
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
    assert all(item.direct_body_index is None for item in header_items)
    body_items = [item for item in reopened.story_items if item.story.kind == "body"]
    assert [item.direct_body_index for item in body_items] == [0, 1, 2, 3, None]

    owned = _replace_document_body(
        rdocx.Document(),
        """
        <w:p><w:r><w:t>first</w:t></w:r>
          <w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>
          <w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:inline><wp:docPr id="41" name="item"/></wp:inline></w:drawing></w:r>
          <w:sdt><w:sdtContent><w:r><w:t>nested control</w:t></w:r></w:sdtContent></w:sdt>
        </w:p>
        <w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
        <w:sdt><w:sdtContent><w:p><w:r><w:t>body control</w:t></w:r></w:p></w:sdtContent></w:sdt>
        <w:p><w:r><w:t>second</w:t></w:r></w:p>
        <w:sectPr/>
        """,
    )
    owned_body_items = [
        item for item in owned.story_items if item.story.kind == "body"
    ]
    assert [item.direct_body_index for item in owned_body_items] == [
        0,
        0,
        0,
        0,
        1,
        2,
        3,
        None,
    ]

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


def _story_paragraph_texts(document, kind):
    return [
        item.text
        for item in document.story_items
        if item.story.kind == kind and item.kind == "paragraph"
    ]


def test_header_footer_and_story_text_edit_the_section_stories():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("body")
    held = document.paragraphs[0]
    document.set_header("Draft")
    document.set_footer("Page footer")
    with pytest.raises(rdocx.StaleElementError):
        held.text
    assert _story_paragraph_texts(document, "header") == ["Draft"]
    assert _story_paragraph_texts(document, "footer") == ["Page footer"]

    header_item = next(
        item
        for item in document.story_items
        if item.story.kind == "header" and item.kind == "paragraph"
    )
    document.set_story_text(header_item, "Final")
    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert _story_paragraph_texts(reopened, "header") == ["Final"]
    assert _story_paragraph_texts(reopened, "body") == ["body"]


def test_set_story_text_rejects_a_story_that_is_not_in_the_document():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("body")
    missing = rdocx.StoryItem(
        story=rdocx.Story(kind="header", part_name="/word/missing.xml", owner_index=0),
        kind="paragraph",
        index_path=(0,),
        text="stale",
        xml=b"",
    )
    held = document.paragraphs[0]
    before = document.to_bytes()
    with pytest.raises(rdocx.RdocxError, match="no header story"):
        document.set_story_text(missing, "edited")
    assert document.to_bytes() == before
    assert held.text == "body"


def test_hyperlinks_are_added_to_paragraphs_and_stories():
    import rdocx

    document = rdocx.Document()
    run = document.add_paragraph("See ").add_hyperlink("docs", "https://example.com/docs")
    run.font.bold = True
    document.set_header("Header")
    header = next(story for story in document.stories if story.kind == "header")
    document.add_hyperlink_to_story(header, "home", "https://example.com/")

    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert [(link.story.kind, link.text, link.url) for link in reopened.hyperlinks] == [
        ("body", "docs", "https://example.com/docs"),
        ("header", "home", "https://example.com/"),
    ]
    assert reopened.paragraphs[0].runs[1].font.bold is True


def test_story_item_xml_is_a_detached_snapshot():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("hello")
    item = next(
        item
        for item in document.story_items
        if item.story.kind == "body" and item.kind == "paragraph"
    )
    assert isinstance(item.xml, bytes)
    assert b"hello" in item.xml
    document.add_paragraph("later")
    assert b"later" not in item.xml


def test_split_run_enables_exact_comment_ranges_without_losing_content():
    import re

    import rdocx

    document = rdocx.Document()
    run = document.add_paragraph("").add_run("Hello 🐝brave world")
    run.font.bold = True
    before_no_ops = document.to_bytes()
    paragraph = document.paragraphs[0]
    assert document.split_run(0, 0, 0) == 0
    assert document.split_run(0, 0, 18) == 1
    assert document.to_bytes() == before_no_ops
    assert paragraph.text == "Hello 🐝brave world"

    brave_start = document.split_run(0, 0, 6)
    with pytest.raises(rdocx.StaleElementError):
        run.text
    brave_end = document.split_run(0, brave_start, 6)
    assert (brave_start, brave_end) == (1, 2)
    runs = document.paragraphs[0].runs
    assert [item.text for item in runs] == ["Hello ", "🐝brave", " world"]
    assert [item.font.bold for item in runs] == [True, True, True]

    brave = rdocx.RunRange(
        start=rdocx.RunPosition(body_index=0, run_index=brave_start),
        end=rdocx.RunPosition(body_index=0, run_index=brave_end),
    )
    comment_id = document.add_comment(brave, author="Ada", text="Which one?")
    reopened = rdocx.Document.from_bytes(document.to_bytes())
    xml = _document_xml(reopened).decode()
    start = xml.index(f'commentRangeStart w:id="{comment_id}"')
    end = xml.index(f'commentRangeEnd w:id="{comment_id}"')
    assert re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml[start:end]) == [
        "🐝brave"
    ]


def test_split_run_rejects_bad_coordinates_without_changing_the_document():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("abc")
    before = document.to_bytes()
    for args in ((1, 0, 1), (0, 1, 1), (0, 0, 4)):
        with pytest.raises(rdocx.RdocxError):
            document.split_run(*args)
        assert document.to_bytes() == before


def test_python_round_three_authoring_and_inspection_is_typed_and_lossless():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("before")
    anchor = next(
        item
        for item in document.story_items
        if item.story.kind == "body" and item.kind == "paragraph"
    )
    picture = document.add_picture(
        _one_pixel_png(), "pixel.png", rdocx.Inches(1), rdocx.Inches(1), after=anchor
    )
    assert picture.kind == "paragraph"
    assert b"<w:drawing>" in picture.xml
    before_bad_picture = document.to_bytes()
    with pytest.raises(rdocx.RdocxError):
        document.add_picture(
            _one_pixel_png(), "pixel.png", width=rdocx.Inches(1), after=picture
        )
    assert document.to_bytes() == before_bad_picture

    table = document.add_table(1, 1)
    table.cell(0, 0).text = "cell text"
    cell_item = next(
        item
        for item in document.story_items
        if item.story.kind == "table_cell"
        and item.kind == "paragraph"
        and item.text == "cell text"
    )
    before_bad_comment = document.to_bytes()
    with pytest.raises(rdocx.RdocxError):
        document.add_comment(
            rdocx.StoryRunRange(
                start=rdocx.StoryRunPosition(item=cell_item, run_index=0),
                end=rdocx.StoryRunPosition(item=cell_item, run_index=9),
            ),
            author="Ada",
            text="invalid",
        )
    assert document.to_bytes() == before_bad_comment
    comment_id = document.add_comment(
        rdocx.StoryRunRange(
            start=rdocx.StoryRunPosition(item=cell_item, run_index=0),
            end=rdocx.StoryRunPosition(item=cell_item, run_index=1),
        ),
        author="Ada",
        text="Check this cell",
    )

    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert reopened.comments[0].id == comment_id
    xml = _document_xml(reopened)
    assert b"pixel.png" not in xml
    assert f'w:id="{comment_id}"'.encode() in xml
    assert b"cell text" in xml
