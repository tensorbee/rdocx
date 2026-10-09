import io
import posixpath
import re
import struct
import subprocess
import time
import zipfile
import zlib

import pytest


def _replace_document_body(document, body, root_default=None):
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
                    if root_default:
                        root = f'<w:document xmlns="{root_default}" '
                        data = data.replace(b"<w:document ", root.encode(), 1)
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


def test_issue_253_python_render_views():
    from pathlib import Path

    import rdocx

    original = rdocx.Document()
    original.add_paragraph("Keep OLDWORD here.")
    edited = rdocx.Document()
    edited.add_paragraph("Keep NEWWORD here.")
    original.compare(edited, "Reviewer", "2026-09-30T12:00:00Z")
    accepted = original.to_pdf()
    assert accepted == original.to_pdf(revision_view="accepted")
    assert accepted != original.to_pdf(revision_view="tracked")
    fonts = Path(__file__).parents[2] / "oxml-layout" / "fonts"
    assert original.to_pdf(font_dir=fonts) == original.to_pdf(
        font_dir=fonts, revision_view="accepted"
    )
    assert original.to_pdf(font_dir=fonts) != original.to_pdf(
        font_dir=fonts, revision_view="tracked"
    )
    for render in (
        lambda **view: original.render_page_to_png(0, 24.0, **view),
        lambda **view: original.render_all_pages(24.0, **view),
        lambda **view: original.render_pages(dpi=24.0, **view),
    ):
        assert render() == render(revision_view="accepted")
        assert render() != render(revision_view="tracked")
        with pytest.raises(ValueError, match="unknown revision view"):
            render(revision_view="final")
    with pytest.raises(ValueError, match="unknown revision view"):
        original.to_pdf(revision_view="final")


def test_issue_253_selected_view_controls_page_count():
    import rdocx

    original = rdocx.Document()
    original.add_paragraph("OLDWORD " * 1500)
    edited = rdocx.Document()
    edited.add_paragraph("NEWWORD " * 1500)
    original.compare(edited, "Reviewer", "2026-09-30T12:00:00Z")
    accepted = original.render_pages(dpi=24.0)
    tracked = original.render_pages(dpi=24.0, revision_view="tracked")
    assert len(tracked) > len(accepted)
    assert tracked == original.render_all_pages(24.0, revision_view="tracked")


def test_issue_253_pdf_text_view(tmp_path):
    import rdocx

    version = subprocess.run(
        ["pdftotext", "-v"], capture_output=True, check=True, text=True
    )
    assert "pdftotext version 26.01.0" in version.stderr.splitlines()

    original = rdocx.Document()
    original.add_paragraph("Keep OLDWORD here.")
    edited = rdocx.Document()
    edited.add_paragraph("Keep NEWWORD here.")
    original.compare(edited, "Reviewer", "2026-09-30T12:00:00Z")

    def text(view):
        path = tmp_path / "redline.pdf"
        path.write_bytes(original.to_pdf(revision_view=view))
        return subprocess.run(
            ["pdftotext", str(path), "-"], capture_output=True, check=True, text=True
        ).stdout

    accepted = text("accepted")
    tracked = text("tracked")
    assert "NEWWORD" in accepted and "OLDWORD" not in accepted
    assert "NEWWORD" in tracked and "OLDWORD" in tracked


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


def test_revisions_list_every_story_and_name_the_story_that_holds_them():
    import rdocx

    def footer_document(text):
        document = rdocx.Document()
        document.add_paragraph("Body.")
        document.set_footer(text)
        return rdocx.Document.from_bytes(document.to_bytes())

    document = footer_document("Footer lorem ipsum")
    document.compare(
        footer_document("Footer lorem IPSUM"), "R", "2026-09-27T12:00:00Z"
    )
    revisions = document.revisions
    footer = next(story for story in document.stories if story.kind == "footer")
    assert footer.part_name == "/word/footer1.xml"
    assert len(revisions) == 2
    assert all(revision.story == footer for revision in revisions)
    assert sorted(revision.kind for revision in revisions) == ["deletion", "insertion"]
    assert {revision.author for revision in revisions} == {"R"}

    body = rdocx.Story(kind="body", part_name="/word/document.xml", owner_index=0)
    assert {revision.story == body for revision in _tracked_document().revisions} == {
        True
    }
    unscoped = rdocx.Revision(id=1, author="R", timestamp=None, kind="insertion")
    assert unscoped.story is None
    scoped = rdocx.Revision(
        id=1, author="R", timestamp=None, kind="insertion", story=footer
    )
    assert scoped.story == footer
    assert scoped != unscoped

    assert document.accept_all() == len(revisions)
    assert document.revisions == ()


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


def test_counted_replacement_checks_expected_counts_before_publishing():
    import pickle

    import rdocx

    document = rdocx.Document()
    document.add_paragraph("{{a}} {{b}} {{b}} {{c}}")
    document.set_header("{{c}} header")
    held = document.paragraphs[0]
    before = document.to_bytes()

    with pytest.raises(rdocx.ReplacementCountError) as raised:
        document.replace_all([("{{a}}", "A", 1), ("{{b}}", "B", 1), ("{{c}}", "C", 2)])
    error = raised.value
    assert isinstance(error, rdocx.RdocxError)
    assert (error.index, error.expected, error.found) == (1, 1, 2)
    assert str(error) == 'pair 1: expected 1 replacement(s) of "{{b}}", found 2'
    # Worker pools pickle exceptions to return them, so the error must survive.
    copy = pickle.loads(pickle.dumps(error))
    assert (type(copy), str(copy), copy.index, copy.expected, copy.found) == (
        rdocx.ReplacementCountError,
        str(error),
        1,
        1,
        2,
    )
    assert document.to_bytes() == before
    assert held.text == "{{a}} {{b}} {{b}} {{c}}"

    with pytest.raises(
        rdocx.ReplacementCountError,
        match=r'^expected 3 replacement\(s\) of "\{\{b\}\}", found 2$',
    ) as raised:
        document.try_replace_text("{{b}}", "B", expect=3)
    assert (raised.value.index, raised.value.expected, raised.value.found) == (None, 3, 2)
    assert document.to_bytes() == before
    assert document.try_replace_text("{{missing}}", "x", expect=0) == 0
    assert document.replace_all([("{{missing}}", "x")]) == (0,)
    assert document.replace_all([]) == ()
    assert held.text == "{{a}} {{b}} {{b}} {{c}}"

    # Pairs run in order, so the second one also replaces what the first wrote.
    assert document.replace_all(
        [("{{a}}", "{{b}}", 1), ("{{b}}", "B", 3), ("{{c}}", "C", None)]
    ) == (1, 3, 2)
    with pytest.raises(rdocx.StaleElementError):
        held.text
    assert document.paragraphs[0].text == "B B B C"
    assert _story_paragraph_texts(document, "header") == ["C header"]
    assert document.try_replace_text("B", "b", expect=3) == 3

    for pairs, message in (
        ([("x",)], "pair 0 must be"),
        ([("x", "y"), ("x", "y", -1)], "pair 1 must be"),
        ("xy", "pair 0 must be"),
    ):
        with pytest.raises(TypeError, match=message):
            document.replace_all(pairs)


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


def _one_pixel_png(pixel=b"\xff\xff\xff"):
    def chunk(kind, data):
        crc = struct.pack(">I", zlib.crc32(kind + data))
        return struct.pack(">I", len(data)) + kind + data + crc

    header = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(b"\x00" + pixel))
        + chunk(b"IEND", b"")
    )


def _body_images(data):
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        xml = package.read("word/document.xml").decode()
        relationships = package.read("word/_rels/document.xml.rels").decode()
        targets = {
            re.search(r'Id="([^"]+)"', element).group(1): re.search(
                r'Target="([^"]+)"', element
            ).group(1)
            for element in re.findall(r"<Relationship [^>]*>", relationships)
        }
        return [
            package.read(posixpath.join("word", targets[embed]))
            for embed in re.findall(r'r:embed="([^"]+)"', xml)
        ]


@pytest.mark.parametrize(
    ("caption", "size"),
    [
        ("Figure 1. Caption.", (1, 0.66)),
        ("Figure 1. New caption.", (1, 0.66)),
        ("Figure 1. Caption.", (1.2, 0.8)),
    ],
)
def test_compare_records_a_picture_whose_image_changed(caption, size):
    import rdocx

    red, blue = _one_pixel_png(b"\xff\x00\x00"), _one_pixel_png(b"\x00\x00\xff")

    def document(image, caption="Figure 1. Caption.", size=(1, 0.66)):
        document = rdocx.Document()
        document.add_paragraph("Before the figure.")
        document.add_picture(
            image, "figure1.png", rdocx.Inches(size[0]), rdocx.Inches(size[1])
        )
        document.add_paragraph(caption)
        return document

    redline = document(red)
    redline.compare(document(blue, caption, size), "Reviewer", "2026-09-30T12:00:00Z")
    assert redline.revisions
    tracked = redline.to_bytes()
    assert _body_images(tracked) == [red, blue]
    for resolve, image, text in (
        ("accept_all", blue, caption),
        ("reject_all", red, "Figure 1. Caption."),
    ):
        resolved = rdocx.Document.from_bytes(tracked)
        getattr(resolved, resolve)()
        assert _body_images(resolved.to_bytes()) == [image]
        assert [paragraph.text for paragraph in resolved.paragraphs][-1] == text


def test_add_picture_beside_a_content_control_ignores_an_unused_root_default():
    import rdocx

    body = (
        "<w:p><w:r><w:t>Before the control.</w:t></w:r></w:p>"
        '<w:sdt><w:sdtPr><w:tag w:val="goog_rdk_0"/></w:sdtPr><w:sdtContent>'
        "<w:p><w:r><w:t>Inside the control.</w:t></w:r></w:p>"
        "</w:sdtContent></w:sdt>"
        "<w:p><w:r><w:t>After the control.</w:t></w:r></w:p><w:sectPr/>"
    )
    tasks = "http://schemas.microsoft.com/office/tasks/2019/documenttasks"
    for root_default in (None, tasks):
        document = _replace_document_body(rdocx.Document(), body, root_default)
        control = next(
            item
            for item in document.story_items
            if item.story.kind == "body" and item.kind == "content_control"
        )
        size = {"width": rdocx.Inches(1), "height": rdocx.Inches(1)}
        after = document.add_picture(_one_pixel_png(), "x.png", after=control, **size)
        appended = document.add_picture(_one_pixel_png(), "x.png", **size)
        assert after.kind == appended.kind == "paragraph"

        xml = _document_xml(rdocx.Document.from_bytes(document.to_bytes()))
        assert xml.count(b"<w:drawing>") == 2
        positions = [
            xml.index(b"Before the control."),
            xml.index(b'<w:tag w:val="goog_rdk_0"/>'),
            xml.index(b"Inside the control."),
            xml.index(b"</w:sdt>"),
            xml.index(b"<w:drawing>"),
            xml.index(b"After the control."),
            xml.rindex(b"<w:drawing>"),
        ]
        assert positions == sorted(positions)


def test_drawings_of_one_part_that_repeat_an_id_open_and_accept_a_new_picture():
    docx = pytest.importorskip("docx")
    import rdocx

    wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    for repeated in ("0", "7"):
        source = docx.Document()
        for name in "AB":
            source.add_paragraph(f"Picture {name}: ").add_run().add_picture(
                io.BytesIO(_one_pixel_png()), width=914400
            )
        for doc_pr in source.element.body.iter(wp + "docPr"):
            doc_pr.set("id", repeated)
        buffer = io.BytesIO()
        source.save(buffer)
        with zipfile.ZipFile(buffer) as archive:
            source_xml = archive.read("word/document.xml")

        document = rdocx.Document.from_bytes(buffer.getvalue())
        assert [p.text for p in document.paragraphs] == ["Picture A: ", "Picture B: "]
        assert _document_xml(document) == source_xml

        document.add_picture(_one_pixel_png(), "added.png", 12700, 12700)
        saved = document.to_bytes()
        rdocx.Document.from_bytes(saved)
        with zipfile.ZipFile(io.BytesIO(saved)) as archive:
            xml = archive.read("word/document.xml")
        ids = re.findall(rb'<wp:docPr\b[^>]*?\bid="(\d+)"', xml)
        assert ids[:2] == [repeated.encode()] * 2
        assert len(ids) == 3 and ids[2] != repeated.encode()


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


def _linked_report_docx():
    """A python-docx report with two pictures of one image and hyperlinks in
    the body, a table cell, and the header, two of them on one relationship."""
    docx = pytest.importorskip("docx")
    from docx.opc.constants import RELATIONSHIP_TYPE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    def add_link(paragraph, text, url, relationship_id=None):
        if relationship_id is None:
            relationship_id = paragraph.part.relate_to(
                url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True
            )
        link = OxmlElement("w:hyperlink")
        link.set(qn("r:id"), relationship_id)
        run = OxmlElement("w:r")
        properties = OxmlElement("w:rPr")
        style = OxmlElement("w:rStyle")
        style.set(qn("w:val"), "Hyperlink")
        properties.append(style)
        properties.append(OxmlElement("w:b"))
        run.append(properties)
        text_element = OxmlElement("w:t")
        text_element.text = text
        run.append(text_element)
        link.append(run)
        paragraph._p.append(link)
        return relationship_id

    source = docx.Document()
    source.add_picture(
        io.BytesIO(_one_pixel_png()), width=docx.shared.Inches(2), height=docx.shared.Inches(1)
    )
    source.add_picture(
        io.BytesIO(_one_pixel_png()), width=docx.shared.Inches(4), height=docx.shared.Inches(2)
    )
    paragraph = source.add_paragraph("See ")
    shared = add_link(paragraph, "first", "https://example.com/shared")
    add_link(paragraph, "second", "https://example.com/shared", shared)
    add_link(
        source.add_table(1, 1).cell(0, 0).paragraphs[0], "cell", "https://example.com/cell"
    )
    add_link(
        source.sections[0].header.paragraphs[0], "header", "https://example.com/header"
    )
    buffer = io.BytesIO()
    source.save(buffer)
    return buffer.getvalue()


def test_hyperlinks_are_retargeted_or_removed_in_every_story():
    docx = pytest.importorskip("docx")
    import rdocx

    document = rdocx.Document.from_bytes(_linked_report_docx())
    held = document.paragraphs[2]
    # The first link shares its relationship, so only it moves to a new one.
    document.set_hyperlink_url(document.hyperlinks[0], "https://example.org/new")
    links = document.hyperlinks
    assert [(link.story.kind, link.text, link.url) for link in links] == [
        ("body", "first", "https://example.org/new"),
        ("body", "second", "https://example.com/shared"),
        ("table_cell", "cell", "https://example.com/cell"),
        ("header", "header", "https://example.com/header"),
    ]
    document.set_hyperlink_url(links[3], "https://example.org/header")
    document.remove_hyperlink(links[2])
    before = document.to_bytes()
    with pytest.raises(rdocx.RdocxError, match="re-fetch"):
        document.remove_hyperlink(links[2])
    assert document.to_bytes() == before
    assert held.text == "See firstsecond"

    saved_bytes = document.to_bytes()
    saved = docx.Document(io.BytesIO(saved_bytes))
    assert [(link.text, link.url) for link in saved.paragraphs[2].hyperlinks] == [
        ("first", "https://example.org/new"),
        ("second", "https://example.com/shared"),
    ]
    cell = saved.tables[0].cell(0, 0).paragraphs[0]
    assert cell.hyperlinks == []
    assert [(run.text, run.bold, run.style.name) for run in cell.runs] == [
        ("cell", True, "Default Paragraph Font")
    ]
    assert [
        (link.text, link.url)
        for link in saved.sections[0].header.paragraphs[0].hyperlinks
    ] == [("header", "https://example.org/header")]
    with zipfile.ZipFile(io.BytesIO(saved_bytes)) as archive:
        assert b"example.com/cell" not in archive.read("word/_rels/document.xml.rels")


def test_a_stale_hyperlink_snapshot_never_edits_another_link():
    import rdocx

    word = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    rel = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    body = (
        '<w:p><w:hyperlink r:id="rIdU"><w:r><w:t>x</w:t></w:r></w:hyperlink>'
        '<w:r><w:t xml:space="preserve"> </w:t></w:r>'
        '<w:hyperlink r:id="rIdU"><w:r><w:rPr><w:b/></w:rPr><w:t>here</w:t></w:r></w:hyperlink>'
        '<w:r><w:t xml:space="preserve"> and </w:t></w:r>'
        '<w:hyperlink r:id="rIdU"><w:r><w:t>here</w:t></w:r></w:hyperlink></w:p><w:sectPr/>'
    )
    source = io.BytesIO(rdocx.Document().to_bytes())
    result = io.BytesIO()
    with zipfile.ZipFile(source) as source_zip, zipfile.ZipFile(result, "w") as result_zip:
        for info in source_zip.infolist():
            data = source_zip.read(info.filename)
            if info.filename == "word/document.xml":
                data = f"<w:document {word} {rel}><w:body>{body}</w:body></w:document>".encode()
            elif info.filename == "word/_rels/document.xml.rels":
                data = data.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rIdU" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://u/" TargetMode="External"/></Relationships>',
                )
            result_zip.writestr(info, data)
    document = rdocx.Document.from_bytes(result.getvalue())

    links = document.hyperlinks
    document.remove_hyperlink(links[0])
    before = document.to_bytes()
    # The bold link moved to position 0, and the plain one now sits at the
    # snapshot's position with the same fields. Neither is edited.
    with pytest.raises(rdocx.RdocxError, match="re-fetch"):
        document.remove_hyperlink(links[1])
    assert document.to_bytes() == before

    # A record rebuilt from the public fields equals the snapshot and resolves
    # only when exactly one link matches it.
    fresh = document.hyperlinks
    rebuilt = [
        rdocx.Hyperlink(
            story=link.story,
            index_path=link.index_path,
            text=link.text,
            url=link.url,
            anchor=link.anchor,
            relationship_id=link.relationship_id,
        )
        for link in fresh
    ]
    assert rebuilt == list(fresh)
    with pytest.raises(rdocx.RdocxError, match="re-fetch"):
        document.set_hyperlink_url(rebuilt[0], "https://z/")
    document.set_hyperlink_url(fresh[1], "https://plain/")
    document.set_hyperlink_url(fresh[0], "https://bold/")
    assert [(link.text, link.url) for link in document.hyperlinks] == [
        ("here", "https://bold/"),
        ("here", "https://plain/"),
    ]


def test_set_picture_size_resizes_every_drawing_of_a_relationship():
    docx = pytest.importorskip("docx")
    import rdocx

    document = rdocx.Document.from_bytes(_linked_report_docx())
    relationship_id = re.search(
        rb'r:embed="([^"]+)"', _document_xml(document)
    ).group(1).decode()
    # Both pictures share one image relationship, so both are resized.
    assert (
        document.set_picture_size(
            relationship_id, width=rdocx.Inches(1), height=rdocx.Inches(1)
        )
        == 2
    )
    before = document.to_bytes()
    with pytest.raises(rdocx.RdocxError):
        document.set_picture_size("rIdMissing", width=1, height=1)
    assert document.to_bytes() == before
    saved = docx.Document(io.BytesIO(document.to_bytes()))
    assert [(shape.width, shape.height) for shape in saved.inline_shapes] == [
        (rdocx.Inches(1), rdocx.Inches(1))
    ] * 2


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


_COMPARE_TIMESTAMP = "2026-09-27T12:00:00Z"
_LOREM = (
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod "
    "tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim "
    "veniam, quis nostrud exercitation ullamco."
)


def _tracked_texts(xml):
    xml = xml.decode()
    deleted = [
        "".join(re.findall(r"<w:delText(?: [^>]*)?>([^<]*)</w:delText>", wrapper))
        for wrapper in re.findall(r"<w:del\b[^>]*(?<!/)>.*?</w:del>", xml)
    ]
    inserted = [
        "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", wrapper))
        for wrapper in re.findall(r"<w:ins\b[^>]*(?<!/)>.*?</w:ins>", xml)
    ]
    return deleted, inserted


def _package_part_bytes(document, name):
    with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as archive:
        return archive.read(name)


def test_compare_granularity_marks_only_the_changed_word():
    import rdocx

    edited_text = _LOREM.replace("magna", "MAGNA")

    def redline(**options):
        original = rdocx.Document()
        original.add_paragraph(_LOREM)
        edited = rdocx.Document()
        edited.add_paragraph(edited_text)
        assert original.compare(edited, "Ada", _COMPARE_TIMESTAMP, **options) == ()
        return original

    whole_run = ([_LOREM], [edited_text])
    for options, expected in [
        ({}, whole_run),
        ({"granularity": "run"}, whole_run),
        ({"granularity": "word"}, (["magna"], ["MAGNA"])),
        ({"granularity": "character"}, (["magna"], ["MAGNA"])),
    ]:
        compared = redline(**options)
        assert _tracked_texts(_document_xml(compared)) == expected, options
        accepted = rdocx.Document.from_bytes(compared.to_bytes())
        accepted.accept_all()
        assert accepted.paragraphs[0].text == edited_text
        rejected = rdocx.Document.from_bytes(compared.to_bytes())
        rejected.reject_all()
        assert rejected.paragraphs[0].text == _LOREM

    explicit_defaults = redline(
        granularity="run",
        ignore_formatting=False,
        ignore_whitespace=False,
        ignore_fields=False,
        ignore_comments=False,
        ignored_stories=(),
    )
    assert explicit_defaults.to_bytes() == redline().to_bytes()


def test_compare_ignore_options_keep_the_original_side():
    import rdocx

    def compared(original, edited, **options):
        work = rdocx.Document.from_bytes(original.to_bytes())
        assert work.compare(edited, "Ada", _COMPARE_TIMESTAMP, **options) == ()
        return work

    plain = rdocx.Document()
    plain.add_paragraph("plain")
    bold = rdocx.Document.from_bytes(plain.to_bytes())
    bold.paragraphs[0].runs[0].font.bold = True
    assert [item.kind for item in compared(plain, bold).revisions] == [
        "run_property_change"
    ]
    unformatted = compared(plain, bold, ignore_formatting=True)
    assert unformatted.revisions == ()
    assert unformatted.paragraphs[0].runs[0].font.bold is None

    spaced = rdocx.Document()
    spaced.add_paragraph("old  tail")
    single = rdocx.Document()
    single.add_paragraph("old tail")
    assert [item.kind for item in compared(spaced, single).revisions] == [
        "deletion",
        "insertion",
    ]
    unspaced = compared(spaced, single, ignore_whitespace=True)
    assert unspaced.revisions == ()
    assert unspaced.paragraphs[0].text == "old  tail"

    def page_field(result):
        document = rdocx.Document()
        document.add_paragraph("placeholder")
        return _replace_document_body(
            document,
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f"<w:r><w:t>{result}</w:t></w:r>"
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
        )

    first_page, second_page = page_field("1"), page_field("2")
    assert [item.kind for item in compared(first_page, second_page).revisions] == [
        "deletion",
        "insertion",
    ]
    unfielded = compared(first_page, second_page, ignore_fields=True)
    assert unfielded.revisions == ()
    assert b"<w:t>1</w:t>" in _document_xml(unfielded)

    reviewed = rdocx.Document()
    reviewed.add_paragraph("review this")
    reviewed.add_paragraph("old ending")
    commented = rdocx.Document.from_bytes(reviewed.to_bytes())
    commented.add_comment(
        rdocx.RunRange(
            start=rdocx.RunPosition(body_index=0, run_index=0),
            end=rdocx.RunPosition(body_index=0, run_index=1),
        ),
        author="Bo",
        text="edited side note",
    )
    commented.paragraphs[1].runs[0].text = "new ending"
    redline = compared(reviewed, commented, ignore_comments=True)
    assert redline.comments == ()
    assert b"<w:comment" not in _document_xml(redline)
    assert _tracked_texts(_document_xml(redline)) == (["old ending"], ["new ending"])

    headed = rdocx.Document()
    headed.add_paragraph("old body")
    headed.set_header("old header")
    rewritten = rdocx.Document.from_bytes(headed.to_bytes())
    for story in ("body", "header"):
        item = next(
            item
            for item in rewritten.story_items
            if item.story.kind == story and item.text
        )
        rewritten.set_story_text(item, item.text.replace("old", "new"))
    header_part = next(
        story.part_name for story in headed.stories if story.kind == "header"
    ).lstrip("/")
    tracked = compared(headed, rewritten)
    assert _tracked_texts(_package_part_bytes(tracked, header_part)) == (
        ["old header"],
        ["new header"],
    )
    ignored = compared(headed, rewritten, ignored_stories=("header",))
    assert _package_part_bytes(ignored, header_part) == _package_part_bytes(headed, header_part)
    assert _tracked_texts(_document_xml(ignored)) == (["old body"], ["new body"])
    unbodied = compared(headed, rewritten, ignored_stories=("body",))
    assert _document_xml(unbodied) == _document_xml(headed)
    assert _tracked_texts(_package_part_bytes(unbodied, header_part)) == (
        ["old header"],
        ["new header"],
    )
    for story in ("footer", "comment", "text_box", "footnote", "endnote"):
        untouched = compared(headed, rewritten, ignored_stories=(story,))
        assert _tracked_texts(_document_xml(untouched)) == (
            ["old body"],
            ["new body"],
        ), story
        assert _tracked_texts(_package_part_bytes(untouched, header_part)) == (
            ["old header"],
            ["new header"],
        ), story


def test_compare_rejects_unknown_options_before_mutation():
    import rdocx

    original = rdocx.Document()
    original.add_paragraph("before")
    edited = rdocx.Document()
    edited.add_paragraph("after")
    live = original.paragraphs[0]
    before = original.to_bytes()
    for options, error, message in [
        ({"granularity": "words"}, rdocx.RdocxError, 'granularity "words"'),
        ({"ignored_stories": ["main"]}, rdocx.RdocxError, 'story "main"'),
        ({"ignored_stories": ["table_cell"]}, rdocx.RdocxError, 'story "table_cell"'),
        (
            {"ignored_stories": ["header", "header"]},
            rdocx.RdocxError,
            "duplicate ignored story",
        ),
        ({"ignored_stories": "header"}, TypeError, None),
    ]:
        with pytest.raises(error, match=message):
            original.compare(edited, "Ada", _COMPARE_TIMESTAMP, **options)
        assert original.to_bytes() == before
        assert live.text == "before"
    with pytest.raises(TypeError):
        original.compare(edited, "Ada", _COMPARE_TIMESTAMP, "word")
    assert live.text == "before"


def test_issue_161_compare_carries_edited_side_comment_thread():
    import rdocx

    original = rdocx.Document()
    original.add_paragraph("Review this heading")
    edited = rdocx.Document.from_bytes(original.to_bytes())
    edited.add_comment(
        rdocx.RunRange(
            start=rdocx.RunPosition(body_index=0, run_index=0),
            end=rdocx.RunPosition(body_index=0, run_index=1),
        ),
        author="Bo",
        text="Edited side note",
    )
    redline = rdocx.Document.from_bytes(original.to_bytes())
    assert redline.compare(edited, "Ada", _COMPARE_TIMESTAMP) == ()
    assert [comment.text for comment in redline.comments] == ["Edited side note"]
    accepted = rdocx.Document.from_bytes(redline.to_bytes())
    accepted.accept_all()
    assert [comment.text for comment in accepted.comments] == ["Edited side note"]
    rejected = rdocx.Document.from_bytes(redline.to_bytes())
    rejected.reject_all()
    assert rejected.comments == ()


@pytest.mark.parametrize("edit", ["add", "remove", "reply", "resolve", "redate"])
def test_issue_161_comment_edits_resolve_to_each_input(edit):
    import rdocx

    original = rdocx.Document()
    original.add_paragraph("Review this heading")
    if edit != "add":
        original.add_comment(
            rdocx.RunRange(
                start=rdocx.RunPosition(body_index=0, run_index=0),
                end=rdocx.RunPosition(body_index=0, run_index=1),
            ),
            author="Bo",
            text="Original note",
            date="2026-09-16T10:15:30Z",
        )
    edited = rdocx.Document.from_bytes(original.to_bytes())
    if edit == "add":
        edited.add_comment(
            rdocx.RunRange(
                start=rdocx.RunPosition(body_index=0, run_index=0),
                end=rdocx.RunPosition(body_index=0, run_index=1),
            ),
            author="Bo",
            text="Added note",
        )
    elif edit == "remove":
        assert edited.remove_comment(edited.comments[0].id)
    elif edit == "reply":
        edited.reply_to(edited.comments[0].id, author="Ada", text="Agreed")
    elif edit == "resolve":
        assert edited.resolve_comment(edited.comments[0].id, resolved=True)
    else:
        source = io.BytesIO(edited.to_bytes())
        result = io.BytesIO()
        with zipfile.ZipFile(source) as source_zip, zipfile.ZipFile(result, "w") as result_zip:
            for info in source_zip.infolist():
                data = source_zip.read(info.filename)
                if info.filename == "word/comments.xml":
                    data = data.replace(b"2026-09-16T10:15:30Z", b"2026-09-17T10:15:30Z")
                result_zip.writestr(info, data)
        edited = rdocx.Document.from_bytes(result.getvalue())

    redline = rdocx.Document.from_bytes(original.to_bytes())
    assert redline.compare(edited, "Ada", _COMPARE_TIMESTAMP) == ()
    assert redline.comments == edited.comments
    if edit == "redate":
        assert len(redline.revisions) == 1
        selected = rdocx.Document.from_bytes(redline.to_bytes())
        assert selected.accept_revision_id(redline.revisions[0].id) == 1
        assert selected.comments == edited.comments
    accepted = rdocx.Document.from_bytes(redline.to_bytes())
    assert accepted.accept_all() > 0
    assert accepted.comments == edited.comments
    rejected = rdocx.Document.from_bytes(redline.to_bytes())
    assert rejected.reject_all() > 0
    assert rejected.comments == original.comments


def test_issue_161_edited_comment_relationship_survives_redline():
    import rdocx

    original = rdocx.Document()
    original.add_paragraph("Review this")
    edited = rdocx.Document.from_bytes(original.to_bytes())
    edited.add_comment(
        rdocx.RunRange(
            start=rdocx.RunPosition(body_index=0, run_index=0),
            end=rdocx.RunPosition(body_index=0, run_index=1),
        ),
        author="Bo",
        text="See link",
    )
    result = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(edited.to_bytes())) as source_zip:
        with zipfile.ZipFile(result, "w") as result_zip:
            for info in source_zip.infolist():
                data = source_zip.read(info.filename)
                if info.filename == "word/comments.xml":
                    data = data.replace(
                        b"</w:p>",
                        b'<w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId42"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:p>',
                        1,
                    )
                if info.filename == "[Content_Types].xml":
                    data = data.replace(
                        b"</Types>",
                        b'<Override PartName="/word/media/review.bin" ContentType="application/octet-stream"/></Types>',
                    )
                result_zip.writestr(info, data)
            result_zip.writestr(
                "word/_rels/comments.xml.rels",
                b'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId42" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/review" TargetMode="External"/><Relationship Id="rId43" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/review.bin"/></Relationships>',
            )
            result_zip.writestr("word/media/review.bin", b"edited comment asset")
    edited = rdocx.Document.from_bytes(result.getvalue())
    redline = rdocx.Document.from_bytes(original.to_bytes())
    assert redline.compare(edited, "Ada", _COMPARE_TIMESTAMP) == ()
    with zipfile.ZipFile(io.BytesIO(redline.to_bytes())) as archive:
        assert b"https://example.com/review" in archive.read("word/_rels/comments.xml.rels")
        assert archive.read("word/media/review.bin") == b"edited comment asset"
    accepted = rdocx.Document.from_bytes(redline.to_bytes())
    accepted.accept_all()
    with zipfile.ZipFile(io.BytesIO(accepted.to_bytes())) as archive:
        assert b"https://example.com/review" in archive.read("word/_rels/comments.xml.rels")
        assert archive.read("word/media/review.bin") == b"edited comment asset"
    rejected = rdocx.Document.from_bytes(redline.to_bytes())
    rejected.reject_all()
    with zipfile.ZipFile(io.BytesIO(rejected.to_bytes())) as archive:
        assert "word/_rels/comments.xml.rels" not in archive.namelist()
        assert "word/media/review.bin" not in archive.namelist()


@pytest.mark.parametrize("options", [{"ignore_comments": True}, {"ignored_stories": ["comment"]}])
def test_issue_161_ignored_comments_skip_missing_related_target(options):
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("Unchanged body")
    document.add_comment(
        rdocx.RunRange(
            start=rdocx.RunPosition(body_index=0, run_index=0),
            end=rdocx.RunPosition(body_index=0, run_index=1),
        ),
        author="Bo",
        text="Ignored note",
    )
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as source_zip:
        with zipfile.ZipFile(output, "w") as result_zip:
            for info in source_zip.infolist():
                result_zip.writestr(info, source_zip.read(info.filename))
            result_zip.writestr(
                "word/_rels/comments.xml.rels",
                b'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId42" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/missing.png"/></Relationships>',
            )
    damaged = rdocx.Document.from_bytes(output.getvalue())
    original = rdocx.Document.from_bytes(damaged.to_bytes())
    assert original.compare(damaged, "Ada", _COMPARE_TIMESTAMP, **options) == ()


def test_issue_161_comment_asset_collision_preserves_body_image():
    import rdocx

    original = rdocx.Document()
    original.add_paragraph("Review this")
    original.add_picture(_one_pixel_png(), "body.png")
    edited = rdocx.Document.from_bytes(original.to_bytes())
    edited.add_comment(
        rdocx.RunRange(
            start=rdocx.RunPosition(body_index=0, run_index=0),
            end=rdocx.RunPosition(body_index=0, run_index=1),
        ),
        author="Bo",
        text="New comment",
    )
    output = io.BytesIO()
    body_asset = None
    with zipfile.ZipFile(io.BytesIO(edited.to_bytes())) as source_zip:
        body_asset = source_zip.read("word/media/image1.png")
        with zipfile.ZipFile(output, "w") as result_zip:
            for info in source_zip.infolist():
                data = source_zip.read(info.filename)
                if info.filename == "word/_rels/document.xml.rels":
                    data = data.replace(b'media/image1.png', b'media/image2.png')
                if info.filename == "word/media/image1.png":
                    data = b"comment asset with a colliding producer name"
                result_zip.writestr(info, data)
            result_zip.writestr("word/media/image2.png", body_asset)
            result_zip.writestr(
                "word/_rels/comments.xml.rels",
                b'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId88" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
            )
    edited = rdocx.Document.from_bytes(output.getvalue())
    redline = rdocx.Document.from_bytes(original.to_bytes())
    assert redline.compare(edited, "Ada", _COMPARE_TIMESTAMP) == ()
    with zipfile.ZipFile(io.BytesIO(redline.to_bytes())) as archive:
        assert archive.read("word/media/image1.png") == body_asset
        assert archive.read("word/media/image1-rdocx-comment-1.png") == b"comment asset with a colliding producer name"
    accepted = rdocx.Document.from_bytes(redline.to_bytes())
    accepted.accept_all()
    assert accepted.compare(edited, "Ada", _COMPARE_TIMESTAMP) == ()
    rejected = rdocx.Document.from_bytes(redline.to_bytes())
    rejected.reject_all()
    assert rejected.compare(original, "Ada", _COMPARE_TIMESTAMP) == ()


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


def test_rebuild_toc_accepts_identity_attributes_on_the_field_runs():
    import rdocx

    # Word writes w:rsidR and w:rsidRPr on the runs it saves, and Google Docs
    # exports write them on every run, the runs of the TOC field code included.
    source = rdocx.Document()
    source.add_paragraph("placeholder")
    document = _replace_document_body(
        source,
        """
        <w:p><w:r w:rsidR="00A1B2C3"><w:fldChar w:fldCharType="begin"/></w:r><w:r w:rsidRPr="00A1B2C3"><w:instrText>TOC \\o "1-1"</w:instrText></w:r><w:r w:rsidDel="00A1B2C3"><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>
        """,
    )

    assert document.rebuild_toc() == rdocx.TocRebuildReport(
        entry_count=1, bookmark_count=1, diagnostics=()
    )
    saved = _document_xml(document)
    for identity in (b'w:rsidR="00A1B2C3"', b'w:rsidRPr="00A1B2C3"', b'w:rsidDel="00A1B2C3"'):
        assert identity in saved


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
        <w:p><w:fldSimple w:instr="SECTION"><w:r><w:t>stale section</w:t></w:r></w:fldSimple><w:fldSimple w:instr="SECTIONPAGES"><w:r><w:t>stale section count</w:t></w:r></w:fldSimple></w:p>
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
        section_fields=1,
        section_pages_fields=1,
    )
    assert report.updated_count == 5
    assert report.diagnostic_count == len(report.diagnostics)
    with pytest.raises(AttributeError):
        report.page_fields = 0
    xml = _document_xml(document)
    assert b"stale page" not in xml
    assert b"stale count" not in xml
    assert b"stale target" not in xml
    assert b"stale section" not in xml
    legacy = rdocx.LayoutBackedFieldUpdateReport(page_fields=0, num_pages_fields=0, page_reference_fields=0, diagnostics=())
    assert legacy.section_fields == legacy.section_pages_fields == 0
    assert legacy.updated_count == 0
    with pytest.raises(AttributeError):
        report.section_fields = 0


def test_pageref_to_a_bookmarked_run_range_is_filled_from_the_layout():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("See page ")
    document.add_table(2, 2)
    target = document.add_paragraph("The target phrase ends here.")
    target.paragraph_format.page_break_before = True
    index = document.find_content_index("The target phrase")
    assert index == 2
    assert document.split_run(index, 0, len("The target phrase")) == 1
    range_ = rdocx.RunRange(
        start=rdocx.RunPosition(body_index=index, run_index=0),
        end=rdocx.RunPosition(body_index=index, run_index=1),
    )
    held = document.paragraphs[0].runs[0]

    bookmark_id = document.add_bookmark("target", range_)
    assert document.bookmarks == (
        rdocx.Bookmark(
            id=bookmark_id,
            name="target",
            # The recursive range counts the four cell paragraphs of the table.
            range=rdocx.RunRange(
                start=rdocx.RunPosition(body_index=5, run_index=0),
                end=rdocx.RunPosition(body_index=5, run_index=1),
            ),
            direct_range=range_,
            text="The target phrase",
            issue=None,
        ),
    )
    later = next(
        item
        for item in document.story_items
        if item.story.kind == "body" and item.text == "The target phrase ends here."
    )
    held.add_field("PAGEREF target \\h", "?")
    # The field is a story item of its own, so later items move and go stale.
    for stale in (lambda: held.text, lambda: document.set_story_text(later, "x")):
        with pytest.raises(rdocx.StaleElementError):
            stale()
    assert document.paragraphs[0].runs[0].text == "See page "

    report = document.update_layout_backed_fields()
    assert report.page_reference_fields == 1
    field = re.search(
        rb'w:instr="PAGEREF target \\h"[^>]*>\s*<w:r>\s*<w:t>([^<]*)</w:t>',
        _document_xml(document),
    )
    assert field is not None and field.group(1) == b"2"

    before = document.to_bytes()
    with pytest.raises(rdocx.RdocxError, match="already exists"):
        document.add_bookmark("target", range_)
    table_range = rdocx.RunRange(
        start=rdocx.RunPosition(body_index=1, run_index=0),
        end=rdocx.RunPosition(body_index=1, run_index=0),
    )
    with pytest.raises(rdocx.RdocxError, match="is not a paragraph"):
        document.add_bookmark("in_table", table_range)
    run = document.paragraphs[0].runs[0]
    with pytest.raises(rdocx.RdocxError, match="field name"):
        run.add_field("  ")
    assert run.text == "See page "
    assert document.to_bytes() == before


def test_bookmarks_report_nested_ranges_without_a_direct_range():
    import rdocx

    source = rdocx.Document()
    source.add_paragraph("placeholder")
    document = _replace_document_body(
        source,
        """
        <w:sdt><w:sdtContent><w:p><w:bookmarkStart w:id="1" w:name="inside"/><w:r><w:t>Inside</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p></w:sdtContent></w:sdt>
        <w:p><w:bookmarkStart w:id="2" w:name="open"/><w:r><w:t>Open</w:t></w:r></w:p>
        """,
    )
    inside, unmatched = document.bookmarks
    assert inside.name == "inside"
    assert inside.range == rdocx.RunRange(
        start=rdocx.RunPosition(body_index=0, run_index=0),
        end=rdocx.RunPosition(body_index=0, run_index=1),
    )
    assert inside.direct_range is None
    assert inside.text == "Inside"
    assert unmatched.range is None and unmatched.direct_range is None
    assert unmatched.issue == "bookmark id 2 has 1 start markers and 0 end markers"


def test_insert_toc_links_headings_and_tab_and_fields_extend_a_run():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("Intro")
    document.add_paragraph("Chapter one").style = "Heading1"
    document.add_paragraph("Section one point one").style = "Heading2"
    document.add_paragraph("Deep detail").style = "Heading3"
    held = document.paragraphs[0]

    before = document.to_bytes()
    with pytest.raises(IndexError):
        document.insert_toc(5)
    with pytest.raises(ValueError, match="max_level"):
        document.insert_toc(0, max_level=0)
    assert document.to_bytes() == before
    assert held.text == "Intro"

    assert document.insert_toc(0, max_level=2) is None
    with pytest.raises(rdocx.StaleElementError):
        held.text
    assert [paragraph.text for paragraph in document.paragraphs[:3]] == [
        "Table of Contents",
        "Chapter one\t",
        "Section one point one\t",
    ]
    assert [(link.text, link.anchor) for link in document.hyperlinks] == [
        ("Chapter one", "_Toc1"),
        ("Section one point one", "_Toc2"),
    ]
    assert [(bookmark.name, bookmark.text) for bookmark in document.bookmarks] == [
        ("_Toc1", "Chapter one"),
        ("_Toc2", "Section one point one"),
    ]

    run = document.paragraphs[3].runs[0]
    run.add_tab()
    assert run.text == "Intro\t"
    run.add_field("PAGE", "1")
    with pytest.raises(rdocx.StaleElementError):
        run.text
    xml = _document_xml(document)
    assert b'<w:fldSimple w:instr="PAGE"' in xml


def test_word_structure_snapshots_preserve_order_ownership_and_types():
    import rdocx

    document = _document_with_structure_snapshots(rdocx.Document())
    reopened = rdocx.Document.from_bytes(document.to_bytes())

    assert len(reopened.sections) == 3
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


def test_update_section_margins_and_orientation_reach_the_layout():
    import rdocx
    from rdocx import Inches

    document = rdocx.Document()
    document.add_paragraph("body text")
    held = document.paragraphs[0]
    assert (document.layout_page(0).width, document.layout_page(0).height) == (612.0, 792.0)

    updated = document.update_section(
        0, orientation="landscape", margin_top=Inches(0.5), margin_left=Inches(2)
    )

    assert updated == document.sections[0]
    assert (updated.orientation, updated.page_width, updated.page_height) == (
        "landscape",
        Inches(11),
        Inches(8.5),
    )
    assert (
        updated.margin_top,
        updated.margin_right,
        updated.margin_bottom,
        updated.margin_left,
    ) == (Inches(0.5), Inches(1), Inches(1), Inches(2))
    assert held.text == "body text"
    assert (document.layout_page(0).width, document.layout_page(0).height) == (792.0, 612.0)
    bounds = document.layout()[0].bounds
    assert (bounds.x, bounds.y, bounds.width) == (144.0, 36.0, 576.0)
    xml = _document_xml(document)
    assert b'<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>' in xml
    assert b'<w:pgMar w:top="720" w:right="1440" w:bottom="1440" w:left="2880"' in xml
    assert rdocx.Document.from_bytes(document.to_bytes()).sections[0] == updated


def test_update_section_keeps_unnamed_partners_and_rejects_atomically():
    import rdocx
    from rdocx import Inches

    document = rdocx.Document()
    document.add_paragraph("body")
    updated = document.update_section(
        0,
        page_height=Inches(14),
        header_distance=Inches(0.25),
        gutter=Inches(0.1),
        column_count=2,
        column_spacing=Inches(0.5),
        page_number_start=5,
        different_first_page=True,
        break_type="oddPage",
    )
    assert (updated.page_width, updated.page_height) == (Inches(8.5), Inches(14))
    assert (updated.header_distance, updated.footer_distance) == (Inches(0.25), Inches(0.5))
    assert (updated.gutter, updated.column_count, updated.column_spacing) == (
        Inches(0.1),
        2,
        Inches(0.5),
    )
    assert (
        updated.page_number_start,
        updated.different_first_page,
        updated.break_type,
    ) == (5, True, "oddPage")
    assert document.update_section(0, column_count=3).column_spacing == Inches(0.5)

    before = document.to_bytes()
    with pytest.raises(ValueError, match="orientation must be portrait or landscape"):
        document.update_section(0, margin_top=0, orientation="sideways")
    with pytest.raises(ValueError, match="break type"):
        document.update_section(0, break_type="page")
    with pytest.raises(rdocx.RdocxError, match="gutter must be nonnegative"):
        document.update_section(0, margin_top=Inches(3), orientation="landscape", gutter=-635)
    with pytest.raises(rdocx.RdocxError, match="page width must be positive"):
        document.update_section(0, page_width=0)
    with pytest.raises(rdocx.RdocxError, match="page-number start must be positive"):
        document.update_section(0, margin_left=0, page_number_start=0)
    with pytest.raises(IndexError, match="section index out of range"):
        document.update_section(1, gutter=0)
    assert document.to_bytes() == before


def test_update_section_never_rewrites_unequal_width_columns():
    import rdocx
    from rdocx import Inches

    document = _replace_document_body(
        rdocx.Document(),
        '<w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr>'
        '<w:cols w:num="2" w:space="720" w:equalWidth="0">'
        '<w:col w:w="3000" w:space="720"/><w:col w:w="5000"/></w:cols>'
        "</w:sectPr>",
    )
    section = document.sections[0]
    assert (section.column_count, section.column_spacing) == (None, None)

    before = document.to_bytes()
    with pytest.raises(ValueError, match="require both"):
        document.update_section(0, column_count=3)
    with pytest.raises(ValueError, match="require both"):
        document.update_section(0, column_spacing=Inches(0.25))
    assert document.to_bytes() == before

    updated = document.update_section(0, column_count=3, column_spacing=Inches(0.25))
    assert (updated.column_count, updated.column_spacing) == (3, Inches(0.25))
    assert b"<w:col " not in _document_xml(document)


def test_update_section_uses_layout_defaults_for_missing_partners():
    import rdocx
    from rdocx import Inches

    document = _replace_document_body(
        rdocx.Document(),
        '<w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr/>',
    )
    updated = document.update_section(
        0,
        page_width=Inches(9),
        margin_left=Inches(2),
        header_distance=Inches(0.25),
        column_count=2,
    )
    assert (updated.page_width, updated.page_height) == (Inches(9), Inches(11))
    assert (
        updated.margin_top,
        updated.margin_right,
        updated.margin_bottom,
        updated.margin_left,
    ) == (Inches(1), Inches(1), Inches(1), Inches(2))
    assert (updated.header_distance, updated.footer_distance) == (
        Inches(0.25), Inches(0.5)
    )
    assert (updated.column_count, updated.column_spacing) == (2, Inches(0.5))
    assert rdocx.Document.from_bytes(document.to_bytes()).sections[0] == updated


def test_insert_and_remove_section_restructure_the_body():
    import rdocx
    from rdocx import Inches

    document = rdocx.Document()
    document.add_paragraph("first")
    held = document.paragraphs[0]
    document.insert_section(1)
    with pytest.raises(rdocx.StaleElementError):
        _ = held.text

    assert [section.is_final for section in document.sections] == [False, True]
    assert document.sections[1].margin_top is None
    assert document.update_section(1, margin_top=Inches(1)).margin_right == Inches(1)
    document.update_section(
        1,
        orientation="landscape",
        page_width=Inches(8.5),
        page_height=Inches(11),
        margin_top=Inches(1),
        margin_right=Inches(1),
        margin_bottom=Inches(1),
        margin_left=Inches(1),
    )
    document.add_paragraph("second")
    assert [fragment.physical_page for fragment in document.layout()] == [1, 1, 2]
    assert document.layout_page(0).width == 612.0
    assert document.layout_page(1).width == 792.0

    with pytest.raises(IndexError, match="section index out of range"):
        document.insert_section(3)
    with pytest.raises(IndexError, match="section index out of range"):
        document.remove_section(2)
    document.remove_section(1)
    assert len(document.sections) == 1
    assert [paragraph.text for paragraph in document.paragraphs] == ["first", "second"]
    before = document.to_bytes()
    with pytest.raises(rdocx.RdocxError, match="sole document section"):
        document.remove_section(0)
    assert document.to_bytes() == before


def test_section_edits_write_the_native_body():
    # section_edits_write_the_body_the_python_binding_pins in
    # crates/rdocx/tests/integration_test.rs pins the same body for the native
    # calls, so CI checks that these calls write what the native ones write.
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("first")
    document.insert_section(1)
    document.update_section(
        0,
        orientation="landscape",
        margin_top=457200,
        margin_left=1828800,
        page_number_start=3,
        footer_distance=228600,
        different_first_page=True,
    )
    document.update_section(
        1,
        page_width=7772400,
        page_height=10058400,
        margin_top=914400,
        margin_right=1143000,
        margin_bottom=685800,
        margin_left=1371600,
        gutter=127000,
        column_count=2,
        column_spacing=457200,
        break_type="continuous",
    )
    document.add_paragraph("second")

    xml = _document_xml(document).decode()
    body = xml[xml.index("<w:body>") : xml.index("</w:body>") + len("</w:body>")]
    assert "".join(line.strip() for line in body.splitlines()) == (
        '<w:body><w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:pPr>'
        '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>'
        '<w:pgMar w:top="720" w:right="1440" w:bottom="1440" w:left="2880"'
        ' w:gutter="0" w:header="720" w:footer="360"/>'
        '<w:pgNumType w:start="3"/><w:titlePg/></w:sectPr></w:pPr></w:p>'
        '<w:p><w:r><w:t>second</w:t></w:r></w:p><w:sectPr>'
        '<w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1800" w:bottom="1080" w:left="2160"'
        ' w:gutter="200" w:header="720" w:footer="720"/>'
        '<w:cols w:num="2" w:space="720"/></w:sectPr></w:body>'
    )


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


def _two_section_document():
    import rdocx

    section = (
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" '
        'w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" '
        'w:gutter="0"/></w:sectPr>'
    )
    source = rdocx.Document()
    source.add_paragraph("placeholder")
    return _replace_document_body(
        source,
        f"<w:p><w:pPr>{section}</w:pPr><w:r><w:t>First section.</w:t></w:r></w:p>"
        f"<w:p><w:r><w:t>Second section.</w:t></w:r></w:p>{section}",
    )


def _default_footer(document, section_index):
    return next(
        variant.story
        for variant in document.header_footer_variants
        if variant.section_index == section_index
        and variant.kind == "footer"
        and variant.variant == "default"
    )


def _story_part(document, story):
    with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as archive:
        return archive.read(story.part_name.lstrip("/"))


def test_footer_with_text_tab_and_page_fields_is_built_for_one_section():
    import rdocx

    document = _two_section_document()
    held = document.paragraphs[0]
    footer = document.create_section_story(1, "footer", "default")
    assert footer.kind == "footer"
    with pytest.raises(rdocx.StaleElementError):
        held.text
    assert _default_footer(document, 0) is None
    assert _default_footer(document, 1) == footer

    # Author the paragraph in the body with the typed API, then move it.
    paragraph = document.add_paragraph("Confidential")
    paragraph.runs[0].add_tab()
    document.paragraphs[2].add_run("Page ").add_field("PAGE", "1")
    document.paragraphs[2].add_run(" of ").add_field("NUMPAGES", "1")
    fragment = document.pop_content(document.find_content_index("Confidential"))
    document.insert_content(footer, fragment)
    assert [paragraph.text for paragraph in document.paragraphs] == [
        "First section.",
        "Second section.",
    ]
    assert _story_paragraph_texts(document, "footer") == ["ConfidentialPage 1 of 1"]

    report = document.update_layout_backed_fields()
    assert (report.page_fields, report.num_pages_fields) == (0, 0)
    xml = _story_part(document, footer)
    assert re.search(rb"<w:t>Confidential</w:t>\s*<w:tab/>", xml)
    for instruction in (b"PAGE", b"NUMPAGES"):
        field = re.search(
            rb'w:instr="' + instruction + rb'"[^>]*>\s*<w:r>\s*<w:t>([^<]*)</w:t>', xml
        )
        assert field is not None and field.group(1) == b"1"
    body = _document_xml(document)
    first_section = body[: body.index(b"First section.")]
    assert b"footerReference" not in first_section
    assert body.count(b"footerReference") == 1
    assert document.to_pdf().startswith(b"%PDF")


def test_section_stories_link_unlink_and_reject_bad_names_atomically():
    import rdocx

    document = _two_section_document()
    footer = document.create_section_story(1, "footer", "default")
    document.add_paragraph("Footer text")
    fragment = document.pop_content(document.find_content_index("Footer text"))
    document.insert_content(footer, fragment)
    before = document.to_bytes()
    for arguments, error in (
        ((1, "side", "default"), ValueError),
        ((1, "footer", "odd"), ValueError),
        ((2, "footer", "default"), IndexError),
    ):
        with pytest.raises(error):
            document.create_section_story(*arguments)
        with pytest.raises(error):
            document.unlink_section_story(*arguments)
        with pytest.raises(error):
            document.link_section_story(*arguments, footer)
    with pytest.raises(rdocx.RdocxError):
        document.link_section_story(0, "header", "default", footer)
    with pytest.raises(TypeError, match="destination must be an int, a StoryItem or a Story"):
        document.insert_content("footer", fragment)
    with pytest.raises(TypeError, match="index must be an int or a StoryItem"):
        document.pop_content(footer)
    assert document.to_bytes() == before

    assert document.link_section_story(0, "footer", "default", footer) == footer
    assert _default_footer(document, 0) == footer
    copy = document.unlink_section_story(0, "footer", "default")
    assert copy.part_name != footer.part_name
    assert _default_footer(document, 0) == copy
    assert _story_paragraph_texts(document, "footer") == ["Footer text", "Footer text"]

    [item] = [item for item in document.story_items if item.story == copy]
    popped = document.pop_content(item)
    assert popped.kind == "paragraph"
    assert [item for item in document.story_items if item.story == copy] == []
    assert _default_footer(document, 1) == footer
    assert _story_paragraph_texts(document, "footer") == ["Footer text"]


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


def test_complex_field_story_snapshots_preserve_cached_text_and_revision():
    import rdocx

    field = (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>7</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )
    seed = rdocx.Document()
    seed.set_header("header")
    seed.set_footer("footer")
    section_xml = "<w:sectPr" + _document_xml(seed).decode().split("<w:sectPr", 1)[1].split("</w:body>", 1)[0]
    document = _replace_document_body(
        seed, f'<w:p>{field}</w:p><w:sdt><w:sdtPr/><w:sdtContent>'
        f'<w:p>{field}</w:p></w:sdtContent></w:sdt>'
        + section_xml,
    )
    related = {story.part_name.lstrip("/"): story.kind for story in document.stories
               if story.kind in ("header", "footer")}
    result = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as source, zipfile.ZipFile(result, "w") as output:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename in related:
                root = "hdr" if related[info.filename] == "header" else "ftr"
                data = (f'<w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                        f'<w:p>{field}</w:p></w:{root}>').encode()
            output.writestr(info, data)
    document = rdocx.Document.from_bytes(result.getvalue())
    before = document.to_bytes()
    held = document.paragraphs[0]
    snapshots = document.story_items
    fields = [item for item in snapshots if item.kind == "field"]
    assert [(item.story.kind, item.text) for item in fields] == [
        ("body", "7"), ("body", "7"), ("header", "7"), ("footer", "7")]
    assert [item.direct_body_index for item in fields] == [0, 1, None, None]
    assert all(field.encode() in item.xml for item in fields)
    assert len({item.revision for item in snapshots}) == 1
    assert document.story_items == snapshots
    assert document.to_bytes() == before
    assert held.text == "7"
    with zipfile.ZipFile(io.BytesIO(before)) as original:
        expected_parts = {name: original.read(name) for name in original.namelist()}
    reopened = rdocx.Document.from_bytes(before)
    assert [(item.story.kind, item.index_path, item.text, item.xml, item.direct_body_index)
            for item in reopened.story_items] == [
        (item.story.kind, item.index_path, item.text, item.xml, item.direct_body_index)
        for item in snapshots]
    with zipfile.ZipFile(io.BytesIO(reopened.to_bytes())) as saved:
        assert {name: saved.read(name) for name in saved.namelist()} == expected_parts


def test_complex_field_nested_snapshots_keep_links_and_read_handles():
    import rdocx

    nested = (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText> IF </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>7</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '<w:r><w:instrText> = 7 yes no </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>yes</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )
    seed = rdocx.Document()
    paragraph = seed.add_paragraph("")
    paragraph.add_hyperlink("seed", "https://example.test/field-owner")
    relationship_id = seed.hyperlinks[0].relationship_id
    document = _replace_document_body(seed,
        f'<w:p><w:hyperlink r:id="{relationship_id}">{nested}</w:hyperlink>'
        '<w:hyperlink w:anchor="neighbor"><w:r><w:t>neighbor</w:t></w:r></w:hyperlink></w:p>')
    before = document.to_bytes()
    held = document.paragraphs[0]
    snapshots = document.story_items
    fields = [item for item in snapshots if item.kind == "field"]
    assert [item.text for item in fields] == ["yes", "7"]
    assert all(item.revision == snapshots[0].revision for item in snapshots)
    links = document.hyperlinks
    assert [(link.text, link.url, link.anchor) for link in links] == [
        ("7yes", "https://example.test/field-owner", None), ("neighbor", None, "neighbor")]
    assert links[0].relationship_id == relationship_id
    assert document.hyperlinks == links
    assert document.story_items == snapshots
    assert held.text == "yesneighbor"
    assert document.to_bytes() == before
    reopened = rdocx.Document.from_bytes(before)
    assert [(item.index_path, item.text, item.xml) for item in reopened.story_items] == [
        (item.index_path, item.text, item.xml) for item in snapshots]
    assert [(link.index_path, link.text, link.url, link.anchor) for link in reopened.hyperlinks] == [
        (link.index_path, link.text, link.url, link.anchor) for link in links]


@pytest.mark.parametrize("default_binding,ancestor_word", [(False, False), (False, True), (True, False), (True, True)])
def test_complex_field_first_run_namespace_lifetimes_do_not_escape(default_binding, ancestor_word):
    import rdocx

    word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ancestor = word if ancestor_word else "urn:producer"
    local = "urn:producer" if ancestor_word else word
    declaration = "xmlns" if default_binding else "xmlns:x"
    prefix = "" if default_binding else "x:"
    content = "8" if ancestor_word else "SHADOW"
    expected = "87" if ancestor_word else "7"
    header_xml = (f'<q:hdr xmlns:q="{word}" {declaration}="{ancestor}"><q:p>'
        f'<q:r {declaration}="{local}"><q:fldChar q:fldCharType="begin"/></q:r>'
        '<q:r><q:instrText> PAGE </q:instrText></q:r>'
        '<q:r><q:fldChar q:fldCharType="separate"/></q:r>'
        f'<q:r><{prefix}t>{content}</{prefix}t><q:t>7</q:t></q:r>'
        '<q:r><q:fldChar q:fldCharType="end"/></q:r></q:p></q:hdr>').encode()
    seed = rdocx.Document()
    seed.add_paragraph("held")
    seed.set_header("source")
    header = next(story for story in seed.stories if story.kind == "header")
    result = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(seed.to_bytes())) as source, zipfile.ZipFile(result, "w") as output:
        for info in source.infolist():
            data = header_xml if info.filename == header.part_name.lstrip("/") else source.read(info.filename)
            output.writestr(info, data)
    document = rdocx.Document.from_bytes(result.getvalue())
    before = document.to_bytes()
    held = document.paragraphs[0]
    snapshots = document.story_items
    fields = [item for item in snapshots if item.kind == "field" and item.story.kind == "header"]
    assert len(fields) == 1
    assert all(item.revision == snapshots[0].revision for item in snapshots)
    assert document.story_items == snapshots
    assert held.text == "held"
    assert document.to_bytes() == before
    with zipfile.ZipFile(io.BytesIO(before)) as archive:
        assert archive.read(header.part_name.lstrip("/")) == header_xml
    reopened = rdocx.Document.from_bytes(before)
    assert [(item.story.kind, item.index_path, item.xml) for item in reopened.story_items] == [
        (item.story.kind, item.index_path, item.xml) for item in snapshots]
    assert fields[0].text == expected


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


def test_split_run_takes_the_direct_body_index_after_a_table():
    import rdocx

    # GitHub issue #163: the index find_content_index returns must address
    # the same paragraph in split_run.
    document = rdocx.Document()
    document.add_paragraph("Alpha paragraph before the table.")
    document.add_table(1, 1).cell(0, 0).text = "cell"
    document.add_paragraph("Beta paragraph after the table.")
    document.add_paragraph("Gamma paragraph at the end.")
    target = next(p for p in document.paragraphs if p.text.startswith("Beta"))
    body_index = document.find_content_index(target)
    assert body_index == 2

    assert document.split_run(body_index, 0, 4) == 1
    assert [[run.text for run in p.runs] for p in document.paragraphs] == [
        ["Alpha paragraph before the table."],
        ["Beta", " paragraph after the table."],
        ["Gamma paragraph at the end."],
    ]

    before = document.to_bytes()
    with pytest.raises(rdocx.RdocxError, match="is a table, not a paragraph"):
        document.split_run(1, 0, 1)
    with pytest.raises(TypeError, match="must be an int or a Paragraph handle"):
        document.split_run("Beta", 0, 1)
    with pytest.raises(OverflowError):
        document.split_run(-1, 0, 1)
    assert document.to_bytes() == before


def test_split_run_accepts_a_paragraph_handle_in_a_control_or_the_body():
    import rdocx

    document = _replace_document_body(
        rdocx.Document(),
        "<w:p><w:r><w:t>Alpha.</w:t></w:r></w:p>"
        "<w:sdt><w:sdtContent>"
        "<w:p><w:r><w:t>Control one.</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Control two.</w:t></w:r></w:p>"
        "</w:sdtContent></w:sdt>"
        '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
        "<w:tr><w:tc>"
        "<w:sdt><w:sdtContent>"
        "<w:p><w:r><w:t>In control.</w:t></w:r></w:p>"
        "</w:sdtContent></w:sdt>"
        "<w:p><w:r><w:t>Cell text.</w:t></w:r></w:p>"
        "</w:tc></w:tr>"
        "</w:tbl>"
        "<w:p><w:r><w:t>Beta.</w:t></w:r></w:p>",
    )
    control_two = document.paragraphs[2]
    assert control_two.text == "Control two."
    run = control_two.runs[0]
    before_no_op = document.to_bytes()
    assert document.split_run(control_two, 0, 0) == 0
    assert document.to_bytes() == before_no_op
    assert run.text == "Control two."

    assert document.split_run(control_two, 0, 7) == 1
    with pytest.raises(rdocx.StaleElementError):
        run.text
    with pytest.raises(rdocx.StaleElementError):
        document.split_run(control_two, 0, 1)
    assert [r.text for r in document.paragraphs[2].runs] == ["Control", " two."]

    # A cell handle is refused, because its index counts the paragraph
    # inside the cell's content control and the cell writer does not.
    cell = document.tables[0].rows[0].cells[0]
    assert [p.text for p in cell.paragraphs] == ["In control.", "Cell text."]
    before_cell = document.to_bytes()
    for cell_paragraph in cell.paragraphs:
        with pytest.raises(ValueError, match="table cell paragraph handle"):
            document.split_run(cell_paragraph, 0, 2)
    assert document.to_bytes() == before_cell

    beta = document.paragraphs[3]
    assert document.find_content_index(beta) == 3
    beta_run = beta.runs[0]
    assert document.split_run(beta, 0, 5) == 1
    assert beta_run.text == "Beta."
    assert document.split_run(beta, 0, 4) == 1
    assert [r.text for r in document.paragraphs[3].runs] == ["Beta", "."]

    other = rdocx.Document()
    other.add_paragraph("elsewhere")
    before = document.to_bytes()
    with pytest.raises(ValueError, match="different document"):
        document.split_run(other.paragraphs[0], 0, 1)
    assert document.to_bytes() == before


def test_run_remove_keeps_markers_and_drops_emptied_wrappers():
    import rdocx

    # GitHub issue #168 asks for the python-docx idiom
    # run._r.getparent().remove(run._r) as a method of the run.
    document = _replace_document_body(
        rdocx.Document(),
        '<w:p><w:bookmarkStart w:id="1" w:name="kept"/>'
        '<w:r><w:t xml:space="preserve">Keep </w:t></w:r>'
        "<w:r><w:t>drop</w:t></w:r><w:bookmarkEnd w:id=\"1\"/>"
        '<w:hyperlink w:anchor="kept"><w:r><w:t>link</w:t></w:r></w:hyperlink>'
        '<w:ins w:id="2" w:author="Ada"><w:r><w:t>inserted</w:t></w:r></w:ins>'
        "</w:p>",
    )
    paragraph = document.paragraphs[0]
    runs = paragraph.runs
    assert [run.text for run in runs] == ["Keep ", "drop", "link", "inserted"]
    kept = runs[0]
    runs[1].remove()
    with pytest.raises(rdocx.StaleElementError):
        kept.text
    with pytest.raises(rdocx.StaleElementError):
        paragraph.text
    assert [run.text for run in document.paragraphs[0].runs] == [
        "Keep ",
        "link",
        "inserted",
    ]
    document.paragraphs[0].runs[1].remove()
    document.paragraphs[0].runs[1].remove()
    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert [run.text for run in reopened.paragraphs[0].runs] == ["Keep "]
    xml = _document_xml(reopened).decode()
    assert '<w:bookmarkStart w:id="1" w:name="kept"/>' in xml
    assert '<w:bookmarkEnd w:id="1"/>' in xml
    assert "hyperlink" not in xml
    assert "w:ins" not in xml

    table = document.add_table(1, 1)
    table.cell(0, 0).text = "cell"
    document.tables[0].cell(0, 0).paragraphs[0].runs[0].remove()
    assert document.tables[0].cell(0, 0).paragraphs[0].text == ""


def test_run_remove_refuses_part_of_a_field_or_a_comment_reference():
    import rdocx

    document = _replace_document_body(
        rdocx.Document(),
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>Entry</w:t></w:r></w:p>"
        '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        "<w:p><w:r><w:t>Commented</w:t></w:r></w:p>",
    )
    comment_range = rdocx.RunRange(
        start=rdocx.RunPosition(body_index=2, run_index=0),
        end=rdocx.RunPosition(body_index=2, run_index=1),
    )
    document.add_comment(comment_range, author="Ada", text="Here")
    before = document.to_bytes()
    for paragraph, run, reason in (
        (0, 0, "part of a complex field"),
        (0, 2, "part of a complex field"),
        (1, 0, "part of a complex field"),
        (2, 1, "the reference of comment 0"),
    ):
        handle = document.paragraphs[paragraph].runs[run]
        with pytest.raises(rdocx.RdocxError, match=reason):
            handle.remove()
        # A refused removal leaves the handles valid.
        assert handle.text == ""
    assert document.to_bytes() == before

    document.paragraphs[0].runs[3].remove()
    assert document.paragraphs[0].text == ""
    assert [comment.text for comment in document.comments] == ["Here"]


def test_story_comment_after_a_block_content_control_anchors_on_its_paragraph():
    import rdocx

    document = _replace_document_body(
        rdocx.Document(),
        "<w:p><w:r><w:t>Alpha.</w:t></w:r></w:p>"
        "<w:sdt><w:sdtContent>"
        "<w:p><w:r><w:t>Control one.</w:t></w:r></w:p>"
        "</w:sdtContent></w:sdt>"
        "<w:p><w:r><w:t>Beta paragraph.</w:t></w:r></w:p>",
    )
    item = next(
        item
        for item in document.story_items
        if item.story.kind == "body"
        and item.kind == "paragraph"
        and item.text == "Beta paragraph."
    )
    comment_id = document.add_comment(
        rdocx.StoryRunRange(
            start=rdocx.StoryRunPosition(item=item, run_index=0),
            end=rdocx.StoryRunPosition(item=item, run_index=1),
        ),
        author="Ada",
        text="Here",
    )
    xml = _document_xml(document).decode()
    start = xml.index(f'commentRangeStart w:id="{comment_id}"')
    end = xml.index(f'commentRangeEnd w:id="{comment_id}"')
    assert re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml[start:end]) == [
        "Beta paragraph."
    ]


# GitHub issue #172: a run index read from Paragraph.runs anchors on that run,
# including the runs inside an inline content control.
_INLINE_CONTROL_PARAGRAPH = (
    '<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r>'
    "<w:sdt><w:sdtPr/><w:sdtContent>{runs}</w:sdtContent></w:sdt>"
    '<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>'
)


def _anchored_texts(document, comment_id):
    xml = _document_xml(document).decode()
    start = xml.index(f'commentRangeStart w:id="{comment_id}"')
    end = xml.index(f'commentRangeEnd w:id="{comment_id}"')
    return re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml[start:end])


def _run_range(start, end, body_index=0):
    import rdocx

    return rdocx.RunRange(
        start=rdocx.RunPosition(body_index=body_index, run_index=start),
        end=rdocx.RunPosition(body_index=body_index, run_index=end),
    )


def test_comment_run_index_counts_the_runs_of_an_inline_content_control():
    import rdocx

    document = _replace_document_body(
        rdocx.Document(),
        _INLINE_CONTROL_PARAGRAPH.format(runs="<w:r><w:t>TARGET</w:t></w:r>"),
    )
    runs = [run.text for run in document.paragraphs[0].runs]
    assert runs == ["before ", "TARGET", " after"]

    target = document.add_comment(_run_range(1, 2), author="A", text="x")
    assert _anchored_texts(document, target) == ["TARGET"]
    # The reference run of the first comment is run 2 now.
    assert [run.text for run in document.paragraphs[0].runs][3] == " after"
    last = document.add_comment(_run_range(3, 4), author="A", text="y")
    assert _anchored_texts(document, last) == [" after"]

    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert [comment.text for comment in reopened.comments] == ["x", "y"]
    assert _anchored_texts(reopened, target) == ["TARGET"]


def test_story_comment_run_index_counts_the_runs_of_an_inline_content_control():
    import rdocx

    document = _replace_document_body(
        rdocx.Document(),
        _INLINE_CONTROL_PARAGRAPH.format(runs="<w:r><w:t>TARGET</w:t></w:r>"),
    )
    item = next(
        item
        for item in document.story_items
        if item.story.kind == "body" and item.kind == "paragraph"
    )
    comment_id = document.add_comment(
        rdocx.StoryRunRange(
            start=rdocx.StoryRunPosition(item=item, run_index=1),
            end=rdocx.StoryRunPosition(item=item, run_index=2),
        ),
        author="A",
        text="x",
    )
    assert _anchored_texts(document, comment_id) == ["TARGET"]


def test_comment_ranges_that_cannot_be_anchored_exactly_are_refused():
    import rdocx

    document = _replace_document_body(
        rdocx.Document(),
        _INLINE_CONTROL_PARAGRAPH.format(
            runs="<w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r>"
        ),
    )
    before = document.to_bytes()
    for start, end in ((0, 2), (2, 4)):
        with pytest.raises(
            rdocx.RdocxError, match="crosses the edge of an inline content control"
        ):
            document.add_comment(_run_range(start, end), author="A", text="x")
    assert document.to_bytes() == before

    comment_id = document.add_comment(_run_range(2, 3), author="A", text="x")
    assert _anchored_texts(document, comment_id) == ["B"]
    assert document.remove_comment(comment_id)
    xml = _document_xml(document).decode()
    assert "commentRange" not in xml and "commentReference" not in xml
    assert [run.text for run in document.paragraphs[0].runs] == [
        "before ",
        "A",
        "B",
        " after",
    ]


def test_story_run_position_takes_a_paragraph_handle_inside_a_block_control():
    import rdocx

    # GitHub issue #163: a paragraph inside a block content control has no
    # story item of its own, and a handle reaches it.
    document = _replace_document_body(
        rdocx.Document(),
        "<w:p><w:r><w:t>Alpha.</w:t></w:r></w:p>"
        "<w:sdt><w:sdtContent>"
        "<w:p><w:r><w:t>Control one.</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Control two.</w:t></w:r></w:p>"
        "</w:sdtContent></w:sdt>"
        '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
        "<w:tr><w:tc><w:p><w:r><w:t>Cell.</w:t></w:r></w:p></w:tc></w:tr>"
        "</w:tbl>"
        "<w:p><w:r><w:t>Beta.</w:t></w:r></w:p>",
    )
    control_two = document.paragraphs[2]
    assert control_two.text == "Control two."
    position = rdocx.StoryRunPosition(paragraph=control_two, run_index=0)
    assert len(position.item.index_path) == 2
    assert position.item.text == "Control two."
    assert position.item.direct_body_index == 1
    comment_id = document.add_comment(
        rdocx.StoryRunRange(
            start=position,
            end=rdocx.StoryRunPosition(paragraph=control_two, run_index=1),
        ),
        author="Ada",
        text="Here",
    )
    assert _anchored_texts(document, comment_id) == ["Control two."]
    xml = _document_xml(document).decode()
    assert xml.index("<w:sdtContent>") < xml.index("commentRangeStart")
    assert xml.index("commentRangeEnd") < xml.index("</w:sdtContent>")
    with pytest.raises(rdocx.StaleElementError):
        rdocx.StoryRunPosition(paragraph=control_two, run_index=0)

    beta = document.paragraphs[3]
    beta_item = next(item for item in document.story_items if item.text == "Beta.")
    assert rdocx.StoryRunPosition(paragraph=beta, run_index=0).item == beta_item

    cell_paragraph = document.tables[0].rows[0].cells[0].paragraphs[0]
    with pytest.raises(ValueError, match="table cell paragraph handle"):
        rdocx.StoryRunPosition(paragraph=cell_paragraph, run_index=0)
    with pytest.raises(TypeError, match="exactly one of item and paragraph"):
        rdocx.StoryRunPosition(item=beta_item, paragraph=beta, run_index=0)
    with pytest.raises(TypeError, match="exactly one of item and paragraph"):
        rdocx.StoryRunPosition(run_index=0)


def test_add_comment_on_text_anchors_the_occurrence_after_a_table():
    import rdocx

    # GitHub issue #163 asks for a helper that comments on a piece of text.
    # This is the fixture of the issue.
    def fixture():
        document = rdocx.Document()
        document.add_paragraph("Alpha paragraph before the table.")
        document.add_table(1, 1).cell(0, 0).text = "cell"
        document.add_paragraph("Beta paragraph after the table.")
        document.add_paragraph("Gamma paragraph at the end.")
        return document

    document = fixture()
    held = document.paragraphs[1]
    comment_id = document.add_comment_on_text("Beta", author="Ada", text="Here")
    assert _anchored_texts(document, comment_id) == ["Beta"]
    assert [run.text for run in document.paragraphs[1].runs if run.text] == [
        "Beta",
        " paragraph after the table.",
    ]
    with pytest.raises(rdocx.StaleElementError):
        held.text

    document = fixture()
    comment_id = document.add_comment_on_text(
        "paragraph",
        author="Ada",
        text="Second",
        occurrence=1,
        initials="AL",
        date="2026-09-27T10:00:00Z",
    )
    assert _anchored_texts(document, comment_id) == ["paragraph"]
    assert [run.text for run in document.paragraphs[1].runs if run.text] == [
        "Beta ",
        "paragraph",
        " after the table.",
    ]
    reopened = rdocx.Document.from_bytes(document.to_bytes())
    [comment] = reopened.comments
    assert (comment.text, comment.initials, comment.date) == (
        "Second",
        "AL",
        "2026-09-27T10:00:00Z",
    )

    before = document.to_bytes()
    for anchor, occurrence in (("beta", 0), ("paragraph", 3), ("", 0)):
        with pytest.raises(rdocx.RdocxError):
            document.add_comment_on_text(
                anchor, author="Ada", text="x", occurrence=occurrence
            )
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


def test_paragraph_text_setter_replaces_runs_and_keeps_format_and_comments():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("Alpha ")
    document.paragraphs[0].add_run("beta").font.bold = True
    document.paragraphs[0].style = "Heading1"
    document.paragraphs[0].alignment = rdocx.WD_ALIGN_PARAGRAPH.CENTER
    comment_id = document.add_comment(
        rdocx.RunRange(
            start=rdocx.RunPosition(body_index=0, run_index=1),
            end=rdocx.RunPosition(body_index=0, run_index=2),
        ),
        author="Ada",
        text="Bold?",
    )
    paragraph = document.paragraphs[0]
    run = paragraph.runs[0]

    paragraph.text = "Omega\tone\ntwo"

    for stale in (lambda: paragraph.text, lambda: run.text):
        with pytest.raises(rdocx.StaleElementError):
            stale()
    paragraph = document.paragraphs[0]
    assert paragraph.text == "Omega\tone\ntwo"
    assert paragraph.style == "Heading1"
    assert paragraph.alignment == rdocx.WD_ALIGN_PARAGRAPH.CENTER
    assert paragraph.runs[0].text == "Omega\tone\ntwo"
    assert paragraph.runs[0].font.bold is None
    xml = re.sub(r">\s+<", "><", _document_xml(document).decode())
    assert "<w:t>Omega</w:t><w:tab/><w:t>one</w:t><w:br/><w:t>two</w:t>" in xml
    start = xml.index(f'<w:commentRangeStart w:id="{comment_id}"/>')
    end = xml.index(f'<w:commentRangeEnd w:id="{comment_id}"/>')
    assert start < xml.index("Omega") < end
    assert f'<w:commentReference w:id="{comment_id}"/>' in xml[end:]
    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert [comment.text for comment in reopened.comments] == ["Bold?"]

    reopened.paragraphs[0].text = None
    assert reopened.paragraphs[0].text == ""

    document.add_table(1, 1)
    document.tables[0].rows[0].cells[0].paragraphs[0].text = "cell"
    assert document.tables[0].rows[0].cells[0].text == "cell"


def test_paragraph_text_setter_rejects_every_paragraph_of_a_multi_paragraph_field():
    import rdocx

    document = _replace_document_body(
        rdocx.Document(),
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3"</w:instrText></w:r></w:p>'
        '<w:p><w:r><w:instrText xml:space="preserve"> \\h </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>Entry</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Last entry</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
    )
    held = list(document.paragraphs)
    texts = [paragraph.text for paragraph in held]
    before = document.to_bytes()

    for paragraph in held:
        with pytest.raises(rdocx.RdocxError, match="field"):
            paragraph.text = "Rewritten"

    assert document.to_bytes() == before
    assert [paragraph.text for paragraph in held] == texts


_CORE_RELATIONSHIP = (
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
)


def _rewrite_package(document, rewrite):
    source = io.BytesIO(document.to_bytes())
    result = io.BytesIO()
    with zipfile.ZipFile(source) as source_zip:
        with zipfile.ZipFile(result, "w") as result_zip:
            for info in source_zip.infolist():
                data = rewrite(info.filename, source_zip.read(info.filename))
                if data is not None:
                    result_zip.writestr(info, data)
    return type(document).from_bytes(result.getvalue())


def _package_part(document, name):
    with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as archive:
        return archive.read(name).decode()


def test_core_properties_use_python_docx_names_and_round_trip(tmp_path):
    import datetime

    import rdocx

    document = rdocx.Document()
    document.add_paragraph("body")
    held = document.paragraphs[0]
    core = document.core_properties
    assert (core.title, core.author, core.revision, core.created) == ("", "", 0, None)

    utc = datetime.timezone.utc
    plus_two = datetime.timezone(datetime.timedelta(hours=2))
    texts = {
        "author": "Ada",
        "category": "Reports",
        "comments": "Reviewed twice",
        "content_status": "Draft",
        "identifier": "DOC-7",
        "keywords": "alpha, beta",
        "language": "en-GB",
        "last_modified_by": "Grace",
        "subject": "Quarterly figures",
        "title": "A title",
        "version": "1.4",
    }
    for name, value in texts.items():
        setattr(core, name, value)
    core.revision = 3
    core.created = datetime.datetime(2026, 9, 1, 10, 0, 0, tzinfo=plus_two)
    core.modified = datetime.datetime(2026, 9, 2, 8, 30, 0)
    core.last_printed = datetime.datetime(2026, 9, 3, 7, 0, 0, tzinfo=utc)
    assert held.text == "body"

    path = tmp_path / "core.docx"
    document.save(path)
    reopened = rdocx.Document(path).core_properties
    for name, value in texts.items():
        assert getattr(reopened, name) == value, name
    assert reopened.revision == 3
    assert reopened.created == datetime.datetime(2026, 9, 1, 8, 0, 0, tzinfo=utc)
    assert reopened.modified == datetime.datetime(2026, 9, 2, 8, 30, 0, tzinfo=utc)
    assert reopened.last_printed == datetime.datetime(2026, 9, 3, 7, 0, 0, tzinfo=utc)
    xml = _package_part(document, "docProps/core.xml")
    assert '<dcterms:created xsi:type="dcterms:W3CDTF">2026-09-01T08:00:00Z<' in xml
    assert "<cp:lastPrinted>2026-09-03T07:00:00Z</cp:lastPrinted>" in xml

    docx = pytest.importorskip("docx")
    oracle = docx.Document(str(path)).core_properties
    for name, value in texts.items():
        assert getattr(oracle, name) == value, name
    assert oracle.revision == 3
    assert oracle.created == reopened.created

    core.title = None
    core.comments = ""
    core.revision = None
    core.created = None
    assert (core.title, core.comments, core.revision, core.created) == ("", "", 0, None)
    xml = _package_part(document, "docProps/core.xml")
    for element in ("dc:title", "dc:description", "cp:revision", "dcterms:created"):
        assert element not in xml, element


def test_core_properties_reject_bad_values_without_changing_the_document():
    import datetime

    import rdocx

    document = rdocx.Document()
    document.core_properties.title = "Kept"
    before = document.to_bytes()
    for name, value, error in (
        ("revision", 0, ValueError),
        ("title", "x" * 256, ValueError),
        ("created", "2026-09-01", TypeError),
        ("last_printed", datetime.date(2026, 9, 1), TypeError),
    ):
        with pytest.raises(error):
            setattr(document.core_properties, name, value)
    assert document.to_bytes() == before
    document.core_properties.title = "x" * 255
    assert len(document.core_properties.title) == 255


def test_core_properties_read_w3cdtf_dates_like_python_docx():
    import datetime

    import rdocx

    def with_core(created, modified, last_printed):
        def core_xml(name, data):
            if name != "docProps/core.xml":
                return data
            return (
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/'
                '2006/metadata/core-properties" xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
                f'<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified>'
                f"<cp:lastPrinted>{last_printed}</cp:lastPrinted>"
                "<cp:revision>seven</cp:revision>"
                "</cp:coreProperties>"
            ).encode()

        return _rewrite_package(rdocx.Document(), core_xml).core_properties

    utc = datetime.timezone.utc
    core = with_core("2024-01-15T10:30:00-05:30", "2024-02", "not a date")
    assert core.created == datetime.datetime(2024, 1, 15, 16, 0, 0, tzinfo=utc)
    assert core.modified == datetime.datetime(2024, 2, 1, tzinfo=utc)
    assert core.last_printed is None
    assert core.revision == 0

    core = with_core(
        "2024-01-15T10:30:00+99:99", "0001-01-01T00:30:00+01:00", "2024-13-01"
    )
    assert (core.created, core.modified, core.last_printed) == (None, None, None)


def test_core_properties_create_a_missing_part_with_its_relationship():
    import rdocx

    def drop_core(name, data):
        if name == "docProps/core.xml":
            return None
        if name == "_rels/.rels":
            return re.sub(
                rb'<Relationship [^>]*metadata/core-properties"[^>]*/>', b"", data
            )
        if name == "[Content_Types].xml":
            return re.sub(rb'<Override PartName="/docProps/core.xml"[^>]*/>', b"", data)
        return data

    document = _rewrite_package(rdocx.Document(), drop_core)
    assert _CORE_RELATIONSHIP not in _package_part(document, "_rels/.rels")
    assert document.core_properties.title == ""

    document.core_properties.title = "Created"

    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert reopened.core_properties.title == "Created"
    relationships = _package_part(document, "_rels/.rels")
    assert f'Type="{_CORE_RELATIONSHIP}" Target="docProps/core.xml"' in relationships
    assert (
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    ) in _package_part(document, "[Content_Types].xml")


def test_fx153_run_remove_survives_save_and_reopen():
    import rdocx

    document = rdocx.Document()
    paragraph = document.add_paragraph("Keep")
    paragraph.add_run(" Drop").remove()
    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert reopened.paragraphs[0].text == "Keep"


def test_issue_168_complete_edit_save_reopen_layout_and_render(tmp_path):
    import importlib.metadata
    from pathlib import Path

    import docx
    from lxml import etree
    import rdocx

    assert importlib.metadata.version("python-docx") == "1.2.0"
    document = rdocx.Document()
    document.add_paragraph("Report {{name}}. See ")
    document.add_paragraph("Target phrase {{date}}")
    table = document.insert_table(1, 3, 3)
    table.set_cell_grid_span(0, 0, 2)
    table = document.tables[0]
    table.set_cell_vertical_merge(1, 2, "restart")
    table.set_cell_vertical_merge(2, 2, "continue")
    table.set_borders("single", size=4, color="000000")
    table.set_cell_margins(top=0, right=63500, bottom=12700, left=127000)
    table.grid_widths = [1371600, 1828800, 1828800]
    table.cell(1, 0).shading = "D9D9D9"
    table.cell(1, 0).set_margins(top=12700, right=0, bottom=25400, left=6350)
    table.cell(1, 0).set_border("bottom", "double", size=6, color="auto")
    table.rows[0].height = 254000
    table.rows[0].cant_split = True
    table.rows[0].is_header = True

    document.add_style("Report Note", based_on="Normal", bold=True)
    definition = document.add_numbering_definition([rdocx.ListLevel(format="decimal")])
    instance = document.add_numbering_instance(definition)
    document.link_style_to_numbering("Report Note", instance, 0)
    document.paragraphs[1].style = "Report Note"
    document.set_default_style("Report Note")
    before_style_render = document.render_page_to_png(0, 72)
    document.set_style("Report Note", bold=False, space_after=rdocx.Pt(6))
    assert document.render_page_to_png(0, 72) != before_style_render
    before_invalid_style = document.to_bytes()
    with pytest.raises(KeyError):
        document.set_style("Missing", bold=True)
    assert document.to_bytes() == before_invalid_style
    assert document.remove_style("Subtitle") is True
    document.update_section(0, orientation="landscape", margin_left=rdocx.Inches(1.5))
    document.insert_section(1)
    document.add_paragraph("Second section {{state}}")
    footer = document.create_section_story(1, "footer", "default")
    document.add_paragraph("Confidential").runs[0].add_tab()
    document.paragraphs[-1].add_run("Page ").add_field("PAGE", "1")
    document.paragraphs[-1].add_run(" of ").add_field("NUMPAGES", "1")
    fragment = document.pop_content(document.find_content_index("Confidential"))
    document.insert_content(footer, fragment)

    target_index = document.find_content_index("Target phrase")
    document.split_run(target_index, 0, len("Target phrase"))
    document.add_bookmark(
        "target",
        rdocx.RunRange(
            start=rdocx.RunPosition(body_index=target_index, run_index=0),
            end=rdocx.RunPosition(body_index=target_index, run_index=1),
        ),
    )
    document.paragraphs[0].runs[0].add_field("PAGEREF target \\h", "?")
    document.paragraphs[0].add_hyperlink("source", "https://example.com/old")
    document.set_hyperlink_url(document.hyperlinks[0], "https://example.com/new")
    document.paragraphs[0].add_hyperlink("remove", "https://example.com/remove")
    document.remove_hyperlink(document.hyperlinks[-1])
    document.add_picture(_one_pixel_png(), "pixel.png")
    relationship_id = re.search(rb'r:embed="([^"]+)"', _document_xml(document)).group(1).decode()
    assert document.set_picture_size(
        relationship_id, width=rdocx.Inches(1), height=rdocx.Inches(1)
    ) == 1

    held = document.paragraphs[0]
    before = document.to_bytes()
    with pytest.raises(rdocx.ReplacementCountError) as mismatch:
        document.replace_all(
            [("{{name}}", "Ada", 1), ("{{date}}", "October", 2), ("{{state}}", "Done", 1)]
        )
    assert mismatch.value.index == 1
    assert document.to_bytes() == before
    assert held.text.startswith("Report {{name}}")
    assert document.try_replace_text("{{missing}}", "x", expect=0) == 0
    assert document.replace_all(
        [("{{name}}", "Ada", 1), ("{{date}}", "October", 1), ("{{state}}", "Done", 1)]
    ) == (1, 1, 1)
    document.paragraphs[-1].text = "Second section Done"
    document.paragraphs[1].add_run(" remove").remove()
    comment_id = document.add_comment_on_text("Target phrase", author="Ada", text="Review")
    assert _anchored_texts(document, comment_id) == ["Target phrase"]
    report = document.update_layout_backed_fields()
    assert (report.page_fields, report.num_pages_fields, report.page_reference_fields) == (
        0, 0, 1
    )

    path = tmp_path / "issue-168.docx"
    document.save(path)
    # The accepted package-write escape hatch is an external lxml ZIP step.
    # Change an existing unmodelled part without adding a typed binding writer.
    changed = tmp_path / "issue-168-package-edited.docx"
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(changed, "w") as output:
        for member in source.infolist():
            data = source.read(member.filename)
            if member.filename == "docProps/app.xml":
                root = etree.fromstring(data)
                application = root.find(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Application"
                )
                assert application is not None
                application.text = "Issue 168 workflow"
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8")
            output.writestr(member, data)
    path = changed
    reopened = rdocx.Document(path)
    with zipfile.ZipFile(path) as package:
        assert package.testzip() is None
        assert b"<w:tab/>" in package.read(footer.part_name.lstrip("/"))
        assert b"Issue 168 workflow" in package.read("docProps/app.xml")
    assert [style.style_id for style in reopened.styles if style.is_default] == ["ReportNote"]
    assert docx.Document(str(path)).styles["Report Note"].font.bold is False
    assert [(link.text, link.url) for link in reopened.hyperlinks] == [
        ("source", "https://example.com/new")
    ]
    assert reopened.comments[0].text == "Review"
    assert reopened.tables[0].rows[0].is_header is True
    assert reopened.layout_page(0).width == 792.0
    fonts = Path(__file__).parents[2] / "oxml-layout" / "fonts"
    assert reopened.to_pdf(font_dir=fonts).startswith(b"%PDF")
    with zipfile.ZipFile(io.BytesIO(reopened.to_bytes())) as package:
        assert b"Issue 168 workflow" in package.read("docProps/app.xml")
    oracle = docx.Document(str(path))
    assert oracle.paragraphs[0].text.startswith("Report Ada. See ")
    assert [(link.text, link.url) for link in oracle.paragraphs[0].hyperlinks] == [
        ("source", "https://example.com/new")
    ]
    assert oracle.tables[0].cell(1, 0).text == ""
    assert oracle.sections[1].footer.paragraphs[0].text.startswith("Confidential")


# Issue 282, hadim: ownership checks happen before publication or invalidation.
def test_comment_removal_paths_preserve_atomicity_and_revisions():
    import rdocx

    def commented(kind, wrapped=False):
        document = rdocx.Document()
        document.add_paragraph("body anchor")
        document.add_paragraph("retained")
        if kind == "cell":
            document.add_table(2, 1).cell(0, 0).text = "cell anchor"
        item = next(item for item in document.story_items
                    if item.kind == "paragraph"
                    and item.story.kind == ("table_cell" if kind == "cell" else "body")
                    and item.text == ("cell anchor" if kind == "cell" else "body anchor"))
        root = document.add_comment(rdocx.StoryRunRange(
            start=rdocx.StoryRunPosition(item=item, run_index=0),
            end=rdocx.StoryRunPosition(item=item, run_index=1),
        ), author="Ada", text="root")
        reply = document.reply_to(root, author="Ben", text="reply")
        document.reply_to(reply, author="Cyd", text="grandchild")
        if wrapped:
            xml = _document_xml(document).decode()
            body = xml[xml.index("<w:body>") + len("<w:body>"):xml.index("</w:body>")]
            start = body.index("<w:p")
            end = body.index("</w:p>", start) + len("</w:p>")
            body = (body[:start] + '<w:sdt><w:sdtPr><w:tag w:val="goog_rdk_ownership"/></w:sdtPr>'
                    '<w:sdtContent>' + body[start:end] + '</w:sdtContent></w:sdt>' + body[end:])
            document = _replace_document_body(document, body)
        return document

    for wrapped in (False, True):
        document = commented("body", wrapped)
        held = document.paragraphs[-1]
        before = document.to_bytes()
        with pytest.raises(rdocx.RdocxError, match="comment"):
            document.pop_content(0)
        assert document.to_bytes() == before
        assert held.text == "retained"
        assert document.remove_content(0) is True
        assert document.comments == ()
        with pytest.raises(rdocx.StaleElementError):
            held.text
        reopened = rdocx.Document.from_bytes(document.to_bytes())
        assert reopened.comments == ()
        assert [p.text for p in reopened.paragraphs] == ["retained"]

    document = commented("cell")
    held = document.paragraphs[0]
    document.tables[0].remove_row(0)
    assert document.comments == ()
    with pytest.raises(rdocx.StaleElementError):
        held.text
    assert rdocx.Document.from_bytes(document.to_bytes()).comments == ()

    document = commented("cell")
    held = document.tables[0].cell(0, 0)
    held.text = "replacement"
    assert len(document.comments) == 3
    with pytest.raises(rdocx.StaleElementError):
        held.text
    assert document.tables[0].cell(0, 0).text == "replacement"
    assert len(rdocx.Document.from_bytes(document.to_bytes()).comments) == 3

    document = rdocx.Document()
    document.add_paragraph("first")
    document.add_paragraph("last")
    document.add_comment(rdocx.RunRange(
        start=rdocx.RunPosition(body_index=0, run_index=0),
        end=rdocx.RunPosition(body_index=1, run_index=1),
    ), author="Ada", text="partial")
    held = document.paragraphs[0]
    before = document.to_bytes()
    with pytest.raises(rdocx.RdocxError, match="comment"):
        document.remove_content(0)
    assert document.to_bytes() == before
    assert held.text == "first"


def test_comment_anchor_fields_preserve_constructor_compatibility():
    import rdocx

    detached = rdocx.Comment(id=7, author="Ada", initials=None, date=None,
                             text="Review", parent_id=None, resolved=False)
    assert detached.anchor_text is None
    assert detached.anchor is None
    document = rdocx.Document()
    document.add_paragraph("Alpha stays.")
    document.add_paragraph("Delta goes away.")
    document.add_comment(rdocx.RunRange(
        start=rdocx.RunPosition(body_index=1, run_index=0),
        end=rdocx.RunPosition(body_index=1, run_index=1),
    ), author="Ada", text="Review Delta")
    comment = document.comments[0]
    assert comment.anchor_text == "Delta goes away."
    assert isinstance(comment.anchor, rdocx.StoryRunRange)
    assert comment.anchor.start.item.direct_body_index == 1

    with pytest.raises(AttributeError):
        comment.anchor_text = "changed"
    explicit = rdocx.Comment(id=8, author=None, initials=None, date=None,
                             text="detached", parent_id=None, resolved=False,
                             anchor_text=comment.anchor_text, anchor=comment.anchor)
    assert explicit.anchor == comment.anchor
    assert explicit.anchor_text == comment.anchor_text
    preserved = comment.anchor_text
    document.add_comment(comment.anchor, author="Ben", text="Same span")
    assert document.comments[-1].anchor_text == preserved
    assert comment.anchor_text == preserved
    with pytest.raises(rdocx.StaleElementError, match="document revision"):
        document.add_comment(comment.anchor, author="Ben", text="Stale span")
    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert [item.anchor_text for item in reopened.comments] == [preserved, preserved]


@pytest.mark.parametrize("kind", ["body", "table_cell", "header", "footer"])
def test_comment_anchor_typed_related_ranges_are_owned_and_reusable(kind):
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("body anchor")
    document.add_table(1, 1).cell(0, 0).text = "cell anchor"
    document.set_header("header anchor")
    document.set_footer("footer anchor")
    item = next(item for item in document.story_items
                if item.story.kind == kind and item.kind == "paragraph")
    span = rdocx.StoryRunRange(
        start=rdocx.StoryRunPosition(item=item, run_index=0),
        end=rdocx.StoryRunPosition(item=item, run_index=1))
    document.add_comment(span, author="Ada", text="Review")
    before = document.to_bytes()
    comment = document.comments[0]
    assert comment.anchor_text == item.text
    assert comment.anchor.start.item.story.kind == kind
    assert comment.anchor.start.item.direct_body_index == (0 if kind == "body" else None)
    assert document.to_bytes() == before
    document.add_comment(comment.anchor, author="Ben", text="Same span")
    assert document.comments[-1].anchor_text == comment.anchor_text


@pytest.mark.parametrize("multi", [False, True])
def test_comment_anchor_block_control_paragraph_endpoints_are_real_typed_ranges(multi):
    import rdocx

    seed = rdocx.Document()
    seed.add_paragraph("seed")
    identity = seed.add_comment(rdocx.RunRange(
        start=rdocx.RunPosition(body_index=0, run_index=0),
        end=rdocx.RunPosition(body_index=0, run_index=1)), author="Ada", text="Review")
    middle = "</w:p><w:p/><w:p>" if multi else ""
    body = (f'<w:sdt><w:sdtPr><w:tag w:val="goog_rdk_test"/></w:sdtPr><w:sdtContent>'
            f'<w:p><w:commentRangeStart w:id="{identity}"/><w:r><w:t>A</w:t></w:r>'
            f'{middle}<w:r><w:t>B</w:t></w:r><w:commentRangeEnd w:id="{identity}"/>'
            f'<w:r><w:commentReference w:id="{identity}"/></w:r></w:p></w:sdtContent></w:sdt>')
    document = _replace_document_body(seed, body)
    before = document.to_bytes()
    comment = document.comments[0]
    assert comment.anchor_text == ("A\n\nB" if multi else "AB")
    assert comment.anchor.start.item.index_path == (0, 0)
    assert comment.anchor.end.item.index_path == (0, 2 if multi else 0)
    assert comment.anchor.start.item.direct_body_index == 0
    assert comment.anchor.start.item.kind == "paragraph"
    assert isinstance(comment.anchor.start.item.xml, bytes)
    assert document.to_bytes() == before
    document.add_comment(comment.anchor, author="Ben", text="Same source")
    assert document.comments[-1].anchor_text == comment.anchor_text

    # Issue289: every path2 endpoint carries the actual paragraph snapshot.
    import xml.etree.ElementTree as ET
    prefixed = _replace_document_body(seed,
        '<w:p><w:r><w:t>Plain paragraph.</w:t></w:r></w:p>' + body)
    original = prefixed.to_bytes()
    for current in [prefixed, rdocx.Document.from_bytes(original)]:
        bounds = current.comments[0].anchor
        for endpoint, path, text in [
            (bounds.start, (1, 0), "A" if multi else "AB"),
            (bounds.end, (1, 2 if multi else 0), "B" if multi else "AB"),
        ]:
            item = endpoint.item
            assert item.story.kind == "body" and item.kind == "paragraph"
            assert item.index_path == path and item.direct_body_index == 1
            assert item.text == text and item.xml
            assert ET.fromstring(item.xml).tag == (
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")
        held = current.paragraphs[1]
        position = rdocx.StoryRunPosition(paragraph=held, run_index=0)
        assert position.item == bounds.start.item
        assert position.item.text == ("A" if multi else "AB")
        assert position.item.xml
        assert current.to_bytes() == original
        current.add_paragraph("Revision changes")
        with pytest.raises(rdocx.StaleElementError, match="document revision"):
            rdocx.StoryRunPosition(paragraph=held, run_index=0)


def test_comment_anchor_header_block_control_preserves_nonbody_identity():
    import rdocx

    seed = rdocx.Document()
    seed.add_paragraph("body")
    seed.set_header("header")
    identity = seed.add_comment(rdocx.RunRange(
        start=rdocx.RunPosition(body_index=0, run_index=0),
        end=rdocx.RunPosition(body_index=0, run_index=1)), author="Ada", text="Review")
    header = next(story for story in seed.stories if story.kind == "header")
    body = _document_xml(seed)
    for local in ["commentRangeStart", "commentRangeEnd", "commentReference"]:
        body = body.replace(f'<w:{local} w:id="{identity}"/>'.encode(), b"")
    xml = (f'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:sdt><w:sdtPr><w:tag w:val="goog_rdk_header"/></w:sdtPr><w:sdtContent>'
           f'<w:p><w:commentRangeStart w:id="{identity}"/><w:r><w:t>HEADER</w:t></w:r>'
           f'<w:commentRangeEnd w:id="{identity}"/><w:r><w:commentReference w:id="{identity}"/>'
           f'</w:r></w:p></w:sdtContent></w:sdt></w:hdr>').encode()
    result = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(seed.to_bytes())) as archive, zipfile.ZipFile(result, "w") as output:
        for info in archive.infolist():
            data = archive.read(info.filename)
            if info.filename == "word/document.xml":
                data = body
            elif info.filename == header.part_name.lstrip("/"):
                data = xml
            output.writestr(info, data)
    document = rdocx.Document.from_bytes(result.getvalue())
    before = document.to_bytes()
    comment = document.comments[0]
    assert comment.anchor_text == "HEADER"
    assert comment.anchor.start.item.index_path == (0, 0)
    assert comment.anchor.start.item.story.kind == "header"
    assert comment.anchor.start.item.story.part_name == header.part_name
    assert comment.anchor.start.item.direct_body_index is None
    assert document.to_bytes() == before
    document.add_comment(comment.anchor, author="Ben", text="Same header")
    assert document.comments[-1].anchor_text == "HEADER"


def test_comment_moves_refuse_unknown_reply_and_invalid_ranges_atomically():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("SOURCE")
    document.add_paragraph("TARGET TARGET")
    identity = document.add_comment_on_text("SOURCE", author="Ada", initials="A",
        text="Review", date="2026-10-08T12:00:00Z")
    reply = document.reply_to(identity, author="Ben", text="Reply")
    document.resolve_comment(identity)
    metadata = document.comments
    target = document.paragraphs[1]
    bounds = rdocx.StoryRunRange(
        start=rdocx.StoryRunPosition(paragraph=target, run_index=0),
        end=rdocx.StoryRunPosition(paragraph=target, run_index=1))
    held = document.paragraphs[0]
    document.move_comment(identity, bounds)
    assert document.comments == metadata
    assert document.comments[0].anchor_text == "TARGET TARGET"
    with pytest.raises(rdocx.StaleElementError):
        held.text
    document.move_comment_to_text(identity, "TARGET", occurrence=1)
    assert document.comments == metadata
    assert document.comments[0].anchor_text == "TARGET"
    before = document.to_bytes()
    current = document.comments[0].anchor
    for selected in [999, reply]:
        with pytest.raises(rdocx.RdocxError):
            document.move_comment(selected, current)
        assert document.to_bytes() == before
    with pytest.raises(rdocx.RdocxError):
        document.move_comment_to_text(identity, "absent")
    assert document.to_bytes() == before
    with pytest.raises(rdocx.StaleElementError):
        document.move_comment(identity, bounds)
    assert document.to_bytes() == before
    reopened = rdocx.Document.from_bytes(before)
    assert reopened.comments == metadata
    assert reopened.comments[0].anchor_text == "TARGET"


def test_scoped_replacement_count_guards_are_atomic():
    import pickle
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("Ver").add_run("sion head")
    document.add_paragraph("Version head")
    held = document.paragraphs[0]
    run = held.runs[0]
    before = document.to_bytes()
    for old, new, expect in [("absent", "x", 0), ("", "x", 0)]:
        assert held.replace_text(old, new, expect=expect) == 0
        assert held.text == "Version head"
        assert run.text == "Ver"
        assert document.to_bytes() == before
    with pytest.raises(rdocx.ReplacementCountError) as raised:
        held.replace_text("Version", "Release", expect=2)
    error = raised.value
    assert (error.index, error.expected, error.found) == (None, 2, 1)
    copied = pickle.loads(pickle.dumps(error))
    assert (copied.index, copied.expected, copied.found) == (None, 2, 1)
    assert document.to_bytes() == before
    assert held.text == "Version head"
    with pytest.raises(rdocx.RdocxError):
        held.replace_text("absent", "\x01")
    assert document.to_bytes() == before
    assert held.replace_text(old="Version", new="Release", expect=1) == 1
    for stale in (held, run):
        with pytest.raises(rdocx.StaleElementError):
            stale.text
    assert [p.text for p in document.paragraphs] == ["Release head", "Version head"]
    with pytest.raises(rdocx.StaleElementError):
        held.replace_text("Release", "Again")
    reopened = rdocx.Document.from_bytes(document.to_bytes())
    assert [p.text for p in reopened.paragraphs] == ["Release head", "Version head"]


def test_scoped_cell_and_story_item_replacements_have_local_counts():
    import rdocx

    document = rdocx.Document()
    document.add_paragraph("Version outside")
    table = document.add_table(1, 2)
    table.cell(0, 0).text = "Version cell"
    document.tables[0].cell(0, 1).text = "Version neighbor"
    cell = document.tables[0].cell(0, 0)
    cell.add_paragraph("Version second")
    cell = document.tables[0].cell(0, 0)
    paragraph = cell.paragraphs[1]
    before = document.to_bytes()
    assert cell.replace_text("absent", "x", expect=0) == 0
    assert cell.text == "Version cell\nVersion second"
    with pytest.raises(rdocx.ReplacementCountError) as raised:
        cell.replace_text("Version", "Release", expect=1)
    assert (raised.value.index, raised.value.found) == (None, 2)
    assert document.to_bytes() == before
    assert paragraph.replace_text("Version", "Release", expect=1) == 1
    assert document.tables[0].cell(0, 0).text == "Version cell\nRelease second"
    cell = document.tables[0].cell(0, 0)
    assert cell.replace_text("Version", "Release", expect=1) == 1
    assert document.tables[0].cell(0, 1).text == "Version neighbor"
    assert document.paragraphs[0].text == "Version outside"
    document.set_header("Version header")
    item = next(item for item in document.story_items if item.story.kind == "header" and item.kind == "paragraph")
    held = document.paragraphs[0]
    before = document.to_bytes()
    assert document.replace_text_at(item, "absent", "x", expect=0) == 0
    assert document.to_bytes() == before
    assert held.text == "Version outside"
    assert document.replace_text_at(item, "Version", "Release", expect=1) == 1
    with pytest.raises(rdocx.StaleElementError):
        document.replace_text_at(item, "Release", "Again")
    assert _story_paragraph_texts(document, "header") == ["Release header"]
    assert document.paragraphs[0].text == "Version outside"


def test_scoped_control_paragraph_handles_use_direct_owner_ordinals():
    import rdocx

    body = '<w:sdt><w:sdtPr/><w:sdtContent><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>Version nested</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>Version A</w:t></w:r></w:p><w:p><w:r><w:t>Version B</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    document = _replace_document_body(rdocx.Document(), body)
    nested = rdocx.Document.from_bytes(document.to_bytes())
    nested_before = nested.to_bytes()
    held_nested = nested.paragraphs[0]
    nested_unsupported = False
    try:
        nested.paragraphs[0].replace_text("Version", "Release", expect=1)
    except IndexError:
        nested_unsupported = True
    count = document.paragraphs[1].replace_text("Version", "Release", expect=1)
    texts = [paragraph.text for paragraph in document.paragraphs]
    print(f"count={count}, nested_unsupported={nested_unsupported}, texts={texts!r}")
    assert count == 1
    assert texts == ["Version nested", "Release A", "Version B"]
    assert nested_unsupported
    assert nested.to_bytes() == nested_before
    assert held_nested.text == "Version nested"


def _word_text_box_body(choice, fallback):
    def paragraphs(texts):
        return "".join(f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>" for text in texts)

    return (
        '<w:p><w:r><w:t>Host NEEDLE</w:t></w:r><w:r><mc:AlternateContent'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
        ' xmlns:v="urn:schemas-microsoft-com:vml">'
        '<mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/>'
        '<wp:docPr id="1" name="Text Box 1"/><a:graphic><a:graphicData'
        ' uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
        f"<wps:txbx><w:txbxContent>{paragraphs(choice)}</w:txbxContent></wps:txbx><wps:bodyPr/>"
        "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>"
        '<mc:Fallback><w:pict><v:shape style="width:72pt;height:36pt"><v:textbox>'
        f"<w:txbxContent>{paragraphs(fallback)}</w:txbxContent></v:textbox></v:shape></w:pict>"
        "</mc:Fallback></mc:AlternateContent></w:r></w:p>"
    )


def test_replace_text_at_edits_both_copies_of_a_word_text_box():
    """#301: the VML copy of a text box kept the old text."""
    import rdocx

    def text_box_paragraph(document):
        return next(item for item in document.story_items
                    if item.story.kind == "text_box" and item.kind == "paragraph")

    texts = ["Box NEEDLE one", "Box NEEDLE two"]
    document = _replace_document_body(rdocx.Document(), _word_text_box_body(texts, texts))
    assert document.replace_text_at(text_box_paragraph(document), "NEEDLE", "PIN", expect=1) == 1
    xml = _document_xml(document).decode()
    choice, fallback = xml[xml.index("<mc:Choice"):xml.index("<mc:Fallback>")], xml[xml.index("<mc:Fallback>"):]
    for copy in (choice, fallback):
        assert "Box PIN one" in copy and "Box NEEDLE two" in copy, copy
    assert "Host NEEDLE" in xml

    document = _replace_document_body(
        rdocx.Document(), _word_text_box_body(texts, ["Box OLD one", "Box NEEDLE two"]))
    before = document.to_bytes()
    with pytest.raises(rdocx.RdocxError, match="copies of the text box"):
        document.replace_text_at(text_box_paragraph(document), "NEEDLE", "PIN")
    assert document.to_bytes() == before


@pytest.mark.parametrize("kind", ["header", "footer"])
def test_whole_story_comment_refusals_preserve_bytes_and_revision(kind):
    import rdocx

    def fixture(partial=False, malformed=False):
        document = rdocx.Document()
        document.add_paragraph("main retained")
        getattr(document, f"set_{kind}")("old story")
        item = next(item for item in document.story_items
                    if item.kind == "paragraph" and item.story.kind == kind)
        root = document.add_comment(rdocx.StoryRunRange(
            start=rdocx.StoryRunPosition(item=item, run_index=0),
            end=rdocx.StoryRunPosition(item=item, run_index=1)),
            author="Ada", text="root")
        reply = document.reply_to(root, author="Ben", text="reply")
        document.reply_to(reply, author="Cyd", text="grandchild")
        if not partial and not malformed:
            return rdocx.Document.from_bytes(document.to_bytes())
        output = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as source:
            with zipfile.ZipFile(output, "w") as target:
                for info in source.infolist():
                    data = source.read(info.filename)
                    if partial:
                        marker = f'<w:commentRangeStart w:id="{root}"/>'.encode()
                        if info.filename.startswith(f"word/{kind}") and info.filename.endswith(".xml"):
                            data = data.replace(marker, b"")
                        if info.filename == "word/document.xml":
                            data = data.replace(b"<w:p>", b"<w:p>" + marker, 1)
                    if malformed and info.filename == "word/commentsExtended.xml":
                        data = re.sub(rb'(w15:paraId=")[^"]+', rb'\g<1>DEADBEEF', data, count=1)
                    target.writestr(info, data)
        return rdocx.Document.from_bytes(output.getvalue())

    document = fixture()
    held = document.paragraphs[0]
    getattr(document, f"set_{kind}")("new story")
    assert document.comments == ()
    with pytest.raises(rdocx.StaleElementError, match="revision 0.*revision 1"):
        held.text
    assert rdocx.Document.from_bytes(document.to_bytes()).comments == ()
    for partial, malformed, text in ((True, False, "new"),
                                     (False, True, "new"),
                                     (False, False, "invalid \x01")):
        document = fixture(partial, malformed)
        held = document.paragraphs[0]
        before = document.to_bytes()
        with pytest.raises(rdocx.RdocxError):
            getattr(document, f"set_{kind}")(text)
        assert document.to_bytes() == before
        assert held.text == "main retained"
