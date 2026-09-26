//! Integration tests for rdocx — end-to-end document creation and round-trip.

use std::borrow::Cow;
use std::collections::HashMap;

use oxml_opc::OpcPackage;
use oxml_opc::relationship::rel_types;
use rdocx::paragraph::Alignment;
use rdocx::table::{
    CellBorderEdge, CellTextDirection, RowHeight, TableBorderEdge, TableCellMargins,
    TableConditionalFormatting, TableLayout, TableLook, TableWidth, VerticalAlignment,
};
use rdocx::{
    BodyItemRef, BorderStyle, Length, ListLevel, MhtmlDiagnostic, ParagraphRef, RunPosition,
    RunRange, SectionBreak, StoryItemKind, StoryKind, StyleBuilder, TabAlignment, TabLeader,
    TableStyleRegion, UnderlineStyle,
};
use rdocx::{Document, PackageReadLimits, RevisionKind, WordCreationProfile, WordPackageClass};
use rdocx_oxml::CT_BorderEdge;
use rdocx_oxml::header_footer::{CT_HdrFtr, HdrFtrType};
use rdocx_oxml::properties::{CT_PPr, CT_RPr, CT_Shd};
use rdocx_oxml::shared::{ST_Border, ST_PageOrientation, ST_SectionType};
use rdocx_oxml::table::{CT_TblBorders, CT_TblCellMar, CT_TblPr, CT_TcPr};

const ODT_ORACLE_VERSION: &str = "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb";
const MHTML_ORACLE_VERSION: &str = "Microsoft Word 16.104 build 16.104.25121423";
const WORD_SECTION_ORACLE: &str = "Microsoft Word 16.112.3 build 16.112.26083020";
const WORD_ROW_CELL_ORACLE: &str = "Microsoft Word 16.112.4 build 16.112.26090911";
const WORD_HTML_FRAGMENT_ORACLE: &str = "Microsoft Word 16.112.4 build 16.112.26090911";
const LIBREOFFICE_HTML_FRAGMENT_ORACLE: &str =
    "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb";
const POPPLER_HTML_FRAGMENT_ORACLE: &str = "pdftotext version 26.09.0";
const HTML_FRAGMENT_RENDER_RECORD: &[&str] = &[
    "header", "fragment", "link", "3.", "outer", "inner", "head", "cell", "rendered", "fallback",
    "body", "cell", "fragment", "link", "3.", "outer", "inner", "head", "cell", "rendered",
    "fallback", "fragment", "link", "3.", "outer", "inner", "head", "cell", "rendered", "fallback",
    "footer", "fragment", "link", "3.", "outer", "inner", "head", "cell", "rendered", "fallback",
];
const WORD_ROW_CELL_RECORDS: &[&str] = &[
    "table | rows=6 | grid=1200,1800,2400",
    "row 0 | cells=3 | before=None | after=None",
    "row 0 format | height=atLeast:480 | header=Some(true) | alignment=Some(Center)",
    "direction | row=0 | cell=1 | value=TopToBottomRightToLeft",
    "row 1 | cells=2 | before=Some(1) | after=None",
    "direction | row=1 | cell=1 | value=BottomToTopLeftToRight",
    "row 2 | cells=2 | before=None | after=None",
    "direction | row=2 | cell=0 | value=LeftToRightTopToBottomVertical",
    "row 3 | cells=3 | before=None | after=None",
    "direction | row=3 | cell=0 | value=TopToBottomRightToLeftVertical",
    "row 4 | cells=3 | before=None | after=None",
    "direction | row=4 | cell=0 | value=TopToBottomLeftToRightVertical",
    "row 5 | cells=3 | before=None | after=None",
    "cell 0:1 | width=1800 | border=single:336699 | margins=60,90,60,90 | shading=D9EAF7 | valign=Some(Center) | nested=1",
    "cell 2:0 | span=Some(2) | width=3000",
    "merge | restart=Some(Restart) | continuation=Some(Continue)",
];
const WORD_SECTION_ENVIRONMENT: &str = "macOS 26.6.2 build 25G83; locale=en-GB; normalization=f251-section-pdf-v1; pdftotext=26.09.0; pdfinfo=26.09.0";
const WORD_SECTION_RECORDS: [&str; 3] = [
    "page | physical=1 | size_pt=612x792 | PAGE=1",
    "page | physical=2 | size_pt=792x612 | PAGE=12",
    "page | physical=3 | size_pt=595x842 | PAGE=27",
];
const MHTML_ORACLE_HTML: &str = "<h1>Oracle title</h1><p><strong>bold</strong> <a href='https://example.test/'>link</a><img src='https://example.test/pixel.png' width='2' height='3'></p><ol><li>one</li><li>two</li></ol><table><tr><td>cell</td></tr></table>";

fn container_neutral_story_fixture() -> Document {
    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let mut seed = Document::new();
    let bytes = seed.to_bytes().expect("serialize seed document");
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:producer"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p><w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>field</w:t></w:r></w:fldSimple><w:r><w:drawing></w:drawing></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p><x:cell x:flag="exact"/></w:tc></w:tr></w:tbl><w:sdt><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>text box</w:t></w:r></w:p><x:textbox x:flag="exact"/></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p><x:keep x:flag="exact"><x:child/></x:keep><w:sectPr><w:headerReference w:type="default" r:id="storyHeader"/><w:footerReference w:type="default" r:id="storyFooter"/></w:sectPr></w:body></w:document>"#
        )
        .into_bytes(),
    );
    let relationships = package.get_or_create_part_rels("/word/document.xml");
    relationships.add_with_id("storyHeader", rel_types::HEADER, "header-story.xml");
    relationships.add_with_id("storyFooter", rel_types::FOOTER, "footer-story.xml");
    relationships.add_with_id(
        "storyFootnotes",
        rel_types::FOOTNOTES,
        "footnotes-story.xml",
    );
    relationships.add_with_id("storyEndnotes", rel_types::ENDNOTES, "endnotes-story.xml");
    relationships.add_with_id("storyComments", rel_types::COMMENTS, "comments-story.xml");
    for (part, content_type, xml) in [
        (
            "/word/header-story.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            format!(
                r#"<a:hdr xmlns:a="{W}" xmlns:x="urn:producer"><a:p><a:r><a:t>header</a:t></a:r></a:p><x:header x:flag="exact"/></a:hdr>"#
            ),
        ),
        (
            "/word/footer-story.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            format!(
                r#"<w:ftr xmlns:w="{W}" xmlns:x="urn:producer"><w:p><w:r><w:t>footer</w:t></w:r></w:p><x:footer x:flag="exact"/></w:ftr>"#
            ),
        ),
        (
            "/word/footnotes-story.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            format!(
                r#"<w:footnotes xmlns:w="{W}" xmlns:x="urn:producer"><w:footnote w:id="2"><w:p><w:r><w:t>footnote</w:t></w:r></w:p><x:footnote x:flag="exact"/></w:footnote></w:footnotes>"#
            ),
        ),
        (
            "/word/endnotes-story.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
            format!(
                r#"<w:endnotes xmlns:w="{W}" xmlns:x="urn:producer"><w:endnote w:id="2"><w:p><w:r><w:t>endnote</w:t></w:r></w:p><x:endnote x:flag="exact"/></w:endnote></w:endnotes>"#
            ),
        ),
        (
            "/word/comments-story.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
            format!(
                r#"<w:comments xmlns:w="{W}" xmlns:x="urn:producer"><w:comment w:id="2" w:author="Ada"><w:p><w:r><w:t>comment</w:t></w:r></w:p><x:comment x:flag="exact"/></w:comment></w:comments>"#
            ),
        ),
    ] {
        package.set_part(part, xml.into_bytes());
        package.content_types.add_override(part, content_type);
    }
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    Document::from_bytes(&bytes.into_inner()).expect("open all-story fixture")
}

#[test]
fn story_item_links_resolve_only_through_the_checked_owner() {
    const LINK_TYPE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    let mut seed = container_neutral_story_fixture();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap()))
        .expect("open story fixture package");
    let body = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    let body = body.replacen(
        "<w:body>",
        r#"<w:body><w:p><w:hyperlink r:id="scopedLink"><w:r><w:t>body link</w:t></w:r></w:hyperlink></w:p><w:sdt><w:sdtContent><w:sdt><w:sdtContent><w:p><w:hyperlink w:anchor="nested-target"><w:r><w:t>nested link</w:t></w:r></w:hyperlink></w:p></w:sdtContent></w:sdt><w:p><w:hyperlink w:anchor="after-target"><w:r><w:t>after link</w:t></w:r></w:hyperlink></w:p></w:sdtContent></w:sdt>"#,
        1,
    );
    package.set_part("/word/document.xml", body.into_bytes());
    let header = std::str::from_utf8(package.get_part("/word/header-story.xml").unwrap()).unwrap();
    let header = header.replacen(
        "</a:hdr>",
        r#"<a:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:hyperlink r:id="scopedLink"><a:r><a:t>header link</a:t></a:r></a:hyperlink></a:p></a:hdr>"#,
        1,
    );
    package.set_part("/word/header-story.xml", header.into_bytes());
    for (owner, target) in [
        ("/word/document.xml", "https://body.example/"),
        ("/word/header-story.xml", "https://header.example/"),
    ] {
        let relationships = package.get_or_create_part_rels(owner);
        relationships.add_with_id("scopedLink", LINK_TYPE, target);
        relationships
            .items
            .last_mut()
            .expect("inserted relationship")
            .target_mode = Some("External".to_owned());
    }
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let document = Document::from_bytes(&bytes.into_inner()).unwrap();
    let stories = document.stories().unwrap();
    let body = stories
        .iter()
        .find(|story| story.kind() == StoryKind::Body)
        .unwrap()
        .clone();
    let header = stories
        .iter()
        .find(|story| story.kind() == StoryKind::Header)
        .unwrap()
        .clone();
    let body_link = document.story_items(&body).unwrap()[0].links().unwrap();
    let header_link = document
        .story_items(&header)
        .unwrap()
        .into_iter()
        .flat_map(|item| item.links().unwrap())
        .collect::<Vec<_>>();

    assert_eq!(body_link[0].rel_id.as_deref(), Some("scopedLink"));
    assert_eq!(header_link[0].rel_id.as_deref(), Some("scopedLink"));
    assert_eq!(body_link[0].url.as_deref(), Some("https://body.example/"));
    assert_eq!(
        header_link[0].url.as_deref(),
        Some("https://header.example/")
    );

    let body_story_links = document.story_links(&body).unwrap();
    assert_eq!(
        body_story_links
            .iter()
            .map(|(location, link)| (
                link.text.as_str(),
                link.anchor.as_deref(),
                location.index_path()
            ))
            .collect::<Vec<_>>(),
        vec![
            ("body link", None, &[0][..]),
            ("nested link", Some("nested-target"), &[2][..]),
            ("after link", Some("after-target"), &[1][..]),
        ]
    );
}

#[test]
fn one_generic_mutation_edits_the_same_shape_in_every_story() {
    let mut document = container_neutral_story_fixture();
    let stories = document.stories().expect("discover document stories");
    let kinds = stories.iter().map(|story| story.kind()).collect::<Vec<_>>();
    assert_eq!(
        kinds,
        [
            StoryKind::Body,
            StoryKind::TableCell,
            StoryKind::TextBox,
            StoryKind::Header,
            StoryKind::Footer,
            StoryKind::Footnote,
            StoryKind::Endnote,
            StoryKind::Comment,
        ]
    );
    for kind in &kinds {
        let story = document
            .stories()
            .expect("re-resolve stories before mutation")
            .into_iter()
            .find(|story| story.kind() == *kind)
            .expect("fixture story survives prior mutation");
        let location = document
            .story_items(&story)
            .expect("traverse story")
            .into_iter()
            .find(|item| item.kind() == StoryItemKind::Paragraph)
            .unwrap_or_else(|| panic!("{:?} fixture story has no paragraph", story.kind()))
            .location()
            .clone();

        let wrong_owner_kind = if story.part_name() == "/word/document.xml" {
            StoryKind::Header
        } else {
            StoryKind::Body
        };
        let wrong_owner = rdocx::ContentLocation::new(
            rdocx::StoryId::new(wrong_owner_kind, story.part_name(), story.owner_index()),
            StoryItemKind::Paragraph,
            vec![0],
        );
        assert!(matches!(
            document.set_story_text(&wrong_owner, "rejected"),
            Err(rdocx::Error::Story(rdocx::StoryError::WrongOwner { .. }))
        ));
        let wrong_kind = rdocx::ContentLocation::new(
            story.clone(),
            StoryItemKind::Table,
            location.index_path().to_vec(),
        );
        assert!(matches!(
            document.set_story_text(&wrong_kind, "rejected"),
            Err(rdocx::Error::Story(rdocx::StoryError::KindMismatch { .. }))
        ));
        let out_of_bounds =
            rdocx::ContentLocation::new(story.clone(), StoryItemKind::Paragraph, vec![usize::MAX]);
        assert!(matches!(
            document.set_story_text(&out_of_bounds, "rejected"),
            Err(rdocx::Error::Story(rdocx::StoryError::OutOfBounds { .. }))
        ));
        document
            .set_story_text(&location, &format!("edited {:?}", story.kind()))
            .expect("generic paragraph mutation");
        assert!(matches!(
            document.set_story_text(&location, "rejected stale reuse"),
            Err(rdocx::Error::Story(rdocx::StoryError::Stale { .. }))
        ));
    }
    for kind in kinds {
        let story = document
            .stories()
            .expect("re-resolve stories after mutation")
            .into_iter()
            .find(|story| story.kind() == kind)
            .expect("edited fixture story");
        let texts = document
            .story_items(&story)
            .expect("re-traverse edited story")
            .into_iter()
            .filter_map(|item| item.text().expect("project story text"))
            .collect::<Vec<_>>();
        assert!(
            texts.contains(&format!("edited {:?}", story.kind())),
            "{:?} was not edited: {texts:?}",
            story.kind()
        );
    }
}

#[test]
fn story_traversal_preserves_owner_order_and_raw_nodes() {
    let mut document = container_neutral_story_fixture();
    let stories = document.stories().expect("discover document stories");
    let body = stories.first().expect("body story");
    let item_kinds = document
        .story_items(body)
        .expect("traverse body story")
        .into_iter()
        .map(|item| item.kind())
        .collect::<Vec<_>>();
    assert_eq!(
        item_kinds,
        vec![
            StoryItemKind::Paragraph,
            StoryItemKind::Paragraph,
            StoryItemKind::Field,
            StoryItemKind::Drawing,
            StoryItemKind::Table,
            StoryItemKind::ContentControl,
            StoryItemKind::Paragraph,
            StoryItemKind::Drawing,
            StoryItemKind::PreservedNode,
            StoryItemKind::PreservedNode,
        ]
    );
    let expected_owners = vec![
        (
            StoryKind::Body,
            vec![
                StoryItemKind::Paragraph,
                StoryItemKind::Paragraph,
                StoryItemKind::Field,
                StoryItemKind::Drawing,
                StoryItemKind::Table,
                StoryItemKind::ContentControl,
                StoryItemKind::Paragraph,
                StoryItemKind::Drawing,
                StoryItemKind::PreservedNode,
                StoryItemKind::PreservedNode,
            ],
        ),
        (
            StoryKind::TableCell,
            vec![StoryItemKind::Paragraph, StoryItemKind::PreservedNode],
        ),
        (
            StoryKind::TextBox,
            vec![StoryItemKind::Paragraph, StoryItemKind::PreservedNode],
        ),
        (
            StoryKind::Header,
            vec![StoryItemKind::Paragraph, StoryItemKind::PreservedNode],
        ),
        (
            StoryKind::Footer,
            vec![StoryItemKind::Paragraph, StoryItemKind::PreservedNode],
        ),
        (
            StoryKind::Footnote,
            vec![StoryItemKind::Paragraph, StoryItemKind::PreservedNode],
        ),
        (
            StoryKind::Endnote,
            vec![StoryItemKind::Paragraph, StoryItemKind::PreservedNode],
        ),
        (
            StoryKind::Comment,
            vec![StoryItemKind::Paragraph, StoryItemKind::PreservedNode],
        ),
    ];
    let before_owners = stories
        .iter()
        .map(|story| {
            (
                story.kind(),
                document
                    .story_items(story)
                    .unwrap()
                    .into_iter()
                    .map(|item| item.kind())
                    .collect::<Vec<_>>(),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(before_owners, expected_owners);
    let before = document.to_bytes().unwrap();
    let mut reopened = Document::from_bytes(&before).unwrap();
    let after_owners = reopened
        .stories()
        .unwrap()
        .into_iter()
        .map(|story| {
            let items = reopened
                .story_items(&story)
                .unwrap()
                .into_iter()
                .map(|item| item.kind())
                .collect::<Vec<_>>();
            (story.kind(), items)
        })
        .collect::<Vec<_>>();
    assert_eq!(after_owners, expected_owners);
    let after = reopened.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(after)).unwrap();
    for (part_name, retained) in [
        (
            "/word/document.xml",
            "<x:keep x:flag=\"exact\"><x:child/></x:keep>",
        ),
        ("/word/document.xml", "<x:cell x:flag=\"exact\"/>"),
        ("/word/document.xml", "<x:textbox x:flag=\"exact\"/>"),
        ("/word/header-story.xml", "<x:header x:flag=\"exact\"/>"),
        ("/word/footer-story.xml", "<x:footer x:flag=\"exact\"/>"),
        (
            "/word/footnotes-story.xml",
            "<x:footnote x:flag=\"exact\"/>",
        ),
        ("/word/endnotes-story.xml", "<x:endnote x:flag=\"exact\"/>"),
        ("/word/comments-story.xml", "<x:comment x:flag=\"exact\"/>"),
    ] {
        let part = package.get_part(part_name).unwrap();
        assert!(
            part.windows(retained.len())
                .any(|window| window == retained.as_bytes()),
            "{part_name} lost {retained}"
        );
    }
}

#[test]
fn story_xml_distinguishes_owned_typed_sources_from_borrowed_package_slices() {
    let document = container_neutral_story_fixture();
    let stories = document.stories().unwrap();
    let body = stories
        .iter()
        .find(|story| story.kind() == StoryKind::Body)
        .unwrap();
    let header = stories
        .iter()
        .find(|story| story.kind() == StoryKind::Header)
        .unwrap();
    let comment = stories
        .iter()
        .find(|story| story.kind() == StoryKind::Comment)
        .unwrap();

    let body_xml = document.story_items(body).unwrap()[0].xml().unwrap();
    assert!(matches!(body_xml, Cow::Owned(_)));
    let header_xml = document.story_items(header).unwrap()[0].xml().unwrap();
    assert!(matches!(header_xml, Cow::Borrowed(_)));
    assert!(
        !std::str::from_utf8(header_xml.as_ref())
            .unwrap()
            .contains("xmlns:a"),
        "borrowed subtree unexpectedly materialized its ancestor namespace"
    );
    let comment_xml = document.story_items(comment).unwrap()[0].xml().unwrap();
    assert!(matches!(comment_xml, Cow::Owned(_)));
}

#[test]
fn section_lookup_is_total_for_every_index() {
    let mut equivalent = Document::new();
    let before_equivalent = equivalent.to_bytes().unwrap();
    equivalent.insert_section(0).unwrap();
    equivalent.remove_section(0).unwrap();
    assert_eq!(equivalent.to_bytes().unwrap(), before_equivalent);

    let mut document = Document::new();
    assert_eq!(document.section_count(), 1);
    assert_eq!(document.sections().count(), 1);
    assert_eq!(document.section(0).unwrap().ordinal(), 0);
    assert!(document.section(0).unwrap().is_final());
    assert!(document.section(1).is_none());
    assert!(document.section_mut(1).is_none());

    document.insert_section(0).unwrap();
    document.insert_section(2).unwrap();
    assert_eq!(document.section_count(), 3);
    for (index, orientation) in [
        ST_PageOrientation::Portrait,
        ST_PageOrientation::Landscape,
        ST_PageOrientation::Portrait,
    ]
    .into_iter()
    .enumerate()
    {
        let mut section = document.section_mut(index).unwrap();
        assert_eq!(section.ordinal(), index);
        assert_eq!(section.is_final(), index == 2);
        section.set_orientation(orientation);
        section.set_different_first_page(index == 0);
    }
    assert_eq!(
        document
            .sections()
            .map(|section| section.orientation())
            .collect::<Vec<_>>(),
        [
            Some(ST_PageOrientation::Portrait),
            Some(ST_PageOrientation::Landscape),
            Some(ST_PageOrientation::Portrait),
        ]
    );
    assert_eq!(
        document
            .sections()
            .map(|section| (section.ordinal(), section.is_final()))
            .collect::<Vec<_>>(),
        [(0, false), (1, false), (2, true)]
    );
    assert_eq!(
        document.section(0).unwrap().properties().title_pg,
        Some(true)
    );
    let landscape = document.section(1).unwrap();
    assert!(
        landscape.properties().page_width.unwrap().0
            > landscape.properties().page_height.unwrap().0
    );

    let before = document.to_bytes().unwrap();
    assert!(document.insert_section(4).is_err());
    assert!(document.remove_section(3).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn section_geometry_round_trips_with_unsupported_children_in_order() {
    let seed = Document::new().to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(seed)).unwrap();
    let source = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:producer"><w:body><w:p><w:r><w:t>geometry</w:t></w:r></w:p><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:paperSrc w:first="7"/><q:pgNumType q:fmt='lowerRoman' x:start='99' q:start = '3' q:chapStyle='2' q:chapSep='hyphen'/><w:cols w:num="2" w:space="360" w:sep="1"/><w:vAlign w:val="center"/><w:titlePg/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>"#;
    package.set_part("/word/document.xml", source.to_vec());
    let mut archive = std::io::Cursor::new(Vec::new());
    package.write_to(&mut archive).unwrap();
    let mut document = Document::from_bytes(archive.get_ref()).unwrap();

    {
        let mut section = document.section_mut(0).unwrap();
        section
            .set_page_size(Length::inches(11.0), Length::inches(8.5))
            .unwrap();
        section.set_orientation(ST_PageOrientation::Landscape);
        section
            .set_margins(
                Length::inches(0.5),
                Length::inches(0.75),
                Length::inches(1.0),
                Length::inches(1.25),
            )
            .unwrap();
        section.set_gutter(Length::inches(0.2)).unwrap();
        section.set_columns(3, Length::inches(0.25)).unwrap();
        section.set_page_number_start(7).unwrap();
        section
            .set_header_footer_distance(Length::inches(0.3), Length::inches(0.4))
            .unwrap();
        section.set_different_first_page(false);
        section.set_break_type(ST_SectionType::OddPage);
    }

    let saved = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&saved).unwrap();
    let section = reopened.section(0).unwrap();
    assert_eq!(section.page_size().unwrap().0.to_twips(), 15840);
    assert_eq!(section.page_size().unwrap().1.to_twips(), 12240);
    assert_eq!(section.orientation(), Some(ST_PageOrientation::Landscape));
    let (top, right, bottom, left) = section.margins().unwrap();
    assert_eq!(
        [
            top.to_twips(),
            right.to_twips(),
            bottom.to_twips(),
            left.to_twips(),
        ],
        [720, 1080, 1440, 1800]
    );
    assert_eq!(section.gutter().unwrap().to_twips(), 288);
    assert_eq!(
        section
            .columns()
            .map(|(count, spacing)| (count, spacing.to_twips())),
        Some((3, 360))
    );
    assert_eq!(section.page_number_start(), Some(7));
    assert_eq!(
        section
            .header_footer_distance()
            .map(|(header, footer)| (header.to_twips(), footer.to_twips())),
        Some((432, 576))
    );
    assert_eq!(section.different_first_page(), Some(false));
    assert_eq!(section.break_type(), Some(ST_SectionType::OddPage));

    let saved_package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let xml = std::str::from_utf8(saved_package.get_part("/word/document.xml").unwrap()).unwrap();
    for retained in [
        r#"<w:paperSrc w:first="7"/>"#,
        r#"q:fmt='lowerRoman'"#,
        r#"x:start='99'"#,
        r#"q:chapStyle='2'"#,
        r#"q:chapSep='hyphen'"#,
        r#"<w:vAlign w:val="center"/>"#,
        r#"<w:docGrid w:linePitch="360"/>"#,
    ] {
        assert!(
            xml.contains(retained),
            "missing retained XML: {retained}\n{xml}"
        );
    }
    let positions = [
        "w:pgMar",
        "w:paperSrc",
        "q:pgNumType",
        "w:cols",
        "w:vAlign",
        "w:titlePg",
        "w:docGrid",
    ]
    .map(|name| xml.find(name).unwrap());
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]), "{xml}");
    assert!(xml.contains("q:start = '7'"), "{xml}");
    assert!(xml.contains(r#"w:sep="1""#), "{xml}");
}

#[test]
fn rejected_section_geometry_is_atomic() {
    let mut document = Document::new();
    let before = document.to_bytes().unwrap();
    {
        let mut section = document.section_mut(0).unwrap();
        assert!(
            section
                .set_page_size(Length::twips(0), Length::inches(11.0))
                .is_err()
        );
        assert!(section.set_columns(0, Length::inches(0.5)).is_err());
        assert!(section.set_columns(2, Length::twips(-1)).is_err());
        let overflow = Length::emu((i32::MAX as i64 + 1) * 635);
        assert!(
            section
                .set_page_size(overflow, Length::inches(11.0))
                .is_err()
        );
        assert!(
            section
                .set_margins(
                    Length::inches(1.0),
                    overflow,
                    Length::inches(1.0),
                    Length::inches(1.0),
                )
                .is_err()
        );
        assert!(section.set_gutter(overflow).is_err());
        assert!(section.set_columns(2, overflow).is_err());
        assert!(
            section
                .set_header_footer_distance(overflow, Length::inches(0.5))
                .is_err()
        );
        assert!(section.set_page_number_start(0).is_err());
    }
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn legacy_section_geometry_setters_preserve_infallible_compatibility() {
    let mut document = Document::new();
    document.set_page_size(Length::twips(0), Length::twips(-1));
    let overflow = Length::emu((i32::MAX as i64 + 1) * 635);
    document.set_margins(overflow, overflow, overflow, overflow);
    document.set_columns(0, Length::twips(-2));
    document.set_gutter(Length::twips(-3));
    document.set_header_footer_distance(Length::twips(-4), Length::twips(-5));

    let section = document.section_properties().unwrap();
    assert_eq!(section.page_width.unwrap().0, 0);
    assert_eq!(section.page_height.unwrap().0, -1);
    assert_eq!(section.margin_top.unwrap().0, i32::MIN);
    assert_eq!(section.margin_right.unwrap().0, i32::MIN);
    assert_eq!(section.margin_bottom.unwrap().0, i32::MIN);
    assert_eq!(section.margin_left.unwrap().0, i32::MIN);
    assert_eq!(section.columns.as_ref().unwrap().num, Some(0));
    assert_eq!(section.columns.as_ref().unwrap().space.unwrap().0, -2);
    assert_eq!(section.gutter.unwrap().0, -3);
    assert_eq!(section.header_distance.unwrap().0, -4);
    assert_eq!(section.footer_distance.unwrap().0, -5);
}

struct F251OracleArtifacts {
    path: std::path::PathBuf,
}

impl F251OracleArtifacts {
    fn create(path: std::path::PathBuf) -> Self {
        std::fs::create_dir_all(&path).unwrap();
        Self { path }
    }

    fn path(&self) -> &std::path::Path {
        &self.path
    }
}

impl Drop for F251OracleArtifacts {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.path);
    }
}

fn f251_oracle_vectors_are_complete(
    page_count: usize,
    size_count: usize,
    displayed_count: usize,
) -> bool {
    page_count == WORD_SECTION_RECORDS.len()
        && size_count == page_count
        && displayed_count == page_count
}

#[test]
fn f251_oracle_completeness_rejects_extra_or_truncated_pages() {
    assert!(f251_oracle_vectors_are_complete(3, 3, 3));
    assert!(!f251_oracle_vectors_are_complete(4, 4, 4));
    assert!(!f251_oracle_vectors_are_complete(3, 2, 3));
    assert!(!f251_oracle_vectors_are_complete(3, 3, 2));
}

#[test]
fn f251_oracle_artifacts_are_removed_during_unwind() {
    let directory = std::env::temp_dir().join(format!(
        "rdocx-f251-cleanup-{}-{:?}",
        std::process::id(),
        std::thread::current().id()
    ));
    let unwind = std::panic::catch_unwind(|| {
        let artifacts = F251OracleArtifacts::create(directory.clone());
        std::fs::write(artifacts.path().join("partial.pdf"), b"partial").unwrap();
        panic!("exercise F-251 cleanup during unwind");
    });
    assert!(unwind.is_err());
    assert!(!directory.exists());
}

fn f251_section_oracle_source() -> Document {
    let mut document = Document::new();
    document.add_paragraph("first section");
    document.insert_section(1).unwrap();
    document.add_paragraph("second section");
    document.insert_section(2).unwrap();
    document.add_paragraph("third section");
    for (index, (width, height, orientation, start)) in [
        (12240, 15840, ST_PageOrientation::Portrait, 1),
        (15840, 12240, ST_PageOrientation::Landscape, 12),
        (11906, 16838, ST_PageOrientation::Portrait, 27),
    ]
    .into_iter()
    .enumerate()
    {
        let mut section = document.section_mut(index).unwrap();
        section
            .set_page_size(Length::twips(width), Length::twips(height))
            .unwrap();
        section.set_orientation(orientation);
        section.set_page_number_start(start).unwrap();
        section.set_break_type(ST_SectionType::NextPage);
    }
    document.set_raw_header_with_images(
        br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>PAGE=</w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>0</w:t></w:r></w:fldSimple></w:p></w:hdr>"#.to_vec(),
        &[],
        HdrFtrType::Default,
    );
    let header_references = document
        .section(2)
        .unwrap()
        .properties()
        .header_refs
        .clone();
    for index in 0..2 {
        document
            .section_mut(index)
            .unwrap()
            .properties_mut()
            .header_refs = header_references.clone();
    }
    document
}

#[test]
fn mixed_orientation_sections_match_word_geometry_and_page_numbers() {
    assert_eq!(
        WORD_SECTION_ORACLE,
        "Microsoft Word 16.112.3 build 16.112.26083020"
    );
    assert_eq!(
        WORD_SECTION_ENVIRONMENT,
        "macOS 26.6.2 build 25G83; locale=en-GB; normalization=f251-section-pdf-v1; pdftotext=26.09.0; pdfinfo=26.09.0"
    );
    let mut document = f251_section_oracle_source();

    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let geometry = reopened
        .sections()
        .map(|section| {
            let (width, height) = section.page_size().unwrap();
            (width.to_twips(), height.to_twips())
        })
        .collect::<Vec<_>>();
    let page_numbers = reopened
        .sections()
        .map(|section| section.page_number_start().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(geometry, [(12240, 15840), (15840, 12240), (11906, 16838)]);
    assert_eq!(page_numbers, [1, 12, 27]);

    let layout = reopened.layout_deterministic().unwrap();
    let page_geometry = layout
        .layout
        .pages
        .iter()
        .map(|page| (page.width.round() as i32, page.height.round() as i32))
        .collect::<Vec<_>>();
    assert_eq!(page_geometry, [(612, 792), (792, 612), (595, 842)]);
    let physical_page_numbers = layout
        .layout
        .pages
        .iter()
        .map(|page| page.page_number)
        .collect::<Vec<_>>();
    let displayed_page_numbers = layout
        .layout
        .pages
        .iter()
        .map(|page| {
            let mut text = String::new();
            oxml_layout::walk(&page.elements, &mut |element, _| match element {
                oxml_layout::PositionedElement::Text(run) => text.push_str(&run.text),
                oxml_layout::PositionedElement::MultilingualText(run) => {
                    text.push_str(&run.logical_text)
                }
                _ => {}
            });
            let value = text.split_once("PAGE=").unwrap().1;
            value
                .chars()
                .take_while(|character| character.is_ascii_digit())
                .collect::<String>()
                .parse::<u32>()
                .unwrap()
        })
        .collect::<Vec<_>>();
    assert_eq!(physical_page_numbers, [1, 2, 3]);
    assert_eq!(displayed_page_numbers, [1, 12, 27]);
    let subject_records = page_geometry
        .iter()
        .zip(&physical_page_numbers)
        .zip(&displayed_page_numbers)
        .map(|(((width, height), physical), displayed)| {
            format!("page | physical={physical} | size_pt={width}x{height} | PAGE={displayed}")
        })
        .collect::<Vec<_>>();
    assert_eq!(subject_records, WORD_SECTION_RECORDS);
}

#[test]
#[ignore = "requires pinned Microsoft Word 16.112.3 GUI automation and Poppler 26.09.0"]
fn regenerate_f251_word_section_oracle() {
    let plist = "/Applications/Microsoft Word.app/Contents/Info.plist";
    let version = std::process::Command::new("plutil")
        .args(["-extract", "CFBundleShortVersionString", "raw", plist])
        .output()
        .unwrap();
    let build = std::process::Command::new("plutil")
        .args(["-extract", "CFBundleVersion", "raw", plist])
        .output()
        .unwrap();
    assert_eq!(String::from_utf8_lossy(&version.stdout).trim(), "16.112.3");
    assert_eq!(
        String::from_utf8_lossy(&build.stdout).trim(),
        "16.112.26083020"
    );
    assert_eq!(
        WORD_SECTION_ORACLE,
        "Microsoft Word 16.112.3 build 16.112.26083020"
    );
    for command in ["pdftotext", "pdfinfo"] {
        let version = std::process::Command::new(command)
            .arg("-v")
            .output()
            .unwrap();
        assert!(version.status.success());
        let expected = format!("{command} version 26.09.0");
        assert_eq!(
            String::from_utf8_lossy(&version.stderr).lines().next(),
            Some(expected.as_str())
        );
    }

    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let directory = std::path::Path::new(
        "/Users/atulsharma/Library/Containers/com.microsoft.Word/Data/Documents/rdocx-f251-word-oracle",
    )
    .join(format!("{}-{nonce}", std::process::id()));
    let artifacts = F251OracleArtifacts::create(directory.clone());
    let source_path = artifacts.path().join("f251-source.docx");
    let pdf_path = artifacts.path().join("f251-word-render.pdf");
    f251_section_oracle_source().save(&source_path).unwrap();

    let script = format!(
        r#"with timeout of 60 seconds
tell application "Microsoft Word"
activate
open file name "{}" read only true add to recent files false
set f251Doc to document 1
save as f251Doc file name "{}" file format format PDF add to recent files false
close f251Doc saving no
end tell
end timeout"#,
        source_path.display(),
        pdf_path.display(),
    );
    let word = std::process::Command::new("osascript")
        .args(["-e", &script])
        .output()
        .unwrap();
    assert!(
        word.status.success(),
        "Word PDF export failed: {}",
        String::from_utf8_lossy(&word.stderr)
    );

    let text_output = std::process::Command::new("pdftotext")
        .args(["-layout", pdf_path.to_str().unwrap(), "-"])
        .output()
        .unwrap();
    assert!(text_output.status.success());
    let page_text = String::from_utf8(text_output.stdout).unwrap();
    let displayed = page_text
        .split('\u{c}')
        .filter(|page| !page.trim().is_empty())
        .map(|page| {
            let value = page.split_once("PAGE=").expect("Word PAGE field").1;
            value
                .chars()
                .take_while(|character| character.is_ascii_digit())
                .collect::<String>()
                .parse::<u32>()
                .unwrap()
        })
        .collect::<Vec<_>>();
    let summary = std::process::Command::new("pdfinfo")
        .arg(&pdf_path)
        .output()
        .unwrap();
    assert!(summary.status.success());
    let page_count = String::from_utf8(summary.stdout)
        .unwrap()
        .lines()
        .find_map(|line| line.strip_prefix("Pages:").map(str::trim))
        .expect("pdfinfo total page count")
        .parse::<usize>()
        .unwrap();
    assert_eq!(
        page_count,
        WORD_SECTION_RECORDS.len(),
        "Word PDF produced an unexpected number of pages"
    );
    let last_page = page_count.to_string();
    let info = std::process::Command::new("pdfinfo")
        .args(["-f", "1", "-l", &last_page, "-box"])
        .arg(&pdf_path)
        .output()
        .unwrap();
    assert!(info.status.success());
    let sizes = String::from_utf8(info.stdout)
        .unwrap()
        .lines()
        .filter(|line| line.starts_with("Page"))
        .filter_map(|line| line.split_once(" size:").map(|(_, size)| size))
        .map(|size| {
            let mut values = size.split_whitespace();
            let width = values.next().unwrap().parse::<f64>().unwrap().round() as i32;
            assert_eq!(values.next(), Some("x"));
            let height = values.next().unwrap().parse::<f64>().unwrap().round() as i32;
            (width, height)
        })
        .collect::<Vec<_>>();
    assert_eq!(sizes.len(), page_count, "one size is required per PDF page");
    assert_eq!(
        displayed.len(),
        page_count,
        "one displayed PAGE value is required per PDF page"
    );
    assert!(f251_oracle_vectors_are_complete(
        page_count,
        sizes.len(),
        displayed.len()
    ));
    let records = sizes
        .iter()
        .zip(&displayed)
        .enumerate()
        .map(|(index, ((width, height), displayed))| {
            format!(
                "page | physical={} | size_pt={width}x{height} | PAGE={displayed}",
                index + 1
            )
        })
        .collect::<Vec<_>>();
    println!("F-251 Word records\n{}", records.join("\n"));
    assert_eq!(records, WORD_SECTION_RECORDS);
    drop(artifacts);
    assert!(!directory.exists());
}

#[test]
fn ordered_section_mutations_preserve_independent_story_references() {
    let mut document = Document::new();
    document.add_paragraph("alpha");
    document.add_paragraph("beta");
    document.set_header("default header");
    document.set_first_page_header("first header");
    document.set_footer("default footer");
    document.insert_section(0).unwrap();
    document.insert_section(1).unwrap();
    document.insert_section(2).unwrap();

    let (mut headers, mut footers) = {
        let mut final_section = document.section_mut(3).unwrap();
        let final_section = final_section.properties_mut();
        (
            std::mem::take(&mut final_section.header_refs),
            std::mem::take(&mut final_section.footer_refs),
        )
    };
    assert_eq!(headers.len(), 2);
    assert_eq!(footers.len(), 1);
    let default_header = headers.remove(
        headers
            .iter()
            .position(|reference| reference.hdr_ftr_type == HdrFtrType::Default)
            .unwrap(),
    );
    let first_header = headers.pop().unwrap();
    let default_footer = footers.pop().unwrap();
    let expected_relationship_ids = [
        default_header.rel_id.clone(),
        first_header.rel_id.clone(),
        default_footer.rel_id.clone(),
    ];
    document
        .section_mut(0)
        .unwrap()
        .properties_mut()
        .header_refs
        .push(default_header);
    document
        .section_mut(2)
        .unwrap()
        .properties_mut()
        .header_refs
        .push(first_header);
    document
        .section_mut(3)
        .unwrap()
        .properties_mut()
        .footer_refs
        .push(default_footer);
    for index in 0..document.section_count() {
        let mut section = document.section_mut(index).unwrap();
        let section = section.properties_mut();
        section
            .extra_xml
            .push(format!(r#"<x:section xmlns:x="urn:producer" x:id="{index}"/>"#).into_bytes());
    }
    document
        .section_mut(0)
        .unwrap()
        .set_orientation(ST_PageOrientation::Portrait);
    document
        .section_mut(2)
        .unwrap()
        .set_orientation(ST_PageOrientation::Landscape);
    document
        .section_mut(3)
        .unwrap()
        .set_orientation(ST_PageOrientation::Portrait);

    document.remove_section(1).unwrap();
    assert_eq!(document.section_count(), 3);
    assert_eq!(
        document
            .text()
            .lines()
            .filter(|line| !line.is_empty())
            .collect::<Vec<_>>(),
        ["alpha", "beta"]
    );

    let bytes = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.section_count(), 3);
    assert_eq!(
        reopened
            .sections()
            .map(|section| {
                section
                    .orientation()
                    .unwrap_or(ST_PageOrientation::Portrait)
            })
            .collect::<Vec<_>>(),
        [
            ST_PageOrientation::Portrait,
            ST_PageOrientation::Landscape,
            ST_PageOrientation::Portrait,
        ]
    );
    for (index, section) in reopened.sections().enumerate() {
        let source_index = [0, 2, 3][index];
        assert_eq!(
            section.properties().extra_xml,
            [
                format!(r#"<x:section xmlns:x="urn:producer" x:id="{source_index}"/>"#)
                    .into_bytes()
            ]
        );
    }
    let package = OpcPackage::from_reader(std::io::Cursor::new(&bytes)).unwrap();
    let relationships = package.get_part_rels("/word/document.xml").unwrap();
    assert_eq!(
        reopened
            .sections()
            .map(|section| {
                (
                    section.properties().header_refs.len(),
                    section.properties().footer_refs.len(),
                )
            })
            .collect::<Vec<_>>(),
        [(1, 0), (2, 0), (0, 1)]
    );
    for (index, is_header, kind, relationship_id, target, text) in [
        (
            0,
            true,
            HdrFtrType::Default,
            expected_relationship_ids[0].as_str(),
            "/word/header1.xml",
            "default header",
        ),
        (
            1,
            true,
            HdrFtrType::Default,
            expected_relationship_ids[0].as_str(),
            "/word/header1.xml",
            "default header",
        ),
        (
            1,
            true,
            HdrFtrType::First,
            expected_relationship_ids[1].as_str(),
            "/word/headerFirst1.xml",
            "first header",
        ),
        (
            2,
            false,
            HdrFtrType::Default,
            expected_relationship_ids[2].as_str(),
            "/word/footer1.xml",
            "default footer",
        ),
    ] {
        let section = reopened.section(index).unwrap();
        let references = if is_header {
            &section.properties().header_refs
        } else {
            &section.properties().footer_refs
        };
        assert!(references.iter().any(|reference| {
            reference.hdr_ftr_type == kind && reference.rel_id == relationship_id
        }));
        let relationship = relationships.get_by_id(relationship_id).unwrap();
        assert_eq!(
            relationship.rel_type,
            if is_header {
                rel_types::HEADER
            } else {
                rel_types::FOOTER
            }
        );
        let resolved = OpcPackage::resolve_rel_target("/word/document.xml", &relationship.target);
        assert_eq!(resolved, target);
        assert_eq!(
            CT_HdrFtr::from_xml(package.get_part(&resolved).unwrap())
                .unwrap()
                .text(),
            text
        );
    }
}

#[test]
fn removing_a_predecessor_materializes_inherited_header_and_footer_references() {
    let mut document = Document::new();
    document.set_header("inherited default header");
    document.set_first_page_header("inherited first header");
    document.set_footer("inherited default footer");
    document.set_first_page_footer("inherited first footer");
    document.insert_section(0).unwrap();

    let (headers, footers) = {
        let mut following = document.section_mut(1).unwrap();
        let following = following.properties_mut();
        (
            std::mem::take(&mut following.header_refs),
            std::mem::take(&mut following.footer_refs),
        )
    };
    document
        .section_mut(0)
        .unwrap()
        .properties_mut()
        .header_refs = headers.clone();
    document
        .section_mut(0)
        .unwrap()
        .properties_mut()
        .footer_refs = footers.clone();
    assert!(
        document
            .section(1)
            .unwrap()
            .properties()
            .header_refs
            .is_empty()
    );
    assert!(
        document
            .section(1)
            .unwrap()
            .properties()
            .footer_refs
            .is_empty()
    );

    document.remove_section(0).unwrap();
    assert_eq!(
        document.section(0).unwrap().properties().header_refs,
        headers
    );
    assert_eq!(
        document.section(0).unwrap().properties().footer_refs,
        footers
    );

    let bytes = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(&bytes)).unwrap();
    let relationships = package.get_part_rels("/word/document.xml").unwrap();
    for (is_header, references, expected) in [
        (
            true,
            &reopened.section(0).unwrap().properties().header_refs,
            [
                (
                    HdrFtrType::Default,
                    "/word/header1.xml",
                    "inherited default header",
                ),
                (
                    HdrFtrType::First,
                    "/word/headerFirst1.xml",
                    "inherited first header",
                ),
            ],
        ),
        (
            false,
            &reopened.section(0).unwrap().properties().footer_refs,
            [
                (
                    HdrFtrType::Default,
                    "/word/footer1.xml",
                    "inherited default footer",
                ),
                (
                    HdrFtrType::First,
                    "/word/footerFirst1.xml",
                    "inherited first footer",
                ),
            ],
        ),
    ] {
        assert_eq!(references.len(), expected.len());
        for (kind, target, text) in expected {
            let reference = references
                .iter()
                .find(|reference| reference.hdr_ftr_type == kind)
                .unwrap();
            let relationship = relationships.get_by_id(&reference.rel_id).unwrap();
            assert_eq!(
                relationship.rel_type,
                if is_header {
                    rel_types::HEADER
                } else {
                    rel_types::FOOTER
                }
            );
            let resolved =
                OpcPackage::resolve_rel_target("/word/document.xml", &relationship.target);
            assert_eq!(resolved, target);
            assert_eq!(
                CT_HdrFtr::from_xml(package.get_part(&resolved).unwrap())
                    .unwrap()
                    .text(),
                text
            );
        }
    }
}

#[test]
fn removing_a_section_never_orphans_a_shared_story() {
    let mut document = Document::new();
    document.set_header("shared header");
    document.insert_section(0).unwrap();
    document.insert_section(1).unwrap();
    let shared = document.section(2).unwrap().properties().header_refs[0].clone();
    document
        .section_mut(0)
        .unwrap()
        .properties_mut()
        .header_refs
        .push(shared);

    document.remove_section(0).unwrap();
    assert_eq!(
        document
            .stories()
            .unwrap()
            .into_iter()
            .filter(|story| story.kind() == StoryKind::Header)
            .count(),
        1
    );

    document
        .section_mut(0)
        .unwrap()
        .properties_mut()
        .header_refs
        .clear();
    document.remove_section(1).unwrap();
    assert_eq!(document.section_count(), 1);
    assert!(
        document
            .stories()
            .unwrap()
            .into_iter()
            .all(|story| story.kind() != StoryKind::Header)
    );
    let owned_package =
        OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
    assert!(owned_package.get_part("/word/header1.xml").is_none());
    assert!(
        owned_package
            .get_part_rels("/word/document.xml")
            .unwrap()
            .items
            .iter()
            .all(|relationship| relationship.rel_type != rel_types::HEADER)
    );
    let before = document.to_bytes().unwrap();
    assert!(document.remove_section(0).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);

    let mut imported = container_neutral_story_fixture();
    imported.insert_section(0).unwrap();
    imported.remove_section(1).unwrap();
    let imported_package =
        OpcPackage::from_reader(std::io::Cursor::new(imported.to_bytes().unwrap())).unwrap();
    for part_name in ["/word/header-story.xml", "/word/footer-story.xml"] {
        assert!(imported_package.get_part(part_name).is_some());
    }
    let imported_relationships = imported_package
        .get_part_rels("/word/document.xml")
        .unwrap();
    assert!(imported_relationships.get_by_id("storyHeader").is_some());
    assert!(imported_relationships.get_by_id("storyFooter").is_some());
}

#[test]
fn chart_rgb_colour_is_reexported_by_all_three_facades() {
    let shared = oxml_chart::RgbColor::new(0x2B, 0x6F, 0xE3);
    let word: rdocx::RgbColor = shared;
    let presentation: rpptx::RgbColor = word;
    assert_eq!(presentation, shared);
}

mod fresh_word_package_profile_tests {
    use super::*;
    use oxml_opc::content_types;

    const STYLES_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    const SETTINGS_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    const FONT_TABLE_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";

    fn all_classes() -> [WordPackageClass; 4] {
        [
            WordPackageClass::Document,
            WordPackageClass::MacroEnabledDocument,
            WordPackageClass::Template,
            WordPackageClass::MacroEnabledTemplate,
        ]
    }

    fn main_content_type(class: WordPackageClass) -> &'static str {
        match class {
            WordPackageClass::Document => content_types::WORD_DOCUMENT,
            WordPackageClass::MacroEnabledDocument => content_types::WORD_DOCUMENT_MACRO_ENABLED,
            WordPackageClass::Template => content_types::WORD_TEMPLATE,
            WordPackageClass::MacroEnabledTemplate => content_types::WORD_TEMPLATE_MACRO_ENABLED,
        }
    }

    fn package_from_profile(profile: WordCreationProfile) -> OpcPackage {
        let mut document = Document::new_with_profile(profile);
        OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap()
    }

    fn package_bytes(package: &OpcPackage) -> Vec<u8> {
        let mut output = std::io::Cursor::new(Vec::new());
        package.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn relationship_types(package: &OpcPackage, source: &str) -> Vec<String> {
        let relationships = if source == "/" {
            &package.package_rels
        } else {
            package.get_part_rels(source).unwrap()
        };
        let mut types: Vec<_> = relationships
            .items
            .iter()
            .map(|relationship| relationship.rel_type.clone())
            .collect();
        types.sort_unstable();
        types
    }

    fn assert_internal_targets_exist(package: &OpcPackage) {
        for (source, relationships) in std::iter::once(("/", &package.package_rels)).chain(
            package
                .part_rels
                .iter()
                .map(|(source, relationships)| (source.as_str(), relationships)),
        ) {
            let mut ids = std::collections::HashSet::new();
            for relationship in &relationships.items {
                assert!(ids.insert(&relationship.id), "duplicate relationship ID");
                if relationship.target_mode.as_deref() != Some("External") {
                    let target = OpcPackage::resolve_rel_target(source, &relationship.target);
                    assert!(
                        package.get_part(&target).is_some(),
                        "{source} targets missing part {target}"
                    );
                }
            }
        }
    }

    #[test]
    fn word_compatible_profiles_reopen_with_the_same_package_class() {
        for class in all_classes() {
            let mut package = package_from_profile(WordCreationProfile::WordCompatible(class));
            const PROBE: &[u8] =
                br#"<probe xmlns="urn:rdocx:f243"><opaque keep="exact"> bytes </opaque></probe>"#;
            package.set_part("/custom/profile-probe.xml", PROBE.to_vec());
            package
                .content_types
                .add_override("/custom/profile-probe.xml", "application/xml");
            package.package_rels.add_with_id(
                "profileProbe",
                "urn:rdocx:relationships/profile-probe",
                "custom/profile-probe.xml",
            );
            let mut reopened =
                Document::from_bytes(&package_bytes(&package)).expect("compatible profile reopens");
            assert_eq!(reopened.package_class().unwrap(), class);
            let saved = reopened.to_bytes().expect("reopened profile serializes");
            let saved = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
            assert_eq!(saved.get_part("/custom/profile-probe.xml"), Some(PROBE));
            let relationship = saved
                .package_rels
                .get_by_id("profileProbe")
                .expect("unmodelled relationship survives");
            assert_eq!(
                relationship.rel_type,
                "urn:rdocx:relationships/profile-probe"
            );
            assert_eq!(relationship.target, "custom/profile-probe.xml");
        }
    }

    #[test]
    fn word_compatible_profiles_have_a_complete_normalized_package_graph() {
        let required_parts = [
            "/docProps/app.xml",
            "/docProps/core.xml",
            "/word/document.xml",
            "/word/fontTable.xml",
            "/word/settings.xml",
            "/word/styles.xml",
            "/word/theme/theme1.xml",
        ];
        for class in all_classes() {
            let package = package_from_profile(WordCreationProfile::WordCompatible(class));
            let mut parts: Vec<_> = package.parts.keys().map(String::as_str).collect();
            parts.sort_unstable();
            assert_eq!(parts, required_parts);
            assert_eq!(package.content_types.overrides.len(), required_parts.len());
            assert_eq!(
                package.content_types.override_for("/word/document.xml"),
                Some(main_content_type(class))
            );
            for (part_name, expected) in [
                ("/word/styles.xml", STYLES_CONTENT_TYPE),
                ("/word/settings.xml", SETTINGS_CONTENT_TYPE),
                ("/word/fontTable.xml", FONT_TABLE_CONTENT_TYPE),
                ("/word/theme/theme1.xml", content_types::THEME),
                ("/docProps/core.xml", content_types::CORE_PROPERTIES),
                ("/docProps/app.xml", content_types::EXTENDED_PROPERTIES),
            ] {
                assert_eq!(
                    package.content_types.override_for(part_name),
                    Some(expected)
                );
            }
            assert_eq!(
                relationship_types(&package, "/"),
                [
                    rel_types::EXTENDED_PROPERTIES,
                    rel_types::DOCUMENT,
                    rel_types::CORE_PROPERTIES,
                ]
            );
            assert_eq!(
                relationship_types(&package, "/word/document.xml"),
                [
                    rel_types::FONT_TABLE,
                    rel_types::SETTINGS,
                    rel_types::STYLES,
                    rel_types::THEME,
                ]
            );
            assert_internal_targets_exist(&package);
            assert!(
                package
                    .package_rels
                    .get_by_type(rel_types::VBA_PROJECT)
                    .is_none()
            );
            assert!(
                package
                    .get_part_rels("/word/document.xml")
                    .unwrap()
                    .get_by_type(rel_types::VBA_PROJECT)
                    .is_none()
            );
            let core =
                std::str::from_utf8(package.get_part("/docProps/core.xml").unwrap()).unwrap();
            assert!(!core.contains("dcterms:created"));
            assert!(!core.contains("dcterms:modified"));
        }
    }

    #[test]
    fn document_new_uses_the_word_compatible_docx_profile() {
        let default = Document::new().to_bytes().unwrap();
        let compatible = Document::new_with_profile(WordCreationProfile::WordCompatible(
            WordPackageClass::Document,
        ))
        .to_bytes()
        .unwrap();
        let minimal =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document))
                .to_bytes()
                .unwrap();
        assert_eq!(default, compatible);
        assert_ne!(minimal, compatible);
        let minimal = OpcPackage::from_reader(std::io::Cursor::new(minimal)).unwrap();
        let mut minimal_parts: Vec<_> = minimal.parts.keys().map(String::as_str).collect();
        minimal_parts.sort_unstable();
        assert_eq!(minimal_parts, ["/word/document.xml", "/word/styles.xml"]);
        assert_eq!(relationship_types(&minimal, "/"), [rel_types::DOCUMENT]);
        assert_eq!(
            relationship_types(&minimal, "/word/document.xml"),
            [rel_types::STYLES]
        );
    }

    #[test]
    fn equivalent_fresh_profiles_serialize_identically() {
        for class in all_classes() {
            let mut first = Document::new_with_profile(WordCreationProfile::WordCompatible(class));
            let mut second = Document::new_with_profile(WordCreationProfile::WordCompatible(class));
            let first_bytes = first.to_bytes().unwrap();
            assert_eq!(first_bytes, second.to_bytes().unwrap());
            assert_eq!(
                first_bytes,
                Document::from_bytes(&first_bytes)
                    .unwrap()
                    .to_bytes()
                    .unwrap()
            );
        }
    }
}

mod theme_and_embedded_font_tests {
    use super::*;
    use oxml_drawing::color::ColorChoice;
    use quick_xml::events::BytesStart;
    use rdocx::{
        CT_OfficeStyleSheet, EmbeddedFont, EmbeddedFontKind, FontDefinition, FontEmbeddingLicense,
        ThemeFontLanguage,
    };

    const FONT_KEY: &str = "{00112233-4455-6677-8899-AABBCCDDEEFF}";
    const FONT_REL_TYPE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
    const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const WORD_THEME_COLOR_ORACLE: [&str; 12] = [
        "030330", "FFFFFF", "030330", "FFFFFF", "D9DBFF", "FF0054", "9482FF", "4A458C", "E0DED9",
        "030330", "9482FF", "FF0054",
    ];
    const CARLITO: &[u8] = include_bytes!("../../oxml-layout/fonts/Carlito-Regular.ttf");
    const CALADEA: &[u8] = include_bytes!("../../oxml-layout/fonts/Caladea-Regular.ttf");

    fn license(authorized: bool, identity: &str) -> FontEmbeddingLicense {
        FontEmbeddingLicense::new(authorized, identity)
    }

    fn corpus_font() -> FontDefinition {
        FontDefinition::new("Corpus Sans")
            .with_alternate_name("Carlito")
            .with_family("swiss")
            .with_pitch("variable")
    }

    fn embedded_font() -> EmbeddedFont {
        EmbeddedFont::new(
            EmbeddedFontKind::Regular,
            CARLITO.to_vec(),
            FONT_KEY,
            license(true, "Apache-2.0: Carlito"),
        )
    }

    fn srgb_color(value: &str) -> ColorChoice {
        let mut element = BytesStart::new("a:srgbClr");
        element.push_attribute(("val", value));
        ColorChoice::from_empty_xml(&element).unwrap()
    }

    fn apply_word_color_projection(theme: &mut CT_OfficeStyleSheet) {
        let colors = &mut theme.theme_elements.color_scheme;
        colors.dark1 = srgb_color(WORD_THEME_COLOR_ORACLE[0]);
        colors.light1 = srgb_color(WORD_THEME_COLOR_ORACLE[1]);
        colors.dark2 = srgb_color(WORD_THEME_COLOR_ORACLE[2]);
        colors.light2 = srgb_color(WORD_THEME_COLOR_ORACLE[3]);
        colors.accent1 = srgb_color(WORD_THEME_COLOR_ORACLE[4]);
        colors.accent2 = srgb_color(WORD_THEME_COLOR_ORACLE[5]);
        colors.accent3 = srgb_color(WORD_THEME_COLOR_ORACLE[6]);
        colors.accent4 = srgb_color(WORD_THEME_COLOR_ORACLE[7]);
        colors.accent5 = srgb_color(WORD_THEME_COLOR_ORACLE[8]);
        colors.accent6 = srgb_color(WORD_THEME_COLOR_ORACLE[9]);
        colors.hyperlink = srgb_color(WORD_THEME_COLOR_ORACLE[10]);
        colors.followed_hyperlink = srgb_color(WORD_THEME_COLOR_ORACLE[11]);
    }

    fn theme_color_projection(theme: &CT_OfficeStyleSheet) -> Vec<String> {
        theme
            .theme_elements
            .color_scheme
            .iter()
            .map(|(_, color)| match color {
                ColorChoice::Srgb { value, .. } => value.to_string(),
                ColorChoice::System {
                    last_color: Some(value),
                    ..
                } => value.to_string(),
                other => panic!("unexpected theme color in Word projection: {other:?}"),
            })
            .collect()
    }

    fn package_bytes(package: &OpcPackage) -> Vec<u8> {
        let mut output = std::io::Cursor::new(Vec::new());
        package.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn authored_document() -> Document {
        let mut document = Document::new();
        let mut theme = CT_OfficeStyleSheet::office_default();
        theme.name = Some("Corpus Theme".to_owned());
        apply_word_color_projection(&mut theme);
        theme.theme_elements.font_scheme.major_font.latin.typeface = "Corpus Sans".to_owned();
        theme.theme_elements.font_scheme.minor_font.latin.typeface = "Corpus Sans".to_owned();
        document.set_theme(theme).unwrap();
        document
            .set_language_defaults(ThemeFontLanguage {
                latin: Some("en-GB".to_owned()),
                east_asia: Some("zh-CN".to_owned()),
                bidi: Some("ar-SA".to_owned()),
            })
            .unwrap();
        document.set_font(corpus_font()).unwrap();
        document.embed_font("Corpus Sans", embedded_font()).unwrap();
        document
    }

    #[test]
    fn authored_theme_font_table_and_embedded_fonts_survive_reopen() {
        let mut document = authored_document();
        let expected_theme = document.theme().unwrap().clone();
        let expected_languages = document.theme_font_language().unwrap().clone();
        let expected_fonts = document.fonts();

        let mut reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();

        assert_eq!(reopened.theme(), Some(&expected_theme));
        assert_eq!(reopened.theme_font_language(), Some(&expected_languages));
        assert_eq!(reopened.fonts(), expected_fonts);
        let corpus = reopened
            .fonts()
            .into_iter()
            .find(|font| font.name == "Corpus Sans")
            .unwrap();
        assert_eq!(corpus.embedded_fonts[0].data, CARLITO);

        assert_eq!(
            reopened
                .remove_embedded_font("Corpus Sans", EmbeddedFontKind::Regular)
                .unwrap(),
            Some(embedded_font())
        );
        assert!(
            reopened
                .fonts()
                .into_iter()
                .find(|font| font.name == "Corpus Sans")
                .unwrap()
                .embedded_fonts
                .is_empty()
        );
        reopened.embed_font("Corpus Sans", embedded_font()).unwrap();
        assert!(reopened.remove_font("Corpus Sans").unwrap().is_some());
        assert!(
            reopened
                .fonts()
                .into_iter()
                .all(|font| font.name != "Corpus Sans")
        );
    }

    #[test]
    fn font_embedding_requires_explicit_authorization_and_license_identity() {
        let mut document = Document::new();
        document.set_font(corpus_font()).unwrap();
        let before = document.to_bytes().unwrap();

        let denied = EmbeddedFont::new(
            EmbeddedFontKind::Regular,
            CARLITO.to_vec(),
            FONT_KEY,
            license(false, "Apache-2.0: Carlito"),
        );
        assert!(document.embed_font("Corpus Sans", denied).is_err());
        assert_eq!(document.to_bytes().unwrap(), before);

        let unidentified = EmbeddedFont::new(
            EmbeddedFontKind::Regular,
            CARLITO.to_vec(),
            FONT_KEY,
            license(true, ""),
        );
        assert!(document.embed_font("Corpus Sans", unidentified).is_err());
        assert_eq!(document.to_bytes().unwrap(), before);

        let invalid_key = EmbeddedFont::new(
            EmbeddedFontKind::Regular,
            CARLITO.to_vec(),
            "00112233-4455-6677-8899-AABBCCDDEEFF",
            license(true, "Apache-2.0: Carlito"),
        );
        assert!(document.embed_font("Corpus Sans", invalid_key).is_err());
        assert_eq!(document.to_bytes().unwrap(), before);

        let mut invalid_family = corpus_font();
        invalid_family.family = Some("unknown".to_owned());
        assert!(document.set_font(invalid_family).is_err());
        assert_eq!(document.to_bytes().unwrap(), before);

        let mut invalid_pitch = corpus_font();
        invalid_pitch.pitch = Some("wide".to_owned());
        assert!(document.set_font(invalid_pitch).is_err());
        assert_eq!(document.to_bytes().unwrap(), before);

        let normalized_identity = EmbeddedFont::new(
            EmbeddedFontKind::Regular,
            CARLITO.to_vec(),
            FONT_KEY,
            license(true, "Apache-2.0\nCarlito"),
        );
        assert!(
            document
                .embed_font("Corpus Sans", normalized_identity)
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), before);

        let exact_identity = "Apache-2.0 & \"Carlito\"";
        document
            .embed_font(
                "Corpus Sans",
                EmbeddedFont::new(
                    EmbeddedFontKind::Regular,
                    CARLITO.to_vec(),
                    FONT_KEY,
                    license(true, exact_identity),
                ),
            )
            .unwrap();
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .fonts()
                .into_iter()
                .find(|font| font.name == "Corpus Sans")
                .unwrap()
                .embedded_fonts[0]
                .license
                .identity,
            exact_identity
        );
    }

    #[test]
    fn font_table_preserves_unknown_children_and_relationship_attributes() {
        let mut document = Document::new();
        let bytes = document.to_bytes().unwrap();
        let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><q:fonts xmlns:q="{WORD_NS}" xmlns:r="{REL_NS}" xmlns:x="urn:producer" xmlns:mc="{MC_NS}" xmlns:w15="urn:producer:w15" mc:Ignorable="w15"><!--root-before--><?root kept?> <x:before x:keep="exact"/><q:font q:name="Corpus Sans"><!--font-before--><?font kept?> <x:fontBefore x:keep="exact"/><q:altName q:val="Original" x:property="kept &amp; &quot;quoted&quot;"> <!--property-note--><?producer kept?></q:altName><q:panose1 q:val="020B0604020202020204"/><x:middle x:keep="exact"/><q:family q:val="swiss" x:property="family-kept"/><q:pitch q:val="variable" x:property="pitch-kept"/><q:embedRegular r:id="producerFont" q:fontKey="{FONT_KEY}" x:keep="relationship-attribute"><!--relationship-note--></q:embedRegular><x:after x:keep="exact"/></q:font><x:between x:keep="exact"/><q:font q:name="Keep Sans"><q:family q:val="swiss"/></q:font><x:tail x:keep="exact"/></q:fonts>"#
        );
        package.set_part("/word/fontTable.xml", xml.into_bytes());
        package.set_part("/word/fonts/producer.odttf", vec![7; 64]);
        package.content_types.add_override(
            "/word/fonts/producer.odttf",
            "application/vnd.openxmlformats-officedocument.obfuscatedFont",
        );
        package
            .get_or_create_part_rels("/word/fontTable.xml")
            .add_with_id("producerFont", FONT_REL_TYPE, "fonts/producer.odttf");

        let mut document = Document::from_bytes(&package_bytes(&package)).unwrap();
        let parsed = document.fonts();
        let corpus = parsed
            .iter()
            .find(|font| font.name == "Corpus Sans")
            .unwrap();
        assert_eq!(corpus.alternate_name.as_deref(), Some("Original"));
        assert_eq!(corpus.embedded_fonts.len(), 1);
        document
            .set_font(corpus_font().with_alternate_name("Updated"))
            .unwrap();
        let saved =
            OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
        let output = std::str::from_utf8(saved.get_part("/word/fontTable.xml").unwrap()).unwrap();

        for fragment in [
            r#"<x:before x:keep="exact"/>"#,
            r#"<!--root-before--><?root kept?> "#,
            r#"<!--font-before--><?font kept?> "#,
            r#"<x:fontBefore x:keep="exact"/>"#,
            r#"<x:middle x:keep="exact"/>"#,
            r#"x:property="kept &amp; &quot;quoted&quot;""#,
            r#"x:property="family-kept""#,
            r#"x:property="pitch-kept""#,
            r#"x:keep="relationship-attribute""#,
            r#"<q:panose1 q:val="020B0604020202020204"/>"#,
            r#"<!--property-note-->"#,
            r#"<?producer kept?>"#,
            r#"<!--relationship-note-->"#,
            r#"<x:after x:keep="exact"/>"#,
            r#"<x:between x:keep="exact"/>"#,
            r#"<x:tail x:keep="exact"/>"#,
        ] {
            assert!(output.contains(fragment), "missing {fragment} in {output}");
        }
        assert!(output.contains(&format!(r#"xmlns:q="{WORD_NS}""#)));
        assert!(output.contains(&format!(r#"xmlns:mc="{MC_NS}""#)));
        assert!(output.contains(r#"mc:Ignorable="w15 rdocx""#));
        assert!(output.contains(
            r#"<w:altName w:val="Updated" x:property="kept &amp; &quot;quoted&quot;"> <!--property-note--><?producer kept?></w:altName>"#
        ));
        assert!(output.find("fontBefore").unwrap() < output.find("altName").unwrap());
        assert!(output.find("altName").unwrap() < output.find("embedRegular").unwrap());

        let mut reopened = Document::from_bytes(&package_bytes(&saved)).unwrap();
        assert!(reopened.remove_font("Corpus Sans").unwrap().is_some());
        let removed =
            OpcPackage::from_reader(std::io::Cursor::new(reopened.to_bytes().unwrap())).unwrap();
        let output = std::str::from_utf8(removed.get_part("/word/fontTable.xml").unwrap()).unwrap();
        assert!(output.contains(r#"<x:before x:keep="exact"/>"#));
        assert!(output.contains(r#"<x:between x:keep="exact"/>"#));
        assert!(output.contains(r#"<x:tail x:keep="exact"/>"#));
        assert!(output.find("before").unwrap() < output.find("between").unwrap());
        assert!(output.find("between").unwrap() < output.find("Keep Sans").unwrap());
        assert!(output.find("Keep Sans").unwrap() < output.find("tail").unwrap());

        let conflicting_xml = format!(
            r#"<q:fonts xmlns:q="{WORD_NS}" xmlns:w="urn:producer"><q:font q:name="Unsafe"/></q:fonts>"#
        );
        let mut conflicting_package =
            OpcPackage::from_reader(std::io::Cursor::new(Document::new().to_bytes().unwrap()))
                .unwrap();
        conflicting_package.set_part("/word/fontTable.xml", conflicting_xml.into_bytes());
        let mut conflicting = Document::from_bytes(&package_bytes(&conflicting_package)).unwrap();
        let before = conflicting.to_bytes().unwrap();
        assert!(conflicting.set_font(corpus_font()).is_err());
        assert_eq!(conflicting.to_bytes().unwrap(), before);
    }

    #[test]
    fn public_authored_theme_and_fonts_match_pinned_word_resolution() {
        const WORD_ORACLE: &str = "Microsoft Word 16.104 build 16.104.25121423";
        const LIBREOFFICE_ORACLE: &str =
            "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb";
        const WORD_FONT_TABLE_ORACLE: &str = concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
            r#"<w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/>"#,
            r#"<w:charset w:val="00"/><w:family w:val="swiss"/>"#,
            r#"<w:pitch w:val="variable"/><w:sig w:usb0="E0002EFF" "#,
            r#"w:usb1="C000785B" w:usb2="00000009" w:usb3="00000000" "#,
            r#"w:csb0="000001FF" w:csb1="00000000"/></w:font></w:fonts>"#,
        );
        assert_eq!(WORD_ORACLE, MHTML_ORACLE_VERSION);
        assert_eq!(LIBREOFFICE_ORACLE, ODT_ORACLE_VERSION);

        let version = std::process::Command::new("soffice")
            .arg("--version")
            .output()
            .expect("pinned LibreOffice is installed");
        assert!(version.status.success());
        assert_eq!(
            String::from_utf8_lossy(&version.stdout).trim(),
            LIBREOFFICE_ORACLE
        );

        let mut authored = authored_document();
        authored
            .add_paragraph("")
            .add_run("Pinned theme and font resolution")
            .font("Corpus Sans")
            .size(18.0);

        let theme = authored.theme().unwrap();
        let font = authored
            .fonts()
            .into_iter()
            .find(|font| font.name == "Corpus Sans")
            .unwrap();
        assert_eq!(
            theme.theme_elements.font_scheme.major_font.latin.typeface,
            "Corpus Sans"
        );
        assert_eq!(
            theme.theme_elements.font_scheme.minor_font.latin.typeface,
            "Corpus Sans"
        );
        assert_eq!(
            theme_color_projection(theme),
            WORD_THEME_COLOR_ORACLE.map(str::to_owned)
        );
        let reopened_theme = Document::from_bytes(&authored.to_bytes().unwrap()).unwrap();
        assert_eq!(
            theme_color_projection(reopened_theme.theme().unwrap()),
            WORD_THEME_COLOR_ORACLE.map(str::to_owned)
        );
        let mut word_package =
            OpcPackage::from_reader(std::io::Cursor::new(Document::new().to_bytes().unwrap()))
                .unwrap();
        word_package.set_part(
            "/word/fontTable.xml",
            WORD_FONT_TABLE_ORACLE.as_bytes().to_vec(),
        );
        let word_oracle = Document::from_bytes(&package_bytes(&word_package)).unwrap();
        let word_font = word_oracle
            .fonts()
            .into_iter()
            .find(|font| font.name == "Arial")
            .unwrap();
        assert_eq!(word_font.family.as_deref(), Some("swiss"));
        assert_eq!(word_font.pitch.as_deref(), Some("variable"));
        assert_eq!(font.family, word_font.family);
        assert_eq!(font.pitch, word_font.pitch);

        let mut oracle = Document::new();
        oracle
            .add_paragraph("")
            .add_run("Pinned theme and font resolution")
            .font("Carlito")
            .size(18.0);

        let authored_render = authored.render_page_to_png_deterministic(0, 150.0).unwrap();
        assert_eq!(
            authored_render,
            oracle.render_page_to_png_deterministic(0, 150.0).unwrap()
        );

        let root = std::env::temp_dir().join(format!(
            "rdocx-theme-font-oracle-{}-{}",
            std::process::id(),
            std::thread::current().name().unwrap_or("test")
        ));
        let _ = std::fs::remove_dir_all(&root);
        let output = root.join("output");
        let profile = root.join("profile");
        std::fs::create_dir_all(&output).unwrap();
        std::fs::create_dir_all(&profile).unwrap();
        let source = root.join("source.docx");
        std::fs::write(&source, authored.to_bytes().unwrap()).unwrap();
        let status = std::process::Command::new("soffice")
            .arg("--headless")
            .arg(format!(
                "-env:UserInstallation=file://{}",
                profile.display()
            ))
            .arg("--convert-to")
            .arg("docx")
            .arg("--outdir")
            .arg(&output)
            .arg(&source)
            .status()
            .expect("LibreOffice conversion starts");
        assert!(status.success());
        let normalized = Document::open(output.join("source.docx")).unwrap();
        assert_eq!(
            authored_render,
            normalized
                .render_page_to_png_deterministic(0, 150.0)
                .unwrap()
        );
        let _ = std::fs::remove_dir_all(&root);
    }

    #[test]
    fn embedded_font_parts_are_packaged_deterministically() {
        let mut first = authored_document();
        let mut second = authored_document();
        let first_bytes = first.to_bytes().unwrap();

        assert_eq!(first_bytes, first.to_bytes().unwrap());
        assert_eq!(first_bytes, second.to_bytes().unwrap());

        let package = OpcPackage::from_reader(std::io::Cursor::new(first_bytes)).unwrap();
        let font_parts: Vec<_> = package
            .parts
            .keys()
            .filter(|part| part.starts_with("/word/fonts/"))
            .collect();
        assert_eq!(font_parts.len(), 1);
        let relationships = package.get_part_rels("/word/fontTable.xml").unwrap();
        assert_eq!(
            relationships
                .items
                .iter()
                .filter(|relationship| relationship.rel_type == FONT_REL_TYPE)
                .count(),
            1
        );

        let original_relationship = relationships
            .items
            .iter()
            .find(|relationship| relationship.rel_type == FONT_REL_TYPE)
            .unwrap();
        let original_id = original_relationship.id.clone();
        let original_target = original_relationship.target.clone();
        let mut package = package;
        package
            .get_or_create_part_rels("/word/fontTable.xml")
            .add_with_id("producerShared", FONT_REL_TYPE, &original_target);
        let font_table = std::str::from_utf8(package.get_part("/word/fontTable.xml").unwrap())
            .unwrap()
            .replace(
                "</w:fonts>",
                &format!(
                    r#"<w:font w:name="Shared Same Id"><w:embedRegular r:id="{original_id}" w:fontKey="{FONT_KEY}"/></w:font><w:font w:name="Shared Same Part"><w:embedRegular r:id="producerShared" w:fontKey="{FONT_KEY}"/></w:font></w:fonts>"#
                ),
            );
        package.set_part("/word/fontTable.xml", font_table.into_bytes());

        let mut shared = Document::from_bytes(&package_bytes(&package)).unwrap();
        shared
            .remove_embedded_font("Corpus Sans", EmbeddedFontKind::Regular)
            .unwrap();
        assert_eq!(
            shared
                .fonts()
                .into_iter()
                .find(|font| font.name == "Shared Same Id")
                .unwrap()
                .embedded_fonts[0]
                .data,
            CARLITO
        );
        shared
            .embed_font(
                "Shared Same Id",
                EmbeddedFont::new(
                    EmbeddedFontKind::Regular,
                    CALADEA.to_vec(),
                    FONT_KEY,
                    license(true, "Apache-2.0: Caladea"),
                ),
            )
            .unwrap();
        let fonts = shared.fonts();
        assert_eq!(
            fonts
                .iter()
                .find(|font| font.name == "Shared Same Id")
                .unwrap()
                .embedded_fonts[0]
                .data,
            CALADEA
        );
        assert_eq!(
            fonts
                .iter()
                .find(|font| font.name == "Shared Same Part")
                .unwrap()
                .embedded_fonts[0]
                .data,
            CARLITO
        );
        let shared_package =
            OpcPackage::from_reader(std::io::Cursor::new(shared.to_bytes().unwrap())).unwrap();
        assert_eq!(
            shared_package
                .get_part_rels("/word/fontTable.xml")
                .unwrap()
                .items
                .iter()
                .filter(|relationship| relationship.rel_type == FONT_REL_TYPE)
                .count(),
            2
        );
        assert_eq!(
            shared_package
                .parts
                .keys()
                .filter(|part| part.starts_with("/word/fonts/"))
                .count(),
            2
        );

        let mut wrong_type_package = OpcPackage::from_reader(std::io::Cursor::new(
            authored_document().to_bytes().unwrap(),
        ))
        .unwrap();
        let wrong_type_relationship = wrong_type_package
            .get_part_rels_mut("/word/fontTable.xml")
            .unwrap()
            .items
            .iter_mut()
            .find(|relationship| relationship.rel_type == FONT_REL_TYPE)
            .unwrap();
        let wrong_type_id = wrong_type_relationship.id.clone();
        wrong_type_relationship.rel_type = rel_types::IMAGE.to_owned();
        let mut wrong_type = Document::from_bytes(&package_bytes(&wrong_type_package)).unwrap();
        wrong_type
            .remove_embedded_font("Corpus Sans", EmbeddedFontKind::Regular)
            .unwrap();
        let wrong_type_saved =
            OpcPackage::from_reader(std::io::Cursor::new(wrong_type.to_bytes().unwrap())).unwrap();
        assert_eq!(
            wrong_type_saved
                .get_part_rels("/word/fontTable.xml")
                .unwrap()
                .get_by_id(&wrong_type_id)
                .unwrap()
                .rel_type,
            rel_types::IMAGE
        );

        let mut cross_owner_package = OpcPackage::from_reader(std::io::Cursor::new(
            authored_document().to_bytes().unwrap(),
        ))
        .unwrap();
        let font_target = cross_owner_package
            .get_part_rels("/word/fontTable.xml")
            .unwrap()
            .items
            .iter()
            .find(|relationship| relationship.rel_type == FONT_REL_TYPE)
            .map(|relationship| {
                OpcPackage::resolve_rel_target("/word/fontTable.xml", &relationship.target)
            })
            .unwrap();
        cross_owner_package.package_rels.add_with_id(
            "producerFontReference",
            "urn:producer:font-reference",
            font_target.trim_start_matches('/'),
        );
        let mut cross_owner = Document::from_bytes(&package_bytes(&cross_owner_package)).unwrap();
        cross_owner
            .remove_embedded_font("Corpus Sans", EmbeddedFontKind::Regular)
            .unwrap();
        let cross_owner_saved =
            OpcPackage::from_reader(std::io::Cursor::new(cross_owner.to_bytes().unwrap())).unwrap();
        assert!(cross_owner_saved.get_part(&font_target).is_some());
        assert!(
            cross_owner_saved
                .package_rels
                .get_by_id("producerFontReference")
                .is_some()
        );

        let mut cross_owner_replace =
            Document::from_bytes(&package_bytes(&cross_owner_package)).unwrap();
        cross_owner_replace
            .embed_font(
                "Corpus Sans",
                EmbeddedFont::new(
                    EmbeddedFontKind::Regular,
                    CALADEA.to_vec(),
                    FONT_KEY,
                    license(true, "Apache-2.0: Caladea"),
                ),
            )
            .unwrap();
        let cross_owner_replace_saved = OpcPackage::from_reader(std::io::Cursor::new(
            cross_owner_replace.to_bytes().unwrap(),
        ))
        .unwrap();
        assert_eq!(
            cross_owner_replace_saved.get_part(&font_target),
            cross_owner_package.get_part(&font_target)
        );

        let raw_reference =
            format!(r#"<rdocx:fontRef r:id="{original_id}" rdocx:keep="exact"/></w:fonts>"#);
        let raw_reference_table =
            std::str::from_utf8(cross_owner_package.get_part("/word/fontTable.xml").unwrap())
                .unwrap()
                .replace("</w:fonts>", &raw_reference);
        let mut raw_reference_package = cross_owner_package;
        raw_reference_package.set_part("/word/fontTable.xml", raw_reference_table.into_bytes());
        let mut raw_reference_document =
            Document::from_bytes(&package_bytes(&raw_reference_package)).unwrap();
        raw_reference_document
            .remove_embedded_font("Corpus Sans", EmbeddedFontKind::Regular)
            .unwrap();
        let raw_reference_saved = OpcPackage::from_reader(std::io::Cursor::new(
            raw_reference_document.to_bytes().unwrap(),
        ))
        .unwrap();
        assert!(
            raw_reference_saved
                .get_part_rels("/word/fontTable.xml")
                .unwrap()
                .get_by_id(&original_id)
                .is_some()
        );
    }
}

mod settings_and_properties_tests {
    use super::*;
    use rdocx::{
        AppProperties, CharacterSpacingControl, CompatibilitySetting, CoreProperties,
        CustomProperty, CustomPropertyValue, ThemeFontLanguage, Twips,
    };

    const CUSTOM_FMTID: &str = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
    const COMPATIBILITY_URI: &str = "http://schemas.microsoft.com/office/word";

    #[test]
    fn authored_settings_and_properties_survive_reopen() {
        let mut document = Document::new();
        document
            .set_core_properties(CoreProperties {
                title: Some("Quarterly proposal".to_owned()),
                creator: Some("Example author".to_owned()),
                subject: Some("Corpus settings".to_owned()),
                description: Some("Source-built metadata".to_owned()),
                keywords: Some("proposal, deterministic".to_owned()),
                last_modified_by: Some("Example reviewer".to_owned()),
                created: Some("2026-09-07T00:00:00Z".to_owned()),
                modified: Some("2026-09-07T01:00:00Z".to_owned()),
            })
            .unwrap();
        let mut application = AppProperties::default();
        application.template = Some("Business.dotx".to_owned());
        application.manager = Some("Example manager".to_owned());
        application.company = Some("Example company".to_owned());
        application.pages = Some(7);
        application.words = Some(420);
        application.application = Some("rdocx test".to_owned());
        application.application_version = Some("1.0".to_owned());
        document
            .set_application_properties(application.clone())
            .unwrap();
        document
            .set_custom_property(CustomProperty {
                fmtid: CUSTOM_FMTID.to_owned(),
                pid: 2,
                name: Some("ClientCode".to_owned()),
                value: CustomPropertyValue::Lpwstr("EXAMPLE-001".to_owned()),
            })
            .unwrap();
        document.set_document_variable("Customer", "Ada").unwrap();
        document
            .set_compatibility_setting("compatibilityMode", COMPATIBILITY_URI, "15")
            .unwrap();
        document.set_default_tab_stop(Twips(720)).unwrap();
        document
            .set_character_spacing_control(CharacterSpacingControl::CompressPunctuation)
            .unwrap();
        document
            .set_language_defaults(ThemeFontLanguage {
                latin: Some("en-GB".to_owned()),
                east_asia: Some("ja-JP".to_owned()),
                bidi: Some("ar-SA".to_owned()),
            })
            .unwrap();

        let bytes = document.to_bytes().unwrap();
        let mut reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(
            reopened.core_properties().unwrap().title.as_deref(),
            Some("Quarterly proposal")
        );
        assert_eq!(reopened.application_properties(), Some(&application));
        assert_eq!(
            reopened.custom_property("ClientCode").unwrap().value,
            CustomPropertyValue::Lpwstr("EXAMPLE-001".to_owned())
        );
        assert_eq!(reopened.document_variable("Customer"), Some("Ada"));
        assert_eq!(
            reopened.compatibility_settings(),
            [CompatibilitySetting {
                name: "compatibilityMode".to_owned(),
                uri: COMPATIBILITY_URI.to_owned(),
                value: "15".to_owned(),
            }]
        );
        assert_eq!(reopened.default_tab_stop(), Some(Twips(720)));
        assert_eq!(
            reopened.character_spacing_control(),
            Some(CharacterSpacingControl::CompressPunctuation)
        );
        assert_eq!(
            reopened.theme_font_language().unwrap(),
            &ThemeFontLanguage {
                latin: Some("en-GB".to_owned()),
                east_asia: Some("ja-JP".to_owned()),
                bidi: Some("ar-SA".to_owned()),
            }
        );

        assert_eq!(
            reopened.remove_document_variable("Customer").unwrap(),
            Some("Ada".to_owned())
        );
        assert!(
            reopened
                .remove_compatibility_setting("compatibilityMode", COMPATIBILITY_URI)
                .unwrap()
                .is_some()
        );
        assert_eq!(
            reopened.remove_default_tab_stop().unwrap(),
            Some(Twips(720))
        );
        assert_eq!(
            reopened.remove_character_spacing_control().unwrap(),
            Some(CharacterSpacingControl::CompressPunctuation)
        );
        assert!(reopened.remove_theme_font_language().unwrap().is_some());
        let removed_bytes = reopened.to_bytes().unwrap();
        let removed = Document::from_bytes(&removed_bytes).unwrap();
        assert_eq!(removed.document_variable("Customer"), None);
        assert!(removed.compatibility_settings().is_empty());
        assert_eq!(removed.default_tab_stop(), None);
        assert_eq!(removed.character_spacing_control(), None);
        assert_eq!(removed.theme_font_language(), None);
    }

    #[test]
    fn removing_one_property_family_prunes_only_its_owned_graph() {
        let mut document =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        document
            .set_core_properties(CoreProperties {
                title: Some("Retained title".to_owned()),
                ..Default::default()
            })
            .unwrap();
        let mut application = AppProperties::default();
        application.company = Some("Retained company".to_owned());
        document.set_application_properties(application).unwrap();
        document
            .set_custom_property(CustomProperty {
                fmtid: CUSTOM_FMTID.to_owned(),
                pid: 2,
                name: Some("RemoveMe".to_owned()),
                value: CustomPropertyValue::Bool(true),
            })
            .unwrap();
        document.set_document_variable("KeepMe", "yes").unwrap();

        assert!(
            document
                .remove_custom_property("RemoveMe")
                .unwrap()
                .is_some()
        );
        let bytes = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(&bytes)).unwrap();
        assert!(
            package
                .package_rels
                .get_by_type(rel_types::CUSTOM_PROPERTIES)
                .is_none()
        );
        assert!(package.get_part("/docProps/custom.xml").is_none());
        assert!(
            package
                .package_rels
                .get_by_type(rel_types::CORE_PROPERTIES)
                .is_some()
        );
        assert!(
            package
                .package_rels
                .get_by_type(rel_types::EXTENDED_PROPERTIES)
                .is_some()
        );
        let mut reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.title(), Some("Retained title"));
        assert_eq!(
            reopened
                .application_properties()
                .unwrap()
                .company
                .as_deref(),
            Some("Retained company")
        );
        assert_eq!(reopened.document_variable("KeepMe"), Some("yes"));

        assert!(reopened.remove_application_properties().unwrap().is_some());
        assert!(reopened.remove_core_properties().unwrap().is_some());
        let bytes = reopened.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(&bytes)).unwrap();
        assert!(
            package
                .package_rels
                .get_by_type(rel_types::EXTENDED_PROPERTIES)
                .is_none()
        );
        assert!(
            package
                .package_rels
                .get_by_type(rel_types::CORE_PROPERTIES)
                .is_none()
        );
        assert!(
            package
                .get_part_rels("/word/document.xml")
                .unwrap()
                .get_by_type(rel_types::SETTINGS)
                .is_some()
        );

        let mut settings_only =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        settings_only
            .set_document_variable("Temporary", "value")
            .unwrap();
        assert_eq!(
            settings_only.remove_document_variable("Temporary").unwrap(),
            Some("value".to_owned())
        );
        let settings_only = settings_only.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(settings_only)).unwrap();
        assert!(
            package
                .get_part_rels("/word/document.xml")
                .unwrap()
                .get_by_type(rel_types::SETTINGS)
                .is_none()
        );
        assert!(package.get_part("/word/settings.xml").is_none());
    }

    #[test]
    fn settings_mutation_preserves_unmodeled_children_in_schema_order() {
        const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let mut seed = Document::new();
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
        let settings = format!(
            r#"<q:settings xmlns:q="{W_NS}" xmlns:x="urn:producer"><q:defaultTabStop q:val="360"/><x:before x:keep="exact"><x:child/></x:before><q:characterSpacingControl q:val="doNotCompress"/><q:compat><q:compatSetting q:name="compatibilityMode" q:uri="{COMPATIBILITY_URI}" q:val="14"/><x:inside x:keep="exact"/></q:compat><q:docVars><q:docVar q:name="Original" q:val="one"/><x:variable x:keep="exact"/></q:docVars><x:after x:keep="exact"/><q:themeFontLang q:val="en-US" q:eastAsia="zh-CN"/></q:settings>"#
        );
        package.set_part("/word/settings.xml", settings.into_bytes());
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let mut document = Document::from_bytes(bytes.get_ref()).unwrap();

        document.set_default_tab_stop(Twips(720)).unwrap();
        document
            .set_character_spacing_control(
                CharacterSpacingControl::CompressPunctuationAndJapaneseKana,
            )
            .unwrap();
        document
            .set_compatibility_setting("compatibilityMode", COMPATIBILITY_URI, "15")
            .unwrap();
        document.set_document_variable("Added", "two").unwrap();
        assert_eq!(
            document.remove_document_variable("Original").unwrap(),
            Some("one".to_owned())
        );
        document
            .set_language_defaults(ThemeFontLanguage {
                latin: Some("en-GB".to_owned()),
                east_asia: Some("ja-JP".to_owned()),
                bidi: None,
            })
            .unwrap();

        let saved =
            OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
        let output = std::str::from_utf8(saved.get_part("/word/settings.xml").unwrap()).unwrap();
        assert!(output.contains(r#"<x:before x:keep="exact"><x:child/></x:before>"#));
        assert!(output.contains(r#"<x:inside x:keep="exact"/>"#));
        assert!(output.contains(r#"<x:variable x:keep="exact"/>"#));
        assert!(output.contains(r#"<x:after x:keep="exact"/>"#));
        assert!(!output.contains("Original"));
        assert!(output.contains("Added"));
        assert!(
            output.find("defaultTabStop").unwrap()
                < output.find("characterSpacingControl").unwrap()
        );
        assert!(
            output.find("characterSpacingControl").unwrap() < output.find("compatSetting").unwrap()
        );
        assert!(output.find("compatSetting").unwrap() < output.find("docVar").unwrap());
        assert!(output.find("docVar").unwrap() < output.find("themeFontLang").unwrap());
    }

    #[test]
    fn fresh_property_output_has_no_clock_or_host_input() {
        fn authored() -> Vec<u8> {
            let mut document = Document::new_with_profile(WordCreationProfile::Minimal(
                WordPackageClass::Document,
            ));
            document
                .set_core_properties(CoreProperties {
                    title: Some("Deterministic metadata".to_owned()),
                    ..Default::default()
                })
                .unwrap();
            let mut application = AppProperties::default();
            application.application = Some("rdocx".to_owned());
            application.application_version = Some("test".to_owned());
            document.set_application_properties(application).unwrap();
            document.to_bytes().unwrap()
        }

        let first = authored();
        let second = authored();
        assert_eq!(first, second);

        let mut invalid =
            Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
        let before = invalid.to_bytes().unwrap();
        assert!(invalid.set_default_tab_stop(Twips(-1)).is_err());
        assert_eq!(invalid.to_bytes().unwrap(), before);
        let package = OpcPackage::from_reader(std::io::Cursor::new(first)).unwrap();
        let core = std::str::from_utf8(package.get_part("/docProps/core.xml").unwrap()).unwrap();
        assert!(!core.contains("dcterms:created"));
        assert!(!core.contains("dcterms:modified"));
    }
}

#[test]
fn identifier_scopes_do_not_alias_or_overreach() {
    let mut document = Document::new();
    document.add_paragraph("scope");
    let range = RunRange {
        start: RunPosition {
            body_index: 0,
            run_index: 0,
        },
        end: RunPosition {
            body_index: 0,
            run_index: 1,
        },
    };
    assert_eq!(document.add_bookmark("Scope", range).unwrap(), 0);
    assert_eq!(
        document
            .add_comment(range, "Author", None, "Scoped comment")
            .unwrap(),
        0
    );
    assert_eq!(document.add_list_definition(&[ListLevel::decimal()]), 1);
    document.add_picture(
        b"scope-image",
        "scope.png",
        Length::inches(1.0),
        Length::inches(1.0),
    );
    let bytes = document.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let document_relationships = package.get_part_rels("/word/document.xml").unwrap();
    let image_relationship = document_relationships
        .items
        .iter()
        .find(|relationship| relationship.rel_type == rel_types::IMAGE)
        .unwrap()
        .id
        .clone();
    package.set_part("/word/header-scope.xml", b"<scope/>".to_vec());
    package
        .content_types
        .add_override("/word/header-scope.xml", "application/xml");
    package
        .get_or_create_part_rels("/word/header-scope.xml")
        .add_with_id(&image_relationship, "urn:scope", "scope-target.xml");
    let mut output = std::io::Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();
    let mut reopened = Document::from_bytes(&output.into_inner()).unwrap();
    let saved = reopened.to_bytes().unwrap();
    let saved = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();

    assert!(
        saved
            .get_part_rels("/word/document.xml")
            .unwrap()
            .get_by_id(&image_relationship)
            .is_some()
    );
    assert!(
        saved
            .get_part_rels("/word/header-scope.xml")
            .unwrap()
            .get_by_id(&image_relationship)
            .is_some()
    );
    let document_xml = std::str::from_utf8(saved.get_part("/word/document.xml").unwrap()).unwrap();
    let comments_xml = std::str::from_utf8(saved.get_part("/word/comments.xml").unwrap()).unwrap();
    let numbering_xml =
        std::str::from_utf8(saved.get_part("/word/numbering.xml").unwrap()).unwrap();
    assert!(document_xml.contains(r#"w:bookmarkStart w:id="0""#));
    assert!(comments_xml.contains(r#"w:id="0""#));
    assert!(document_xml.contains(r#"wp:docPr id="1""#));
    assert!(numbering_xml.contains(r#"w:num w:numId="1""#));
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct MhtmlOracleRecord {
    text: String,
    bold_text: Vec<String>,
    numbered_text: Vec<String>,
    table_cells: Vec<Vec<Vec<String>>>,
    image_sizes: Vec<(i64, i64)>,
    links: Vec<(String, Option<String>, Option<String>)>,
    diagnostics: Vec<MhtmlDiagnostic>,
}

mod public_authoring_conformance_harness {
    use std::path::PathBuf;
    use std::process::Command;

    fn repository_root() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .and_then(|path| path.parent())
            .expect("rdocx crate must live under the workspace crates directory")
            .to_path_buf()
    }

    #[test]
    fn sanitized_public_authoring_fixture_passes_every_conformance_stage() {
        let root = repository_root();
        let script = root.join("scripts/docx_authoring_conformance.py");
        assert!(
            script.is_file(),
            "missing public authoring conformance harness"
        );

        let output = Command::new("python3")
            .arg(&script)
            .arg("--public")
            .current_dir(&root)
            .output()
            .expect("public authoring conformance harness must start");
        assert!(
            output.status.success(),
            "public authoring conformance failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }

    #[test]
    fn public_fixture_rejects_base_package_raw_xml_or_private_oxml_dependency() {
        let root = repository_root();
        let script = root.join("scripts/docx_authoring_conformance.py");
        assert!(
            script.is_file(),
            "missing public authoring conformance harness"
        );

        let output = Command::new("python3")
            .arg(&script)
            .arg("--self-test")
            .current_dir(&root)
            .output()
            .expect("authoring conformance self-test must start");
        assert!(
            output.status.success(),
            "authoring conformance self-test failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }

    #[test]
    #[ignore = "requires the ignored M23 private corpus and pinned renderers"]
    fn m23_private_from_scratch_corpus_passes_required_mode() {
        let root = repository_root();
        let script = root.join("scripts/docx_authoring_conformance.py");
        let output = Command::new("python3")
            .arg(&script)
            .arg("--private-generate-required")
            .current_dir(&root)
            .output()
            .expect("private authoring conformance harness must start");
        assert!(
            output.status.success(),
            "private from-scratch authoring conformance failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }
}

mod flat_opc_package_class_tests {
    use super::*;
    use base64::Engine;
    use base64::engine::general_purpose::STANDARD as BASE64;
    use oxml_opc::content_types;

    const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const WORD_ORACLE_VERSION: &str = "Microsoft Word 16.104 build 16.104.25121423";
    const VBA_BYTES: &[u8] = b"source-built-vba-project";

    fn content_type(class: WordPackageClass) -> &'static str {
        match class {
            WordPackageClass::Document => content_types::WORD_DOCUMENT,
            WordPackageClass::MacroEnabledDocument => content_types::WORD_DOCUMENT_MACRO_ENABLED,
            WordPackageClass::Template => content_types::WORD_TEMPLATE,
            WordPackageClass::MacroEnabledTemplate => content_types::WORD_TEMPLATE_MACRO_ENABLED,
        }
    }

    fn package_bytes(package: &OpcPackage) -> Vec<u8> {
        let mut output = std::io::Cursor::new(Vec::new());
        package.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn source_package(class: WordPackageClass) -> OpcPackage {
        let mut package = OpcPackage::with_main_part("word/document.xml", content_type(class));
        package.set_part(
            "/word/document.xml",
            format!(
                r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p><w:r><w:t>class fixture</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#,
            )
            .into_bytes(),
        );
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::VBA_PROJECT, "vbaProject.bin");
        package.set_part("/word/vbaProject.bin", VBA_BYTES.to_vec());
        package.content_types.add_override(
            "/word/vbaProject.bin",
            "application/vnd.ms-office.vbaProject",
        );
        package.set_part(
            "/custom/preserved.xml",
            br#"<preserved xmlns="urn:rdocx:f238"><opaque a="1"> bytes </opaque></preserved>"#
                .to_vec(),
        );
        package
            .content_types
            .add_override("/custom/preserved.xml", "application/xml");
        package.set_part(
            "/custom/text-xml.xml",
            br#"<text-xml xmlns="urn:rdocx:f238:text">preserved</text-xml>"#.to_vec(),
        );
        package
            .content_types
            .add_override("/custom/text-xml.xml", "text/xml");
        package.set_part("/custom/empty.bin", Vec::new());
        package
            .content_types
            .add_override("/custom/empty.bin", "application/octet-stream");
        package
    }

    pub(super) fn add_package_signature_graph(package: &mut OpcPackage) {
        package.set_part("/_xmlsignatures/origin.sigs", Vec::new());
        package.set_part(
            "/_xmlsignatures/sig1.xml",
            br#"<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>"#.to_vec(),
        );
        package.content_types.add_override(
            "/_xmlsignatures/origin.sigs",
            "application/vnd.openxmlformats-package.digital-signature-origin",
        );
        package.content_types.add_override(
            "/_xmlsignatures/sig1.xml",
            "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
        );
        package.package_rels.add_with_id(
            "package-signature-origin",
            rel_types::DIGITAL_SIGNATURE_ORIGIN,
            "_xmlsignatures/origin.sigs",
        );
        package
            .get_or_create_part_rels("/_xmlsignatures/origin.sigs")
            .add_with_id(
                "package-signature",
                rel_types::DIGITAL_SIGNATURE,
                "sig1.xml",
            );
    }

    pub(super) fn has_package_signature_invalidation_marker(package: &OpcPackage) -> bool {
        package.package_rels.items.iter().any(|relationship| {
            relationship.rel_type == "urn:rdocx:relationships/invalidated-package-signature"
        })
    }

    fn all_classes() -> [WordPackageClass; 4] {
        [
            WordPackageClass::Document,
            WordPackageClass::MacroEnabledDocument,
            WordPackageClass::Template,
            WordPackageClass::MacroEnabledTemplate,
        ]
    }

    #[test]
    fn flat_opc_and_modern_word_package_classes_reopen_without_repair_and_preserve_payloads() {
        for class in all_classes() {
            let document = Document::from_bytes(&package_bytes(&source_package(class))).unwrap();
            assert_eq!(document.package_class().unwrap(), class);
            let flat = document.to_flat_opc_bytes().unwrap();
            let imported = Document::from_flat_opc_bytes(&flat).unwrap();
            assert_eq!(imported.package_class().unwrap(), class);
            let zip = imported.to_bytes_as(class).unwrap();
            let reopened = Document::from_bytes(&zip).unwrap();
            assert_eq!(reopened.package_class().unwrap(), class);
            let package = OpcPackage::from_reader(std::io::Cursor::new(zip)).unwrap();
            assert_eq!(package.get_part("/word/vbaProject.bin"), Some(VBA_BYTES));
            assert_eq!(
                package.get_part("/custom/preserved.xml"),
                source_package(class).get_part("/custom/preserved.xml")
            );
            assert_eq!(
                package
                    .get_part_rels("/word/document.xml")
                    .unwrap()
                    .get_by_type(rel_types::VBA_PROJECT)
                    .unwrap()
                    .target,
                "vbaProject.bin"
            );
        }
    }

    #[test]
    fn flat_opc_xml_and_binary_parts_round_trip_without_loss() {
        let document = Document::from_bytes(&package_bytes(&source_package(
            WordPackageClass::MacroEnabledDocument,
        )))
        .unwrap();
        let canonical = document.to_flat_opc_bytes().unwrap();
        let aliased = String::from_utf8(canonical)
            .unwrap()
            .replace("pkg:", "alias:")
            .replace("xmlns:pkg=", "xmlns:alias=");
        let aliased = aliased
            .replace(
                "<alias:binaryData></alias:binaryData>",
                "<alias:binaryData/>",
            )
            .replacen(
                "<alias:part ",
                "<alias:part xmlns:local=\"urn:rdocx:f238:local\" ",
                1,
            )
            .replacen(
                "<Relationship ",
                "<Relationship xmlns:localRelationship=\"urn:rdocx:f238:relationship\" ",
                1,
            );
        let relationship_start = aliased.find("<Relationship ").unwrap();
        let relationship_end =
            relationship_start + aliased[relationship_start..].find("/>").unwrap();
        let mut equivalent_relationship_syntax = aliased;
        equivalent_relationship_syntax
            .replace_range(relationship_end..relationship_end + 2, "></Relationship>");
        let imported =
            Document::from_flat_opc_bytes(equivalent_relationship_syntax.as_bytes()).unwrap();
        let output = imported.to_flat_opc_bytes().unwrap();
        let xml = std::str::from_utf8(&output).unwrap();
        assert!(xml.contains("<pkg:package"));
        assert!(xml.contains("<pkg:xmlData>"));
        assert!(xml.contains("<pkg:binaryData>"));
        assert!(xml.contains("pkg:contentType=\"text/xml\"><pkg:xmlData>"));
        assert!(!xml.contains("alias:"));
        let reopened = OpcPackage::from_reader(std::io::Cursor::new(
            imported
                .to_bytes_as(WordPackageClass::MacroEnabledDocument)
                .unwrap(),
        ))
        .unwrap();
        assert_eq!(reopened.get_part("/word/vbaProject.bin"), Some(VBA_BYTES));
        assert_eq!(
            reopened.get_part("/custom/preserved.xml"),
            source_package(WordPackageClass::MacroEnabledDocument)
                .get_part("/custom/preserved.xml")
        );
        assert_eq!(reopened.get_part("/custom/empty.bin"), Some(&[][..]));
    }

    #[test]
    fn flat_opc_accepts_sole_mixed_case_special_names_and_relationship_owner() {
        let document =
            Document::from_bytes(&package_bytes(&source_package(WordPackageClass::Document)))
                .unwrap();
        let flat = String::from_utf8(document.to_flat_opc_bytes().unwrap()).unwrap();
        let mixed = flat
            .replace("pkg:name=\"/_rels/.rels\"", "pkg:name=\"/_RELS/.RELS\"")
            .replace(
                "pkg:name=\"/word/_rels/document.xml.rels\"",
                "pkg:name=\"/WORD/_RELS/DOCUMENT.XML.RELS\"",
            )
            .replace(
                "pkg:name=\"/word/document.xml\"",
                "pkg:name=\"/WORD/DOCUMENT.XML\"",
            );

        let mut imported = Document::from_flat_opc_bytes(mixed.as_bytes()).unwrap();
        let package =
            OpcPackage::from_reader(std::io::Cursor::new(imported.to_bytes().unwrap())).unwrap();
        assert!(package.parts.contains_key("/WORD/DOCUMENT.XML"));
        assert!(package.part_rels.contains_key("/WORD/DOCUMENT.XML"));
    }

    #[test]
    fn flat_opc_import_materializes_inherited_payload_namespaces() {
        let document = Document::from_bytes(&package_bytes(&source_package(
            WordPackageClass::MacroEnabledDocument,
        )))
        .unwrap();
        let canonical = String::from_utf8(document.to_flat_opc_bytes().unwrap()).unwrap();

        let inherited_prefix = canonical
            .replacen(
                &format!("<w:document xmlns:w=\"{WORD_NS}\" "),
                "<w:document ",
                1,
            )
            .replacen(
                "<pkg:part pkg:name=\"/word/document.xml\"",
                &format!("<pkg:part xmlns:w=\"{WORD_NS}\" pkg:name=\"/word/document.xml\""),
                1,
            );
        assert_ne!(inherited_prefix, canonical);
        let imported = Document::from_flat_opc_bytes(inherited_prefix.as_bytes()).unwrap();
        let reopened = OpcPackage::from_reader(std::io::Cursor::new(
            imported
                .to_bytes_as(WordPackageClass::MacroEnabledDocument)
                .unwrap(),
        ))
        .unwrap();
        assert!(
            std::str::from_utf8(reopened.get_part("/word/document.xml").unwrap())
                .unwrap()
                .contains(&format!("xmlns:w=\"{WORD_NS}\""))
        );

        let inherited_default = canonical.replacen(
            "<pkg:xmlData><preserved xmlns=\"urn:rdocx:f238\">",
            "<pkg:xmlData xmlns=\"urn:rdocx:f238\"><preserved>",
            1,
        );
        assert_ne!(inherited_default, canonical);
        let imported = Document::from_flat_opc_bytes(inherited_default.as_bytes()).unwrap();
        let reopened = OpcPackage::from_reader(std::io::Cursor::new(
            imported
                .to_bytes_as(WordPackageClass::MacroEnabledDocument)
                .unwrap(),
        ))
        .unwrap();
        assert!(
            std::str::from_utf8(reopened.get_part("/custom/preserved.xml").unwrap())
                .unwrap()
                .contains("xmlns=\"urn:rdocx:f238\"")
        );

        let inherited_mc_value_prefix = canonical.replacen(
            "<pkg:xmlData><preserved xmlns=\"urn:rdocx:f238\"><opaque a=\"1\"> bytes </opaque></preserved></pkg:xmlData>",
            "<pkg:xmlData xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"urn:word:w14\" xmlns:w15=\"urn:word:w15\" xmlns:w16=\"urn:word:w16\" xmlns:w17=\"urn:word:w17\" xmlns:w18=\"urn:word:w18\" xmlns:w19=\"urn:word:w19\"><preserved xmlns=\"urn:rdocx:f238\" mc:Ignorable=\"w14\" mc:MustUnderstand=\"w15\" mc:ProcessContent=\"w16:item\" mc:PreserveElements=\"w17:*\" mc:PreserveAttributes=\"w18:attribute\"><mc:AlternateContent><mc:Choice Requires=\"w19\"><opaque a=\"1\"> bytes </opaque></mc:Choice><mc:Fallback/></mc:AlternateContent></preserved></pkg:xmlData>",
            1,
        );
        assert_ne!(inherited_mc_value_prefix, canonical);
        let imported = Document::from_flat_opc_bytes(inherited_mc_value_prefix.as_bytes()).unwrap();
        let reopened = OpcPackage::from_reader(std::io::Cursor::new(
            imported
                .to_bytes_as(WordPackageClass::MacroEnabledDocument)
                .unwrap(),
        ))
        .unwrap();
        assert_eq!(
            reopened.get_part("/custom/preserved.xml"),
            Some(
                br#"<preserved xmlns="urn:rdocx:f238" mc:Ignorable="w14" mc:MustUnderstand="w15" mc:ProcessContent="w16:item" mc:PreserveElements="w17:*" mc:PreserveAttributes="w18:attribute" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="urn:word:w14" xmlns:w15="urn:word:w15" xmlns:w16="urn:word:w16" xmlns:w17="urn:word:w17" xmlns:w18="urn:word:w18" xmlns:w19="urn:word:w19"><mc:AlternateContent><mc:Choice Requires="w19"><opaque a="1"> bytes </opaque></mc:Choice><mc:Fallback/></mc:AlternateContent></preserved>"#
                    .as_slice()
            )
        );
    }

    #[test]
    fn flat_opc_treats_transitional_and_strict_alt_chunks_as_opaque() {
        const ALT_CHUNK_RELATIONSHIPS: [&str; 2] = [
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
            "http://purl.oclc.org/ooxml/officeDocument/relationships/aFChunk",
        ];
        const XHTML: &[u8] =
            br#"<html xmlns="http://www.w3.org/1999/xhtml"><body> <p>chunk</p><!--opaque--></body></html>"#;
        const ORDINARY_XHTML: &[u8] =
            br#"<html xmlns="http://www.w3.org/1999/xhtml"><body><p>ordinary XML</p></body></html>"#;

        for relationship_type in ALT_CHUNK_RELATIONSHIPS {
            let mut package = source_package(WordPackageClass::Document);
            package.set_part(
                "/word/document.xml",
                format!(
                    r#"<w:document xmlns:w="{WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="altChunk1"/><w:sectPr/></w:body></w:document>"#,
                )
                .into_bytes(),
            );
            package
                .get_or_create_part_rels("/word/document.xml")
                .add_with_id("altChunk1", relationship_type, "afchunk1.xhtml");
            package.set_part("/word/afchunk1.xhtml", XHTML.to_vec());
            package
                .content_types
                .add_override("/word/afchunk1.xhtml", "application/xhtml+xml");
            package.set_part("/word/ordinary.xhtml", ORDINARY_XHTML.to_vec());
            package
                .content_types
                .add_override("/word/ordinary.xhtml", "application/xhtml+xml");

            let document = Document::from_bytes(&package_bytes(&package)).unwrap();
            let flat = String::from_utf8(document.to_flat_opc_bytes().unwrap()).unwrap();
            let part_start = flat
                .find("<pkg:part pkg:name=\"/word/afchunk1.xhtml\"")
                .unwrap();
            let part_end = part_start + flat[part_start..].find("</pkg:part>").unwrap();
            let part = &flat[part_start..part_end];
            assert!(part.contains("<pkg:binaryData>"), "{part}");
            assert!(part.contains(&BASE64.encode(XHTML)), "{part}");
            let ordinary_start = flat
                .find("<pkg:part pkg:name=\"/word/ordinary.xhtml\"")
                .unwrap();
            let ordinary_end = ordinary_start + flat[ordinary_start..].find("</pkg:part>").unwrap();
            assert!(
                flat[ordinary_start..ordinary_end].contains("<pkg:xmlData>"),
                "{}",
                &flat[ordinary_start..ordinary_end]
            );

            let imported = Document::from_flat_opc_bytes(flat.as_bytes()).unwrap();
            let reopened = OpcPackage::from_reader(std::io::Cursor::new(
                imported.to_bytes_as(WordPackageClass::Document).unwrap(),
            ))
            .unwrap();
            assert_eq!(reopened.get_part("/word/afchunk1.xhtml"), Some(XHTML));
            assert_eq!(
                reopened.get_part("/word/ordinary.xhtml"),
                Some(ORDINARY_XHTML)
            );
        }
    }

    #[test]
    fn representative_m22_document_composes_the_complete_milestone_gate() {
        let mut package = source_package(WordPackageClass::MacroEnabledTemplate);
        let header_id = package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::HEADER, "header1.xml");
        package.set_part(
            "/word/header1.xml",
            format!(
                r#"<w:hdr xmlns:w="{WORD_NS}"><w:p><w:r><w:t>original header</w:t></w:r></w:p></w:hdr>"#,
            )
            .into_bytes(),
        );
        package.content_types.add_override(
            "/word/header1.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        );
        package.set_part(
            "/word/document.xml",
            format!(
                r#"<w:document xmlns:w="{WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:rdocx:m22"><w:body>
<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \o "1-1" \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
<w:p><w:r><w:t>stale entry</w:t></w:r></w:p>
<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Milestone heading</w:t></w:r></w:p>
<w:p><w:fldSimple w:instr="MERGEFIELD Name"><w:r><w:t>stored name</w:t></w:r></w:fldSimple></w:p>
<x:unsupported x:token="preserve-me"/>
<w:sectPr><w:headerReference w:type="default" r:id="{header_id}"/></w:sectPr></w:body></w:document>"#,
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(&package)).unwrap();

        document
            .add_paragraph("Authored equation")
            .add_equation(rdocx::OfficeMath::inline(vec![
                rdocx::MathRun::new("x").into(),
                rdocx::MathExpression::Fraction(rdocx::MathFraction::new(
                    rdocx::MathArgument::text("1"),
                    rdocx::MathArgument::text("2"),
                )),
            ]))
            .unwrap();
        let rendered = document
            .render_page_to_svg_deterministic(0)
            .unwrap()
            .expect("the representative document renders page zero");
        assert!(rendered.svg.contains(">x</text>"), "{}", rendered.svg);
        assert!(rendered.svg.contains(">1</text>"), "{}", rendered.svg);
        assert!(rendered.svg.contains(">2</text>"), "{}", rendered.svg);

        let toc = document.rebuild_toc().unwrap();
        assert_eq!(toc.entry_count, 1);
        let post_toc_bytes = document
            .to_bytes_as(WordPackageClass::MacroEnabledTemplate)
            .unwrap();
        let post_toc_package =
            OpcPackage::from_reader(std::io::Cursor::new(&post_toc_bytes)).unwrap();
        let post_toc_xml =
            std::str::from_utf8(post_toc_package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(!post_toc_xml.contains("stale entry"), "{post_toc_xml}");
        assert!(post_toc_xml.contains("Milestone heading"), "{post_toc_xml}");
        assert!(post_toc_xml.contains("PAGEREF _Toc"), "{post_toc_xml}");
        let mut field_updated = Document::from_bytes(&post_toc_bytes).unwrap();
        let mut field_context = rdocx::FieldEvaluationContext::default();
        field_context
            .merge_fields
            .insert("Name".to_owned(), "Ada".to_owned());
        assert!(field_updated.update_fields(&field_context).unwrap() >= 1);
        let field_updated_package = OpcPackage::from_reader(std::io::Cursor::new(
            field_updated
                .to_bytes_as(WordPackageClass::MacroEnabledTemplate)
                .unwrap(),
        ))
        .unwrap();
        assert!(
            std::str::from_utf8(
                field_updated_package
                    .get_part("/word/document.xml")
                    .unwrap()
            )
            .unwrap()
            .contains(">Ada<")
        );

        let records = [
            std::collections::BTreeMap::from([("Name".to_owned(), "Ada".to_owned())]),
            std::collections::BTreeMap::from([("Name".to_owned(), "Grace".to_owned())]),
        ];
        let merge_source = Document::from_bytes(&post_toc_bytes).unwrap();
        let merged = merge_source.mail_merge_sections(&records).unwrap();
        let merged_bytes = merged
            .to_bytes_as(WordPackageClass::MacroEnabledTemplate)
            .unwrap();
        let merged_package = OpcPackage::from_reader(std::io::Cursor::new(&merged_bytes)).unwrap();
        let merged_xml =
            std::str::from_utf8(merged_package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(merged_xml.contains(">Ada<"), "{merged_xml}");
        assert!(merged_xml.contains(">Grace<"), "{merged_xml}");
        assert_eq!(
            merged_xml.matches(r#"<w:type w:val="nextPage"/>"#).count(),
            1,
            "{merged_xml}"
        );
        let inventory = merged.embedded_content().unwrap();
        assert_eq!(inventory.len(), 1);
        assert_eq!(inventory[0].kind, rdocx::EmbeddedContentKind::VbaProject);

        let mut compared = Document::from_bytes(&merged_bytes).unwrap();
        let mut edited_package =
            OpcPackage::from_reader(std::io::Cursor::new(&merged_bytes)).unwrap();
        let edited_header =
            std::str::from_utf8(edited_package.get_part("/word/header1.xml").unwrap())
                .unwrap()
                .replace("original header", "edited header");
        edited_package.set_part("/word/header1.xml", edited_header.into_bytes());
        let mut edited = Document::from_bytes(&package_bytes(&edited_package)).unwrap();
        edited.add_paragraph("comparison addition");
        compared
            .compare(&edited, "M22 gate", "2026-09-06T09:00:00Z")
            .unwrap();
        assert!(!compared.revisions().is_empty());

        let flat = compared.to_flat_opc_bytes().unwrap();
        let reopened = Document::from_flat_opc_bytes(&flat).unwrap();
        assert_eq!(
            reopened.package_class().unwrap(),
            WordPackageClass::MacroEnabledTemplate
        );
        assert!(
            reopened
                .paragraphs()
                .iter()
                .any(|paragraph| paragraph.equations().next().is_some())
        );
        let final_zip = reopened
            .to_bytes_as(WordPackageClass::MacroEnabledTemplate)
            .unwrap();
        let final_package = OpcPackage::from_reader(std::io::Cursor::new(final_zip)).unwrap();
        assert_eq!(
            final_package.get_part("/word/vbaProject.bin"),
            Some(VBA_BYTES)
        );
        assert_eq!(
            final_package.get_part("/custom/preserved.xml"),
            package.get_part("/custom/preserved.xml")
        );
        let compared_document_xml =
            std::str::from_utf8(final_package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(
            compared_document_xml.contains("preserve-me"),
            "{compared_document_xml}"
        );
        assert!(
            compared_document_xml.contains("comparison addition"),
            "{compared_document_xml}"
        );
        assert!(
            compared_document_xml.contains("<w:ins"),
            "{compared_document_xml}"
        );
        assert!(
            std::str::from_utf8(final_package.get_part("/word/header1.xml").unwrap())
                .unwrap()
                .contains("<w:ins")
        );
    }

    #[test]
    fn ordinary_save_preserves_opened_word_template_and_macro_classes() {
        for class in all_classes() {
            let mut document =
                Document::from_bytes(&package_bytes(&source_package(class))).unwrap();
            let ordinary = document.to_bytes().unwrap();
            assert_eq!(
                Document::from_bytes(&ordinary)
                    .unwrap()
                    .package_class()
                    .unwrap(),
                class
            );
            let flat = document.to_flat_opc_bytes().unwrap();
            assert_eq!(
                Document::from_flat_opc_bytes(&flat)
                    .unwrap()
                    .package_class()
                    .unwrap(),
                class
            );
        }
    }

    #[test]
    fn word_package_class_conversion_changes_only_the_main_content_type() {
        let source = source_package(WordPackageClass::MacroEnabledTemplate);
        let document = Document::from_bytes(&package_bytes(&source)).unwrap();
        let baseline = OpcPackage::from_reader(std::io::Cursor::new(
            document
                .to_bytes_as(WordPackageClass::MacroEnabledTemplate)
                .unwrap(),
        ))
        .unwrap();
        for class in all_classes() {
            let converted = document.to_bytes_as(class).unwrap();
            let package = OpcPackage::from_reader(std::io::Cursor::new(converted)).unwrap();
            assert_eq!(
                package.content_types.override_for("/word/document.xml"),
                Some(content_type(class))
            );
            assert_eq!(package.parts, baseline.parts);
            assert_eq!(
                package.package_rels.to_xml().unwrap(),
                baseline.package_rels.to_xml().unwrap()
            );
            assert_eq!(package.part_rels.len(), baseline.part_rels.len());
            for (part_name, expected) in &baseline.part_rels {
                assert_eq!(
                    package.part_rels[part_name].to_xml().unwrap(),
                    expected.to_xml().unwrap()
                );
            }
            let mut expected_types = baseline.content_types.clone();
            expected_types.add_override("/word/document.xml", content_type(class));
            assert_eq!(package.content_types, expected_types);
        }
        assert_eq!(
            document.package_class().unwrap(),
            WordPackageClass::MacroEnabledTemplate
        );

        let mut signed_package = baseline;
        add_package_signature_graph(&mut signed_package);
        let signed_document = Document::from_bytes(&package_bytes(&signed_package)).unwrap();
        let same_class = OpcPackage::from_reader(std::io::Cursor::new(
            signed_document
                .to_bytes_as(WordPackageClass::MacroEnabledTemplate)
                .unwrap(),
        ))
        .unwrap();
        let changed_class = OpcPackage::from_reader(std::io::Cursor::new(
            signed_document
                .to_bytes_as(WordPackageClass::Document)
                .unwrap(),
        ))
        .unwrap();
        assert!(!has_package_signature_invalidation_marker(&same_class));
        assert!(has_package_signature_invalidation_marker(&changed_class));
        assert!(changed_class.parts.contains_key("/_xmlsignatures/sig1.xml"));
        let flat = signed_document.to_flat_opc_bytes().unwrap();
        let imported_flat = Document::from_flat_opc_bytes(&flat).unwrap();
        let flat_reopened = OpcPackage::from_reader(std::io::Cursor::new(
            imported_flat
                .to_bytes_as(WordPackageClass::MacroEnabledTemplate)
                .unwrap(),
        ))
        .unwrap();
        assert!(has_package_signature_invalidation_marker(&flat_reopened));
    }

    #[test]
    fn unknown_word_main_content_type_fails_closed() {
        let mut unknown = source_package(WordPackageClass::Document);
        unknown
            .content_types
            .add_override("/word/document.xml", "application/x-unknown-word-main+xml");
        assert!(Document::from_bytes(&package_bytes(&unknown)).is_err());

        let mut fallback = source_package(WordPackageClass::Document);
        fallback
            .content_types
            .overrides
            .remove("/word/document.xml");
        assert!(Document::from_bytes(&package_bytes(&fallback)).is_err());

        let mut ambiguous = source_package(WordPackageClass::Document);
        ambiguous
            .package_rels
            .add(rel_types::DOCUMENT, "word/second.xml");
        ambiguous.set_part(
            "/word/second.xml",
            format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body/></w:document>"#).into_bytes(),
        );
        ambiguous
            .content_types
            .add_override("/word/second.xml", content_types::WORD_DOCUMENT);
        assert!(Document::from_bytes(&package_bytes(&ambiguous)).is_err());

        let mut external = source_package(WordPackageClass::Document);
        external.package_rels.items[0].target_mode = Some("External".to_owned());
        assert!(Document::from_bytes(&package_bytes(&external)).is_err());

        let mut unsafe_target = source_package(WordPackageClass::Document);
        unsafe_target.package_rels.items[0].target = "word/../document.xml".to_owned();
        assert!(Document::from_bytes(&package_bytes(&unsafe_target)).is_err());
    }

    #[test]
    fn office_document_relationship_modes_accept_internal_and_reject_others() {
        let mut internal = source_package(WordPackageClass::Document);
        internal.package_rels.items[0].target_mode = Some("Internal".to_owned());
        let opened = Document::from_bytes(&package_bytes(&internal)).unwrap();
        assert_eq!(opened.package_class().unwrap(), WordPackageClass::Document);

        for rejected_mode in ["External", "ProducerDefined"] {
            let mut rejected = source_package(WordPackageClass::Document);
            rejected.package_rels.items[0].target_mode = Some(rejected_mode.to_owned());
            assert!(Document::from_bytes(&package_bytes(&rejected)).is_err());
        }
    }

    #[test]
    fn malformed_or_unsafe_flat_opc_fails_before_document_publication() {
        let mut source = source_package(WordPackageClass::MacroEnabledDocument);
        source.set_part(
            "/word/sub/document.xml",
            format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body/></w:document>"#).into_bytes(),
        );
        source
            .content_types
            .add_override("/word/sub/document.xml", content_types::WORD_DOCUMENT);
        let document = Document::from_bytes(&package_bytes(&source)).unwrap();
        let valid = String::from_utf8(document.to_flat_opc_bytes().unwrap()).unwrap();
        let first_part_end = valid.find("</pkg:part>").unwrap() + "</pkg:part>".len();
        let duplicate = format!(
            "{}{}{}",
            &valid[..first_part_end],
            valid[..first_part_end]
                .rsplit_once("<pkg:part")
                .map(|(_, part)| format!("<pkg:part{part}"))
                .unwrap(),
            &valid[first_part_end..]
        );
        let case_variant_duplicate = duplicate.replacen("/_rels/.rels", "/_RELS/.RELS", 1);
        let mismatched_data = valid
            .replacen("<pkg:xmlData>", "<pkg:binaryData>", 1)
            .replacen("</pkg:xmlData>", "</pkg:binaryData>", 1);
        let extra_data = valid.replacen(
            "</pkg:xmlData></pkg:part>",
            "</pkg:xmlData><pkg:binaryData></pkg:binaryData></pkg:part>",
            1,
        );
        let first_part = &valid[valid.find("<pkg:part").unwrap()..first_part_end];
        let malformed_relationship = valid.replacen(
            "</pkg:package>",
            &format!(
                "{}</pkg:package>",
                first_part.replacen("/_rels/.rels", "/bad/_rels/.rels", 1)
            ),
            1,
        );
        let nested_relationship_filename = valid.replacen(
            "</pkg:package>",
            &format!(
                "{}</pkg:package>",
                first_part.replacen("/_rels/.rels", "/word/_rels/sub/document.xml.rels", 1,)
            ),
            1,
        );
        let mutations = [
            valid.replacen(
                "http://schemas.microsoft.com/office/2006/xmlPackage",
                "urn:wrong-package",
                1,
            ),
            valid.replacen("/word/document.xml", "/word/../document.xml", 1),
            valid.replacen("/word/document.xml", "/word/%2e%2e/document.xml", 1),
            valid.replacen("<pkg:binaryData>", "<pkg:binaryData>!", 1),
            mismatched_data,
            extra_data,
            duplicate,
            case_variant_duplicate,
            malformed_relationship,
            nested_relationship_filename,
            valid.replacen(
                "application/vnd.ms-office.vbaProject",
                "application/vnd.ms-office.vbaProject/extra",
                1,
            ),
            valid.replacen(
                "http://schemas.openxmlformats.org/package/2006/relationships",
                "urn:wrong-relationships",
                1,
            ),
        ];
        for malformed in mutations {
            assert!(Document::from_flat_opc_bytes(malformed.as_bytes()).is_err());
        }
        let lexical = valid.replacen("class fixture", "class\u{1}fixture", 1);
        assert!(Document::from_flat_opc_bytes(lexical.as_bytes()).is_err());
        for limits in [
            PackageReadLimits {
                max_entries: 1,
                max_part_uncompressed_bytes: u64::MAX,
                max_total_uncompressed_bytes: u64::MAX,
            },
            PackageReadLimits {
                max_entries: usize::MAX,
                max_part_uncompressed_bytes: 8,
                max_total_uncompressed_bytes: u64::MAX,
            },
            PackageReadLimits {
                max_entries: usize::MAX,
                max_part_uncompressed_bytes: u64::MAX,
                max_total_uncompressed_bytes: 8,
            },
        ] {
            assert!(Document::from_flat_opc_bytes_with_limits(valid.as_bytes(), limits).is_err());
        }
    }

    #[test]
    fn flat_opc_and_package_class_path_apis_publish_reopenable_files() {
        let document = Document::from_bytes(&package_bytes(&source_package(
            WordPackageClass::MacroEnabledDocument,
        )))
        .unwrap();
        let directory = std::env::temp_dir().join(format!(
            "rdocx-f238-paths-{}-{:?}",
            std::process::id(),
            std::thread::current().id()
        ));
        std::fs::create_dir_all(&directory).unwrap();
        let flat_path = directory.join("document.xml");
        let template_path = directory.join("extension-is-not-authority.docx");
        document.save_flat_opc(&flat_path).unwrap();
        let reopened = Document::open_flat_opc(&flat_path).unwrap();
        assert_eq!(
            reopened.package_class().unwrap(),
            WordPackageClass::MacroEnabledDocument
        );
        reopened
            .save_as_package_class(&template_path, WordPackageClass::MacroEnabledTemplate)
            .unwrap();
        assert_eq!(
            Document::open(&template_path)
                .unwrap()
                .package_class()
                .unwrap(),
            WordPackageClass::MacroEnabledTemplate
        );
        std::fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    #[ignore = "requires installed Microsoft Word 16.104 GUI automation"]
    fn flat_opc_and_modern_word_package_classes_open_in_pinned_word_without_repair() {
        let plist = "/Applications/Microsoft Word.app/Contents/Info.plist";
        let version = std::process::Command::new("plutil")
            .args(["-extract", "CFBundleShortVersionString", "raw", plist])
            .output()
            .unwrap();
        let build = std::process::Command::new("plutil")
            .args(["-extract", "CFBundleVersion", "raw", plist])
            .output()
            .unwrap();
        assert_eq!(String::from_utf8_lossy(&version.stdout).trim(), "16.104");
        assert_eq!(
            String::from_utf8_lossy(&build.stdout).trim(),
            "16.104.25121423"
        );
        assert_eq!(
            WORD_ORACLE_VERSION,
            "Microsoft Word 16.104 build 16.104.25121423"
        );

        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let directory = std::path::Path::new(
            "/Users/atulsharma/Library/Containers/com.microsoft.Word/Data/Documents/rdocx-f238-word-oracle",
        )
        .join(format!("{}-{nonce}", std::process::id()));
        std::fs::create_dir_all(&directory).unwrap();
        let document = Document::new();
        let mut paths = Vec::new();
        for (class, extension) in [
            (WordPackageClass::Document, "docx"),
            (WordPackageClass::MacroEnabledDocument, "docm"),
            (WordPackageClass::Template, "dotx"),
            (WordPackageClass::MacroEnabledTemplate, "dotm"),
        ] {
            let path = directory.join(format!("candidate.{extension}"));
            document.save_as_package_class(&path, class).unwrap();
            paths.push(path);
        }
        let flat_path = directory.join("candidate.xml");
        document.save_flat_opc(&flat_path).unwrap();
        paths.push(flat_path);

        for path in &paths {
            let script = format!(
                r#"with timeout of 60 seconds
tell application "Microsoft Word"
activate
set candidatePath to (POSIX file "{}") as text
set candidateDocument to open file name candidatePath read only true add to recent files false
delay 1
close candidateDocument saving no
end tell
end timeout"#,
                path.display()
            );
            let opened = std::process::Command::new("osascript")
                .args(["-e", &script])
                .output()
                .unwrap();
            assert!(
                opened.status.success(),
                "Word rejected {}: {}",
                path.display(),
                String::from_utf8_lossy(&opened.stderr)
            );
        }
        std::fs::remove_dir_all(directory).unwrap();
    }
}

#[test]
fn encoded_drawing_relationship_ids_reopen_extract_and_render() {
    let image = mhtml_pixel_png();
    let mut source = Document::new();
    source.add_picture(
        &image,
        "encoded.png",
        Length::inches(1.0),
        Length::inches(1.0),
    );
    source
        .add_chart(
            oxml_chart::ChartKind::Bar,
            Length::inches(3.0),
            Length::inches(2.0),
            &oxml_chart::ChartData {
                categories: vec!["North".to_owned(), "South".to_owned()],
                series: vec![("Revenue".to_owned(), vec![12.0, 18.0])],
                number_format: Some("0".to_owned()),
                ..oxml_chart::ChartData::default()
            },
        )
        .unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(source.to_bytes().unwrap()))
        .expect("authored drawing package");
    let relationships = package
        .get_part_rels("/word/document.xml")
        .expect("document relationships");
    let image_id = relationships
        .get_by_type(rel_types::IMAGE)
        .expect("image relationship")
        .id
        .clone();
    let chart_id = relationships
        .get_by_type(rel_types::CHART)
        .expect("chart relationship")
        .id
        .clone();
    let document_xml = String::from_utf8(
        package
            .get_part("/word/document.xml")
            .expect("document part")
            .to_vec(),
    )
    .unwrap()
    .replace(
        &format!(r#"r:embed="{image_id}""#),
        &format!(r#"r:embed="{}""#, image_id.replacen('I', "&#73;", 1)),
    )
    .replace(
        &format!(r#"r:id="{chart_id}""#),
        &format!(r#"r:id="{}""#, chart_id.replacen('I', "&#x49;", 1)),
    );
    assert!(document_xml.contains("r&#73;d"));
    assert!(document_xml.contains("r&#x49;d"));
    package.set_part("/word/document.xml", document_xml.into_bytes());

    let mut producer_bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut producer_bytes).unwrap();
    let mut document =
        Document::from_bytes(producer_bytes.get_ref()).expect("encoded drawing package reopens");
    assert_eq!(document.images()[0].embed_id, image_id);
    assert_eq!(document.image_data(&image_id), Some(image.clone()));

    let page = document
        .layout_page(0)
        .unwrap()
        .expect("encoded drawings produce a page");
    fn has_group(elements: &[oxml_layout::PositionedElement]) -> bool {
        elements.iter().any(|element| match element {
            oxml_layout::PositionedElement::Group(_) => true,
            oxml_layout::PositionedElement::MarkedContent { children, .. } => has_group(children),
            _ => false,
        })
    }
    let mut images = 0;
    let mut paths = 0;
    oxml_layout::walk(&page.elements, &mut |element, _| match element {
        oxml_layout::PositionedElement::Image { .. } => images += 1,
        oxml_layout::PositionedElement::Path(_) => paths += 1,
        _ => {}
    });
    assert_eq!(images, 1);
    assert!(
        has_group(&page.elements),
        "the decoded chart relationship should render a group"
    );
    assert!(paths > 0, "the decoded chart should render vector geometry");
    let rendered = document
        .render_page_to_png_deterministic(0, 72.0)
        .unwrap()
        .expect("encoded drawing page renders");
    assert!(rendered.starts_with(b"\x89PNG\r\n\x1a\n"));

    let saved = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&saved).expect("saved drawing package reopens");
    assert_eq!(reopened.images()[0].embed_id, image_id);
    assert_eq!(reopened.image_data(&image_id), Some(image));
}

fn normalized_mhtml_record(
    document: &Document,
    diagnostics: &[MhtmlDiagnostic],
) -> MhtmlOracleRecord {
    let paragraphs = document.paragraphs();
    let bold_text = paragraphs
        .iter()
        .flat_map(|paragraph| paragraph.runs())
        .filter(|run| run.bold_value() == Some(true) || run.style_id() == Some("Strong"))
        .map(|run| run.text())
        .collect();
    let numbered_text = paragraphs
        .iter()
        .filter(|paragraph| paragraph.numbering().is_some())
        .map(|paragraph| paragraph.text())
        .collect();
    let mut table_cells = Vec::new();
    for table_index in 0..document.table_count() {
        let table = document.table(table_index).unwrap();
        let mut rows = Vec::new();
        for row_index in 0..table.row_count() {
            let row = table.row(row_index).unwrap();
            rows.push(
                (0..row.cell_count())
                    .map(|cell_index| row.cell(cell_index).unwrap().text())
                    .collect(),
            );
        }
        table_cells.push(rows);
    }
    MhtmlOracleRecord {
        text: document.text().trim_end().to_owned(),
        bold_text,
        numbered_text,
        table_cells,
        image_sizes: document
            .images()
            .iter()
            .map(|image| (image.width_emu, image.height_emu))
            .collect(),
        links: document
            .links()
            .into_iter()
            .map(|link| (link.text.trim().to_owned(), link.url, link.anchor))
            .collect(),
        diagnostics: diagnostics.to_vec(),
    }
}

fn pinned_rdocx_mhtml_record() -> MhtmlOracleRecord {
    MhtmlOracleRecord {
        text: "Oracle title\nbold link\none\ntwo\ncell".to_owned(),
        bold_text: vec!["bold".to_owned()],
        numbered_text: vec!["one".to_owned(), "two".to_owned()],
        table_cells: vec![vec![vec!["cell".to_owned()]]],
        image_sizes: vec![(19_050, 28_575)],
        links: vec![(
            "link".to_owned(),
            Some("https://example.test/".to_owned()),
            None,
        )],
        diagnostics: Vec::new(),
    }
}

fn pinned_word_mhtml_record() -> MhtmlOracleRecord {
    MhtmlOracleRecord {
        image_sizes: Vec::new(),
        ..pinned_rdocx_mhtml_record()
    }
}

fn mhtml_oracle_accepts(rdocx: &MhtmlOracleRecord, word: &MhtmlOracleRecord) -> bool {
    // Word 16.104 drops this contained PNG while rdocx retains it by contract.
    // Compare every shared field, then assert each side of that pinned difference.
    let mut common_rdocx = rdocx.clone();
    common_rdocx.image_sizes.clear();
    common_rdocx == *word
        && *rdocx == pinned_rdocx_mhtml_record()
        && *word == pinned_word_mhtml_record()
}

fn mhtml_pixel_png() -> Vec<u8> {
    vec![
        137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6,
        0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 8, 29, 99, 96, 96, 96, 248, 15, 0,
        1, 4, 1, 0, 30, 115, 156, 64, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
    ]
}

fn source_built_mhtml_with_pixel(html: &str) -> Vec<u8> {
    use base64::Engine as _;
    let encoded = base64::engine::general_purpose::STANDARD.encode(mhtml_pixel_png());
    format!(
        "MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary=integration; start=\"<root@rdocx>\"\r\n\r\n--integration\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root@rdocx>\r\nContent-Location: https://example.test/index.html\r\n\r\n{html}\r\n--integration\r\nContent-Type: image/png\r\nContent-ID: <pixel@rdocx>\r\nContent-Location: https://example.test/pixel.png\r\nContent-Transfer-Encoding: base64\r\n\r\n{encoded}\r\n--integration--\r\n"
    )
    .into_bytes()
}

fn mhtml_word_oracle_source() -> Vec<u8> {
    source_built_mhtml_with_pixel(MHTML_ORACLE_HTML)
}

#[test]
fn mhtml_import_and_export_preserve_supported_word_structure() {
    use base64::Engine as _;
    let png = mhtml_pixel_png();
    let encoded = base64::engine::general_purpose::STANDARD.encode(&png);
    let html = "<h1>Title</h1><p><strong>body</strong> <a href='https://example.test/'>link</a><img src='cid:pixel@rdocx' width='2' height='3'></p><ol><li>one</li><li>two</li></ol><table><tr><th colspan='2'>head</th></tr><tr><td>a</td><td>b</td></tr></table>";
    let input = format!(
        "MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary=integration; start=\"<root@rdocx>\"\r\n\r\n--integration\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root@rdocx>\r\nContent-Location: https://example.test/index.html\r\n\r\n{html}\r\n--integration\r\nContent-Type: image/png\r\nContent-ID: <pixel@rdocx>\r\nContent-Location: pixel.png\r\nContent-Transfer-Encoding: base64\r\n\r\n{encoded}\r\n--integration--\r\n"
    );
    let imported = Document::from_mhtml_bytes(input.as_bytes()).expect("source-built MHTML");
    assert_eq!(
        imported.document.text(),
        "Title\nbody link\none\ntwo\nhead\t\na\tb\t\n"
    );
    assert_eq!(
        imported
            .document
            .paragraph(1)
            .unwrap()
            .run(0)
            .unwrap()
            .bold_value(),
        Some(true)
    );
    let first_list = imported.document.paragraph(2).unwrap().numbering().unwrap();
    let second_list = imported.document.paragraph(3).unwrap().numbering().unwrap();
    assert_eq!(first_list, second_list);
    assert_eq!(
        imported
            .document
            .table(0)
            .unwrap()
            .cell(0, 0)
            .unwrap()
            .grid_span(),
        Some(2)
    );
    assert_eq!(imported.document.images()[0].width_emu, 19_050);
    assert_eq!(imported.document.images()[0].height_emu, 28_575);
    assert_eq!(
        imported
            .document
            .image_data(&imported.document.images()[0].embed_id),
        Some(png)
    );
    assert_eq!(
        imported.document.links()[0].url.as_deref(),
        Some("https://example.test/")
    );
    let written = imported.document.to_mhtml_bytes().expect("MHTML export");
    let reopened = Document::from_mhtml_bytes(&written.bytes).expect("MHTML reimport");
    assert_eq!(reopened.document.text(), imported.document.text());
    assert_eq!(reopened.document.links().len(), 1);
    assert_eq!(reopened.document.images()[0].width_emu, 19_050);
    assert_eq!(reopened.document.images()[0].height_emu, 28_575);
    let mut docx_candidate = reopened.document;
    let docx = docx_candidate.to_bytes().expect("projected DOCX");
    Document::from_bytes(&docx).expect("projected DOCX reopens");

    let directory = std::env::temp_dir().join(format!("rdocx-mhtml-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&directory);
    std::fs::create_dir(&directory).unwrap();
    let input_path = directory.join("input.mhtml");
    std::fs::write(&input_path, input).unwrap();
    assert_eq!(
        Document::open_mhtml(&input_path).unwrap().document.text(),
        imported.document.text()
    );
    let output_path = directory.join("output.mhtml");
    std::fs::write(&output_path, b"old incomplete bytes").unwrap();
    assert!(
        imported
            .document
            .save_mhtml(&output_path)
            .unwrap()
            .is_empty()
    );
    assert!(Document::open_mhtml(&output_path).is_ok());
    std::fs::remove_dir_all(directory).unwrap();
}

#[test]
fn mhtml_conversions_match_the_pinned_word_structure() {
    assert_eq!(
        MHTML_ORACLE_VERSION,
        "Microsoft Word 16.104 build 16.104.25121423"
    );
    let imported = Document::from_mhtml_bytes(&mhtml_word_oracle_source())
        .expect("source-built MHTML imported by rdocx");
    let word_record = pinned_word_mhtml_record();
    let rdocx_record = normalized_mhtml_record(&imported.document, &imported.diagnostics);
    assert!(mhtml_oracle_accepts(&rdocx_record, &word_record));

    let perturbations = [
        MHTML_ORACLE_HTML.replace("Oracle title", "Changed title"),
        MHTML_ORACLE_HTML.replace("<strong>bold</strong>", "<span>bold</span>"),
        MHTML_ORACLE_HTML.replace("<td>cell</td>", "<td>changed</td>"),
        MHTML_ORACLE_HTML.replace("<li>two</li>", ""),
        MHTML_ORACLE_HTML.replace(
            "href='https://example.test/'",
            "href='https://changed.test/'",
        ),
        MHTML_ORACLE_HTML.replace(
            "<img src='https://example.test/pixel.png' width='2' height='3'>",
            "",
        ),
        MHTML_ORACLE_HTML.replace("<strong>", "<object></object><strong>"),
    ];
    for html in perturbations {
        let candidate = Document::from_mhtml_bytes(&source_built_mhtml_with_pixel(&html)).unwrap();
        let candidate = normalized_mhtml_record(&candidate.document, &candidate.diagnostics);
        assert!(
            !mhtml_oracle_accepts(&candidate, &word_record),
            "accepted perturbed MHTML source {html:?}"
        );
    }

    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    package.set_part(
        "/word/document.xml",
        br#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:producer"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p><x:raw keep="yes"/><w:p><w:r><w:t>after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#.to_vec(),
    );
    let mut packaged = std::io::Cursor::new(Vec::new());
    package.write_to(&mut packaged).unwrap();
    let lossy = Document::from_bytes(packaged.get_ref()).unwrap();
    let written = lossy.to_mhtml_bytes().unwrap();
    assert_eq!(written.diagnostics.len(), 1);
    assert_eq!(written.diagnostics[0].location, "body[1]");
    assert_eq!(
        Document::from_mhtml_bytes(&written.bytes)
            .unwrap()
            .document
            .text(),
        "before\nafter\n"
    );
}

#[test]
#[ignore = "requires pinned Microsoft Word 16.104 for oracle regeneration"]
fn regenerate_mhtml_word_oracle_authenticates_exact_build() {
    let plist = "/Applications/Microsoft Word.app/Contents/Info.plist";
    let version = std::process::Command::new("plutil")
        .args(["-extract", "CFBundleShortVersionString", "raw", plist])
        .output()
        .unwrap();
    let build = std::process::Command::new("plutil")
        .args(["-extract", "CFBundleVersion", "raw", plist])
        .output()
        .unwrap();
    assert_eq!(String::from_utf8_lossy(&version.stdout).trim(), "16.104");
    assert_eq!(
        String::from_utf8_lossy(&build.stdout).trim(),
        "16.104.25121423"
    );
    let source = format!(
        "/private/tmp/F-239-word-16.104-{}-source.mhtml",
        std::process::id()
    );
    let output = format!(
        "/private/tmp/F-239-word-16.104-{}-output.docx",
        std::process::id()
    );
    std::fs::write(&source, mhtml_word_oracle_source()).unwrap();
    let script = format!(
        r#"with timeout of 120 seconds
tell application "Microsoft Word"
activate
open POSIX file "{source}"
delay 3
set oracleDocument to active document
save as oracleDocument file name "{output}" file format format document default add to recent files false
close oracleDocument saving no
end tell
end timeout
"#
    );
    let conversion = std::process::Command::new("osascript")
        .args(["-e", &script])
        .output()
        .unwrap();
    assert!(
        conversion.status.success(),
        "Word conversion failed: {}",
        String::from_utf8_lossy(&conversion.stderr)
    );
    let oracle = Document::open(&output).expect("Word-produced DOCX");
    let imported = Document::from_mhtml_bytes(&mhtml_word_oracle_source())
        .expect("same source imported by rdocx");
    let word_record = normalized_mhtml_record(&oracle, &[]);
    let rdocx_record = normalized_mhtml_record(&imported.document, &imported.diagnostics);
    assert_eq!(word_record, pinned_word_mhtml_record());
    assert_eq!(rdocx_record, pinned_rdocx_mhtml_record());
    assert!(mhtml_oracle_accepts(&rdocx_record, &word_record));
    std::fs::remove_file(source).unwrap();
    std::fs::remove_file(output).unwrap();
}

#[derive(Debug, PartialEq)]
struct OdtStructuralRecord {
    body_order: Vec<String>,
    paragraphs: Vec<OdtOracleParagraphRecord>,
    tables: Vec<OdtOracleTableRecord>,
    images: Vec<(i64, i64)>,
    media: Vec<Vec<u8>>,
}

#[test]
fn paragraph_items_keep_runs_equations_controls_and_raw_xml_in_source_order() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let math_namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:q="{math_namespace}" xmlns:x="urn:producer"><w:body><w:p><w:r><w:t>before</w:t></w:r><q:oMath><q:r><q:t>x</q:t></q:r></q:oMath><q:oMathPara><q:oMathParaPr><q:jc q:val="centerGroup"/></q:oMathParaPr><q:oMath><q:r><q:t>display one</q:t></q:r></q:oMath><q:oMath><q:r><q:t>display two</q:t></q:r></q:oMath></q:oMathPara><w:sdt><w:sdtContent><w:r><w:t>control</w:t></w:r></w:sdtContent></w:sdt><x:raw keep="yes"/><w:r><w:t>after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#
    );
    package.set_part("/word/document.xml", xml.into_bytes());
    let mut saved = std::io::Cursor::new(Vec::new());
    package.write_to(&mut saved).unwrap();
    let document = Document::from_bytes(saved.get_ref()).unwrap();
    let paragraph = document.paragraph(0).unwrap();
    let items = paragraph.items().collect::<Vec<_>>();
    assert!(matches!(items[0], rdocx::ParagraphItemRef::Run(_)));
    assert!(matches!(items[1], rdocx::ParagraphItemRef::Equation(_)));
    assert!(matches!(items[2], rdocx::ParagraphItemRef::Equation(_)));
    assert!(matches!(
        items[3],
        rdocx::ParagraphItemRef::ContentControl(_)
    ));
    assert!(matches!(
        items[4],
        rdocx::ParagraphItemRef::UnsupportedXml(_)
    ));
    assert!(matches!(items[5], rdocx::ParagraphItemRef::Run(_)));
    let rdocx::OfficeMath::Display(display) = paragraph.equation(1).unwrap() else {
        panic!("display equation")
    };
    assert_eq!(display.equations.len(), 2);
    assert_eq!(
        display.properties.justification,
        Some(rdocx::MathJustification::CenterGroup)
    );
}

#[test]
fn public_equation_authoring_saves_reopens_and_remains_mutable() {
    let mut document = Document::new();
    let mut paragraph = document.add_paragraph("before");
    let text = |value| rdocx::MathArgument::text(value);
    paragraph
        .add_equation(rdocx::OfficeMath::inline(vec![
            rdocx::MathRun::new("x").into(),
            rdocx::MathExpression::Fraction(rdocx::MathFraction::new(text("1"), text("2"))),
            rdocx::MathExpression::Subscript(rdocx::MathScript::new(text("x"), text("i"))),
            rdocx::MathExpression::Superscript(rdocx::MathScript::new(text("x"), text("2"))),
            rdocx::MathExpression::SubSuperscript(rdocx::MathSubSuperscript::new(
                text("x"),
                text("i"),
                text("2"),
            )),
            rdocx::MathExpression::PreSubSuperscript(rdocx::MathPreSubSuperscript::new(
                text("x"),
                text("i"),
                text("2"),
            )),
            rdocx::MathExpression::Radical(rdocx::MathRadical::with_degree(text("3"), text("x"))),
            rdocx::MathExpression::Matrix(rdocx::MathMatrix::new(vec![rdocx::MathMatrixRow::new(
                vec![text("a"), text("b")],
            )])),
            rdocx::MathExpression::LowerLimit(rdocx::MathLimit::new(text("lim"), text("0"))),
            rdocx::MathExpression::UpperLimit(rdocx::MathLimit::new(text("max"), text("n"))),
            rdocx::MathExpression::Nary(rdocx::MathNary::new("∑", text("x"))),
            rdocx::MathExpression::Delimiter(rdocx::MathDelimiter::new("(", ")", vec![text("x")])),
            rdocx::MathExpression::Accent(rdocx::MathAccent::new("̂", text("x"))),
        ]))
        .unwrap();
    let mut display = rdocx::CT_OMathPara::new(vec![
        rdocx::CT_OMath::new(vec![rdocx::MathRun::new("display one").into()]),
        rdocx::CT_OMath::new(vec![rdocx::MathRun::new("display two").into()]),
    ]);
    display.properties.justification = Some(rdocx::MathJustification::CenterGroup);
    paragraph
        .add_equation(rdocx::OfficeMath::Display(display))
        .unwrap();
    paragraph.add_run("after");
    let mut defaults = rdocx::MathProperties::new();
    defaults.math_font = Some("Cambria Math".to_owned());
    defaults.justification = Some(rdocx::MathJustification::CenterGroup);
    document.set_math_properties(defaults).unwrap();

    let bytes = document.to_bytes().unwrap();
    let mut reopened = Document::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened.math_properties().unwrap().math_font.as_deref(),
        Some("Cambria Math")
    );
    let mut paragraph = reopened.paragraph_mut(0).unwrap();
    let equation = paragraph.equation_mut(0).unwrap();
    let rdocx::OfficeMath::Inline(equation) = equation else {
        panic!("inline equation")
    };
    assert_eq!(equation.expressions.len(), 13);
    let rdocx::MathExpression::Fraction(fraction) = &mut equation.expressions[1] else {
        panic!("fraction")
    };
    fraction.fraction_type = rdocx::FractionType::Linear;
    let rdocx::OfficeMath::Display(display) = paragraph.equation_mut(1).unwrap() else {
        panic!("display equation")
    };
    assert_eq!(display.equations.len(), 2);
    assert_eq!(
        display.properties.justification,
        Some(rdocx::MathJustification::CenterGroup)
    );
    let rdocx::MathExpression::Run(run) = &mut display.equations[1].expressions[0] else {
        panic!("display math run")
    };
    run.text = "changed display".to_owned();
    let mutated = reopened.to_bytes().unwrap();
    let final_document = Document::from_bytes(&mutated).unwrap();
    let rdocx::OfficeMath::Inline(equation) =
        final_document.paragraph(0).unwrap().equation(0).unwrap()
    else {
        panic!("inline equation")
    };
    assert_eq!(equation.expressions.len(), 13);
    let rdocx::MathExpression::Fraction(fraction) = &equation.expressions[1] else {
        panic!("fraction")
    };
    assert_eq!(fraction.fraction_type, rdocx::FractionType::Linear);
    let rdocx::OfficeMath::Display(display) =
        final_document.paragraph(0).unwrap().equation(1).unwrap()
    else {
        panic!("display equation")
    };
    let rdocx::MathExpression::Run(run) = &display.equations[1].expressions[0] else {
        panic!("display math run")
    };
    assert_eq!(run.text, "changed display");
}

#[derive(Debug, PartialEq)]
struct OdtOracleParagraphRecord {
    text: String,
    alignment: String,
    numbering: Option<(bool, u32)>,
    runs: Vec<(String, bool, bool, Option<String>)>,
}

#[derive(Debug, PartialEq)]
struct OdtOracleTableRecord {
    rows: usize,
    columns: usize,
    cells: Vec<(String, Option<u32>, Option<String>)>,
}

#[derive(Debug, PartialEq)]
struct OdtSupportedRecord {
    body_order: Vec<String>,
    paragraphs: Vec<OdtParagraphRecord>,
    tables: Vec<OdtTableRecord>,
    images: Vec<(i64, i64)>,
    media: Vec<Vec<u8>>,
}

#[test]
fn svg_export_preserves_searchable_text_geometry_fonts_images_links_and_clips() {
    let mut document = Document::new();
    document.add_paragraph("Searchable SVG text");
    document.append_hyperlink("safe link", "https://example.test/?a=1&b=2");
    document.add_picture(
        PNG_2_BY_3,
        "two-by-three.png",
        Length::emu(914_400),
        Length::emu(1_371_600),
    );
    {
        let mut table = document.add_table(1, 1);
        table.cell(0, 0).unwrap().set_text("clipped cell");
        table.row(0).unwrap().set_height_exact(Length::pt(8.0));
    }

    let source_before_render = document.to_bytes().unwrap();
    let result = document
        .render_page_to_svg_deterministic(0)
        .expect("deterministic SVG layout")
        .expect("page zero");
    assert_eq!(document.to_bytes().unwrap(), source_before_render);
    let mut reader = quick_xml::Reader::from_str(&result.svg);
    loop {
        if matches!(
            reader.read_event().expect("generated SVG is parseable XML"),
            quick_xml::events::Event::Eof
        ) {
            break;
        }
    }
    assert!(
        result
            .svg
            .starts_with("<svg xmlns=\"http://www.w3.org/2000/svg\"")
    );
    assert!(result.svg.contains("viewBox=\"0 0 612 792\""));
    assert!(result.svg.contains(">Searchable </text>"));
    assert!(result.svg.contains(">SVG </text>"));
    assert!(result.svg.contains(">text</text>"));
    assert!(result.svg.contains("@font-face"));
    assert!(result.svg.contains("data:font/ttf;base64,"));
    assert!(result.svg.contains("data:image/png;base64,"));
    assert!(result.svg.contains("clip-path=\"url(#rdocx-def-"));
    assert!(
        result
            .svg
            .contains("href=\"https://example.test/?a=1&amp;b=2\"")
    );
    assert!(!result.svg.contains("file://"));
    assert!(!result.svg.contains("href=\"http://"));
    assert!(!result.svg.contains("url('http"));
}

#[derive(Debug, PartialEq)]
struct OdtParagraphRecord {
    text: String,
    alignment: String,
    numbering: Option<(bool, u32)>,
    space_before: Option<i32>,
    space_after: Option<i32>,
    indent_left: Option<i32>,
    indent_right: Option<i32>,
    first_line_indent: Option<i32>,
    line_spacing: Option<(String, i32)>,
    runs: Vec<OdtRunRecord>,
}

#[derive(Debug, PartialEq)]
struct OdtRunRecord {
    text: String,
    font: Option<String>,
    size_half_points: Option<u32>,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    color: Option<String>,
    highlight: Option<String>,
    vertical: Option<String>,
}

#[derive(Debug, PartialEq)]
struct OdtTableRecord {
    rows: usize,
    columns: usize,
    cells: Vec<(Vec<OdtParagraphRecord>, Option<u32>, Option<String>)>,
}

fn odt_paragraph_record(
    document: &Document,
    paragraph: ParagraphRef<'_>,
    numbering_kinds: &HashMap<(u32, u32), bool>,
) -> OdtParagraphRecord {
    let paragraph_style = document.resolve_paragraph_properties(paragraph.style_id());
    let alignment = paragraph
        .alignment()
        .map(|value| format!("{value:?}"))
        .or_else(|| paragraph_style.jc.map(|value| format!("{value:?}")))
        .unwrap_or_else(|| "Left".to_string());
    let numbering = paragraph.numbering().and_then(|(id, level)| {
        numbering_kinds
            .get(&(id, level))
            .map(|bullet| (*bullet, level))
    });
    let direct_line = paragraph
        .line_spacing_multiple()
        .map(|value| ("auto".to_string(), (value * 240.0).round() as i32))
        .or_else(|| {
            paragraph
                .line_spacing()
                .map(|value| ("exact".to_string(), value.to_twips()))
        });
    let style_line = paragraph_style.line_spacing.map(|value| {
        (
            paragraph_style
                .line_rule
                .clone()
                .unwrap_or_else(|| "auto".to_string()),
            value.0,
        )
    });
    let runs = (0..paragraph.run_count())
        .filter_map(|index| paragraph.run(index))
        .filter(|run| !run.text().is_empty())
        .map(|run| {
            let resolved = document.resolve_run_properties(paragraph.style_id(), run.style_id());
            OdtRunRecord {
                text: run.text(),
                font: run.font_name().map(str::to_string).or_else(|| {
                    resolved
                        .font_ascii
                        .or(resolved.font_hansi)
                        .or(resolved.font_east_asia)
                        .or(resolved.font_cs)
                }),
                size_half_points: run
                    .size()
                    .map(|value| (value * 2.0).round() as u32)
                    .or_else(|| resolved.sz.map(|value| value.0)),
                bold: run.bold_value().or(resolved.bold).unwrap_or(false),
                italic: run.italic_value().or(resolved.italic).unwrap_or(false),
                underline: run
                    .underline_code_value()
                    .map(|value| value != 0)
                    .or_else(|| resolved.underline.map(|value| value.to_str() != "none"))
                    .unwrap_or(false),
                strike: run.strike_value().or(resolved.strike).unwrap_or(false),
                color: run
                    .color()
                    .map(str::to_string)
                    .or(resolved.color)
                    .or_else(|| Some("000000".to_string())),
                highlight: run.highlight().or_else(|| {
                    resolved
                        .shading
                        .and_then(|shading| shading.fill)
                        .or_else(|| resolved.highlight.map(|value| value.to_str().to_string()))
                }),
                vertical: run.vert_align().map(str::to_string).or(resolved.vert_align),
            }
        })
        .collect();
    OdtParagraphRecord {
        text: paragraph.text(),
        alignment,
        numbering,
        space_before: paragraph
            .space_before()
            .map(Length::to_twips)
            .or_else(|| paragraph_style.space_before.map(|value| value.0)),
        space_after: paragraph
            .space_after()
            .map(Length::to_twips)
            .or_else(|| paragraph_style.space_after.map(|value| value.0)),
        indent_left: paragraph
            .indent_left()
            .map(Length::to_twips)
            .or_else(|| paragraph_style.ind_left.map(|value| value.0)),
        indent_right: paragraph
            .indent_right()
            .map(Length::to_twips)
            .or_else(|| paragraph_style.ind_right.map(|value| value.0)),
        first_line_indent: paragraph
            .first_line_indent()
            .map(Length::to_twips)
            .or_else(|| {
                paragraph_style
                    .ind_first_line
                    .map(|value| value.0)
                    .or_else(|| {
                        paragraph_style
                            .ind_hanging
                            .map(|value| value.0.saturating_neg())
                    })
            }),
        line_spacing: direct_line.or(style_line),
        runs,
    }
}

fn source_built_odt() -> Vec<u8> {
    use std::io::Write;
    use zip::write::SimpleFileOptions;
    use zip::{CompressionMethod, ZipWriter};

    let content = br##"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3">
<office:automatic-styles>
<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>
<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:font-style="italic" fo:color="#123456"/></style:style>
<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1"/></text:list-style>
</office:automatic-styles>
<office:body><office:text>
<text:p text:style-name="P1">Alpha <text:span text:style-name="T1">formatted</text:span></text:p>
<text:list text:style-name="L1"><text:list-item><text:p>one</text:p></text:list-item></text:list>
<table:table table:name="Table1"><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell table:number-columns-spanned="2"><text:p>wide</text:p><text:p>second</text:p></table:table-cell><table:covered-table-cell/></table:table-row><table:table-row><table:table-cell><text:p>left</text:p></table:table-cell><table:table-cell><text:p>right</text:p></table:table-cell></table:table-row></table:table>
<text:p><draw:frame draw:name="Picture1" svg:width="1in" svg:height="0.5in"><draw:image xlink:href="Pictures/pixel.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></text:p>
</office:text></office:body></office:document-content>"##;
    let styles = br#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3"><office:styles><style:default-style style:family="paragraph"><style:text-properties fo:font-size="11pt"/></style:default-style></office:styles></office:document-styles>"#;
    let manifest = br#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Pictures/pixel.png" manifest:media-type="image/png"/></manifest:manifest>"#;

    let mut output = std::io::Cursor::new(Vec::new());
    let mut archive = ZipWriter::new(&mut output);
    for (name, bytes, method) in [
        (
            "mimetype",
            b"application/vnd.oasis.opendocument.text".as_slice(),
            CompressionMethod::Stored,
        ),
        (
            "content.xml",
            content.as_slice(),
            CompressionMethod::Deflated,
        ),
        ("styles.xml", styles.as_slice(), CompressionMethod::Deflated),
        (
            "META-INF/manifest.xml",
            manifest.as_slice(),
            CompressionMethod::Deflated,
        ),
        (
            "Pictures/pixel.png",
            PNG_2_BY_3,
            CompressionMethod::Deflated,
        ),
    ] {
        archive
            .start_file(
                name,
                SimpleFileOptions::default().compression_method(method),
            )
            .unwrap();
        archive.write_all(bytes).unwrap();
    }
    archive.finish().unwrap();
    output.into_inner()
}

fn odt_supported_record(mut document: Document) -> OdtSupportedRecord {
    let bytes = document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let numbering = package
        .get_part("/word/numbering.xml")
        .map(rdocx_oxml::numbering::CT_Numbering::from_xml)
        .transpose()
        .unwrap();
    let mut numbering_kinds = HashMap::new();
    if let Some(numbering) = numbering.as_ref() {
        for instance in &numbering.nums {
            let Some(definition) = numbering.get_abstract_num_for(instance.num_id) else {
                continue;
            };
            for level in &definition.levels {
                if let Some(format) = level.num_fmt.as_ref() {
                    numbering_kinds.insert(
                        (instance.num_id, level.ilvl),
                        matches!(format, rdocx_oxml::numbering::ST_NumberFormat::Bullet),
                    );
                }
            }
        }
    }
    let body_order = document
        .body_items()
        .map(|item| match item {
            BodyItemRef::Paragraph(paragraph) => format!("p:{}", paragraph.text()),
            BodyItemRef::Table(table) => {
                format!("table:{}x{}", table.row_count(), table.column_count())
            }
            BodyItemRef::ContentControl(_) => "content-control".to_string(),
            BodyItemRef::UnsupportedXml(_) => "unsupported".to_string(),
        })
        .collect();
    let paragraphs = document
        .paragraphs()
        .into_iter()
        .map(|paragraph| odt_paragraph_record(&document, paragraph, &numbering_kinds))
        .collect();
    let tables = (0..document.table_count())
        .map(|index| {
            let table = document.table(index).unwrap();
            let mut cells = Vec::new();
            for row_index in 0..table.row_count() {
                let row = table.row(row_index).unwrap();
                for cell_index in 0..row.cell_count() {
                    let cell = row.cell(cell_index).unwrap();
                    cells.push((
                        cell.paragraphs()
                            .map(|paragraph| {
                                odt_paragraph_record(&document, paragraph, &numbering_kinds)
                            })
                            .collect(),
                        cell.grid_span(),
                        cell.v_merge().map(|value| format!("{value:?}")),
                    ));
                }
            }
            OdtTableRecord {
                rows: table.row_count(),
                columns: table.column_count(),
                cells,
            }
        })
        .collect();
    let images = document
        .images()
        .into_iter()
        .map(|image| (image.width_emu, image.height_emu))
        .collect();
    let mut media: Vec<(String, Vec<u8>)> = package
        .parts
        .iter()
        .filter(|(name, _)| name.starts_with("/word/media/"))
        .map(|(name, bytes)| (name.clone(), bytes.clone()))
        .collect();
    media.sort_by(|left, right| left.0.cmp(&right.0));
    OdtSupportedRecord {
        body_order,
        paragraphs,
        tables,
        images,
        media: media.into_iter().map(|(_, bytes)| bytes).collect(),
    }
}

fn odt_structural_record(mut document: Document) -> OdtStructuralRecord {
    let body_order = document
        .body_items()
        .map(|item| match item {
            BodyItemRef::Paragraph(paragraph) => format!("p:{}", paragraph.text()),
            BodyItemRef::Table(table) => {
                format!("table:{}x{}", table.row_count(), table.column_count())
            }
            BodyItemRef::ContentControl(_) => "content-control".to_string(),
            BodyItemRef::UnsupportedXml(_) => "unsupported".to_string(),
        })
        .collect();
    let paragraphs = document
        .paragraphs()
        .into_iter()
        .map(|paragraph| {
            let paragraph_style = document.resolve_paragraph_properties(paragraph.style_id());
            let alignment = paragraph
                .alignment()
                .map(|value| format!("{value:?}"))
                .or_else(|| paragraph_style.jc.map(|value| format!("{value:?}")))
                .unwrap_or_else(|| "Left".to_string());
            let numbering = paragraph.numbering().and_then(|(id, level)| {
                document
                    .numbering_is_bullet(id)
                    .map(|bullet| (bullet, level))
            });
            let runs = (0..paragraph.run_count())
                .filter_map(|index| paragraph.run(index))
                .filter(|run| !run.text().is_empty())
                .map(|run| {
                    let resolved =
                        document.resolve_run_properties(paragraph.style_id(), run.style_id());
                    (
                        run.text(),
                        run.bold_value().or(resolved.bold).unwrap_or(false),
                        run.italic_value().or(resolved.italic).unwrap_or(false),
                        Some(
                            run.color()
                                .map(str::to_string)
                                .or(resolved.color)
                                .unwrap_or_else(|| "000000".to_string()),
                        ),
                    )
                })
                .collect();
            OdtOracleParagraphRecord {
                text: paragraph.text(),
                alignment,
                numbering,
                runs,
            }
        })
        .collect();
    let tables = (0..document.table_count())
        .map(|index| {
            let table = document.table(index).unwrap();
            let mut cells = Vec::new();
            for row_index in 0..table.row_count() {
                let row = table.row(row_index).unwrap();
                for cell_index in 0..row.cell_count() {
                    let cell = row.cell(cell_index).unwrap();
                    cells.push((
                        cell.text(),
                        cell.grid_span(),
                        cell.v_merge().map(|value| format!("{value:?}")),
                    ));
                }
            }
            OdtOracleTableRecord {
                rows: table.row_count(),
                columns: table.column_count(),
                cells,
            }
        })
        .collect();
    let images = document
        .images()
        .into_iter()
        .map(|image| (image.width_emu, image.height_emu))
        .collect();
    let bytes = document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let mut media: Vec<(String, Vec<u8>)> = package
        .parts
        .iter()
        .filter(|(name, _)| name.starts_with("/word/media/"))
        .map(|(name, bytes)| (name.clone(), bytes.clone()))
        .collect();
    media.sort_by(|left, right| left.0.cmp(&right.0));
    OdtStructuralRecord {
        body_order,
        paragraphs,
        tables,
        images,
        media: media.into_iter().map(|(_, bytes)| bytes).collect(),
    }
}

#[test]
fn odt_reader_matches_pinned_libreoffice_structure() {
    let version = std::process::Command::new("soffice")
        .arg("--version")
        .output()
        .expect("pinned LibreOffice is installed");
    assert!(version.status.success());
    assert_eq!(
        String::from_utf8_lossy(&version.stdout).trim(),
        ODT_ORACLE_VERSION
    );

    let root = std::env::temp_dir().join(format!(
        "rdocx-odt-oracle-{}-{}",
        std::process::id(),
        std::thread::current().name().unwrap_or("test")
    ));
    let _ = std::fs::remove_dir_all(&root);
    let output = root.join("output");
    let profile = root.join("profile");
    std::fs::create_dir_all(&output).unwrap();
    std::fs::create_dir_all(&profile).unwrap();
    let source = root.join("source.odt");
    let odt = source_built_odt();
    std::fs::write(&source, &odt).unwrap();

    let status = std::process::Command::new("soffice")
        .arg("--headless")
        .arg(format!(
            "-env:UserInstallation=file://{}",
            profile.display()
        ))
        .arg("--convert-to")
        .arg("docx")
        .arg("--outdir")
        .arg(&output)
        .arg(&source)
        .status()
        .expect("LibreOffice conversion starts");
    assert!(status.success());

    let ours = Document::from_odt_bytes(&odt).unwrap().document;
    let oracle = Document::open(output.join("source.docx")).unwrap();
    let ours = odt_structural_record(ours);
    let oracle = odt_structural_record(oracle);
    let _ = std::fs::remove_dir_all(&root);
    assert_eq!(ours, oracle);
}

#[test]
fn odt_writer_round_trip_preserves_supported_document_content() {
    let mut document = Document::new();
    let mut paragraph = document.add_paragraph("");
    paragraph.set_alignment(Alignment::Center);
    paragraph.set_space_before(Length::twips(60));
    paragraph.set_space_after(Length::twips(120));
    paragraph.set_indent_left(Length::twips(720));
    paragraph.set_indent_right(Length::twips(360));
    paragraph.set_first_line_indent(Length::twips(240));
    paragraph.set_line_spacing_multiple(1.5);
    paragraph
        .add_run("Round")
        .font("Arial")
        .size(11.0)
        .bold(true)
        .underline(true)
        .strike(true)
        .color("123456")
        .highlight("ABCDEF")
        .superscript();
    paragraph
        .add_run(" trip")
        .font("Liberation Serif")
        .size(13.0)
        .italic(true)
        .subscript();

    let list_id = document.add_list_definition(&[ListLevel::bullet(), ListLevel::decimal()]);
    document.add_paragraph("top item").set_numbering(list_id, 0);
    document
        .add_paragraph("nested item")
        .set_numbering(list_id, 1);

    {
        let mut table = document.add_table(2, 2);
        table.cell(0, 0).unwrap().set_text("vertical");
        table.cell(0, 0).unwrap().set_v_merge_restart();
        table.cell(0, 1).unwrap().set_text("top right");
        table
            .cell(0, 1)
            .unwrap()
            .paragraph_mut(0)
            .unwrap()
            .set_space_after(Length::twips(80));
        let mut top_right = table.cell(0, 1).unwrap();
        let mut cell_list = top_right.add_paragraph("cell list item");
        cell_list.set_alignment(Alignment::Right);
        cell_list.set_numbering(list_id, 1);
        cell_list
            .add_run(" formatted")
            .font("Arial")
            .size(10.0)
            .underline(true)
            .highlight("FEDCBA");
        table.cell(1, 0).unwrap().set_v_merge_continue();
        table.cell(1, 1).unwrap().set_text("bottom right");
    }
    document.add_picture(
        PNG_2_BY_3,
        "two-by-three.png",
        Length::emu(914_400),
        Length::emu(457_200),
    );

    let expected_docx = document.to_bytes().unwrap();
    let expected = odt_supported_record(Document::from_bytes(&expected_docx).unwrap());
    assert_eq!(expected.paragraphs[2].numbering, Some((false, 1)));
    let written = document.to_odt_bytes().unwrap();
    assert!(
        written
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.path == "body[3]/tblPr")
    );
    assert!(
        written
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.path == "body[3]/tblGrid")
    );
    let actual = odt_supported_record(Document::from_odt_bytes(&written.bytes).unwrap().document);
    assert_eq!(actual, expected);
}

#[test]
fn html_import_projects_a_reopenable_word_document() {
    let parsed = Document::from_html(
        "<h1>Title</h1><p>A <b>browser</b> fragment.</p><table><tr><th>Head</th></tr><tr><td>Cell<p>Second</p></td></tr></table>",
    )
    .expect("supported HTML");
    let mut document = parsed.document;
    let bytes = document.to_bytes().expect("generated DOCX");
    let reopened = Document::from_bytes(&bytes).expect("generated DOCX reopens");
    assert_eq!(reopened.paragraph(0).unwrap().text(), "Title");
    assert_eq!(reopened.paragraph(1).unwrap().text(), "A browser fragment.");
    assert_eq!(
        reopened.table(0).unwrap().cell(0, 0).unwrap().text(),
        "Head"
    );
    assert_eq!(
        reopened
            .table(0)
            .unwrap()
            .cell(1, 0)
            .unwrap()
            .paragraph_count(),
        2
    );

    let root = std::env::temp_dir().join(format!("rdocx-html-import-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&root);
    std::fs::create_dir(&root).unwrap();
    let html_path = root.join("source.html");
    std::fs::write(&html_path, "<p>path input</p>").unwrap();
    let opened = Document::open_html(&html_path).expect("bounded UTF-8 path input");
    assert_eq!(opened.document.text(), "path input\n");

    let preformatted = Document::from_html("<pre>first\nsecond</pre>").unwrap();
    let mut preformatted_document = preformatted.document;
    let preformatted_bytes = preformatted_document.to_bytes().unwrap();
    let package =
        oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(preformatted_bytes)).unwrap();
    let document_xml =
        std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(document_xml.contains("<w:br"));

    let invalid_path = root.join("invalid.html");
    std::fs::write(&invalid_path, [0xff, 0xfe]).unwrap();
    assert!(matches!(
        Document::open_html(&invalid_path),
        Err(rdocx::Error::Html { location, .. }) if location == "input"
    ));

    let oversized_path = root.join("oversized.html");
    std::fs::File::create(&oversized_path)
        .unwrap()
        .set_len(64 * 1024 * 1024 + 1)
        .unwrap();
    assert!(Document::open_html(&oversized_path).is_err());
    std::fs::remove_dir_all(root).unwrap();
}

#[test]
fn rich_html_fragments_match_word_in_every_supported_container() {
    use base64::Engine as _;

    let mut document = container_neutral_story_fixture();
    let png = mhtml_pixel_png();
    let data_uri = format!(
        "data:image/png;base64,{}",
        base64::engine::general_purpose::STANDARD.encode(&png)
    );
    let html = format!(
        "<style>strong {{ color: #336699; }}</style><p><strong>fragment</strong> <a href='https://example.test/path'>link</a><img src='resolved.png' width='2' height='3'><img src='{data_uri}' width='1' height='1'></p><ol start='3'><li>outer<ul><li>inner</li></ul></li></ol><table><tr><th>head</th></tr><tr><td>cell</td></tr></table><p><a href='javascript:alert(1)'>unsafe</a><img src='missing.png' alt='fallback'><iframe>lost</iframe></p>"
    );
    let images = [rdocx::HtmlImageResource {
        source: "resolved.png",
        bytes: &png,
        filename: "resolved.png",
    }];

    for kind in [
        StoryKind::Body,
        StoryKind::TableCell,
        StoryKind::Header,
        StoryKind::Footer,
    ] {
        let story = document
            .stories()
            .expect("discover fragment destination stories")
            .into_iter()
            .find(|story| story.kind() == kind)
            .unwrap_or_else(|| panic!("{kind:?} story"));
        let direct_start = if kind == StoryKind::Body { 6 } else { 2 };
        let destination = rdocx::ContentLocation::end(story.clone());
        let result = document
            .insert_html_fragment(&destination, &html, &images)
            .unwrap_or_else(|error| panic!("{kind:?} HTML fragment insertion: {error}"));

        assert_eq!(result.story.kind(), story.kind());
        assert_eq!(result.story.part_name(), story.part_name());
        assert_eq!(result.story.owner_index(), story.owner_index());
        document
            .story_items(&result.story)
            .expect("result returns the refreshed story identity");
        assert_eq!(result.direct_range, direct_start..direct_start + 5);
        assert_eq!(
            result
                .diagnostics
                .iter()
                .map(|diagnostic| diagnostic.message.as_str())
                .collect::<Vec<_>>(),
            [
                "dropped HTML link target and retained anchor text",
                "dropped unresolved HTML image `missing.png` and retained alternate text",
                "dropped HTML iframe content",
            ]
        );

        let bytes = document.to_bytes().expect("fragment document serializes");
        document = Document::from_bytes(&bytes).expect("fragment document reopens");
        let reopened_story = document
            .stories()
            .unwrap()
            .into_iter()
            .find(|story| story.kind() == kind)
            .unwrap();
        let projected_text = document
            .story_items(&reopened_story)
            .unwrap()
            .into_iter()
            .filter_map(|item| item.text().unwrap())
            .collect::<Vec<_>>()
            .join("|");
        for expected in ["fragment", "link", "outer", "inner", "unsafe", "fallback"] {
            assert!(
                projected_text.contains(expected),
                "{kind:?} omitted {expected:?}: {projected_text}"
            );
        }
        let projected_xml = document
            .story_items(&reopened_story)
            .unwrap()
            .into_iter()
            .map(|item| String::from_utf8_lossy(item.xml().unwrap().as_ref()).into_owned())
            .collect::<String>();
        assert!(projected_xml.contains(">head<"));
        assert!(projected_xml.contains(">cell<"));
        let links = document
            .story_items(&reopened_story)
            .unwrap()
            .into_iter()
            .flat_map(|item| item.links().unwrap())
            .collect::<Vec<_>>();
        assert!(
            links.iter().any(|link| {
                link.text.trim() == "link"
                    && link.url.as_deref() == Some("https://example.test/path")
            }),
            "{kind:?} links: {links:?}"
        );
    }

    let relationship_bytes = document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(relationship_bytes)).unwrap();
    for (owner, expected_images, expected_links) in [
        ("/word/document.xml", 4, 2),
        ("/word/header-story.xml", 2, 1),
        ("/word/footer-story.xml", 2, 1),
    ] {
        let relationships = package.get_part_rels(owner).unwrap();
        assert_eq!(
            relationships
                .items
                .iter()
                .filter(|relationship| relationship.rel_type == rel_types::IMAGE)
                .count(),
            expected_images,
            "image relationship owner {owner}"
        );
        assert_eq!(
            relationships
                .items
                .iter()
                .filter(|relationship| relationship.rel_type == rel_types::HYPERLINK)
                .count(),
            expected_links,
            "hyperlink relationship owner {owner}"
        );
    }
    assert_eq!(
        WORD_HTML_FRAGMENT_ORACLE,
        "Microsoft Word 16.112.4 build 16.112.26090911"
    );
    assert_eq!(
        LIBREOFFICE_HTML_FRAGMENT_ORACLE,
        "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb"
    );
    assert_eq!(POPPLER_HTML_FRAGMENT_ORACLE, "pdftotext version 26.09.0");
    let rendered = document
        .to_pdf_deterministic()
        .expect("fragment document renders through the production layout path");
    assert!(rendered.starts_with(b"%PDF-"));
    assert!(rendered.ends_with(b"%%EOF"));

    let stable = document.to_bytes().expect("serialize atomicity baseline");
    let body = document
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::Body)
        .unwrap();
    let error = document
        .insert_html_fragment(
            &rdocx::ContentLocation::end(body),
            "<p><img src='broken.png'></p>",
            &[rdocx::HtmlImageResource {
                source: "broken.png",
                bytes: b"not an image",
                filename: "broken.png",
            }],
        )
        .expect_err("malformed explicit image must fail");
    assert!(error.to_string().contains("unsupported or malformed"));
    assert_eq!(
        document
            .to_bytes()
            .expect("serialize after rejected fragment"),
        stable,
        "a rejected fragment changed the live document"
    );
}

#[test]
fn html_fragment_insertion_respects_a_cell_boundary_and_preserved_raw_siblings() {
    let mut document = container_neutral_story_fixture();
    let cell = document
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::TableCell)
        .unwrap();
    let destination = rdocx::ContentLocation::new(cell, StoryItemKind::Paragraph, vec![0]);
    let result = document
        .insert_html_fragment(&destination, "<p>before cell</p>", &[])
        .unwrap();
    assert_eq!(result.direct_range, 0..1);

    let items = document.story_items(&result.story).unwrap();
    assert_eq!(items[0].text().unwrap().as_deref(), Some("before cell"));
    assert_eq!(items[1].text().unwrap().as_deref(), Some("cell"));
    assert_eq!(items[2].kind(), StoryItemKind::PreservedNode);
    assert!(String::from_utf8_lossy(items[2].xml().unwrap().as_ref()).contains("x:flag=\"exact\""));
}

#[test]
fn html_fragment_cell_routing_uses_the_direct_body_container() {
    let mut seed = container_neutral_story_fixture();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let xml = String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
    let xml = xml.replacen(
        "<w:tbl>",
        "<w:sdt><w:sdtContent><w:tbl><w:tr><w:tc><w:p><w:r><w:t>earlier controlled cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:sdtContent></w:sdt><w:tbl>",
        1,
    );
    package.set_part("/word/document.xml", xml.into_bytes());
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(bytes.get_ref()).unwrap();
    let cells = document
        .stories()
        .unwrap()
        .into_iter()
        .filter(|story| story.kind() == StoryKind::TableCell)
        .collect::<Vec<_>>();
    assert_eq!(cells.len(), 2);
    assert_eq!(cells[1].owner_index(), 1);

    let result = document
        .insert_html_fragment(
            &rdocx::ContentLocation::end(cells[1].clone()),
            "<p>ordinary body cell insertion</p>",
            &[],
        )
        .unwrap();
    let text = document
        .story_items(&result.story)
        .unwrap()
        .into_iter()
        .filter_map(|item| item.text().unwrap())
        .collect::<String>();
    assert!(text.contains("ordinary body cell insertion"));
    let earlier = document
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::TableCell && story.owner_index() == 0)
        .unwrap();
    assert_eq!(
        document.story_items(&earlier).unwrap()[0]
            .text()
            .unwrap()
            .as_deref(),
        Some("earlier controlled cell")
    );
}

#[test]
fn html_fragment_rejects_unreviewed_story_kinds_atomically() {
    let mut document = container_neutral_story_fixture();
    let stable = document.to_bytes().unwrap();
    let comment = document
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::Comment)
        .unwrap();
    let error = document
        .insert_html_fragment(
            &rdocx::ContentLocation::end(comment),
            "<p>not allowed</p>",
            &[],
        )
        .expect_err("comments are outside the F-261 contract");
    assert!(
        error
            .to_string()
            .contains("support body, table-cell, header, and footer")
    );
    assert_eq!(document.to_bytes().unwrap(), stable);
}

#[test]
fn m23_drawings_text_boxes_and_watermarks_match_word() {
    let mut document = Document::new();
    document.add_paragraph("drawing matrix");
    let body = document
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::Body)
        .unwrap();
    let png = mhtml_pixel_png();

    document
        .add_picture_with_options(
            &body,
            &png,
            "cropped.png",
            rdocx::PictureOptions {
                width: Length::pt(72.0),
                height: Length::pt(36.0),
                crop: Some(rdocx::PictureCrop {
                    left: 10_000,
                    top: 5_000,
                    right: 20_000,
                    bottom: 0,
                }),
                anchor: None,
                name: Some("Cropped inline".to_owned()),
                description: Some("inline corpus picture".to_owned()),
            },
        )
        .unwrap();
    let body = document
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::Body)
        .unwrap();
    document
        .add_picture_with_options(
            &body,
            &png,
            "floating.png",
            rdocx::PictureOptions {
                width: Length::pt(90.0),
                height: Length::pt(45.0),
                crop: None,
                anchor: Some(rdocx::PictureAnchor {
                    horizontal_relative_from: rdocx::DrawingHorizontalRelativeFrom::Margin,
                    horizontal_offset: Length::pt(18.0),
                    horizontal_alignment: None,
                    vertical_relative_from: rdocx::DrawingVerticalRelativeFrom::Paragraph,
                    vertical_offset: Length::pt(6.0),
                    vertical_alignment: None,
                    wrap: rdocx::DrawingWrap::Square,
                    distance_top: Length::pt(2.0),
                    distance_bottom: Length::pt(3.0),
                    distance_left: Length::pt(4.0),
                    distance_right: Length::pt(5.0),
                    relative_height: 7,
                    behind_text: false,
                }),
                name: Some("Floating picture".to_owned()),
                description: None,
            },
        )
        .unwrap();
    let body = document
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::Body)
        .unwrap();
    document
        .add_text_box_to_story(
            &body,
            " rotated corpus text ",
            rdocx::TextBoxOptions {
                width: Length::pt(144.0),
                height: Length::pt(54.0),
                anchor: rdocx::PictureAnchor {
                    horizontal_relative_from: rdocx::DrawingHorizontalRelativeFrom::Column,
                    horizontal_offset: Length::pt(24.0),
                    horizontal_alignment: None,
                    vertical_relative_from: rdocx::DrawingVerticalRelativeFrom::Paragraph,
                    vertical_offset: Length::pt(12.0),
                    vertical_alignment: None,
                    wrap: rdocx::DrawingWrap::TopAndBottom,
                    distance_top: Length::pt(2.0),
                    distance_bottom: Length::pt(2.0),
                    distance_left: Length::pt(0.0),
                    distance_right: Length::pt(0.0),
                    relative_height: 8,
                    behind_text: false,
                },
                rotation_degrees: 15.0,
                text_direction: rdocx::TextBoxDirection::Vertical,
                fill_color: Some("D9EAF7".to_owned()),
            },
        )
        .unwrap();
    document
        .set_text_watermark_for(
            0,
            HdrFtrType::Default,
            "CONFIDENTIAL",
            rdocx::TextWatermarkOptions {
                width: Length::pt(360.0),
                height: Length::pt(90.0),
                rotation_degrees: 315.0,
                color: "D9D9D9".to_owned(),
                font_family: Some("Calibri".to_owned()),
                opacity: 0.5,
            },
        )
        .unwrap();

    let bytes = document.to_bytes().unwrap();
    let mut reopened = Document::from_bytes(&bytes).unwrap();
    let round_trip = reopened.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(round_trip)).unwrap();
    let xml = String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
    assert!(xml.contains("<a:srcRect l=\"10000\" t=\"5000\" r=\"20000\" b=\"0\"/>"));
    assert!(xml.contains("<wp:extent cx=\"914400\" cy=\"457200\"/>"));
    assert!(xml.contains("relativeFrom=\"margin\""));
    assert!(xml.contains("<wp:wrapSquare"));
    for distance in [
        "distT=\"25400\"",
        "distB=\"38100\"",
        "distL=\"50800\"",
        "distR=\"63500\"",
    ] {
        assert!(xml.contains(distance), "{distance}");
    }
    assert!(xml.contains("<mc:AlternateContent"));
    assert!(xml.contains("<wps:wsp"));
    assert!(xml.contains("<w:txbxContent"));
    assert_eq!(
        xml.matches("<w:t xml:space=\"preserve\"> rotated corpus text </w:t>")
            .count(),
        2
    );
    assert!(xml.contains("<a:xfrm rot=\"900000\""));
    assert!(xml.contains("<wps:bodyPr vert=\"vert\"/>"));
    assert!(xml.contains("<mc:Fallback"));
    assert!(xml.contains("<v:shapetype id=\"rdocx-textbox-type-"));
    assert!(xml.contains("type=\"#rdocx-textbox-type-"));
    assert!(xml.contains("<v:textbox"));
    assert!(
        xml.contains(
            "<v:textbox style=\"layout-flow:vertical;mso-layout-flow-alt:top-to-bottom\">"
        )
    );
    let shape_xml = &xml[xml.find("<wps:wsp").unwrap()..xml.find("</wps:wsp>").unwrap()];
    assert!(shape_xml.find("<a:prstGeom").unwrap() < shape_xml.find("<a:solidFill").unwrap());
    assert!(shape_xml.find("<wps:txbx").unwrap() < shape_xml.find("<wps:bodyPr").unwrap());
    assert_eq!(reopened.paragraphs()[0].text(), "drawing matrix");
    let native_pdf = reopened.to_pdf_deterministic().unwrap();
    assert!(native_pdf.starts_with(b"%PDF-"));
}

#[test]
fn drawing_option_matrix_round_trips_with_story_relationships() {
    let mut document = container_neutral_story_fixture();
    let header = document
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::Header)
        .unwrap();
    let before_items = document.story_items(&header).unwrap().len();
    let before_paragraphs = document
        .story_items(&header)
        .unwrap()
        .iter()
        .filter(|item| item.kind() == StoryItemKind::Paragraph)
        .count();
    document
        .add_picture_with_options(
            &header,
            &mhtml_pixel_png(),
            "header-crop.png",
            rdocx::PictureOptions {
                width: Length::pt(24.0),
                height: Length::pt(12.0),
                crop: Some(rdocx::PictureCrop {
                    left: 1,
                    top: 2,
                    right: 3,
                    bottom: 4,
                }),
                anchor: None,
                name: None,
                description: None,
            },
        )
        .unwrap();
    let bytes = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    let reopened_header = reopened
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::Header)
        .unwrap();
    let reopened_items = reopened.story_items(&reopened_header).unwrap();
    assert_eq!(reopened_items.len(), before_items + 2);
    assert_eq!(
        reopened_items
            .iter()
            .filter(|item| item.kind() == StoryItemKind::Paragraph)
            .count(),
        before_paragraphs + 1
    );
    let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    assert_eq!(
        package
            .get_part_rels(reopened_header.part_name())
            .unwrap()
            .items
            .iter()
            .filter(|relationship| relationship.rel_type == rel_types::IMAGE)
            .count(),
        1
    );
}

#[test]
fn section_aware_watermark_pages_select_the_requested_variant() {
    let mut document = Document::new();
    document
        .set_text_watermark_for(
            0,
            HdrFtrType::Even,
            "EVEN MARK",
            rdocx::TextWatermarkOptions::default(),
        )
        .unwrap();
    assert!(!document.even_and_odd_headers());
    document
        .section_mut(0)
        .unwrap()
        .set_different_first_page(true);
    document.set_even_and_odd_headers(true).unwrap();
    for (kind, text) in [
        (HdrFtrType::Default, "DEFAULT MARK"),
        (HdrFtrType::First, "FIRST MARK"),
    ] {
        document
            .set_text_watermark_for(0, kind, text, rdocx::TextWatermarkOptions::default())
            .unwrap();
    }
    let bytes = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let header_text = package
        .parts
        .iter()
        .filter(|(part, _)| part.starts_with("/word/header"))
        .map(|(_, bytes)| String::from_utf8_lossy(bytes).into_owned())
        .collect::<String>();
    for text in ["DEFAULT MARK", "FIRST MARK", "EVEN MARK"] {
        assert_eq!(header_text.matches(text).count(), 1, "{text}");
    }
    for (kind, text) in [
        (HdrFtrType::Default, "DEFAULT MARK"),
        (HdrFtrType::First, "FIRST MARK"),
        (HdrFtrType::Even, "EVEN MARK"),
    ] {
        let story = reopened
            .section_story(0, rdocx::HeaderFooterKind::Header, kind)
            .unwrap()
            .unwrap();
        let xml = String::from_utf8_lossy(package.get_part(story.story().part_name()).unwrap());
        assert!(xml.contains(text), "{kind:?} selected {xml}");
    }
}

#[test]
fn staged_drawing_invariants_reject_invalid_inputs_atomically() {
    let mut document = Document::new();
    let body = document.stories().unwrap()[0].clone();
    let stable = document.to_bytes().unwrap();
    let invalid = rdocx::PictureOptions {
        width: Length::pt(0.0),
        height: Length::pt(10.0),
        crop: Some(rdocx::PictureCrop {
            left: 80_000,
            top: 0,
            right: 30_000,
            bottom: 0,
        }),
        anchor: None,
        name: None,
        description: None,
    };
    assert!(
        document
            .add_picture_with_options(&body, &mhtml_pixel_png(), "invalid.png", invalid)
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), stable);

    let invalid_text_box = rdocx::TextBoxOptions {
        width: Length::pt(10.0),
        height: Length::pt(10.0),
        anchor: rdocx::PictureAnchor {
            horizontal_relative_from: rdocx::DrawingHorizontalRelativeFrom::Page,
            horizontal_offset: Length::pt(0.0),
            horizontal_alignment: None,
            vertical_relative_from: rdocx::DrawingVerticalRelativeFrom::Page,
            vertical_offset: Length::pt(0.0),
            vertical_alignment: None,
            wrap: rdocx::DrawingWrap::None,
            distance_top: Length::pt(0.0),
            distance_bottom: Length::pt(0.0),
            distance_left: Length::pt(0.0),
            distance_right: Length::pt(0.0),
            relative_height: 0,
            behind_text: false,
        },
        rotation_degrees: f64::from(i32::MAX) / 60_000.0 + 1.0,
        text_direction: rdocx::TextBoxDirection::Horizontal,
        fill_color: None,
    };
    assert!(
        document
            .add_text_box_to_story(&body, "invalid", invalid_text_box)
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), stable);

    let valid = rdocx::PictureOptions {
        width: Length::pt(10.0),
        height: Length::pt(10.0),
        crop: None,
        anchor: None,
        name: None,
        description: None,
    };
    assert!(
        document
            .add_picture_with_options(&body, b"not an image", "invalid.png", valid.clone())
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), stable);

    document.add_paragraph("invalidate the retained story identity");
    let stable_after_edit = document.to_bytes().unwrap();
    assert!(
        document
            .add_picture_with_options(&body, &mhtml_pixel_png(), "stale.png", valid)
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), stable_after_edit);
}

#[test]
#[ignore = "requires pinned Word, LibreOffice, and Poppler render artifacts"]
fn regenerate_f261_html_fragment_render_oracle() {
    use base64::Engine as _;

    let plist = "/Applications/Microsoft Word.app/Contents/Info.plist";
    for (key, expected) in [
        ("CFBundleShortVersionString", "16.112.4"),
        ("CFBundleVersion", "16.112.26090911"),
    ] {
        let output = std::process::Command::new("plutil")
            .args(["-extract", key, "raw", plist])
            .output()
            .unwrap();
        assert!(output.status.success());
        assert_eq!(String::from_utf8_lossy(&output.stdout).trim(), expected);
    }
    assert_eq!(
        WORD_HTML_FRAGMENT_ORACLE,
        "Microsoft Word 16.112.4 build 16.112.26090911"
    );
    let libreoffice = std::process::Command::new("soffice")
        .arg("--version")
        .output()
        .unwrap();
    assert!(libreoffice.status.success());
    assert_eq!(
        String::from_utf8_lossy(&libreoffice.stdout).trim(),
        LIBREOFFICE_HTML_FRAGMENT_ORACLE
    );
    for command in ["pdftotext", "pdftoppm"] {
        let poppler = std::process::Command::new(command)
            .arg("-v")
            .output()
            .unwrap();
        assert!(poppler.status.success());
        assert_eq!(
            String::from_utf8_lossy(&poppler.stderr).lines().next(),
            Some(format!("{command} version 26.09.0").as_str())
        );
    }
    assert_eq!(POPPLER_HTML_FRAGMENT_ORACLE, "pdftotext version 26.09.0");

    let output = std::env::var("RDOCX_F261_ORACLE_DOCX")
        .expect("set RDOCX_F261_ORACLE_DOCX to a temporary output path");
    let png = mhtml_pixel_png();
    let data_uri = format!(
        "data:image/png;base64,{}",
        base64::engine::general_purpose::STANDARD.encode(&png)
    );
    let html = format!(
        "<style>strong {{ color: #336699; }}</style><p><strong>fragment</strong> <a href='https://example.test/path'>link</a><img src='resolved.png' width='2' height='3'><img src='{data_uri}' width='1' height='1'></p><ol start='3'><li>outer<ul><li>inner</li></ul></li></ol><table><tr><th>head</th></tr><tr><td>cell</td></tr></table><p>rendered fallback</p>"
    );
    let images = [rdocx::HtmlImageResource {
        source: "resolved.png",
        bytes: &png,
        filename: "resolved.png",
    }];
    let mut document = Document::new();
    document.add_paragraph("body");
    document
        .add_table(1, 1)
        .cell(0, 0)
        .unwrap()
        .set_text("cell");
    document.set_header("header");
    document.set_footer("footer");
    for kind in [
        StoryKind::Body,
        StoryKind::TableCell,
        StoryKind::Header,
        StoryKind::Footer,
    ] {
        let story = document
            .stories()
            .unwrap()
            .into_iter()
            .find(|story| story.kind() == kind)
            .unwrap();
        document
            .insert_html_fragment(&rdocx::ContentLocation::end(story), &html, &images)
            .unwrap();
    }
    document.save(output).unwrap();
    let native_pdf = std::env::var("RDOCX_F261_NATIVE_PDF")
        .expect("set RDOCX_F261_NATIVE_PDF to a temporary output path");
    std::fs::write(native_pdf, document.to_pdf_deterministic().unwrap()).unwrap();

    let pdfs = [
        (
            "RDOCX_F261_WORD_PDF",
            std::env::var("RDOCX_F261_WORD_PDF").expect("set RDOCX_F261_WORD_PDF"),
        ),
        (
            "RDOCX_F261_LIBREOFFICE_PDF",
            std::env::var("RDOCX_F261_LIBREOFFICE_PDF").expect("set RDOCX_F261_LIBREOFFICE_PDF"),
        ),
    ];
    for (variable, path) in &pdfs {
        let info = std::process::Command::new("pdfinfo")
            .arg(path)
            .output()
            .unwrap();
        assert!(info.status.success());
        let info = String::from_utf8(info.stdout).unwrap();
        assert!(info.lines().any(|line| line == "Pages:           1"));
        assert!(
            info.lines()
                .any(|line| line == "Page size:       612 x 792 pts (letter)")
        );
        let text = std::process::Command::new("pdftotext")
            .args(["-layout", path, "-"])
            .output()
            .unwrap();
        assert!(text.status.success());
        let record = String::from_utf8(text.stdout)
            .unwrap()
            .split_whitespace()
            .filter(|token| *token != "◦")
            .map(str::to_owned)
            .collect::<Vec<_>>();
        assert_eq!(record, HTML_FRAGMENT_RENDER_RECORD, "{variable}");
    }

    let raster_root = std::env::temp_dir().join(format!(
        "rdocx-f261-render-{}-{:?}",
        std::process::id(),
        std::thread::current().id()
    ));
    std::fs::create_dir_all(&raster_root).unwrap();
    let mut pngs = Vec::new();
    for (index, (_, pdf)) in pdfs.iter().enumerate() {
        let prefix = raster_root.join(format!("viewer-{index}"));
        let raster = std::process::Command::new("pdftoppm")
            .args(["-f", "1", "-singlefile", "-png", "-r", "150"])
            .arg(pdf)
            .arg(&prefix)
            .output()
            .unwrap();
        assert!(
            raster.status.success(),
            "pdftoppm failed: {}",
            String::from_utf8_lossy(&raster.stderr)
        );
        pngs.push(prefix.with_extension("png"));
    }
    let scripts = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../scripts");
    let score = std::process::Command::new("python3")
        .args([
            "-c",
            "import sys; sys.path.insert(0, sys.argv[1]); from pathlib import Path; from golden_png_harness import decode_png; from pptx_ssim_harness import structural_similarity; print(structural_similarity(decode_png(Path(sys.argv[2])), decode_png(Path(sys.argv[3]))))",
        ])
        .arg(scripts)
        .args(&pngs)
        .output()
        .unwrap();
    assert!(
        score.status.success(),
        "SSIM comparison failed: {}",
        String::from_utf8_lossy(&score.stderr)
    );
    let score = String::from_utf8(score.stdout)
        .unwrap()
        .trim()
        .parse::<f64>()
        .unwrap();
    assert!(score >= 0.75, "Word and LibreOffice render SSIM {score}");
    std::fs::remove_dir_all(raster_root).unwrap();
}

#[test]
fn settings_relationship_target_is_resolved_instead_of_assumed() {
    let settings = br#"<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:documentProtection w:edit="comments" w:enforcement="true" w:hash="custom-hash" w:salt="custom-salt"/></w:settings>"#;
    let mut seed =
        Document::new_with_profile(WordCreationProfile::Minimal(WordPackageClass::Document));
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes.clone())).unwrap();
    package.set_part("/word/config/protection.xml", settings.to_vec());
    package.content_types.add_override(
        "/word/config/protection.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    );
    package
        .get_or_create_part_rels("/word/document.xml")
        .add(rel_types::SETTINGS, "config/protection.xml");
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();

    let mut document = Document::from_bytes(input.get_ref()).unwrap();
    let protection = document.document_protection().unwrap();
    assert_eq!(protection.mode, rdocx::ProtectionMode::Comments);
    assert_eq!(protection.hash.as_deref(), Some("custom-hash"));

    let saved = document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    assert_eq!(
        package.get_part("/word/config/protection.xml").unwrap(),
        settings
    );
    assert!(package.get_part("/word/settings.xml").is_none());
    let relationship = package
        .get_part_rels("/word/document.xml")
        .and_then(|relationships| relationships.get_by_type(rel_types::SETTINGS))
        .unwrap();
    assert_eq!(relationship.target, "config/protection.xml");
}

#[test]
fn rtf_reader_projects_word_text_formatting_tables_lists_and_images() {
    let input = br"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0 Arial;}}{\colortbl;\red18\green52\blue86;}\f0\fs24\cf1 Heading\par{\listtext\'b7\tab}Item\par\trowd\cellx1440\cellx2880 Left\cell Right\cell\row{\pict\pngblip\picw1\pich1 89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d49444154789c6360f8cff00000040101089d1de10000000049454e44ae426082}}";
    let parsed = Document::from_rtf_bytes(input).expect("Word-style RTF subset");

    assert_eq!(parsed.document.paragraph(0).unwrap().text(), "Heading");
    assert_eq!(
        parsed
            .document
            .paragraph(0)
            .unwrap()
            .run(0)
            .unwrap()
            .font_name(),
        Some("Arial")
    );
    assert_eq!(
        parsed.document.paragraph(0).unwrap().run(0).unwrap().size(),
        Some(12.0)
    );
    assert_eq!(
        parsed
            .document
            .paragraph(0)
            .unwrap()
            .run(0)
            .unwrap()
            .color(),
        Some("123456")
    );
    let numbering = parsed.document.paragraph(1).unwrap().numbering().unwrap();
    assert_eq!(numbering.1, 0);
    assert_eq!(parsed.document.numbering_is_bullet(numbering.0), Some(true));
    assert_eq!(parsed.document.table_count(), 1);
    let table = parsed.document.table(0).unwrap();
    assert_eq!(table.cell(0, 0).unwrap().text(), "Left");
    assert_eq!(table.cell(0, 1).unwrap().text(), "Right");
    let images = parsed.document.images();
    assert_eq!(images.len(), 1);
    assert_eq!(
        (images[0].width_emu, images[0].height_emu),
        (12_700, 12_700)
    );
}

const WORD_RTF_ORACLE_VERSION: &str = "Microsoft Word 16.104 build 16.104.25121423";
const WORD_RTF_SOURCE: &[u8] = br"{\rtf1\ansi\deff0{\fonttbl{\f0 Calibri;}}{\colortbl;\red18\green52\blue86;}{\*\listtable{\list{\listlevel\levelnfc23\levelstartat1}\listid10}}{\*\listoverridetable{\listoverride\listid10\ls5}}\pard\qc\li720\ri360\fi-240\sb60\sa120\sl-360\slmult0\f0\fs22\cf1 Oracle{\b B}{\i I}{\ul U}{\strike S}{\highlight1 H}{\super P}{\sub D}{\caps C}{\scaps M}{\v V}X\line Y\tab Z\par\pard\plain\f0\fs22\ls5\ilvl0 List item\par\trowd\cellx1440\cellx4320 left\cell right\cell\row\pard{\pict\pngblip\picwgoal100\pichgoal200\picscalex200\picscaley50 89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d49444154789c6360f8cff00000040101089d1de10000000049454e44ae426082}{\*\wordprivate ignored}}";

fn rtf_text(bytes: Vec<u8>) -> String {
    String::from_utf8(bytes).expect("RTF writer emits ASCII")
}

fn tiny_jpeg() -> Vec<u8> {
    let mut jpeg = vec![0xff, 0xd8];
    jpeg.extend_from_slice(&[0xff, 0xe0, 0x00, 0x02]);
    jpeg.extend_from_slice(&[
        0xff, 0xc0, 0x00, 0x11, 0x08, 0x00, 0x02, 0x00, 0x03, 0x03, 0x01, 0x11, 0x00, 0x02, 0x11,
        0x00, 0x03, 0x11, 0x00,
    ]);
    jpeg.extend_from_slice(&[0xff, 0xd9]);
    jpeg
}

#[test]
fn rtf_writer_emits_stable_header_tables_before_body() {
    let mut document = Document::new();
    let list_id = document.add_list_definition(&[
        ListLevel::bullet(),
        ListLevel::decimal().start(3),
        ListLevel::new(rdocx::ListNumberFormat::UpperRoman).start(5),
    ]);
    let mut first = document.add_paragraph("");
    first
        .add_run("header order")
        .font("Arial")
        .color("123456")
        .highlight("ABCDEF");
    document
        .add_paragraph("list item")
        .set_numbering(list_id, 2);

    let rtf = rtf_text(document.to_rtf_bytes().unwrap().bytes);

    assert!(rtf.starts_with(
        "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0\\fcharset0 Calibri;}{\\f1\\fcharset0 Arial;}}"
    ));
    let font_table = rtf.find("{\\fonttbl").unwrap();
    let color_table = rtf
        .find("{\\colortbl;\\red18\\green52\\blue86;\\red171\\green205\\blue239;}")
        .unwrap();
    let list_table = rtf.find("{\\*\\listtable").unwrap();
    let body = rtf.find("\\pard").unwrap();
    assert!(font_table < color_table && color_table < list_table && list_table < body);
    assert!(rtf.contains("{\\listlevel\\levelnfc23\\levelnfcn23\\levelstartat1}"));
    assert!(rtf.contains("{\\listlevel\\levelnfc0\\levelnfcn0\\levelstartat3}"));
    assert!(rtf.contains("{\\listlevel\\levelnfc1\\levelnfcn1\\levelstartat5}"));
    assert!(rtf.contains("\\ls1\\ilvl2"));
}

#[test]
fn rtf_writer_round_trip_preserves_supported_document_content() {
    let mut document = Document::new();
    let list_id = document.add_list_definition(&[ListLevel::bullet(), ListLevel::decimal()]);
    let mut paragraph = document.add_paragraph("");
    paragraph.set_alignment(Alignment::Center);
    paragraph.set_indent_left(Length::twips(720));
    paragraph.set_indent_right(Length::twips(360));
    paragraph.set_signed_first_line_indent_value(Some(Length::twips(-240)));
    paragraph.set_space_before(Length::twips(60));
    paragraph.set_space_after(Length::twips(120));
    paragraph.set_line_spacing(18.0);
    paragraph
        .add_run("Round")
        .font("Arial")
        .size(11.0)
        .bold(true)
        .color("123456");
    paragraph.add_run(" trip").italic(true).underline(true);
    paragraph.add_run(" strike").strike(true);
    paragraph.add_run("\tline\nend");
    document
        .add_paragraph("nested item")
        .set_numbering(list_id, 1);
    {
        let mut table = document.add_table(1, 2);
        table.cell(0, 0).unwrap().set_text("left");
        table.cell(0, 1).unwrap().set_text("right");
    }
    document.add_picture(
        PNG_2_BY_3,
        "two-by-three.png",
        Length::emu(12_700),
        Length::emu(25_400),
    );
    let jpeg = tiny_jpeg();
    document.add_picture(
        &jpeg,
        "three-by-two.jpg",
        Length::emu(38_100),
        Length::emu(50_800),
    );

    let written = document.to_rtf_bytes().unwrap();
    assert!(written.diagnostics.is_empty());
    let mut reparsed = Document::from_rtf_bytes(&written.bytes).unwrap().document;

    assert_eq!(
        reparsed.paragraph(0).unwrap().text(),
        "Round trip strike\tline\nend"
    );
    assert_eq!(
        reparsed.paragraph(0).unwrap().alignment(),
        Some(Alignment::Center)
    );
    assert_eq!(
        reparsed
            .paragraph(0)
            .unwrap()
            .indent_left()
            .map(Length::to_emu),
        Some(457_200)
    );
    assert_eq!(
        reparsed.paragraph(0).unwrap().run(0).unwrap().font_name(),
        Some("Arial")
    );
    assert_eq!(
        reparsed.paragraph(0).unwrap().run(0).unwrap().bold_value(),
        Some(true)
    );
    assert_eq!(
        reparsed.paragraph(0).unwrap().run(0).unwrap().color(),
        Some("123456")
    );
    assert_eq!(
        reparsed
            .paragraph(0)
            .unwrap()
            .run(1)
            .unwrap()
            .italic_value(),
        Some(true)
    );
    assert!(
        reparsed
            .paragraph(0)
            .unwrap()
            .run(1)
            .unwrap()
            .is_underline()
    );
    assert_eq!(
        reparsed
            .paragraph(0)
            .unwrap()
            .run(2)
            .unwrap()
            .strike_value(),
        Some(true)
    );
    let numbering = reparsed.paragraph(1).unwrap().numbering().unwrap();
    assert_eq!(numbering.1, 1);
    assert_eq!(reparsed.numbering_is_bullet(numbering.0), Some(true));
    assert_eq!(
        reparsed.table(0).unwrap().cell(0, 0).unwrap().text(),
        "left"
    );
    assert_eq!(
        reparsed.table(0).unwrap().cell(0, 1).unwrap().text(),
        "right"
    );
    let images = reparsed.images();
    assert_eq!(images.len(), 2);
    assert_eq!(
        (images[0].width_emu, images[0].height_emu),
        (12_700, 25_400)
    );
    assert_eq!(
        reparsed.image_data(&images[0].embed_id).unwrap(),
        PNG_2_BY_3
    );
    assert_eq!(
        (images[1].width_emu, images[1].height_emu),
        (38_100, 50_800)
    );
    assert_eq!(reparsed.image_data(&images[1].embed_id).unwrap(), jpeg);
    assert!(
        normalize_word_rtf_oracle(&mut reparsed)
            .document_markers
            .contains(&"break")
    );
}

#[test]
fn rtf_writer_preserves_truncating_image_goal_dimensions() {
    let mut document = Document::new();
    document.add_picture(
        PNG_2_BY_3,
        "two-by-three.png",
        Length::emu(12_699),
        Length::emu(-12_699),
    );
    let jpeg = tiny_jpeg();
    document.add_picture(&jpeg, "tiny.jpg", Length::emu(12_700), Length::emu(19_049));

    let rtf = rtf_text(document.to_rtf_bytes().unwrap().bytes);

    assert!(rtf.contains("\\pict\\pngblip\\picwgoal19\\pichgoal-19 "));
    assert!(rtf.contains("\\pict\\jpegblip\\picwgoal20\\pichgoal29 "));
}

#[test]
fn rtf_writer_resets_table_cell_paragraph_state() {
    let mut document = Document::new();
    let list_id = document.add_list_definition(&[ListLevel::decimal()]);
    {
        let mut table = document.add_table(1, 2);
        let mut first_cell = table.cell(0, 0).unwrap();
        first_cell.remove_first_empty_paragraph();
        let mut first = first_cell.add_paragraph("center");
        first.set_alignment(Alignment::Center);
        let mut second = first_cell.add_paragraph("list");
        second.set_numbering(list_id, 0);
        table.cell(0, 1).unwrap().set_text("default");
    }

    let rtf = rtf_text(document.to_rtf_bytes().unwrap().bytes);
    let center = rtf
        .find("\\pard\\intbl\\qc")
        .expect("centered cell paragraph resets");
    let list = rtf
        .find("\\par \\pard\\intbl\\ls1\\ilvl0")
        .expect("list paragraph resets");
    let next_cell = rtf
        .find("\\cell \\pard\\intbl{\\plain\\f0 default")
        .expect("next cell resets");

    assert!(center < list && list < next_cell, "{rtf}");
}

#[test]
fn rtf_writer_serializes_grid_span_with_one_boundary_per_output_cell() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes.clone())).unwrap();
    let source_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
          <w:p><w:r><w:t>merged</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    package.set_part("/word/document.xml", source_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let document = Document::from_bytes(input.get_ref()).unwrap();

    let written = document.to_rtf_bytes().unwrap();
    assert!(written.diagnostics.is_empty());
    let rtf = rtf_text(written.bytes.clone());
    let row = rtf
        .split("\\trowd")
        .nth(1)
        .and_then(|tail| tail.split("\\row").next())
        .expect("row RTF");

    assert_eq!(
        row.matches("\\cellx").count(),
        row.matches("\\cell ").count()
    );
    assert!(row.contains("\\cellx3000"), "{row}");
    assert!(!row.contains("\\cellx1000\\cellx3000"), "{row}");
    let reparsed = Document::from_rtf_bytes(&written.bytes).unwrap().document;
    assert_eq!(reparsed.table(0).unwrap().row(0).unwrap().cell_count(), 1);
    assert_eq!(
        reparsed.table(0).unwrap().cell(0, 0).unwrap().text(),
        "merged"
    );

    let mut public_document = Document::new();
    {
        let mut table = public_document.add_table(1, 1);
        let mut cell = table.cell(0, 0).unwrap();
        cell.set_grid_span(2);
        cell.set_text("public");
    }
    let public_rtf = rtf_text(public_document.to_rtf_bytes().unwrap().bytes);
    let public_row = public_rtf
        .split("\\trowd")
        .nth(1)
        .and_then(|tail| tail.split("\\row").next())
        .expect("public row RTF");
    assert_eq!(
        public_row.matches("\\cellx").count(),
        public_row.matches("\\cell ").count()
    );
}

#[test]
fn rtf_writer_reports_public_unsupported_paragraph_properties_once() {
    let mut document = Document::new();
    let mut paragraph = document.add_paragraph("lossy paragraph");
    paragraph.set_keep_with_next(true);
    paragraph.set_keep_together(true);
    paragraph.set_page_break_before(true);
    paragraph.set_widow_control(false);
    paragraph.set_border_bottom(BorderStyle::Single, 8, "112233");
    paragraph.set_add_tab_stop_with_leader(TabAlignment::Right, Length::twips(720), TabLeader::Dot);
    paragraph.set_outline_level(2);
    paragraph.set_shading("FFFF00");
    paragraph.set_section_break(SectionBreak::Continuous);

    let messages = document
        .to_rtf_bytes()
        .unwrap()
        .diagnostics
        .into_iter()
        .map(|diagnostic| {
            (
                diagnostic.destination.expect("diagnostic location"),
                diagnostic.message,
            )
        })
        .collect::<Vec<_>>();

    assert_eq!(
        messages,
        [
            (
                "body[0]/ppr/keepNext".to_owned(),
                "keep-with-next paragraph property was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/ppr/keepLines".to_owned(),
                "keep-lines paragraph property was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/ppr/pageBreakBefore".to_owned(),
                "page-break-before paragraph property was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/ppr/widowControl".to_owned(),
                "widow-control paragraph property was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/ppr/outlineLvl".to_owned(),
                "paragraph outline level was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/ppr/borders".to_owned(),
                "paragraph borders were dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/ppr/tabs".to_owned(),
                "paragraph tab stops were dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/ppr/shading".to_owned(),
                "paragraph shading was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/ppr/sectPr".to_owned(),
                "paragraph section properties were dropped during RTF export".to_owned(),
            ),
        ]
    );
}

#[test]
fn rtf_writer_reports_table_and_row_property_diagnostics() {
    let mut document = Document::new();
    {
        let mut table = document.add_table(1, 1);
        table.set_style("TableGrid");
        table.set_width(Length::twips(4000));
        table.set_alignment(Alignment::Center);
        table.set_borders(BorderStyle::Single, 8, "112233");
        table.set_cell_margins(
            Length::twips(10),
            Length::twips(20),
            Length::twips(30),
            Length::twips(40),
        );
        table.set_layout_fixed();
        table.set_indent(Length::twips(120));
        let mut row = table.row(0).unwrap();
        row.set_height_exact(Length::twips(360));
        row.set_header();
        row.set_cant_split();
    }

    let messages = document
        .to_rtf_bytes()
        .unwrap()
        .diagnostics
        .into_iter()
        .map(|diagnostic| {
            (
                diagnostic.destination.expect("diagnostic location"),
                diagnostic.message,
            )
        })
        .collect::<Vec<_>>();

    assert_eq!(
        messages,
        [
            (
                "body[0]/tblPr/tblStyle".to_owned(),
                "table style was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/tblPr/tblW".to_owned(),
                "table width was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/tblPr/jc".to_owned(),
                "table alignment was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/tblPr/tblBorders".to_owned(),
                "table borders were dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/tblPr/tblCellMar".to_owned(),
                "table cell margins were dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/tblPr/tblLayout".to_owned(),
                "table layout was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/tblPr/tblInd".to_owned(),
                "table indent was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/row[0]/trPr/trHeight".to_owned(),
                "table-row height was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/row[0]/trPr/hRule".to_owned(),
                "table-row height rule was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/row[0]/trPr/tblHeader".to_owned(),
                "table-row repeat header property was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/row[0]/trPr/cantSplit".to_owned(),
                "table-row cant-split property was dropped during RTF export".to_owned(),
            ),
        ]
    );
}

#[test]
fn rtf_writer_reports_raw_table_and_row_property_diagnostics() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let source_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:shd w:val="clear" w:fill="FFFF00"/>
        <w:tblLook w:firstRow="1"/>
        <w:tblPrChange w:id="1" w:author="A" w:date="2026-08-24T00:00:00Z"/>
        <ext:tblPrChange xmlns:ext="urn:producer"/>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
      <w:tr>
        <w:trPr>
          <w:cnfStyle w:val="100000000000"/>
          <w:jc w:val="center"/>
          <w:ins w:id="2" w:author="A" w:date="2026-08-24T00:00:00Z"/>
          <ext:del xmlns:ext="urn:producer"/>
        </w:trPr>
        <w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    package.set_part("/word/document.xml", source_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let document = Document::from_bytes(input.get_ref()).unwrap();

    let messages = document
        .to_rtf_bytes()
        .unwrap()
        .diagnostics
        .into_iter()
        .map(|diagnostic| {
            (
                diagnostic.destination.expect("diagnostic location"),
                diagnostic.message,
            )
        })
        .collect::<Vec<_>>();

    assert_eq!(
        messages,
        [
            (
                "body[0]/tblPr/shd".to_owned(),
                "table shading was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/tblPr/tblLook".to_owned(),
                "table look was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/tblPr/tblPrChange".to_owned(),
                "table property revision was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/tblPr/revisionXml".to_owned(),
                "unmodelled table property revision XML was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/row[0]/trPr/jc".to_owned(),
                "table-row alignment was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/row[0]/trPr/cnfStyle".to_owned(),
                "table-row conditional style property was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/row[0]/trPr/revisions".to_owned(),
                "table-row revision markers were dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/row[0]/trPr/raw[13]".to_owned(),
                "unmodelled table-row property del was dropped during RTF export".to_owned(),
            ),
        ]
    );
}

#[test]
fn rtf_writer_reports_raw_table_cell_property_xml_once() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let source_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:hMerge w:val="restart"/>
            <ext:cellProp xmlns:ext="urn:producer"/>
          </w:tcPr>
          <w:p><w:r><w:t>cell</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    package.set_part("/word/document.xml", source_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let document = Document::from_bytes(input.get_ref()).unwrap();

    let messages = document
        .to_rtf_bytes()
        .unwrap()
        .diagnostics
        .into_iter()
        .map(|diagnostic| {
            (
                diagnostic.destination.expect("diagnostic location"),
                diagnostic.message,
            )
        })
        .collect::<Vec<_>>();

    assert_eq!(
        messages,
        [(
            "body[0]/row[0]/cell[0]/tcPr/raw[4]".to_owned(),
            "unmodelled table-cell property cellProp was dropped during RTF export".to_owned(),
        )]
    );
}

#[test]
fn rtf_writer_uses_cell_widths_or_reports_width_loss() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let source_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2222" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>dxa</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="50" w:type="pct"/></w:tcPr><w:p><w:r><w:t>pct</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>missing</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:tbl>
      <w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>grid</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>short</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    package.set_part("/word/document.xml", source_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let document = Document::from_bytes(input.get_ref()).unwrap();

    let written = document.to_rtf_bytes().unwrap();
    let rtf = rtf_text(written.bytes);

    assert!(rtf.contains("\\cellx2222"), "{rtf}");
    let messages = written
        .diagnostics
        .into_iter()
        .map(|diagnostic| {
            (
                diagnostic.destination.expect("diagnostic location"),
                diagnostic.message,
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(
        messages,
        [
            (
                "body[0]/row[0]/cell[1]/tcPr/tcW".to_owned(),
                "unsupported table-cell width type was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/row[0]/cell[2]/tcPr/tcW".to_owned(),
                "table-cell width could not be preserved because the table grid is missing"
                    .to_owned(),
            ),
            (
                "body[1]/row[0]/cell[1]/tcPr/tcW".to_owned(),
                "table-cell width could not be preserved because the table grid is too short"
                    .to_owned(),
            ),
        ]
    );
}

#[test]
fn rtf_writer_rejects_invalid_cell_width_boundaries() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes.clone())).unwrap();
    let source_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="0" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>zero</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="-5" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>negative</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    package.set_part("/word/document.xml", source_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let document = Document::from_bytes(input.get_ref()).unwrap();

    let written = document.to_rtf_bytes().unwrap();
    let rtf = rtf_text(written.bytes);
    assert!(rtf.contains("\\cellx1440\\cellx2880"), "{rtf}");
    let messages = written
        .diagnostics
        .into_iter()
        .map(|diagnostic| {
            (
                diagnostic.destination.expect("diagnostic location"),
                diagnostic.message,
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(
        messages,
        [
            (
                "body[0]/row[0]/cell[0]/tcPr/tcW".to_owned(),
                "invalid table-cell width was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/row[0]/cell[1]/tcPr/tcW".to_owned(),
                "invalid table-cell width was dropped during RTF export".to_owned(),
            ),
        ]
    );

    let mut overflow_package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let overflow_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2147483000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>overflow</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    overflow_package.set_part("/word/document.xml", overflow_xml.to_vec());
    let mut overflow_input = std::io::Cursor::new(Vec::new());
    overflow_package.write_to(&mut overflow_input).unwrap();
    let overflow_document = Document::from_bytes(overflow_input.get_ref()).unwrap();
    let error = match overflow_document.to_rtf_bytes() {
        Ok(_) => panic!("overflowing table boundaries should fail"),
        Err(error) => error,
    };
    assert!(
        error
            .to_string()
            .contains("RTF table cell boundaries exceed the supported range")
    );
}

#[test]
fn rtf_writer_preserves_none_numbering_without_coercion() {
    let mut seed = Document::new();
    seed.add_list_definition(&[ListLevel::decimal(), ListLevel::bullet()]);
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let document_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="9"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>too deep</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>unsupported format</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>supported sibling</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    let numbering_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="none"/><w:lvlText w:val="%1."/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="*"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"#;
    package.set_part("/word/document.xml", document_xml.to_vec());
    package.set_part("/word/numbering.xml", numbering_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let document = Document::from_bytes(input.get_ref()).unwrap();

    let written = document.to_rtf_bytes().unwrap();
    let rtf = rtf_text(written.bytes);
    assert!(!rtf.contains("\\ilvl8"), "{rtf}");
    assert!(
        rtf.contains("\\ls1\\ilvl0{\\plain\\f0 unsupported format}"),
        "{rtf}"
    );
    assert!(rtf.contains("{\\plain\\f0 too deep}"), "{rtf}");
    assert!(rtf.contains("{\\plain\\f0 unsupported format}"), "{rtf}");
    assert!(
        rtf.contains("\\ls1\\ilvl1{\\plain\\f0 supported sibling}"),
        "{rtf}"
    );
    let messages = written
        .diagnostics
        .into_iter()
        .map(|diagnostic| {
            (
                diagnostic.destination.expect("diagnostic location"),
                diagnostic.message,
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(
        messages,
        [(
            "body[0]/ppr/numPr/ilvl".to_owned(),
            "numbering level above 8 was dropped during RTF export".to_owned(),
        )]
    );
}

#[test]
fn rtf_writer_reports_run_properties_individually_without_false_supported_diagnostics() {
    let mut supported = Document::new();
    let mut paragraph = supported.add_paragraph("");
    paragraph
        .add_run("supported")
        .font("Arial")
        .size(11.0)
        .bold(true)
        .italic(true)
        .underline(true)
        .strike(true)
        .color("123456")
        .highlight("ABCDEF")
        .all_caps(true)
        .small_caps(true)
        .hidden(true)
        .superscript();
    assert!(supported.to_rtf_bytes().unwrap().diagnostics.is_empty());

    let mut lossy = Document::new();
    let mut paragraph = lossy.add_paragraph("");
    let mut run = paragraph.add_run("lossy");
    run.set_style("Emphasis");
    run.set_underline_style(UnderlineStyle::Double);
    run.set_double_strike(true);
    run.set_character_spacing(Length::twips(20));
    run.set_width_scale(80);
    run.set_position(4);

    let messages = lossy
        .to_rtf_bytes()
        .unwrap()
        .diagnostics
        .into_iter()
        .map(|diagnostic| {
            (
                diagnostic.destination.expect("diagnostic location"),
                diagnostic.message,
            )
        })
        .collect::<Vec<_>>();

    assert_eq!(
        messages,
        [
            (
                "body[0]/run[0]/rPr/rStyle".to_owned(),
                "run style was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/u".to_owned(),
                "non-basic underline style was simplified during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/dstrike".to_owned(),
                "double strikethrough was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/spacing".to_owned(),
                "run character spacing was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/w".to_owned(),
                "run width scale was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/position".to_owned(),
                "run text position was dropped during RTF export".to_owned(),
            ),
        ]
    );
}

#[test]
fn rtf_writer_reports_raw_run_property_diagnostics_individually() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let source_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Arial" w:hAnsi="Courier New" w:eastAsia="MS Mincho" w:cs="Arial" w:asciiTheme="minorHAnsi" w:hAnsiTheme="majorHAnsi"/>
          <w:b/>
          <w:bCs w:val="0"/>
          <w:i/>
          <w:iCs w:val="0"/>
          <w:sz w:val="22"/>
          <w:szCs w:val="24"/>
          <w:color w:val="123456" w:themeColor="accent1"/>
          <w:highlight w:val="yellow"/>
          <w:ins w:id="1" w:author="A" w:date="2026-08-24T00:00:00Z"/>
          <w:rPrChange w:id="2" w:author="A" w:date="2026-08-24T00:00:00Z"/>
          <ext:rPrChange xmlns:ext="urn:producer"/>
        </w:rPr>
        <w:t>raw</w:t>
      </w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    package.set_part("/word/document.xml", source_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let document = Document::from_bytes(input.get_ref()).unwrap();

    let messages = document
        .to_rtf_bytes()
        .unwrap()
        .diagnostics
        .into_iter()
        .map(|diagnostic| {
            (
                diagnostic.destination.expect("diagnostic location"),
                diagnostic.message,
            )
        })
        .collect::<Vec<_>>();

    assert_eq!(
        messages,
        [
            (
                "body[0]/run[0]/rPr/hAnsi".to_owned(),
                "alternate run font was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/eastAsia".to_owned(),
                "alternate run font was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/asciiTheme".to_owned(),
                "theme run font was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/hAnsiTheme".to_owned(),
                "theme run font was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/bCs".to_owned(),
                "complex-script bold was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/iCs".to_owned(),
                "complex-script italic was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/szCs".to_owned(),
                "complex-script font size was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/themeColor".to_owned(),
                "theme run colour was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/highlight".to_owned(),
                "keyword highlight colour was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/revisions".to_owned(),
                "run revision markers were dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/rPrChange".to_owned(),
                "run property revision was dropped during RTF export".to_owned(),
            ),
            (
                "body[0]/run[0]/rPr/revisionXml".to_owned(),
                "unmodelled run property revision XML was dropped during RTF export".to_owned(),
            ),
        ]
    );
}

#[test]
fn rtf_writer_reports_each_lossy_item_without_dropping_supported_siblings() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let source_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:p="urn:producer">
  <w:body>
    <w:p><w:r><w:t>first</w:t></w:r></w:p>
    <w:sdt><w:sdtPr><w:tag w:val="drop"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt>
    <p:opaque p:flag="drop"><p:child/></p:opaque>
    <w:p><w:bookmarkStart w:id="1" w:name="mark"/><w:r><w:t>last</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    package.set_part("/word/document.xml", source_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let document = Document::from_bytes(input.get_ref()).unwrap();

    let written = document.to_rtf_bytes().unwrap();
    let messages = written
        .diagnostics
        .iter()
        .map(|diagnostic| {
            (
                diagnostic.destination.as_deref().unwrap(),
                diagnostic.message.as_str(),
            )
        })
        .collect::<Vec<_>>();

    assert_eq!(
        messages,
        [
            (
                "body[1]",
                "body content control was dropped during RTF export",
            ),
            (
                "body[2]",
                "unmodelled body XML was dropped during RTF export",
            ),
            (
                "body[3]/bookmark[0]",
                "bookmark marker was dropped during RTF export",
            ),
            (
                "body[3]/run[1]/content[0]",
                "footnote reference was dropped during RTF export",
            ),
        ]
    );
    let reparsed = Document::from_rtf_bytes(&written.bytes).unwrap().document;
    assert_eq!(reparsed.text(), "first\nlast\n");
}

#[test]
fn rtf_writer_leaves_docx_unmodelled_xml_unchanged() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let source_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:p="urn:producer">
  <w:body>
    <w:p><w:r><w:t>kept</w:t></w:r></w:p>
    <p:opaque p:flag="keep"><p:child/></p:opaque>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    package.set_part("/word/document.xml", source_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let mut document = Document::from_bytes(input.get_ref()).unwrap();

    let before = document_xml(&mut document);
    assert!(!document.to_rtf_bytes().unwrap().diagnostics.is_empty());
    let after = document_xml(&mut document);

    assert_eq!(before, after);
    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    assert!(reopened.body_items().any(|item| matches!(
        item,
        BodyItemRef::UnsupportedXml(raw)
            if std::str::from_utf8(raw).unwrap()
                == "<p:opaque p:flag=\"keep\"><p:child/></p:opaque>"
    )));
}

#[derive(Debug, PartialEq)]
struct WordRtfRunRecord {
    text: String,
    font: Option<String>,
    size_points: Option<f64>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline_code: Option<i32>,
    strike: Option<bool>,
    color: Option<String>,
    highlight: Option<String>,
    position: Option<i32>,
    vert_align: Option<String>,
    all_caps: Option<bool>,
    small_caps: Option<bool>,
    hidden: Option<bool>,
    inline_image: bool,
}

#[derive(Debug, PartialEq)]
struct WordRtfParagraphRecord {
    text: String,
    alignment: Option<Alignment>,
    indent_left_emu: Option<i64>,
    indent_right_emu: Option<i64>,
    first_line_indent_emu: Option<i64>,
    space_before_emu: Option<i64>,
    space_after_emu: Option<i64>,
    line_spacing_emu: Option<i64>,
    line_spacing_multiple: Option<f64>,
    numbering: Option<(bool, u32)>,
    runs: Vec<WordRtfRunRecord>,
}

#[derive(Debug, PartialEq)]
struct WordRtfTableRecord {
    cells: Vec<Vec<(String, Option<i64>)>>,
}

#[derive(Debug, PartialEq)]
struct WordRtfStructure {
    body_order: Vec<&'static str>,
    paragraphs: Vec<WordRtfParagraphRecord>,
    tables: Vec<WordRtfTableRecord>,
    images: Vec<(i64, i64)>,
    document_markers: Vec<&'static str>,
}

fn normalize_word_rtf_oracle(document: &mut Document) -> WordRtfStructure {
    let run_formats = word_rtf_body_run_formats(document);
    let paragraphs = document
        .paragraphs()
        .into_iter()
        .enumerate()
        .map(|(paragraph_index, paragraph)| WordRtfParagraphRecord {
            text: paragraph.text(),
            alignment: paragraph.alignment(),
            indent_left_emu: paragraph.indent_left().map(Length::to_emu),
            indent_right_emu: paragraph.indent_right().map(Length::to_emu),
            first_line_indent_emu: paragraph.first_line_indent().map(Length::to_emu),
            space_before_emu: paragraph.space_before().map(Length::to_emu),
            space_after_emu: paragraph.space_after().map(Length::to_emu),
            line_spacing_emu: paragraph.line_spacing().map(Length::to_emu),
            line_spacing_multiple: paragraph.line_spacing_multiple(),
            numbering: paragraph
                .numbering()
                .map(|(id, level)| (document.numbering_is_bullet(id).unwrap_or(false), level)),
            runs: paragraph
                .runs()
                .enumerate()
                .map(|run| WordRtfRunRecord {
                    text: run.1.text(),
                    font: run.1.font_name().map(str::to_owned),
                    size_points: run.1.size(),
                    bold: run.1.bold_value(),
                    italic: run.1.italic_value(),
                    underline_code: run.1.underline_code_value(),
                    strike: run.1.strike_value(),
                    color: run.1.color().map(str::to_owned),
                    highlight: run.1.highlight(),
                    position: run.1.position(),
                    vert_align: run.1.vert_align().map(str::to_owned),
                    all_caps: run_formats
                        .get(paragraph_index)
                        .and_then(|runs| runs.get(run.0))
                        .and_then(|format| format.all_caps),
                    small_caps: run_formats
                        .get(paragraph_index)
                        .and_then(|runs| runs.get(run.0))
                        .and_then(|format| format.small_caps),
                    hidden: run_formats
                        .get(paragraph_index)
                        .and_then(|runs| runs.get(run.0))
                        .and_then(|format| format.hidden),
                    inline_image: run.1.inline_image().is_some(),
                })
                .collect(),
        })
        .collect();
    let tables = (0..document.table_count())
        .map(|table_index| {
            let table = document.table(table_index).unwrap();
            WordRtfTableRecord {
                cells: (0..table.row_count())
                    .map(|row| {
                        (0..table.column_count())
                            .map(|column| {
                                let cell = table.cell(row, column).unwrap();
                                (cell.text(), cell.width().map(Length::to_emu))
                            })
                            .collect()
                    })
                    .collect(),
            }
        })
        .collect();
    WordRtfStructure {
        body_order: document
            .body_items()
            .map(|item| match item {
                BodyItemRef::Paragraph(_) => "paragraph",
                BodyItemRef::Table(_) => "table",
                BodyItemRef::ContentControl(_) => "content-control",
                BodyItemRef::UnsupportedXml(_) => "unsupported",
            })
            .collect(),
        paragraphs,
        tables,
        images: document
            .images()
            .into_iter()
            .map(|image| (image.width_emu, image.height_emu))
            .collect(),
        document_markers: word_rtf_document_markers(document),
    }
}

#[derive(Clone, Copy, Default)]
struct WordRtfRunFormat {
    all_caps: Option<bool>,
    small_caps: Option<bool>,
    hidden: Option<bool>,
}

fn word_rtf_body_run_formats(document: &mut Document) -> Vec<Vec<WordRtfRunFormat>> {
    let bytes = document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let document_xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap())
        .unwrap()
        .to_owned();
    body_paragraph_xml(&document_xml)
        .into_iter()
        .map(|paragraph_xml| {
            run_xml(paragraph_xml)
                .into_iter()
                .map(|run_xml| WordRtfRunFormat {
                    all_caps: word_toggle(run_xml, "caps"),
                    small_caps: word_toggle(run_xml, "smallCaps"),
                    hidden: word_toggle(run_xml, "vanish"),
                })
                .collect()
        })
        .collect()
}

fn body_paragraph_xml(document_xml: &str) -> Vec<&str> {
    let Some(body_start) = document_xml.find("<w:body") else {
        return Vec::new();
    };
    let body =
        &document_xml[body_start..document_xml.find("</w:body>").unwrap_or(document_xml.len())];
    let mut paragraphs = Vec::new();
    let mut cursor = 0;
    let mut table_depth = 0_usize;
    while let Some(relative) = body[cursor..].find('<') {
        let open = cursor + relative;
        let Some(close) = body[open..].find('>').map(|close| open + close) else {
            break;
        };
        let tag = &body[open + 1..close];
        if tag.starts_with("w:tbl") && !tag.ends_with('/') {
            table_depth += 1;
        } else if tag.starts_with("/w:tbl") {
            table_depth = table_depth.saturating_sub(1);
        } else if table_depth == 0 && tag.starts_with("w:p") {
            if tag.ends_with('/') {
                paragraphs.push(&body[open..=close]);
                cursor = close + 1;
                continue;
            }
            let Some(end_relative) = body[close + 1..].find("</w:p>") else {
                break;
            };
            let end = close + 1 + end_relative + "</w:p>".len();
            paragraphs.push(&body[open..end]);
            cursor = end;
            continue;
        }
        cursor = close + 1;
    }
    paragraphs
}

fn run_xml(paragraph_xml: &str) -> Vec<&str> {
    let mut runs = Vec::new();
    let mut cursor = 0;
    while let Some(relative) = paragraph_xml[cursor..].find("<w:r") {
        let open = cursor + relative;
        if paragraph_xml[open..].starts_with("<w:rPr") {
            cursor = open + "<w:rPr".len();
            continue;
        }
        let Some(tag_close) = paragraph_xml[open..].find('>').map(|close| open + close) else {
            break;
        };
        if paragraph_xml[open + 1..tag_close].ends_with('/') {
            runs.push(&paragraph_xml[open..=tag_close]);
            cursor = tag_close + 1;
            continue;
        }
        let Some(end_relative) = paragraph_xml[tag_close + 1..].find("</w:r>") else {
            break;
        };
        let end = tag_close + 1 + end_relative + "</w:r>".len();
        runs.push(&paragraph_xml[open..end]);
        cursor = end;
    }
    runs
}

fn word_toggle(run_xml: &str, name: &str) -> Option<bool> {
    let pattern = format!("<w:{name}");
    let start = run_xml.find(&pattern)?;
    let end = run_xml[start..]
        .find('>')
        .map(|end| start + end)
        .unwrap_or(run_xml.len());
    let tag = &run_xml[start..end];
    if tag.contains("w:val=\"0\"") || tag.contains("w:val=\"false\"") {
        Some(false)
    } else {
        Some(true)
    }
}

fn word_rtf_document_markers(document: &mut Document) -> Vec<&'static str> {
    let bytes = document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let document_xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap())
        .unwrap()
        .to_owned();
    [
        ("break", "<w:br"),
        ("tab", "<w:tab"),
        ("caps", "<w:caps"),
        ("small-caps", "<w:smallCaps"),
        ("hidden", "<w:vanish"),
    ]
    .into_iter()
    .filter_map(|(name, marker)| document_xml.contains(marker).then_some(name))
    .collect()
}

fn word_rtf_run(text: &str) -> WordRtfRunRecord {
    WordRtfRunRecord {
        text: text.to_owned(),
        font: Some("Calibri".to_owned()),
        size_points: Some(11.0),
        bold: None,
        italic: None,
        underline_code: None,
        strike: None,
        color: Some("123456".to_owned()),
        highlight: None,
        position: None,
        vert_align: None,
        all_caps: None,
        small_caps: None,
        hidden: None,
        inline_image: false,
    }
}

fn captured_word_rtf_records() -> WordRtfStructure {
    // Captured by importing WORD_RTF_SOURCE into WORD_RTF_ORACLE_VERSION and
    // saving as DOCX. The ignored test below repeats and validates that capture.
    WordRtfStructure {
        body_order: vec!["paragraph", "paragraph", "table", "paragraph"],
        paragraphs: vec![
            WordRtfParagraphRecord {
                text: "OracleBIUSHPDCMVX\nY\tZ".to_owned(),
                alignment: Some(Alignment::Center),
                indent_left_emu: Some(457_200),
                indent_right_emu: Some(228_600),
                first_line_indent_emu: Some(-152_400),
                space_before_emu: Some(38_100),
                space_after_emu: Some(76_200),
                line_spacing_emu: Some(228_600),
                line_spacing_multiple: None,
                numbering: None,
                runs: vec![
                    word_rtf_run("Oracle"),
                    WordRtfRunRecord {
                        bold: Some(true),
                        ..word_rtf_run("B")
                    },
                    WordRtfRunRecord {
                        italic: Some(true),
                        ..word_rtf_run("I")
                    },
                    WordRtfRunRecord {
                        underline_code: Some(1),
                        ..word_rtf_run("U")
                    },
                    WordRtfRunRecord {
                        strike: Some(true),
                        ..word_rtf_run("S")
                    },
                    WordRtfRunRecord {
                        highlight: Some("123456".to_owned()),
                        ..word_rtf_run("H")
                    },
                    WordRtfRunRecord {
                        vert_align: Some("superscript".to_owned()),
                        ..word_rtf_run("P")
                    },
                    WordRtfRunRecord {
                        vert_align: Some("subscript".to_owned()),
                        ..word_rtf_run("D")
                    },
                    WordRtfRunRecord {
                        all_caps: Some(true),
                        ..word_rtf_run("C")
                    },
                    WordRtfRunRecord {
                        small_caps: Some(true),
                        ..word_rtf_run("M")
                    },
                    WordRtfRunRecord {
                        hidden: Some(true),
                        ..word_rtf_run("V")
                    },
                    word_rtf_run("X"),
                    WordRtfRunRecord {
                        text: "\n".to_owned(),
                        font: None,
                        size_points: None,
                        bold: None,
                        italic: None,
                        underline_code: None,
                        strike: None,
                        color: None,
                        highlight: None,
                        position: None,
                        vert_align: None,
                        all_caps: None,
                        small_caps: None,
                        hidden: None,
                        inline_image: false,
                    },
                    word_rtf_run("Y"),
                    WordRtfRunRecord {
                        text: "\t".to_owned(),
                        font: None,
                        size_points: None,
                        bold: None,
                        italic: None,
                        underline_code: None,
                        strike: None,
                        color: None,
                        highlight: None,
                        position: None,
                        vert_align: None,
                        all_caps: None,
                        small_caps: None,
                        hidden: None,
                        inline_image: false,
                    },
                    word_rtf_run("Z"),
                ],
            },
            WordRtfParagraphRecord {
                text: "List item".to_owned(),
                alignment: None,
                indent_left_emu: None,
                indent_right_emu: None,
                first_line_indent_emu: None,
                space_before_emu: None,
                space_after_emu: None,
                line_spacing_emu: None,
                line_spacing_multiple: None,
                numbering: Some((true, 0)),
                runs: vec![WordRtfRunRecord {
                    color: None,
                    ..word_rtf_run("List item")
                }],
            },
            WordRtfParagraphRecord {
                text: String::new(),
                alignment: None,
                indent_left_emu: None,
                indent_right_emu: None,
                first_line_indent_emu: None,
                space_before_emu: None,
                space_after_emu: None,
                line_spacing_emu: None,
                line_spacing_multiple: None,
                numbering: None,
                runs: vec![WordRtfRunRecord {
                    text: String::new(),
                    font: None,
                    size_points: None,
                    bold: None,
                    italic: None,
                    underline_code: None,
                    strike: None,
                    color: None,
                    highlight: None,
                    position: None,
                    vert_align: None,
                    all_caps: None,
                    small_caps: None,
                    hidden: None,
                    inline_image: true,
                }],
            },
        ],
        tables: vec![WordRtfTableRecord {
            cells: vec![vec![
                ("left".to_owned(), Some(914_400)),
                ("right".to_owned(), Some(1_828_800)),
            ]],
        }],
        images: vec![(127_000, 63_500)],
        document_markers: vec!["break", "tab", "caps", "small-caps", "hidden"],
    }
}

#[test]
fn rtf_reader_matches_the_pinned_word_docx_structure() {
    let mut parsed = Document::from_rtf_bytes(WORD_RTF_SOURCE).expect("oracle source RTF");

    assert_eq!(
        WORD_RTF_ORACLE_VERSION,
        "Microsoft Word 16.104 build 16.104.25121423"
    );
    assert_eq!(
        normalize_word_rtf_oracle(&mut parsed.document),
        captured_word_rtf_records()
    );
    assert_eq!(
        parsed
            .diagnostics
            .iter()
            .map(|diagnostic| (
                diagnostic.destination.as_deref(),
                diagnostic.message.as_str()
            ))
            .collect::<Vec<_>>(),
        vec![(Some("wordprivate"), "unsupported RTF destination skipped")]
    );

    let generated = parsed.document.to_bytes().unwrap();
    let mut reopened = Document::from_bytes(&generated).unwrap();
    assert_eq!(
        normalize_word_rtf_oracle(&mut reopened),
        captured_word_rtf_records()
    );
}

#[test]
#[ignore = "requires the pinned Microsoft Word RTF-to-DOCX oracle"]
fn regenerate_pinned_word_rtf_structure() {
    let actual_version = std::env::var("RDOCX_WORD_RTF_ORACLE_VERSION")
        .expect("set RDOCX_WORD_RTF_ORACLE_VERSION from Microsoft Word");
    assert_eq!(actual_version, WORD_RTF_ORACLE_VERSION);
    let docx_path = std::env::var("RDOCX_WORD_RTF_ORACLE_DOCX")
        .expect("set RDOCX_WORD_RTF_ORACLE_DOCX to Word's saved DOCX");
    let mut word_document = Document::open(docx_path).expect("open Word-produced DOCX");
    let actual = normalize_word_rtf_oracle(&mut word_document);
    println!("{actual:#?}");
    assert_eq!(actual, captured_word_rtf_records());
}

#[test]
fn unsupported_rtf_destinations_are_diagnosed_without_dropping_supported_siblings() {
    let input = br"{\rtf1 before{\*\producerprivate secret}after}";
    let parsed = Document::from_rtf_bytes(input).expect("skippable destination");

    assert_eq!(parsed.document.text(), "beforeafter\n");
    assert_eq!(parsed.diagnostics.len(), 1);
    assert_eq!(parsed.diagnostics[0].offset, 16);
    assert_eq!(
        parsed.diagnostics[0].destination.as_deref(),
        Some("producerprivate")
    );
    assert_eq!(
        parsed.diagnostics[0].message,
        "unsupported RTF destination skipped"
    );
}

#[test]
fn rtf_reader_projects_word_paragraph_indents_spacing_and_line_height() {
    let parsed = Document::from_rtf_bytes(
        br"{\rtf1\ansi\li720\ri360\fi-240\sb120\sa240\sl-360\slmult0 formatted\par\pard\sl360\slmult1 multiplied}",
    )
    .unwrap();
    let paragraph = parsed.document.paragraph(0).unwrap();
    assert_eq!(paragraph.indent_left(), Some(Length::twips(720)));
    assert_eq!(paragraph.indent_right(), Some(Length::twips(360)));
    assert_eq!(paragraph.first_line_indent(), Some(Length::twips(-240)));
    assert_eq!(paragraph.space_before(), Some(Length::twips(120)));
    assert_eq!(paragraph.space_after(), Some(Length::twips(240)));
    assert_eq!(paragraph.line_spacing(), Some(Length::twips(360)));
    assert_eq!(
        parsed
            .document
            .paragraph(1)
            .unwrap()
            .line_spacing_multiple(),
        Some(1.5)
    );
}

#[test]
fn rtf_breaks_tabs_and_line_spacing_survive_docx_projection() {
    let mut parsed = Document::from_rtf_bytes(
        br"{\rtf1\ansi before\tab between\line after\par\sl360\slmult0 at-least\par\sl-240\slmult0 exact\par\sl0\slmult0 automatic}",
    )
    .unwrap();
    let bytes = parsed.document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(&bytes)).unwrap();
    let document_xml =
        std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(document_xml.contains("<w:tab/>"));
    assert!(document_xml.contains("<w:br"));
    assert!(document_xml.contains("w:line=\"360\" w:lineRule=\"atLeast\""));
    assert!(document_xml.contains("w:line=\"240\" w:lineRule=\"exact\""));

    let reopened = Document::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened.paragraph(0).unwrap().text(),
        "before\tbetween\nafter"
    );
    assert_eq!(
        reopened.paragraph(1).unwrap().line_spacing(),
        Some(Length::twips(360))
    );
    assert_eq!(
        reopened.paragraph(2).unwrap().line_spacing(),
        Some(Length::twips(240))
    );
    assert_eq!(reopened.paragraph(3).unwrap().line_spacing(), None);
    assert_eq!(reopened.paragraph(3).unwrap().line_spacing_multiple(), None);
}

#[test]
fn rtf_table_boundaries_and_cell_numbering_survive_projection() {
    let input = br"{\rtf1\ansi{\listtable{\list{\listlevel\levelnfc23\levelstartat1}\listid10}}{\listoverridetable{\listoverride\listid10\listoverridecount1\ls5{\lfolevel\listoverridestartat\levelstartat3}}}\trowd\cellx1440\cellx4320\intbl\ls5\ilvl0 Item\cell\pard Other\cell\row}";
    let mut parsed = Document::from_rtf_bytes(input).unwrap();
    let table = parsed.document.table(0).unwrap();
    assert_eq!(table.cell(0, 0).unwrap().width(), Some(Length::twips(1440)));
    assert_eq!(table.cell(0, 1).unwrap().width(), Some(Length::twips(2880)));
    let numbering = table
        .cell(0, 0)
        .unwrap()
        .paragraph(0)
        .unwrap()
        .numbering()
        .unwrap();
    assert_eq!(parsed.document.numbering_is_bullet(numbering.0), Some(true));

    let bytes = parsed.document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let numbering_xml = package.get_part("/word/numbering.xml").unwrap();
    assert!(
        numbering_xml
            .windows(18)
            .any(|window| window == br#"<w:start w:val="3""#)
    );
}

#[test]
fn rtf_levelnfcn_is_authoritative_regardless_of_control_order() {
    for controls in ["\\levelnfcn23\\levelnfc0", "\\levelnfc0\\levelnfcn23"] {
        let input = format!(
            "{{\\rtf1{{\\*\\listtable{{\\list{{\\listlevel{controls}}}\\listid10}}}}{{\\*\\listoverridetable{{\\listoverride\\listid10\\ls5}}}}\\ls5 item}}"
        );
        let parsed = Document::from_rtf_bytes(input.as_bytes()).unwrap();
        let numbering = parsed.document.paragraph(0).unwrap().numbering().unwrap();
        assert_eq!(parsed.document.numbering_is_bullet(numbering.0), Some(true));
    }
}

#[test]
fn unsupported_and_missing_rtf_lists_are_never_silent_decimal_fallbacks() {
    let unsupported = br"{\rtf1\ansi{\listtable{\list{\listlevel\levelnfc255}\listid10}}{\listoverridetable{\listoverride\listid10\ls5}}\ls5 item}";
    let parsed = Document::from_rtf_bytes(unsupported).unwrap();
    assert!(parsed.diagnostics.iter().any(|diagnostic| {
        diagnostic.message == "unsupported RTF list number format 255 converted to decimal"
    }));
    assert!(Document::from_rtf_bytes(br"{\rtf1\ansi\ls99 missing}").is_err());
}

#[test]
fn rtf_pictures_keep_run_and_cell_order_while_scaling_and_diagnosing_crop() {
    const PNG: &str = "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d49444154789c6360f8cff00000040101089d1de10000000049454e44ae426082";
    let body = format!(
        "{{\\rtf1\\ansi before{{\\pict\\pngblip\\picwgoal100\\pichgoal200\\picscalex200\\picscaley50\\piccropl10 {PNG}}}after\\par}}"
    );
    let parsed = Document::from_rtf_bytes(body.as_bytes()).unwrap();
    let paragraph = parsed.document.paragraph(0).unwrap();
    assert_eq!(paragraph.run_count(), 3);
    assert_eq!(paragraph.run(0).unwrap().text(), "before");
    assert!(paragraph.run(1).unwrap().inline_image().is_some());
    assert_eq!(paragraph.run(2).unwrap().text(), "after");
    let image = &parsed.document.images()[0];
    assert_eq!((image.width_emu, image.height_emu), (127_000, 63_500));
    assert!(
        parsed
            .diagnostics
            .iter()
            .any(|diagnostic| { diagnostic.message == "RTF picture cropping was dropped" })
    );

    let table_source = format!(
        "{{\\rtf1\\ansi\\trowd\\cellx1440\\intbl before{{\\pict\\pngblip {PNG}}}after\\cell\\row}}"
    );
    let parsed = Document::from_rtf_bytes(table_source.as_bytes()).unwrap();
    assert_eq!(parsed.document.paragraph_count(), 0);
    let table = parsed.document.table(0).unwrap();
    let cell = table.cell(0, 0).unwrap();
    let paragraph = cell.paragraph(0).unwrap();
    assert_eq!(paragraph.run_count(), 3);
    assert!(paragraph.run(1).unwrap().inline_image().is_some());

    let after_table =
        format!("{{\\rtf1\\ansi\\trowd\\cellx1440 cell\\cell\\row{{\\pict\\pngblip {PNG}}}}}");
    let parsed = Document::from_rtf_bytes(after_table.as_bytes()).unwrap();
    let items = parsed.document.body_items().collect::<Vec<_>>();
    assert!(matches!(items[0], BodyItemRef::Table(_)));
    assert!(matches!(items[1], BodyItemRef::Paragraph(_)));
    assert!(
        parsed
            .document
            .paragraph(0)
            .unwrap()
            .run(0)
            .unwrap()
            .inline_image()
            .is_some()
    );
}

#[test]
fn differing_rtf_row_boundaries_are_reported_as_lossy() {
    let parsed = Document::from_rtf_bytes(
        br"{\rtf1\trowd\cellx1440 one\cell\row\trowd\cellx2880 two\cell\row}",
    )
    .unwrap();
    assert!(parsed.diagnostics.iter().any(|diagnostic| {
        diagnostic.message == "RTF table row boundaries differ from the first row"
    }));
}

fn document_xml(document: &mut Document) -> Vec<u8> {
    let bytes = document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    package.get_part("/word/document.xml").unwrap().to_vec()
}

#[test]
fn m23_layout_and_data_tables_match_word() {
    let mut document = Document::new();
    {
        let mut table = document.add_table(2, 3);
        table
            .set_grid_widths(&[
                Length::twips(1_200),
                Length::twips(2_400),
                Length::twips(1_800),
            ])
            .unwrap();
        table.set_width_mode(TableWidth::Percentage(72.5)).unwrap();
        table.set_indent_checked(Length::twips(240)).unwrap();
        table.set_alignment(Alignment::Center);
        table.set_layout(TableLayout::Fixed);
        table.set_shading_checked("D9EAF7").unwrap();
        table
            .set_all_borders_checked(BorderStyle::Single, 8, "336699")
            .unwrap();
        table
            .set_border_checked(TableBorderEdge::Top, BorderStyle::None, 0, "auto")
            .unwrap();
        table
            .set_cell_margins_checked(
                Length::twips(80),
                Length::twips(120),
                Length::twips(80),
                Length::twips(120),
            )
            .unwrap();
        table.set_look(TableLook {
            first_row: true,
            last_row: false,
            first_column: true,
            last_column: false,
            horizontal_banding: true,
            vertical_banding: false,
        });
        table.cell(0, 0).unwrap().set_text("Capability");
        table.cell(0, 1).unwrap().set_text("Owner");
        table.cell(0, 2).unwrap().set_text("Status");
        table.cell(1, 0).unwrap().set_text("Public facade");
        table.cell(1, 1).unwrap().set_text("M23");
        table.cell(1, 2).unwrap().set_text("Pass");
    }

    let bytes = document.to_bytes().unwrap();
    let mut reopened = Document::from_bytes(&bytes).unwrap();
    let table = reopened.table(0).unwrap();
    assert_eq!(table.width_mode(), Some(TableWidth::Percentage(72.5)));
    assert_eq!(table.indent(), Some(Length::twips(240)));
    assert_eq!(table.layout(), Some(TableLayout::Fixed));
    assert_eq!(table.shading_fill(), Some("D9EAF7"));
    assert_eq!(
        table.cell_margins(),
        Some(TableCellMargins {
            top: Some(Length::twips(80)),
            right: Some(Length::twips(120)),
            bottom: Some(Length::twips(80)),
            left: Some(Length::twips(120)),
        })
    );
    assert_eq!(table.border(TableBorderEdge::Top).unwrap().style(), "none");
    assert_eq!(
        table.grid_widths(),
        vec![
            Length::twips(1_200),
            Length::twips(2_400),
            Length::twips(1_800),
        ]
    );
    assert_eq!(
        table.look(),
        Some(TableLook {
            first_row: true,
            last_row: false,
            first_column: true,
            last_column: false,
            horizontal_banding: true,
            vertical_banding: false,
        })
    );
    assert_eq!(
        table.cell(0, 0).unwrap().width(),
        Some(Length::twips(1_200))
    );
    assert_eq!(
        table.cell(0, 1).unwrap().width(),
        Some(Length::twips(2_400))
    );
    assert_eq!(
        table.cell(0, 2).unwrap().width(),
        Some(Length::twips(1_800))
    );

    let layout = reopened.layout_deterministic().unwrap();
    let fragments = layout.body_layout_fragments(0).unwrap();
    assert_eq!(fragments.len(), 1);
    assert!(
        (fragments[0].width - 270.0).abs() < 0.01,
        "{}",
        fragments[0].width
    );
    let first_render = reopened
        .render_page_to_png_deterministic(0, 72.0)
        .unwrap()
        .unwrap();
    let second_render = reopened
        .render_page_to_png_deterministic(0, 72.0)
        .unwrap()
        .unwrap();
    assert_eq!(first_render, second_render);
    assert!(first_render.starts_with(b"\x89PNG\r\n\x1a\n"));

    let xml = String::from_utf8(document_xml(&mut reopened)).unwrap();
    let table_properties = &xml[xml.find("<w:tblPr>").unwrap()..xml.find("</w:tblPr>").unwrap()];
    let positions = [
        "<w:tblW",
        "<w:jc",
        "<w:tblInd",
        "<w:tblBorders",
        "<w:shd",
        "<w:tblLayout",
        "<w:tblCellMar",
        "<w:tblLook",
    ]
    .map(|element| table_properties.find(element).unwrap());
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
    assert!(
        table_properties.contains(r#"<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>"#)
    );

    let mut paginated = Document::new();
    {
        let mut table = paginated.add_table(80, 1);
        table.set_grid_widths(&[Length::twips(4_000)]).unwrap();
        table
            .set_width_mode(TableWidth::Fixed(Length::twips(4_000)))
            .unwrap();
        table.set_layout(TableLayout::Fixed);
        for row in 0..80 {
            table
                .cell(row, 0)
                .unwrap()
                .set_text(&format!("Reviewed data row {row:02}"));
        }
    }
    let pagination = paginated.layout_deterministic().unwrap();
    assert_eq!(pagination.layout.pages.len(), 3);
    assert_eq!(pagination.body_layout_fragments(0).unwrap().len(), 3);
}

#[test]
fn complete_table_width_modes_round_trip() {
    let mut document = Document::new();
    {
        let mut table = document.add_table(1, 1);
        table.set_width_mode(TableWidth::Auto).unwrap();
        table.set_layout(TableLayout::AutoFit);
    }
    document
        .add_table(1, 1)
        .set_width_mode(TableWidth::Fixed(Length::twips(3_600)))
        .unwrap();
    document
        .add_table(1, 1)
        .set_width_mode(TableWidth::Percentage(55.5))
        .unwrap();

    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table(0).unwrap().width_mode(),
        Some(TableWidth::Auto)
    );
    assert_eq!(
        reopened.table(0).unwrap().layout(),
        Some(TableLayout::AutoFit)
    );
    assert_eq!(
        reopened.table(1).unwrap().width_mode(),
        Some(TableWidth::Fixed(Length::twips(3_600)))
    );
    assert_eq!(
        reopened.table(2).unwrap().width_mode(),
        Some(TableWidth::Percentage(55.5))
    );
}

#[test]
fn checked_table_setters_are_atomic() {
    let mut document = Document::new();
    document.add_table(1, 2);
    let before = document.to_bytes().unwrap();

    {
        let mut table = document.table_mut(0).unwrap();
        assert!(
            table
                .set_width_mode(TableWidth::Percentage(f64::NAN))
                .is_err()
        );
        assert!(
            table
                .set_width_mode(TableWidth::Fixed(Length::emu(i64::MAX)))
                .is_err()
        );
        assert!(table.set_indent_checked(Length::twips(-1)).is_err());
        assert!(table.set_shading_checked("not-a-color").is_err());
        assert!(
            table
                .set_all_borders_checked(BorderStyle::Single, 0, "000000")
                .is_err()
        );
        assert!(
            table
                .set_border_checked(TableBorderEdge::Left, BorderStyle::Single, 8, "12345G")
                .is_err()
        );
        assert!(
            table
                .set_cell_margins_checked(
                    Length::twips(0),
                    Length::twips(0),
                    Length::twips(-1),
                    Length::twips(0),
                )
                .is_err()
        );
        assert!(table.set_grid_widths(&[Length::twips(1_000)]).is_err());
        assert!(
            table
                .set_grid_widths(&[Length::twips(i32::MAX), Length::twips(i32::MAX),])
                .is_err()
        );
    }

    assert_eq!(document.to_bytes().unwrap(), before);

    document
        .table_mut(0)
        .unwrap()
        .cell(0, 1)
        .unwrap()
        .set_grid_span(0);
    let zero_span = document.to_bytes().unwrap();
    assert!(
        document
            .table_mut(0)
            .unwrap()
            .set_grid_widths(&[Length::twips(1_000), Length::twips(2_000)])
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), zero_span);

    document
        .table_mut(0)
        .unwrap()
        .cell(0, 1)
        .unwrap()
        .set_grid_span(2);
    let invalid_coverage = document.to_bytes().unwrap();
    assert!(
        document
            .table_mut(0)
            .unwrap()
            .set_grid_widths(&[Length::twips(1_000), Length::twips(2_000)])
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), invalid_coverage);
}

#[test]
fn checked_table_mutation_preserves_raw_property_slots_and_border_extensions() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    package.set_part(
        "/word/document.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ext="urn:m23-table">
  <q:body>
    <q:tbl>
      <q:tblPr>
        <ext:pre ext:value="exact"/>
        <q:tblW q:w="2400" q:type="dxa"/>
        <q:tblBorders><q:top q:val="single" q:sz="8" q:color="112233"/><ext:diagonal ext:value="exact"/></q:tblBorders>
        <ext:after ext:value="exact"/>
        <q:tblLook q:val="04A0"/>
      </q:tblPr>
      <q:tblGrid><q:gridCol q:w="1200"/><q:gridCol q:w="1200"/></q:tblGrid>
      <q:tr><q:tc><q:tcPr><q:tcW q:w="2400" q:type="dxa"/><q:gridSpan q:val="2"/></q:tcPr><q:p/></q:tc></q:tr>
    </q:tbl>
    <q:sectPr/>
  </q:body>
</q:document>"#
            .to_vec(),
    );
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();

    let mut document = Document::from_bytes(input.get_ref()).unwrap();
    assert_eq!(
        document.table(0).unwrap().look(),
        Some(TableLook {
            first_row: true,
            last_row: false,
            first_column: true,
            last_column: false,
            horizontal_banding: true,
            vertical_banding: false,
        })
    );
    let first_xml = String::from_utf8(document_xml(&mut document)).unwrap();
    {
        let mut table = document.table_mut(0).unwrap();
        table
            .set_grid_widths(&[Length::twips(1_200), Length::twips(2_400)])
            .unwrap();
        table
            .set_border_checked(TableBorderEdge::Bottom, BorderStyle::None, 0, "auto")
            .unwrap();
        table.set_shading_checked("ABCDEF").unwrap();
    }
    let second_xml = String::from_utf8(document_xml(&mut document)).unwrap();
    assert_eq!(
        document.table(0).unwrap().cell(0, 0).unwrap().width(),
        Some(Length::twips(3_600))
    );

    let raw_element = |xml: &str, marker: &str| {
        let start = xml.find(&format!("<{marker}")).unwrap();
        let end = xml[start..].find("/>").unwrap() + start + 2;
        xml[start..end].to_owned()
    };
    for marker in ["ext:pre", "ext:diagonal", "ext:after"] {
        assert_eq!(first_xml.matches(marker).count(), 1, "{first_xml}");
        assert_eq!(second_xml.matches(marker).count(), 1, "{second_xml}");
        assert_eq!(
            raw_element(&first_xml, marker),
            raw_element(&second_xml, marker)
        );
    }
    assert!(second_xml.contains(r#"<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>"#));
    let width = second_xml.find("<w:tblW").unwrap();
    let borders = second_xml.find("<w:tblBorders").unwrap();
    let after = second_xml.find("<ext:after").unwrap();
    let shading = second_xml.find("<w:shd").unwrap();
    let look = second_xml.find("<w:tblLook").unwrap();
    assert!(
        width < borders && borders < after && after < shading && shading < look,
        "{second_xml}"
    );
}

fn m23_row_cell_oracle_source() -> Document {
    let mut document = Document::new();
    {
        let mut table = document.add_table(6, 3);
        table
            .set_grid_widths(&[
                Length::twips(1_200),
                Length::twips(1_800),
                Length::twips(2_400),
            ])
            .unwrap();
        table.set_row_grid_omissions(1, Some(1), None).unwrap();
        table.set_cell_grid_span_checked(2, 0, Some(2)).unwrap();
        table
            .set_cell_vertical_merge(0, 1, Some(rdocx::table::VMerge::Restart))
            .unwrap();
        table
            .set_cell_vertical_merge(1, 0, Some(rdocx::table::VMerge::Continue))
            .unwrap();

        {
            let mut row = table.row(0).unwrap();
            row.set_height_checked(RowHeight::AtLeast(Length::twips(480)))
                .unwrap();
            row.set_header_value(Some(true));
            row.set_cant_split_value(Some(false));
            row.set_alignment(Alignment::Center);
            row.set_conditional_formatting(TableConditionalFormatting {
                first_row: true,
                ..Default::default()
            });
        }
        {
            let mut cell = table.cell(0, 1).unwrap();
            cell.set_width_checked(Length::twips(1_800)).unwrap();
            cell.set_border_checked(CellBorderEdge::Bottom, BorderStyle::Single, 8, "336699")
                .unwrap();
            cell.set_margins_checked(
                Length::twips(60),
                Length::twips(90),
                Length::twips(60),
                Length::twips(90),
            )
            .unwrap();
            cell.set_shading_checked("D9EAF7").unwrap();
            cell.set_vertical_alignment(VerticalAlignment::Center);
            cell.set_text_direction(Some(CellTextDirection::TopToBottomRightToLeft));
            cell.set_conditional_formatting(TableConditionalFormatting {
                first_row: true,
                first_column: true,
                ..Default::default()
            });
            cell.set_no_wrap_value(Some(false));
            let mut nested = cell.add_table_checked(1, 2).unwrap();
            nested
                .set_grid_widths(&[Length::twips(700), Length::twips(900)])
                .unwrap();
        }
        for (row, cell, direction) in [
            (0, 0, CellTextDirection::LeftToRightTopToBottom),
            (0, 1, CellTextDirection::TopToBottomRightToLeft),
            (1, 1, CellTextDirection::BottomToTopLeftToRight),
            (2, 0, CellTextDirection::LeftToRightTopToBottomVertical),
            (3, 0, CellTextDirection::TopToBottomRightToLeftVertical),
            (4, 0, CellTextDirection::TopToBottomLeftToRightVertical),
        ] {
            table
                .cell(row, cell)
                .unwrap()
                .set_text_direction(Some(direction));
        }
    }
    document
}

/// Normalize the stable facts retained by Word's writer. Exact local checks
/// separately cover explicit false values, default direction, contextual
/// conditional markers, and absent heights that Word canonicalizes.
fn m23_row_cell_oracle_records(document: &Document) -> Vec<String> {
    let table = document.table(0).unwrap();
    let mut records = vec![format!(
        "table | rows={} | grid={}",
        table.row_count(),
        table
            .grid_widths()
            .iter()
            .map(|width| width.to_twips().to_string())
            .collect::<Vec<_>>()
            .join(",")
    )];
    for row_index in 0..table.row_count() {
        let row = table.row(row_index).unwrap();
        records.push(format!(
            "row {row_index} | cells={} | before={:?} | after={:?}",
            row.cell_count(),
            row.grid_before(),
            row.grid_after(),
        ));
        if row_index == 0 {
            let height = match row.height() {
                Some(RowHeight::AtLeast(value)) => format!("atLeast:{}", value.to_twips()),
                Some(RowHeight::Exact(value)) => format!("exact:{}", value.to_twips()),
                None => "none".to_owned(),
            };
            records.push(format!(
                "row 0 format | height={height} | header={:?} | alignment={:?}",
                row.header_value(),
                row.alignment(),
            ));
        }
        for cell_index in 0..row.cell_count() {
            let cell = row.cell(cell_index).unwrap();
            if let Some(direction) = cell
                .text_direction()
                .filter(|value| *value != CellTextDirection::LeftToRightTopToBottom)
            {
                records.push(format!(
                    "direction | row={row_index} | cell={cell_index} | value={direction:?}"
                ));
            }
        }
    }
    let cell = table.cell(0, 1).unwrap();
    let margins = cell.margins().unwrap();
    records.push(format!(
        "cell 0:1 | width={} | border={}:{} | margins={},{},{},{} | shading={} | valign={:?} | nested={}",
        cell.width().unwrap().to_twips(),
        cell.border(CellBorderEdge::Bottom).unwrap().style(),
        cell.border(CellBorderEdge::Bottom).unwrap().color().unwrap(),
        margins.top.unwrap().to_twips(),
        margins.right.unwrap().to_twips(),
        margins.bottom.unwrap().to_twips(),
        margins.left.unwrap().to_twips(),
        cell.shading_fill().unwrap(),
        cell.vertical_alignment(),
        cell.items()
            .filter(|item| matches!(item, rdocx::CellItemRef::Table(_)))
            .count(),
    ));
    let span = table.cell(2, 0).unwrap();
    records.push(format!(
        "cell 2:0 | span={:?} | width={}",
        span.grid_span(),
        span.width().unwrap().to_twips(),
    ));
    records.push(format!(
        "merge | restart={:?} | continuation={:?}",
        table.cell(0, 1).unwrap().v_merge(),
        table.cell(1, 0).unwrap().v_merge(),
    ));
    records
}

#[test]
fn m23_nested_rows_and_cells_match_word() {
    assert_eq!(
        WORD_ROW_CELL_ORACLE,
        "Microsoft Word 16.112.4 build 16.112.26090911"
    );
    let mut document = m23_row_cell_oracle_source();

    let mut reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    assert_eq!(
        m23_row_cell_oracle_records(&reopened),
        WORD_ROW_CELL_RECORDS
    );
    {
        let table = reopened.table(0).unwrap();
        let row = table.row(0).unwrap();
        assert_eq!(row.height(), Some(RowHeight::AtLeast(Length::twips(480))));
        assert_eq!(row.header_value(), Some(true));
        assert_eq!(row.cant_split_value(), Some(false));
        assert_eq!(row.alignment(), Some(Alignment::Center));
        assert_eq!(
            row.conditional_formatting(),
            Some(TableConditionalFormatting {
                first_row: true,
                ..Default::default()
            })
        );
        assert_eq!(table.row(1).unwrap().grid_before(), Some(1));
        assert_eq!(table.row(1).unwrap().cell_count(), 2);
        assert_eq!(table.row(2).unwrap().cell_count(), 2);
        assert_eq!(table.cell(2, 0).unwrap().grid_span(), Some(2));
        assert_eq!(
            table.cell(2, 0).unwrap().width(),
            Some(Length::twips(3_000))
        );
        assert_eq!(
            table.cell(0, 1).unwrap().v_merge(),
            Some(&rdocx::table::VMerge::Restart)
        );
        assert_eq!(
            table.cell(1, 0).unwrap().v_merge(),
            Some(&rdocx::table::VMerge::Continue)
        );
        let cell = table.cell(0, 1).unwrap();
        assert_eq!(cell.width(), Some(Length::twips(1_800)));
        assert_eq!(
            cell.border(CellBorderEdge::Bottom).unwrap().color(),
            Some("336699")
        );
        assert_eq!(
            cell.margins(),
            Some(TableCellMargins {
                top: Some(Length::twips(60)),
                right: Some(Length::twips(90)),
                bottom: Some(Length::twips(60)),
                left: Some(Length::twips(90)),
            })
        );
        assert_eq!(cell.shading_fill(), Some("D9EAF7"));
        assert_eq!(cell.vertical_alignment(), Some(VerticalAlignment::Center));
        assert_eq!(
            cell.text_direction(),
            Some(CellTextDirection::TopToBottomRightToLeft)
        );
        assert_eq!(cell.no_wrap_value(), Some(false));
        assert_eq!(
            cell.conditional_formatting(),
            Some(TableConditionalFormatting {
                first_row: true,
                first_column: true,
                ..Default::default()
            })
        );
        assert_eq!(
            cell.items()
                .filter(|item| matches!(item, rdocx::CellItemRef::Table(_)))
                .count(),
            1
        );
        for (row, cell, direction) in [
            (0, 0, CellTextDirection::LeftToRightTopToBottom),
            (0, 1, CellTextDirection::TopToBottomRightToLeft),
            (1, 1, CellTextDirection::BottomToTopLeftToRight),
            (2, 0, CellTextDirection::LeftToRightTopToBottomVertical),
            (3, 0, CellTextDirection::TopToBottomRightToLeftVertical),
            (4, 0, CellTextDirection::TopToBottomLeftToRightVertical),
        ] {
            assert_eq!(
                table.cell(row, cell).unwrap().text_direction(),
                Some(direction)
            );
        }
    }

    let xml = String::from_utf8(document_xml(&mut reopened)).unwrap();
    assert!(
        xml.contains(r#"<w:cantSplit w:val="false"/>"#),
        "{WORD_ROW_CELL_ORACLE}\n{xml}"
    );
    assert!(xml.contains(r#"<w:noWrap w:val="false"/>"#), "{xml}");
    assert!(xml.contains(r#"<w:gridBefore w:val="1"/>"#), "{xml}");
    assert!(xml.contains(r#"<w:gridSpan w:val="2"/>"#), "{xml}");
    let row_properties = &xml[xml.find("<w:trPr>").unwrap()..xml.find("</w:trPr>").unwrap()];
    let row_positions = [
        "<w:cnfStyle",
        "<w:cantSplit",
        "<w:trHeight",
        "<w:tblHeader",
        "<w:jc",
    ]
    .map(|element| row_properties.find(element).unwrap());
    assert!(row_positions.windows(2).all(|pair| pair[0] < pair[1]));
    let cell_marker = xml.find("D9EAF7").unwrap();
    let cell_start = xml[..cell_marker].rfind("<w:tcPr>").unwrap();
    let cell_end = xml[cell_marker..].find("</w:tcPr>").unwrap() + cell_marker;
    let cell_properties = &xml[cell_start..cell_end];
    let cell_positions = [
        "<w:cnfStyle",
        "<w:tcW",
        "<w:vMerge",
        "<w:tcBorders",
        "<w:shd",
        "<w:noWrap",
        "<w:tcMar",
        "<w:textDirection",
        "<w:vAlign",
    ]
    .map(|element| cell_properties.find(element).unwrap());
    assert!(cell_positions.windows(2).all(|pair| pair[0] < pair[1]));

    let first = reopened
        .render_page_to_png_deterministic(0, 72.0)
        .unwrap()
        .unwrap();
    let second = reopened
        .render_page_to_png_deterministic(0, 72.0)
        .unwrap()
        .unwrap();
    assert_eq!(first, second);
}

#[test]
#[ignore = "requires pinned Microsoft Word 16.112.4 GUI automation"]
fn regenerate_f258_word_row_cell_oracle() {
    let plist = "/Applications/Microsoft Word.app/Contents/Info.plist";
    let version = std::process::Command::new("plutil")
        .args(["-extract", "CFBundleShortVersionString", "raw", plist])
        .output()
        .unwrap();
    let build = std::process::Command::new("plutil")
        .args(["-extract", "CFBundleVersion", "raw", plist])
        .output()
        .unwrap();
    assert_eq!(String::from_utf8_lossy(&version.stdout).trim(), "16.112.4");
    assert_eq!(
        String::from_utf8_lossy(&build.stdout).trim(),
        "16.112.26090911"
    );
    assert_eq!(
        WORD_ROW_CELL_ORACLE,
        "Microsoft Word 16.112.4 build 16.112.26090911"
    );

    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let directory = std::path::Path::new(
        "/Users/atulsharma/Library/Containers/com.microsoft.Word/Data/Documents/rdocx-f258-word-oracle",
    )
    .join(format!("{}-{nonce}", std::process::id()));
    let artifacts = F251OracleArtifacts::create(directory);
    let source = artifacts.path().join("f258-source.docx");
    let output = artifacts.path().join("f258-word.docx");
    m23_row_cell_oracle_source().save(&source).unwrap();

    let script = format!(
        r#"with timeout of 60 seconds
tell application "Microsoft Word"
activate
open POSIX file "{}"
delay 3
set oracleDocument to active document
save as oracleDocument file name "{}" file format format document default add to recent files false
close oracleDocument saving no
end tell
end timeout"#,
        source.display(),
        output.display(),
    );
    let word = std::process::Command::new("osascript")
        .args(["-e", &script])
        .output()
        .unwrap();
    assert!(
        word.status.success(),
        "Word DOCX save failed: {}",
        String::from_utf8_lossy(&word.stderr)
    );

    let oracle = Document::open(&output).unwrap();
    let records = m23_row_cell_oracle_records(&oracle);
    for record in &records {
        println!("{record:?},");
    }
    assert_eq!(records, WORD_ROW_CELL_RECORDS);
}

#[test]
fn complete_row_and_cell_properties_round_trip() {
    let mut document = Document::new();
    let mut table = document.add_table(6, 2);
    table
        .row(0)
        .unwrap()
        .set_height_checked(RowHeight::Exact(Length::twips(720)))
        .unwrap();
    table.row(0).unwrap().set_header_value(Some(false));
    table.row(0).unwrap().set_cant_split_value(None);
    let directions = [
        CellTextDirection::LeftToRightTopToBottom,
        CellTextDirection::TopToBottomRightToLeft,
        CellTextDirection::BottomToTopLeftToRight,
        CellTextDirection::LeftToRightTopToBottomVertical,
        CellTextDirection::TopToBottomRightToLeftVertical,
        CellTextDirection::TopToBottomLeftToRightVertical,
    ];
    let edges = [
        CellBorderEdge::Top,
        CellBorderEdge::Bottom,
        CellBorderEdge::Left,
        CellBorderEdge::Right,
        CellBorderEdge::InsideHorizontal,
        CellBorderEdge::InsideVertical,
    ];
    for (index, (direction, edge)) in directions.into_iter().zip(edges).enumerate() {
        let mut cell = table.cell(index, 0).unwrap();
        cell.set_text_direction(Some(direction));
        cell.set_border_checked(edge, BorderStyle::Single, 8, "123456")
            .unwrap();
    }
    table.cell(0, 0).unwrap().set_no_wrap_value(None);
    table.cell(0, 1).unwrap().set_no_wrap_value(Some(true));

    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    let reopened_table = reopened.table(0).unwrap();
    let row = reopened_table.row(0).unwrap();
    assert_eq!(row.height(), Some(RowHeight::Exact(Length::twips(720))));
    assert_eq!(row.header_value(), Some(false));
    assert_eq!(row.cant_split_value(), None);
    for (index, (direction, edge)) in directions.into_iter().zip(edges).enumerate() {
        let cell = reopened_table.cell(index, 0).unwrap();
        assert_eq!(cell.text_direction(), Some(direction));
        assert_eq!(cell.border(edge).unwrap().style(), "single");
    }
    assert_eq!(reopened_table.cell(0, 0).unwrap().no_wrap_value(), None);
    assert_eq!(
        reopened_table.cell(0, 1).unwrap().no_wrap_value(),
        Some(true)
    );
}

#[test]
fn checked_row_cell_topology_is_atomic() {
    let mut document = Document::new();
    document.add_table(2, 2);
    let before = document.to_bytes().unwrap();

    assert!(
        document
            .table_mut(0)
            .unwrap()
            .set_row_grid_omissions(0, Some(2), Some(1))
            .is_err()
    );
    assert!(
        document
            .table_mut(0)
            .unwrap()
            .set_cell_grid_span_checked(0, 0, Some(0))
            .is_err()
    );
    assert!(
        document
            .table_mut(0)
            .unwrap()
            .set_cell_vertical_merge(1, 0, Some(rdocx::table::VMerge::Continue))
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), before);

    {
        let mut table = document.table_mut(0).unwrap();
        assert!(
            table
                .row(0)
                .unwrap()
                .set_height_checked(RowHeight::Exact(Length::twips(-1)))
                .is_err()
        );
        assert!(
            table
                .cell(0, 0)
                .unwrap()
                .set_width_checked(Length::emu(i64::MAX))
                .is_err()
        );
        assert!(
            table
                .cell(0, 0)
                .unwrap()
                .set_border_checked(CellBorderEdge::Top, BorderStyle::Single, 0, "000000")
                .is_err()
        );
        assert!(
            table
                .cell(0, 0)
                .unwrap()
                .set_shading_checked("invalid")
                .is_err()
        );
        assert!(table.cell(0, 0).unwrap().add_table_checked(0, 1).is_err());
    }
    assert_eq!(document.to_bytes().unwrap(), before);

    document
        .table_mut(0)
        .unwrap()
        .cell(0, 1)
        .unwrap()
        .set_text("keep");
    let nonempty = document.to_bytes().unwrap();
    assert!(
        document
            .table_mut(0)
            .unwrap()
            .set_cell_grid_span_checked(0, 0, Some(2))
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), nonempty);

    let mut span_document = Document::new();
    span_document.add_table(1, 3);
    let original_widths = span_document.table(0).unwrap().grid_widths();
    let merged_width = original_widths
        .iter()
        .map(|width| width.to_twips())
        .sum::<i32>();
    span_document
        .table_mut(0)
        .unwrap()
        .set_cell_grid_span_checked(0, 0, Some(3))
        .unwrap();
    assert_eq!(
        span_document.table(0).unwrap().cell(0, 0).unwrap().width(),
        Some(Length::twips(merged_width))
    );
    span_document
        .table_mut(0)
        .unwrap()
        .set_cell_grid_span_checked(0, 0, None)
        .unwrap();
    let table = span_document.table(0).unwrap();
    assert_eq!(table.row(0).unwrap().cell_count(), 3);
    for (cell, width) in original_widths.into_iter().enumerate() {
        assert_eq!(table.cell(0, cell).unwrap().width(), Some(width));
    }
}

#[test]
fn content_measurement_reuses_production_layout_and_is_pure() {
    let mut document = Document::new();
    document.add_paragraph(
        "A caller-width paragraph must wrap through the production deterministic line breaker.",
    );
    let story = document
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::Body)
        .unwrap();
    let location = document
        .story_items(&story)
        .unwrap()
        .into_iter()
        .find(|item| item.kind() == StoryItemKind::Paragraph)
        .unwrap()
        .location()
        .clone();
    let before = document.to_bytes().unwrap();
    let cached = document.layout_deterministic().unwrap();

    let narrow = document
        .measure_content(&location, Length::pt(90.0), rdocx::RenderOptions::default())
        .unwrap();
    let wide = document
        .measure_content(
            &location,
            Length::pt(360.0),
            rdocx::RenderOptions::default(),
        )
        .unwrap();

    assert!(narrow.height_points > wide.height_points);
    assert_eq!(narrow.diagnostics, wide.diagnostics);
    assert!(
        document
            .measure_content(&location, Length::emu(0), rdocx::RenderOptions::default(),)
            .is_err()
    );
    let wrong_kind =
        rdocx::ContentLocation::new(story, StoryItemKind::Table, location.index_path().to_vec());
    assert!(
        document
            .measure_content(
                &wrong_kind,
                Length::pt(90.0),
                rdocx::RenderOptions::default(),
            )
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), before);
    assert!(std::sync::Arc::ptr_eq(
        &cached,
        &document.layout_deterministic().unwrap()
    ));
}

#[test]
fn content_measurement_preserves_layout_diagnostic_order() {
    const MATH: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    let mut document = Document::new();
    let mut paragraph = document.add_paragraph("");
    for name in ["firstUnsupported", "secondUnsupported"] {
        let equation = rdocx::CT_OMath::from_xml(
            format!(r#"<m:oMath xmlns:m="{MATH}"><m:{name}/></m:oMath>"#).as_bytes(),
        )
        .unwrap();
        paragraph
            .add_equation(rdocx::OfficeMath::Inline(equation))
            .unwrap();
    }
    let body = document
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::Body)
        .unwrap();
    let location = document.story_items(&body).unwrap()[0].location().clone();

    let measurement = document
        .measure_content(
            &location,
            Length::pt(180.0),
            rdocx::RenderOptions::default(),
        )
        .unwrap();
    assert_eq!(
        measurement
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.as_str())
            .collect::<Vec<_>>(),
        [
            "OfficeMath content at paragraph/run-boundary/0/raw-child/0 was preserved but could not be rendered",
            "OfficeMath content at paragraph/run-boundary/0/raw-child/1 was preserved but could not be rendered",
        ]
    );
}

#[test]
fn independent_nested_tables_measure_to_one_final_height() {
    let margin = Length::pt(6.0);
    let mut document = Document::new();
    {
        let mut table = document.add_table(1, 2);
        table.set_cell_grid_span_checked(0, 0, Some(2)).unwrap();
        table
            .cell(0, 0)
            .unwrap()
            .set_margins_checked(margin, margin, margin, margin)
            .unwrap();
        table
            .cell(0, 0)
            .unwrap()
            .set_border_checked(CellBorderEdge::Bottom, BorderStyle::Single, 8, "4472C4")
            .unwrap();
        let mut cell = table.cell(0, 0).unwrap();
        let mut nested = cell.add_table_checked(1, 1).unwrap();
        nested.cell(0, 0).unwrap().set_text("Short nested table");
    }
    {
        let mut table = document.add_table(1, 2);
        table.set_cell_grid_span_checked(0, 0, Some(2)).unwrap();
        table
            .cell(0, 0)
            .unwrap()
            .set_margins_checked(margin, margin, margin, margin)
            .unwrap();
        table
            .cell(0, 0)
            .unwrap()
            .set_border_checked(CellBorderEdge::Bottom, BorderStyle::Single, 8, "4472C4")
            .unwrap();
        let mut cell = table.cell(0, 0).unwrap();
        let mut nested = cell.add_table_checked(2, 1).unwrap();
        nested.cell(0, 0).unwrap().set_text(
            "A longer nested table cell wraps at the caller width and establishes the maximum.",
        );
        nested.cell(1, 0).unwrap().set_text("Second nested row");
    }
    let body = document
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == StoryKind::Body)
        .unwrap();
    let locations = document
        .story_items(&body)
        .unwrap()
        .into_iter()
        .filter(|item| item.kind() == StoryItemKind::Table)
        .map(|item| item.location().clone())
        .collect::<Vec<_>>();
    assert_eq!(locations.len(), 2);
    let measurements = locations
        .iter()
        .map(|location| {
            document
                .measure_content(location, Length::pt(180.0), rdocx::RenderOptions::default())
                .unwrap()
        })
        .collect::<Vec<_>>();
    let final_height = measurements
        .iter()
        .map(|measurement| measurement.height_points)
        .fold(0.0_f64, f64::max);
    let final_height_twips = (final_height * 20.0).ceil() as i32;
    assert!(measurements[0].height_points < final_height);
    assert!(
        measurements
            .iter()
            .all(|result| result.diagnostics.is_empty())
    );

    for index in 0..2 {
        document
            .table_mut(index)
            .unwrap()
            .row(0)
            .unwrap()
            .set_height_checked(RowHeight::AtLeast(Length::twips(final_height_twips)))
            .unwrap();
    }
    let layout = document
        .layout_deterministic_with_options(rdocx::RenderOptions::default())
        .unwrap();
    let first = layout.body_layout_fragments(0).unwrap();
    let second = layout.body_layout_fragments(1).unwrap();
    assert_eq!(first.len(), 1);
    assert_eq!(second.len(), 1);
    assert!((first[0].height - second[0].height).abs() < 0.001);
    // A minimum row height excludes the horizontal border bands and the cell's
    // top and bottom margins, as in Word, so each table keeps its 6 point
    // margins and its 1 point bottom border outside that minimum.
    assert!(
        (first[0].height - (f64::from(final_height_twips) / 20.0 + 12.0 + 1.0)).abs() < 0.001,
        "measured {final_height}, rounded to {final_height_twips} twips, laid out {}",
        first[0].height
    );
}

#[test]
fn checked_row_cell_mutation_preserves_raw_slots_and_aliases() {
    let mut seed = Document::new();
    let mut package =
        OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
    package.set_part(
        "/word/document.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ext="urn:m23-row-cell">
  <q:body>
    <q:tbl>
      <q:tblPr><q:tblW q:w="2400" q:type="dxa"/></q:tblPr>
      <q:tblGrid><q:gridCol q:w="1200"/><q:gridCol q:w="1200"/></q:tblGrid>
      <q:tr>
        <q:trPr><ext:rowPre ext:value="exact"/><ext:cnfStyle ext:val="000000000000"/><q:cnfStyle q:val="100000000000"/><ext:rowMid ext:value="exact"/><ext:trHeight ext:val="999"/><q:tblHeader q:val="false"/><ext:jc ext:val="left"/><ext:rowAfter ext:value="exact"/></q:trPr>
        <q:tc><q:tcPr><ext:cellPre ext:value="exact"/><q:tcW q:w="1200" q:type="dxa"/><q:tcBorders><q:top q:val="single" q:sz="8" q:color="112233"/><ext:diagonal ext:value="exact"/></q:tcBorders><ext:cellMid ext:value="exact"/><ext:noWrap ext:val="true"/><q:noWrap q:val="0"/><ext:textDirection ext:val="lrTb"/><q:textDirection q:val="btLr"/><ext:cellAfter ext:value="exact"/></q:tcPr><q:p/></q:tc>
        <q:tc><q:tcPr><q:tcW q:w="1200" q:type="dxa"/></q:tcPr><q:p/></q:tc>
      </q:tr>
    </q:tbl>
    <q:sectPr/>
  </q:body>
</q:document>"#
            .to_vec(),
    );
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();

    let mut document = Document::from_bytes(input.get_ref()).unwrap();
    assert_eq!(
        document.table(0).unwrap().row(0).unwrap().header_value(),
        Some(false)
    );
    assert_eq!(
        document
            .table(0)
            .unwrap()
            .cell(0, 0)
            .unwrap()
            .text_direction(),
        Some(CellTextDirection::BottomToTopLeftToRight)
    );
    let first_xml = String::from_utf8(document_xml(&mut document)).unwrap();
    {
        let mut table = document.table_mut(0).unwrap();
        let mut row = table.row(0).unwrap();
        row.set_height_checked(RowHeight::Exact(Length::twips(540)))
            .unwrap();
        row.set_alignment(Alignment::Right);
        row.set_conditional_formatting(TableConditionalFormatting {
            last_row: true,
            ..Default::default()
        });
    }
    {
        let mut table = document.table_mut(0).unwrap();
        let mut cell = table.cell(0, 0).unwrap();
        cell.set_border_checked(CellBorderEdge::Bottom, BorderStyle::None, 0, "auto")
            .unwrap();
        cell.set_margins_checked(
            Length::twips(40),
            Length::twips(50),
            Length::twips(60),
            Length::twips(70),
        )
        .unwrap();
        cell.set_shading_checked("ABCDEF").unwrap();
        cell.set_vertical_alignment(VerticalAlignment::Bottom);
        cell.set_text_direction(Some(CellTextDirection::TopToBottomRightToLeft));
        cell.set_conditional_formatting(TableConditionalFormatting {
            last_column: true,
            ..Default::default()
        });
        cell.set_no_wrap_value(Some(false));
    }
    let second_xml = String::from_utf8(document_xml(&mut document)).unwrap();

    let raw_element = |xml: &str, marker: &str| {
        let start = xml.find(&format!("<{marker}")).unwrap();
        let end = xml[start..].find("/>").unwrap() + start + 2;
        xml[start..end].to_owned()
    };
    for marker in [
        "ext:rowPre",
        "ext:cnfStyle",
        "ext:rowMid",
        "ext:trHeight",
        "ext:jc",
        "ext:rowAfter",
        "ext:cellPre",
        "ext:diagonal",
        "ext:cellMid",
        "ext:noWrap",
        "ext:textDirection",
        "ext:cellAfter",
    ] {
        assert_eq!(
            raw_element(&first_xml, marker),
            raw_element(&second_xml, marker)
        );
    }
    assert!(second_xml.contains(r#"<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>"#));
    assert!(second_xml.contains(r#"<w:noWrap w:val="false"/>"#));
    assert!(second_xml.contains(r#"<w:textDirection w:val="tbRl"/>"#));
    assert!(second_xml.contains(r#"<w:vAlign w:val="bottom"/>"#));
}

#[test]
fn row_cell_pagination_matrix_is_deterministic() {
    let mut document = Document::new();
    {
        let mut table = document.add_table(70, 1);
        table.row(0).unwrap().set_header_value(Some(true));
        for index in 0..70 {
            let mut row = table.row(index).unwrap();
            row.set_height_checked(if index % 2 == 0 {
                RowHeight::Exact(Length::twips(720))
            } else {
                RowHeight::AtLeast(Length::twips(480))
            })
            .unwrap();
            row.set_cant_split_value(Some(index % 3 == 0));
            let mut cell = row.cell(0).unwrap();
            let text = if index == 0 {
                "REPEATED HEADER".to_owned()
            } else {
                format!("reviewed row {index:02}")
            };
            cell.set_text(&text);
            cell.set_no_wrap_value(Some(index % 2 == 0));
            cell.set_text_direction(Some(if index % 2 == 0 {
                CellTextDirection::TopToBottomRightToLeft
            } else {
                CellTextDirection::LeftToRightTopToBottom
            }));
        }
    }
    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table(0).unwrap().row(0).unwrap().height(),
        Some(RowHeight::Exact(Length::twips(720)))
    );
    assert_eq!(
        reopened.table(0).unwrap().row(1).unwrap().height(),
        Some(RowHeight::AtLeast(Length::twips(480)))
    );
    assert_eq!(
        reopened
            .table(0)
            .unwrap()
            .cell(0, 0)
            .unwrap()
            .no_wrap_value(),
        Some(true)
    );
    assert_eq!(
        reopened
            .table(0)
            .unwrap()
            .cell(1, 0)
            .unwrap()
            .text_direction(),
        Some(CellTextDirection::LeftToRightTopToBottom)
    );

    let first = reopened.layout_deterministic().unwrap();
    let second = reopened.layout_deterministic().unwrap();
    let page_texts = |layout: &rdocx_layout::WordLayoutResult| {
        layout
            .layout
            .pages
            .iter()
            .map(|page| {
                let mut text = String::new();
                oxml_layout::walk(&page.elements, &mut |element, _| match element {
                    oxml_layout::PositionedElement::Text(run) => text.push_str(&run.text),
                    oxml_layout::PositionedElement::MultilingualText(run) => {
                        text.push_str(&run.logical_text)
                    }
                    _ => {}
                });
                text
            })
            .collect::<Vec<_>>()
    };
    let first_text = page_texts(&first);
    let second_text = page_texts(&second);
    assert_eq!(first_text, second_text);
    assert_eq!(first_text.len(), 4);
    assert!(
        first_text
            .iter()
            .all(|page| page.matches("REPEATED HEADER").count() == 1)
    );
    let joined = first_text.join("|");
    for index in 1..70 {
        assert_eq!(
            joined.matches(&format!("reviewed row {index:02}")).count(),
            1
        );
    }
    for page_index in 0..first_text.len() {
        let first_png = reopened
            .render_page_to_png_deterministic(page_index, 72.0)
            .unwrap()
            .unwrap();
        let second_png = reopened
            .render_page_to_png_deterministic(page_index, 72.0)
            .unwrap()
            .unwrap();
        assert_eq!(first_png, second_png);
    }
}

#[test]
fn public_body_items_preserve_opened_document_order() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let document_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:p="urn:producer">
  <w:body>
    <w:p><w:r><w:t>first</w:t></w:r></w:p>
    <w:p/>
    <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:sdt><w:sdtPr><w:tag w:val="public"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt>
    <p:opaque p:flag="keep"><p:child/></p:opaque>
    <w:p><w:r><w:t>last</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    package.set_part("/word/document.xml", document_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();

    let document = Document::from_bytes(input.get_ref()).unwrap();
    let items = document
        .body_items()
        .map(|item| match item {
            BodyItemRef::Paragraph(paragraph) => format!("paragraph:{}", paragraph.text()),
            BodyItemRef::Table(table) => format!(
                "table:{}:{}",
                table.row_count(),
                table.cell(0, 0).unwrap().text()
            ),
            BodyItemRef::ContentControl(control) => {
                format!("control:{}:{}", control.tag().unwrap(), control.text())
            }
            BodyItemRef::UnsupportedXml(raw) => {
                format!("raw:{}", std::str::from_utf8(raw).unwrap())
            }
        })
        .collect::<Vec<_>>();

    assert_eq!(
        items,
        [
            "paragraph:first",
            "paragraph:",
            "table:1:cell",
            "control:public:control",
            "raw:<p:opaque p:flag=\"keep\"><p:child/></p:opaque>",
            "paragraph:last",
        ]
    );
    assert_eq!(
        document
            .paragraphs()
            .iter()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>(),
        ["first", "", "control", "last"]
    );
    assert_eq!(document.tables().len(), 1);
}

#[test]
fn word_fractional_line_spacing_opens_with_nearest_twip_values() {
    let mut seed = Document::new();
    let bytes = seed.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let document_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:spacing w:line="257.1432" w:lineRule="auto"/></w:pPr><w:r><w:t>one</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:line="320.00879999999995" w:lineRule="auto"/></w:pPr><w:r><w:t>two</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:line="342.8616" w:lineRule="auto"/></w:pPr><w:r><w:t>three</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>"#;
    package.set_part("/word/document.xml", document_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();

    let mut document = Document::from_bytes(input.get_ref()).unwrap();
    let multiples = document
        .paragraphs()
        .iter()
        .map(|paragraph| paragraph.line_spacing_multiple().unwrap())
        .collect::<Vec<_>>();

    assert_eq!(multiples, [257.0 / 240.0, 320.0 / 240.0, 343.0 / 240.0]);

    let saved = document.to_bytes().unwrap();
    let saved_package = OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    let saved_xml =
        std::str::from_utf8(saved_package.get_part("/word/document.xml").unwrap()).unwrap();
    for value in ["257", "320", "343"] {
        assert!(saved_xml.contains(&format!(r#"w:line="{value}""#)));
    }
    for value in ["257.1432", "320.00879999999995", "342.8616"] {
        assert!(!saved_xml.contains(value));
    }
    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(
        reopened
            .paragraphs()
            .iter()
            .map(|paragraph| paragraph.line_spacing_multiple().unwrap())
            .collect::<Vec<_>>(),
        multiples
    );

    let integer_xml = String::from_utf8(document_xml.to_vec())
        .unwrap()
        .replace("257.1432", "257")
        .replace("320.00879999999995", "320")
        .replace("342.8616", "343");
    let mut integer_seed = Document::new();
    let mut integer_package =
        OpcPackage::from_reader(std::io::Cursor::new(integer_seed.to_bytes().unwrap())).unwrap();
    integer_package.set_part("/word/document.xml", integer_xml.into_bytes());
    let mut integer_bytes = std::io::Cursor::new(Vec::new());
    integer_package.write_to(&mut integer_bytes).unwrap();
    let integer_document = Document::from_bytes(integer_bytes.get_ref()).unwrap();
    assert_eq!(
        document.render_page_to_png_deterministic(0, 72.0).unwrap(),
        integer_document
            .render_page_to_png_deterministic(0, 72.0)
            .unwrap()
    );
}

#[test]
fn bounded_document_reader_rejects_package_expansion() {
    let bytes = Document::new().to_bytes().unwrap();
    let result = Document::from_bytes_with_limits(
        &bytes,
        PackageReadLimits {
            max_entries: 1,
            max_part_uncompressed_bytes: 1_024 * 1_024,
            max_total_uncompressed_bytes: 2 * 1_024 * 1_024,
        },
    );

    assert!(matches!(
        result,
        Err(rdocx::Error::Opc(
            oxml_opc::OpcError::PackageLimitExceeded {
                kind: "entry count",
                limit: 1
            }
        ))
    ));
}

#[test]
fn nested_revisions_are_reported_once_in_document_order() {
    let mut source = Document::new();
    let bytes = source.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let document_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:ins w:id="1" w:author="Ada"><w:r><w:t>top</w:t></w:r></w:ins></w:p>
    <w:tbl>
      <w:tblPr><w:tblPrChange w:id="2" w:author="Ben"><w:tblPr><w:jc w:val="center"/></w:tblPr></w:tblPrChange></w:tblPr>
      <w:tblGrid/><w:tr><w:trPr><w:ins w:id="3" w:author="Cy"/></w:trPr><w:tc>
        <w:p><w:del w:id="4" w:author="Dee"><w:r><w:delText>cell</w:delText></w:r></w:del></w:p>
      </w:tc></w:tr>
    </w:tbl>
    <w:sdt><w:sdtPr><w:tag w:val="outer"/></w:sdtPr><w:sdtContent><w:p>
      <w:moveFrom w:id="5" w:author="Eve"><w:r><w:t>control</w:t></w:r></w:moveFrom>
      <w:sdt><w:sdtPr><w:tag w:val="inner"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:del w:id="6" w:author="Fox"/></w:rPr><w:t>nested</w:t></w:r></w:sdtContent></w:sdt>
    </w:p></w:sdtContent></w:sdt>
    <w:sectPr><w:sectPrChange w:id="7" w:author="Gia"><w:sectPr><w:titlePg/></w:sectPr></w:sectPrChange></w:sectPr>
  </w:body>
</w:document>"#;
    package.set_part("/word/document.xml", document_xml.to_vec());
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();

    let document = Document::from_bytes(input.get_ref()).unwrap();
    let revisions = document.revisions();
    assert_eq!(
        revisions
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>(),
        vec![1, 2, 3, 4, 5, 6, 7]
    );
    assert_eq!(revisions[0].author(), "Ada");
    assert_eq!(revisions[0].timestamp(), None);
    assert_eq!(revisions[0].kind(), RevisionKind::Insertion);
    assert_eq!(revisions[3].kind(), RevisionKind::Deletion);
    assert_eq!(revisions[4].kind(), RevisionKind::MoveFrom);
    assert_eq!(revisions[6].kind(), RevisionKind::SectionPropertyChange);
}

#[test]
fn m14_collaboration_models_coexist_and_preserve_unmodelled_xml() {
    const PRESERVED_PART: &[u8] = b"producer-private-bytes\x00\xff";
    const OPAQUE_CONTROL: &[u8] =
        br#"<p:opaque-control p:flag="keep"><p:child/></p:opaque-control>"#;
    const OPAQUE_COMMENT: &[u8] = br#"<p:opaque-comment p:flag="keep"/>"#;

    let mut source = Document::new();
    let bytes = source.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let document_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:p="urn:producer">
  <w:body>
    <w:p><w:r><w:t>plain </w:t></w:r><w:ins w:id="10" w:author="Ada"><w:r><w:t>inserted</w:t></w:r></w:ins><w:del w:id="11" w:author="Ben"><w:r><w:delText>deleted</w:delText></w:r></w:del></w:p>
    <w:sdt><w:sdtPr><w:tag w:val="customer"/><p:opaque-control p:flag="keep"><p:child/></p:opaque-control></w:sdtPr><w:sdtContent><w:p><w:r><w:t>old value</w:t></w:r></w:p></w:sdtContent></w:sdt>
    <w:p><w:bookmarkStart w:id="30" w:name="existing"/><w:commentRangeStart w:id="20"/><w:r><w:t>anchor</w:t></w:r><w:bookmarkEnd w:id="30"/><w:commentRangeEnd w:id="20"/><w:r><w:commentReference w:id="20"/></w:r><w:r><w:t> tail</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>"#;
    let comments_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:p="urn:producer"><w:comment w:id="20" w:author="Ada"><w:p w14:paraId="A1B2C3D4"><w:r><w:t>existing comment</w:t></w:r><p:opaque-comment p:flag="keep"/></w:p></w:comment></w:comments>"#;
    package.set_part("/word/document.xml", document_xml.to_vec());
    package.set_part("/word/comments.xml", comments_xml.to_vec());
    package.content_types.add_override(
        "/word/comments.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    );
    package
        .get_or_create_part_rels("/word/document.xml")
        .add(rel_types::COMMENTS, "comments.xml");
    package.set_part("/custom/producer.bin", PRESERVED_PART.to_vec());
    package
        .content_types
        .add_override("/custom/producer.bin", "application/octet-stream");
    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();

    let mut document = Document::from_bytes(input.get_ref()).unwrap();
    assert_eq!(
        document
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>(),
        vec![10, 11]
    );
    assert_eq!(document.comments()[0].text(), "existing comment");
    assert_eq!(document.content_controls()[0].tag(), Some("customer"));
    assert_eq!(document.content_controls()[0].text(), "old value");
    assert_eq!(document.bookmarks()[0].name(), Some("existing"));

    assert_eq!(document.accept_revision_id(10).unwrap(), 1);
    let added_comment = document
        .add_comment(
            RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 0,
                    run_index: 1,
                },
            },
            "Cy",
            Some("CY"),
            "added comment",
        )
        .unwrap();
    assert_eq!(
        document
            .set_content_control_value_by_tag("customer", "new value")
            .unwrap(),
        1
    );
    document
        .add_bookmark(
            "added_bookmark",
            RunRange {
                start: RunPosition {
                    body_index: 2,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 2,
                    run_index: 1,
                },
            },
        )
        .unwrap();

    let saved = document.to_bytes().unwrap();
    let saved_package = OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    assert_eq!(
        saved_package.get_part("/custom/producer.bin"),
        Some(PRESERVED_PART)
    );
    assert!(
        saved_package
            .get_part("/word/document.xml")
            .unwrap()
            .windows(OPAQUE_CONTROL.len())
            .any(|window| window == OPAQUE_CONTROL)
    );
    assert!(
        saved_package
            .get_part("/word/comments.xml")
            .unwrap()
            .windows(OPAQUE_COMMENT.len())
            .any(|window| window == OPAQUE_COMMENT)
    );

    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(
        reopened
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>(),
        vec![11]
    );
    assert_eq!(
        reopened
            .comments()
            .iter()
            .find(|comment| comment.id() == added_comment)
            .unwrap()
            .text(),
        "added comment"
    );
    assert_eq!(reopened.content_controls()[0].text(), "new value");
    assert!(
        reopened
            .bookmarks()
            .iter()
            .any(|bookmark| bookmark.name() == Some("added_bookmark"))
    );
}

const PNG_2_BY_3: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x03, 0x08, 0x02, 0x00, 0x00, 0x00, 0x36, 0x88, 0x49,
    0xd6, 0x00, 0x00, 0x00, 0x10, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
    0x44, 0x0c, 0x28, 0x14, 0x00, 0x44, 0xd0, 0x05, 0xfb, 0xa4, 0xcf, 0xde, 0x80, 0x00, 0x00, 0x00,
    0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
];

#[test]
fn rdocx_error_opc_wraps_the_shared_error_type() {
    let shared_error = oxml_opc::OpcError::PartNotFound("missing.xml".to_string());
    let error: rdocx::Error = shared_error.into();

    match error {
        rdocx::Error::Opc(inner) => {
            let _: oxml_opc::OpcError = inner;
        }
        other => panic!("expected shared OPC error, got {other}"),
    }
}

#[test]
fn new_document_uses_the_shared_word_package_setup() {
    let mut document = Document::new();
    let bytes = document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();

    assert_eq!(
        package.main_document_part().as_deref(),
        Some("/word/document.xml")
    );
    assert_eq!(
        package.content_types.content_type_for("/word/document.xml"),
        Some("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
    );
    assert_eq!(
        package.content_types.content_type_for("/word/styles.xml"),
        Some("application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")
    );
    assert!(package.get_part("/word/styles.xml").is_some());
    let styles = package
        .get_part_rels("/word/document.xml")
        .and_then(|rels| rels.get_by_type(rel_types::STYLES))
        .expect("new document must relate its styles part");
    assert_eq!(styles.target, "styles.xml");
}

#[test]
fn create_and_round_trip_simple_document() {
    let mut doc = Document::new();
    doc.add_paragraph("Hello, World!");
    doc.add_paragraph("This is a test document.");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 2);
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Hello, World!");
    assert_eq!(paras[1].text(), "This is a test document.");
}

#[test]
fn create_and_round_trip_formatted_document() {
    let mut doc = Document::new();

    // Title paragraph
    doc.add_paragraph("Document Title")
        .style("Heading1")
        .alignment(Alignment::Center);

    // Normal paragraph with multiple formatted runs
    let mut para = doc.add_paragraph("");
    para.add_run("This is ").font("Arial").size(11.0);
    para.add_run("bold").bold(true).font("Arial").size(11.0);
    para.add_run(" and this is ").font("Arial").size(11.0);
    para.add_run("italic").italic(true).font("Arial").size(11.0);
    para.add_run(".").font("Arial").size(11.0);

    // Justified paragraph with indentation
    doc.add_paragraph("This paragraph has special formatting.")
        .alignment(Alignment::Justify)
        .indent_left(Length::inches(0.5))
        .space_before(Length::pt(12.0))
        .space_after(Length::pt(6.0));

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 3);

    // Check title
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Document Title");
    assert_eq!(paras[0].style_id(), Some("Heading1"));
    assert_eq!(paras[0].alignment(), Some(Alignment::Center));

    // Check formatted runs
    let runs: Vec<_> = paras[1].runs().collect();
    assert_eq!(runs.len(), 5);
    assert!(!runs[0].is_bold());
    assert!(runs[1].is_bold());
    assert!(!runs[2].is_italic());
    assert!(runs[3].is_italic());

    // Check justified paragraph
    assert_eq!(paras[2].alignment(), Some(Alignment::Justify));
}

#[test]
fn round_trip_preserves_styles() {
    let mut doc = Document::new();
    doc.add_paragraph("Normal paragraph");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Should have the default styles
    assert!(doc2.style("Normal").is_some());
    assert!(doc2.style("Heading1").is_some());

    let normal = doc2.style("Normal").unwrap();
    assert!(normal.is_default());
    assert_eq!(normal.name(), Some("Normal"));
}

#[test]
fn save_and_load_file() {
    let dir = std::env::temp_dir();
    let process_id = std::process::id().to_string();
    let path = dir.join(format!("rdocx_test_output_{process_id}.docx"));
    assert!(
        path.file_name()
            .and_then(|name| name.to_str())
            .is_some_and(|name| name.contains(&process_id)),
        "temporary output path must contain the test process ID"
    );

    // Create and save
    let mut doc = Document::new();
    doc.add_paragraph("Saved to disk");
    doc.save(&path).unwrap();

    // Load back
    let doc2 = Document::open(&path).unwrap();
    assert_eq!(doc2.paragraph_count(), 1);
    assert_eq!(doc2.paragraphs()[0].text(), "Saved to disk");

    // Clean up
    std::fs::remove_file(&path).ok();
}

#[test]
fn section_properties_preserved() {
    let doc = Document::new();
    let sect = doc.section_properties().unwrap();

    // Default US Letter
    assert_eq!(sect.page_width.unwrap().0, 12240); // 8.5"
    assert_eq!(sect.page_height.unwrap().0, 15840); // 11"
    assert_eq!(sect.margin_top.unwrap().0, 1440); // 1"

    // Round-trip
    let mut doc2 = Document::new();
    let bytes = doc2.to_bytes().unwrap();
    let doc3 = Document::from_bytes(&bytes).unwrap();
    let sect3 = doc3.section_properties().unwrap();
    assert_eq!(sect3.page_width.unwrap().0, 12240);
}

#[test]
fn empty_document_round_trip() {
    let mut doc = Document::new();
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraph_count(), 0);
}

#[test]
fn run_color_and_font_round_trip() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("Red text")
        .color("FF0000")
        .font("Times New Roman")
        .size(16.0);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    let runs: Vec<_> = paras[0].runs().collect();
    assert_eq!(runs[0].color(), Some("FF0000"));
    assert_eq!(runs[0].font_name(), Some("Times New Roman"));
    assert_eq!(runs[0].size(), Some(16.0));
}

#[test]
fn facade_table_and_tristate_accessors_are_total() {
    let mut document = Document::new();
    document.add_paragraph("").add_run("value");

    let paragraph = document.paragraph(0).unwrap();
    let run = paragraph.run(0).unwrap();
    assert_eq!(run.bold_value(), None);
    assert_eq!(run.italic_value(), None);
    assert_eq!(run.strike_value(), None);

    {
        let mut paragraph = document.paragraph_mut(0).unwrap();
        let mut run = paragraph.run_mut(0).unwrap();
        run.set_bold_value(Some(false));
        run.set_italic_value(Some(true));
        run.set_strike_value(Some(false));
        assert!(run.set_underline_code_value(Some(9)));
        assert!(!run.set_underline_code_value(Some(5)));
    }
    assert_eq!(
        document
            .paragraph(0)
            .unwrap()
            .run(0)
            .unwrap()
            .underline_code_value(),
        Some(9)
    );
    {
        let mut paragraph = document.paragraph_mut(0).unwrap();
        let mut run = paragraph.run_mut(0).unwrap();
        assert!(run.set_underline_code_value(Some(10)));
    }

    assert_eq!(document.table_count(), 0);
    assert!(document.table(0).is_none());
    assert!(document.table_mut(0).is_none());

    document.add_table(1, 1);
    assert_eq!(document.table_count(), 1);
    assert_eq!(document.table(0).unwrap().row_count(), 1);
    assert_eq!(
        document
            .table(0)
            .unwrap()
            .cell(0, 0)
            .unwrap()
            .paragraph_count(),
        1
    );
    assert!(
        document
            .table_mut(0)
            .unwrap()
            .cell(0, 0)
            .unwrap()
            .paragraph_mut(0)
            .is_some()
    );

    let bytes = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    let paragraph = reopened.paragraph(0).unwrap();
    let run = paragraph.run(0).unwrap();
    assert_eq!(run.bold_value(), Some(false));
    assert_eq!(run.italic_value(), Some(true));
    assert_eq!(run.strike_value(), Some(false));
    assert_eq!(run.underline_code_value(), Some(10));
}

#[test]
fn established_underline_enum_and_first_line_indent_remain_compatible() {
    fn established_underline_code(style: UnderlineStyle) -> i32 {
        match style {
            UnderlineStyle::None => 0,
            UnderlineStyle::Single => 1,
            UnderlineStyle::Double => 3,
            UnderlineStyle::Thick => 6,
            UnderlineStyle::Dotted => 4,
            UnderlineStyle::Dash => 7,
            UnderlineStyle::Wave => 11,
            UnderlineStyle::Words => 2,
        }
    }

    assert_eq!(established_underline_code(UnderlineStyle::Wave), 11);

    let mut document = Document::new();
    document
        .add_paragraph("legacy")
        .set_first_line_indent(Length::inches(-0.25));
    let xml = String::from_utf8(document_xml(&mut document)).unwrap();
    assert!(xml.contains("w:firstLine=\"-360\""));
    assert!(!xml.contains("w:hanging=\"360\""));
}

// ---- Phase 2 Integration Tests ----

#[test]
fn paragraph_borders_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Bordered paragraph")
        .border_all(BorderStyle::Single, 4, "000000");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    assert!(paras[0].has_borders());
}

#[test]
fn paragraph_tab_stops_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Tab text")
        .add_tab_stop(TabAlignment::Right, Length::inches(6.0))
        .add_tab_stop_with_leader(TabAlignment::Right, Length::inches(6.5), TabLeader::Dot);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    assert_eq!(paras[0].tab_stop_count(), 2);
}

#[test]
fn paragraph_shading_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Highlighted paragraph").shading("FFFF00");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    assert_eq!(paras[0].shading_fill(), Some("FFFF00"));
}

#[test]
fn paragraph_spacing_and_indent_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Indented text")
        .indent_left(Length::inches(1.0))
        .indent_right(Length::inches(0.5))
        .first_line_indent(Length::inches(0.25))
        .space_before(Length::pt(12.0))
        .space_after(Length::pt(6.0))
        .line_spacing_multiple(1.5);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Verify it round-trips without error
    assert_eq!(doc2.paragraph_count(), 1);
    assert_eq!(doc2.paragraphs()[0].text(), "Indented text");
}

#[test]
fn paragraph_pagination_controls_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Keep with next")
        .keep_with_next(true)
        .keep_together(true)
        .widow_control(true);
    doc.add_paragraph("Page break before")
        .page_break_before(true);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 2);
}

#[test]
fn run_underline_styles_round_trip() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("Simple underline").underline(true);
    para.add_run("Wave underline")
        .underline_style(UnderlineStyle::Wave);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Simple underlineWave underline");
}

#[test]
fn run_advanced_formatting_round_trip() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("Strike").strike(true);
    para.add_run("DStrike").double_strike(true);
    para.add_run("CAPS").all_caps(true);
    para.add_run("SmallCaps").small_caps(true);
    para.add_run("Super").superscript();
    para.add_run("Sub").subscript();
    para.add_run("Hidden").hidden(true);
    para.add_run("Spaced").character_spacing(Length::pt(2.0));
    para.add_run("Wide").width_scale(150);
    para.add_run("Raised").position(6);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    let runs: Vec<_> = paras[0].runs().collect();
    assert_eq!(runs.len(), 10);
    assert!(runs[0].is_strike());
    assert_eq!(runs[4].vert_align(), Some("superscript"));
    assert_eq!(runs[5].vert_align(), Some("subscript"));
}

#[test]
fn run_style_assignment_round_trip() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("Styled run").style("Heading1Char");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    let runs: Vec<_> = paras[0].runs().collect();
    assert_eq!(runs[0].style_id(), Some("Heading1Char"));
}

#[test]
fn custom_style_round_trip() {
    let mut doc = Document::new();

    doc.add_style(
        StyleBuilder::paragraph("CustomHeading", "Custom Heading")
            .based_on("Heading1")
            .next_style("Normal"),
    )
    .unwrap();

    doc.add_style(StyleBuilder::character("Emphasis", "Emphasis Style"))
        .unwrap();

    doc.add_paragraph("Custom styled").style("CustomHeading");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let s = doc2.style("CustomHeading").unwrap();
    assert_eq!(s.name(), Some("Custom Heading"));
    assert_eq!(s.based_on(), Some("Heading1"));

    assert!(doc2.style("Emphasis").is_some());

    let paras = doc2.paragraphs();
    assert_eq!(paras[0].style_id(), Some("CustomHeading"));
}

#[test]
fn source_built_style_graph_matches_pinned_word_effective_formatting() {
    const WORD_ORACLE: &str = "Microsoft Word 16.104 build 16.104.25121423";
    const LIBREOFFICE_ORACLE: &str =
        "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb";
    const MAX_DIFFERING_RENDER_BYTES: usize = 0;
    const WORD_STYLES_ORACLE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="character" w:styleId="CorpusBodyChar"><w:name w:val="Corpus Body Char"/><w:link w:val="CorpusBody"/><w:rPr><w:b/><w:i/><w:color w:val="2E5A88"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="CorpusBody" w:default="1"><w:name w:val="Corpus Body"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:link w:val="CorpusBodyChar"/><w:autoRedefine w:val="0"/><w:hidden w:val="0"/><w:uiPriority w:val="17"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:locked w:val="0"/><w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="CorpusTable" w:default="1"><w:name w:val="Corpus Table"/><w:tblPr><w:shd w:val="clear" w:fill="F2F2F2"/></w:tblPr><w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="D9EAF7"/></w:tcPr></w:tblStylePr></w:style>
</w:styles>"#;
    assert_eq!(WORD_ORACLE, MHTML_ORACLE_VERSION);
    assert_eq!(LIBREOFFICE_ORACLE, ODT_ORACLE_VERSION);

    let mut authored = corpus_style_document();
    let authored_paragraph = authored.resolve_paragraph_properties(None);
    let authored_run = authored.resolve_run_properties(Some("CorpusBody"), Some("CorpusBody"));
    assert_eq!(authored_paragraph.space_after, Some(rdocx::Twips(0)));
    assert_eq!(authored_run.bold, Some(true));
    assert_eq!(authored_run.italic, Some(true));
    assert_eq!(authored_run.color.as_deref(), Some("2E5A88"));

    let mut package =
        OpcPackage::from_reader(std::io::Cursor::new(authored.to_bytes().unwrap())).unwrap();
    package.set_part("/word/styles.xml", WORD_STYLES_ORACLE.as_bytes().to_vec());
    let mut oracle_bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut oracle_bytes).unwrap();
    let oracle = Document::from_bytes(oracle_bytes.get_ref()).unwrap();
    oracle.validate_style_graph().unwrap();
    assert_eq!(
        oracle.resolve_paragraph_properties(None),
        authored_paragraph
    );
    assert_eq!(
        oracle.resolve_run_properties(Some("CorpusBody"), Some("CorpusBody")),
        authored_run
    );
    let authored_render = authored
        .render_page_to_png_deterministic(0, 150.0)
        .unwrap()
        .unwrap();
    let oracle_render = oracle
        .render_page_to_png_deterministic(0, 150.0)
        .unwrap()
        .unwrap();
    let differing = authored_render
        .iter()
        .zip(&oracle_render)
        .filter(|(left, right)| left != right)
        .count()
        + authored_render.len().abs_diff(oracle_render.len());
    assert_eq!(differing, MAX_DIFFERING_RENDER_BYTES);

    if std::env::var_os("RDOCX_RUN_PINNED_STYLE_ORACLE").is_some() {
        let version = std::process::Command::new("soffice")
            .arg("--version")
            .output()
            .expect("pinned LibreOffice is installed");
        assert!(version.status.success());
        assert_eq!(
            String::from_utf8_lossy(&version.stdout).trim(),
            LIBREOFFICE_ORACLE
        );

        let root = std::env::temp_dir().join(format!(
            "rdocx-style-oracle-{}-{}",
            std::process::id(),
            std::thread::current().name().unwrap_or("test")
        ));
        let _ = std::fs::remove_dir_all(&root);
        let output = root.join("output");
        let profile = root.join("profile");
        std::fs::create_dir_all(&output).unwrap();
        std::fs::create_dir_all(&profile).unwrap();
        let source = root.join("source.docx");
        std::fs::write(&source, authored.to_bytes().unwrap()).unwrap();
        let status = std::process::Command::new("soffice")
            .arg("--headless")
            .arg(format!(
                "-env:UserInstallation=file://{}",
                profile.display()
            ))
            .arg("--convert-to")
            .arg("docx")
            .arg("--outdir")
            .arg(&output)
            .arg(&source)
            .status()
            .expect("pinned LibreOffice style normalization starts");
        assert!(status.success());
        let normalized = Document::open(output.join("source.docx")).unwrap();
        normalized.validate_style_graph().unwrap();
        let normalized_paragraph = normalized.resolve_paragraph_properties(None);
        assert_eq!(
            normalized_paragraph.space_after,
            authored_paragraph.space_after
        );
        let normalized_run =
            normalized.resolve_run_properties(Some("CorpusBody"), Some("CorpusBody"));
        assert_eq!(normalized_run.bold, authored_run.bold);
        assert_eq!(normalized_run.italic, authored_run.italic);
        assert_eq!(normalized_run.color, authored_run.color);
        assert_eq!(
            normalized.style("CorpusBody").unwrap().linked_style(),
            Some("CorpusBodyChar")
        );
        assert_eq!(
            normalized.style("CorpusBodyChar").unwrap().linked_style(),
            Some("CorpusBody")
        );
        let _ = std::fs::remove_dir_all(&root);
    }
}

#[test]
fn invalid_style_graph_never_publishes_a_partial_mutation() {
    let mut document = Document::new();
    let before = document.to_bytes().unwrap();

    let error = document
        .add_style(StyleBuilder::paragraph("Broken", "Broken").based_on("MissingStyle"))
        .unwrap_err();

    assert!(error.to_string().contains("MissingStyle"));
    assert!(document.style("Broken").is_none());
    assert_eq!(document.to_bytes().unwrap(), before);

    document
        .add_style(StyleBuilder::paragraph("CycleA", "Cycle A").based_on("Normal"))
        .unwrap();
    document
        .add_style(StyleBuilder::paragraph("CycleB", "Cycle B").based_on("CycleA"))
        .unwrap();
    let before_cycle = document.to_bytes().unwrap();
    assert!(
        document
            .set_style(StyleBuilder::paragraph("CycleA", "Cycle A").based_on("CycleB"))
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), before_cycle);

    document
        .add_style(StyleBuilder::character("WrongLink", "Wrong Link"))
        .unwrap();
    let before_link = document.to_bytes().unwrap();
    assert!(
        document
            .add_style(
                StyleBuilder::character("WrongTarget", "Wrong Target").linked_style("WrongLink")
            )
            .is_err()
    );
    assert!(document.style("WrongTarget").is_none());
    assert_eq!(document.to_bytes().unwrap(), before_link);

    document
        .add_style(StyleBuilder::table("DuplicateRegion", "Duplicate Region"))
        .unwrap();
    let before_regions = document.to_bytes().unwrap();
    assert!(
        document
            .set_style(
                StyleBuilder::table("DuplicateRegion", "Duplicate Region")
                    .conditional_table_style(
                        TableStyleRegion::FirstRow,
                        None,
                        None,
                        None,
                        None,
                        None
                    )
                    .conditional_table_style(
                        TableStyleRegion::FirstRow,
                        None,
                        None,
                        None,
                        None,
                        None
                    ),
            )
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), before_regions);

    let mut package =
        OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
    let styles = String::from_utf8(package.get_part("/word/styles.xml").unwrap().to_vec()).unwrap();
    let duplicate_defaults = styles.replace(
        r#"<w:style w:type="paragraph" w:styleId="Heading1">"#,
        r#"<w:style w:type="paragraph" w:styleId="Heading1" w:default="1">"#,
    );
    package.set_part("/word/styles.xml", duplicate_defaults.into_bytes());
    let mut invalid_bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut invalid_bytes).unwrap();
    let mut invalid = Document::from_bytes(invalid_bytes.get_ref()).unwrap();
    assert!(invalid.validate_style_graph().is_err());
    let invalid_before = invalid.to_bytes().unwrap();
    assert!(
        invalid
            .add_style(StyleBuilder::paragraph("Rejected", "Rejected"))
            .is_err()
    );
    assert_eq!(invalid.to_bytes().unwrap(), invalid_before);
}

#[test]
fn authored_style_graph_survives_save_and_reopen() {
    let mut document = corpus_style_document();
    let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
    reopened.validate_style_graph().unwrap();
    let paragraph = reopened.style("CorpusBody").unwrap();
    assert!(paragraph.is_default());
    assert_eq!(paragraph.based_on(), Some("Normal"));
    assert_eq!(paragraph.next_style(), Some("Normal"));
    assert_eq!(paragraph.linked_style(), Some("CorpusBodyChar"));
    assert_eq!(paragraph.priority(), Some(17));
    assert_eq!(paragraph.auto_redefine(), Some(false));
    assert_eq!(paragraph.hidden(), Some(false));
    assert_eq!(paragraph.semi_hidden(), Some(true));
    assert_eq!(paragraph.unhide_when_used(), Some(true));
    assert_eq!(paragraph.quick_format(), Some(true));
    assert_eq!(paragraph.locked(), Some(false));
    assert_eq!(
        paragraph
            .paragraph_properties()
            .and_then(|properties| properties.space_after),
        Some(rdocx::Twips(0))
    );
    assert_eq!(
        reopened.style("CorpusBodyChar").unwrap().linked_style(),
        Some("CorpusBody")
    );
    let table = reopened.style("CorpusTable").unwrap();
    assert!(table.is_default());
    assert_eq!(table.conditional_table_styles().len(), 1);
    assert_eq!(
        table.conditional_table_styles()[0].region(),
        Some(TableStyleRegion::Band1Horz)
    );
    assert_eq!(
        table.conditional_table_styles()[0]
            .cell_properties()
            .and_then(|properties| properties.shading.as_ref())
            .and_then(|shading| shading.fill.as_deref()),
        Some("D9EAF7")
    );

    let mut updated = reopened;
    updated
        .set_style(
            StyleBuilder::paragraph("CorpusBody", "Corpus Body")
                .clear_based_on()
                .clear_next_style()
                .clear_linked_style()
                .clear_priority()
                .clear_auto_redefine()
                .clear_hidden()
                .clear_semi_hidden()
                .clear_unhide_when_used()
                .clear_quick_format()
                .clear_locked()
                .clear_paragraph_properties()
                .clear_run_properties(),
        )
        .unwrap();
    let paragraph = updated.style("CorpusBody").unwrap();
    assert_eq!(paragraph.based_on(), None);
    assert_eq!(paragraph.next_style(), None);
    assert_eq!(paragraph.linked_style(), None);
    assert_eq!(paragraph.priority(), None);
    assert_eq!(paragraph.paragraph_properties(), None);
    assert_eq!(paragraph.run_properties(), None);
    assert_eq!(
        updated.style("CorpusBodyChar").unwrap().linked_style(),
        None
    );
    updated
        .set_style(
            StyleBuilder::table("CorpusTable", "Corpus Table")
                .clear_table_properties()
                .clear_conditional_table_styles(),
        )
        .unwrap();
    assert_eq!(
        updated.style("CorpusTable").unwrap().table_properties(),
        None
    );
    assert!(
        updated
            .style("CorpusTable")
            .unwrap()
            .conditional_table_styles()
            .is_empty()
    );
}

#[test]
fn conditional_table_style_updates_preserve_siblings_and_existing_groups() {
    let mut document = corpus_style_document();
    document
        .set_style(
            StyleBuilder::table("CorpusTable", "Corpus Table")
                .conditional_table_style(
                    TableStyleRegion::Band1Horz,
                    Some(CT_PPr {
                        space_after: Some(rdocx::Twips(60)),
                        ..CT_PPr::default()
                    }),
                    None,
                    None,
                    None,
                    None,
                )
                .conditional_table_style(
                    TableStyleRegion::FirstRow,
                    None,
                    None,
                    None,
                    None,
                    Some(CT_TcPr {
                        shading: Some(CT_Shd {
                            val: "clear".to_owned(),
                            color: None,
                            fill: Some("112233".to_owned()),
                            ..Default::default()
                        }),
                        ..CT_TcPr::default()
                    }),
                ),
        )
        .unwrap();

    let table = document.style("CorpusTable").unwrap();
    assert_eq!(table.conditional_table_styles().len(), 2);
    let regions = table.conditional_table_styles();
    let band = regions
        .iter()
        .find(|region| region.region() == Some(TableStyleRegion::Band1Horz))
        .unwrap();
    assert_eq!(
        band.paragraph_properties()
            .and_then(|properties| properties.space_after),
        Some(rdocx::Twips(60))
    );
    assert_eq!(
        band.cell_properties()
            .and_then(|properties| properties.shading.as_ref())
            .and_then(|shading| shading.fill.as_deref()),
        Some("D9EAF7")
    );

    document
        .set_style(
            StyleBuilder::table("CorpusTable", "Corpus Table")
                .clear_conditional_table_styles()
                .conditional_table_style(TableStyleRegion::LastRow, None, None, None, None, None),
        )
        .unwrap();
    let table = document.style("CorpusTable").unwrap();
    let regions = table.conditional_table_styles();
    assert_eq!(regions.len(), 1);
    assert_eq!(regions[0].region(), Some(TableStyleRegion::LastRow));
}

#[test]
fn table_style_updates_merge_nested_borders_and_margins() {
    let mut document = Document::new();
    document
        .add_style(StyleBuilder::table("LayeredTable", "Layered Table"))
        .unwrap();
    let mut package =
        OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
    let styles = String::from_utf8(package.get_part("/word/styles.xml").unwrap().to_vec()).unwrap();
    let styles = styles.replace(
        r#"<w:name w:val="Layered Table"/>"#,
        r#"<w:name w:val="Layered Table"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8" w:color="111111"/><w:bottom w:val="double" w:sz="12" w:color="222222"/><x:diagonal xmlns:x="urn:producer" x:keep="exact"/></w:tblBorders><w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/></w:tblCellMar></w:tblPr>"#,
    );
    package.set_part("/word/styles.xml", styles.into_bytes());
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(bytes.get_ref()).unwrap();

    document
        .set_style(
            StyleBuilder::table("LayeredTable", "Layered Table").table_properties(CT_TblPr {
                borders: Some(CT_TblBorders {
                    top: Some(CT_BorderEdge {
                        val: ST_Border::Thick,
                        sz: Some(16),
                        space: None,
                        color: Some("AABBCC".to_owned()),
                        extra_attributes: Vec::new(),
                    }),
                    ..CT_TblBorders::default()
                }),
                cell_margin: Some(CT_TblCellMar {
                    top: Some(rdocx::Twips(240)),
                    ..CT_TblCellMar::default()
                }),
                ..CT_TblPr::default()
            }),
        )
        .unwrap();

    let package =
        OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
    let output = std::str::from_utf8(package.get_part("/word/styles.xml").unwrap()).unwrap();
    assert!(output.contains(r#"<w:top w:val="thick" w:sz="16" w:color="AABBCC"/>"#));
    assert!(output.contains(r#"<w:bottom w:val="double" w:sz="12" w:color="222222"/>"#));
    assert_eq!(output.matches(r#"x:diagonal"#).count(), 1);
    assert!(output.contains(r#"<w:top w:w="240" w:type="dxa"/>"#));
    assert!(output.contains(r#"<w:left w:w="140" w:type="dxa"/>"#));
}

#[test]
fn style_removal_rejects_live_references_and_preserves_unknown_xml() {
    let mut document = Document::new();
    document
        .add_style(StyleBuilder::paragraph("Disposable", "Disposable"))
        .unwrap();
    document
        .add_style(StyleBuilder::paragraph("Unused", "Unused"))
        .unwrap();
    document
        .add_style(StyleBuilder::paragraph("Keeper", "Keeper"))
        .unwrap();
    document
        .add_paragraph("live style reference")
        .style("Disposable");

    let mut package =
        OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
    let styles = String::from_utf8(package.get_part("/word/styles.xml").unwrap().to_vec()).unwrap();
    let extension = r#"<x:producer xmlns:x="urn:producer" x:exact="&amp; &quot;kept&quot;"/>"#;
    let paragraph_extension = r#"<x:pKeep xmlns:x="urn:producer" x:exact="paragraph"/>"#;
    let run_extension = r#"<x:rKeep xmlns:x="urn:producer" x:exact="run"/>"#;
    let styles = styles.replace(
        r#"<w:name w:val="Keeper"/>"#,
        &format!(
            r#"<w:name w:val="Keeper"/><w:pPr>{paragraph_extension}<w:spacing w:after="100"/></w:pPr><w:rPr>{run_extension}<w:b/></w:rPr>{extension}"#
        ),
    );
    package.set_part("/word/styles.xml", styles.into_bytes());
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(bytes.get_ref()).unwrap();

    assert!(document.remove_style("Disposable").is_err());
    assert!(document.style("Disposable").is_some());
    document
        .set_style(
            StyleBuilder::paragraph("Keeper", "Updated Keeper")
                .priority(9)
                .paragraph_properties(CT_PPr {
                    space_before: Some(rdocx::Twips(40)),
                    ..CT_PPr::default()
                })
                .run_properties(CT_RPr {
                    color: Some("335577".to_owned()),
                    ..CT_RPr::default()
                }),
        )
        .unwrap();
    assert!(document.remove_style("Unused").unwrap());
    document.paragraph_mut(0).unwrap().set_style("Normal");
    assert!(document.remove_style("Disposable").unwrap());

    let saved =
        OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
    let output = std::str::from_utf8(saved.get_part("/word/styles.xml").unwrap()).unwrap();
    assert_eq!(output.matches(extension).count(), 1);
    assert_eq!(output.matches(paragraph_extension).count(), 1);
    assert_eq!(output.matches(run_extension).count(), 1);
    let keeper = document.style("Keeper").unwrap();
    assert_eq!(
        keeper
            .paragraph_properties()
            .and_then(|properties| properties.space_after),
        Some(rdocx::Twips(100))
    );
    assert_eq!(
        keeper
            .paragraph_properties()
            .and_then(|properties| properties.space_before),
        Some(rdocx::Twips(40))
    );
    assert_eq!(
        keeper
            .run_properties()
            .and_then(|properties| properties.bold),
        Some(true)
    );
    assert_eq!(
        keeper
            .run_properties()
            .and_then(|properties| properties.color.as_deref()),
        Some("335577")
    );
    assert!(document.style("Unused").is_none());
    assert!(document.style("Disposable").is_none());

    document
        .add_style(StyleBuilder::paragraph("NoteStyle", "Note Style"))
        .unwrap();
    document.add_footnote("styled note");
    let mut package =
        OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
    let notes = String::from_utf8(package.get_part("/word/footnotes.xml").unwrap().to_vec())
        .unwrap()
        .replacen(
            "<w:p>",
            r#"<w:p><x:producer xmlns:x="urn:producer"><w:pStyle w:val="NoteStyle"/></x:producer>"#,
            1,
        );
    package.set_part("/word/footnotes.xml", notes.into_bytes());
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut with_note_style = Document::from_bytes(bytes.get_ref()).unwrap();
    assert!(with_note_style.remove_style("NoteStyle").is_err());
    assert!(with_note_style.style("NoteStyle").is_some());
}

fn corpus_style_document() -> Document {
    let mut document = Document::new();
    document
        .add_style(
            StyleBuilder::character("CorpusBodyChar", "Corpus Body Char").run_properties(CT_RPr {
                bold: Some(true),
                italic: Some(true),
                color: Some("2E5A88".to_owned()),
                ..CT_RPr::default()
            }),
        )
        .unwrap();
    document
        .add_style(
            StyleBuilder::paragraph("CorpusBody", "Corpus Body")
                .based_on("Normal")
                .next_style("Normal")
                .linked_style("CorpusBodyChar")
                .priority(17)
                .auto_redefine(false)
                .hidden(false)
                .semi_hidden(true)
                .unhide_when_used(true)
                .quick_format(true)
                .locked(false)
                .paragraph_properties(CT_PPr {
                    space_after: Some(rdocx::Twips(0)),
                    ..CT_PPr::default()
                })
                .run_properties(CT_RPr {
                    sz: Some(rdocx_oxml::HalfPoint(28)),
                    ..CT_RPr::default()
                }),
        )
        .unwrap();
    document
        .add_style(
            StyleBuilder::table("CorpusTable", "Corpus Table")
                .table_properties(CT_TblPr {
                    shading: Some(CT_Shd {
                        val: "clear".to_owned(),
                        color: None,
                        fill: Some("F2F2F2".to_owned()),
                        ..Default::default()
                    }),
                    ..CT_TblPr::default()
                })
                .conditional_table_style(
                    TableStyleRegion::Band1Horz,
                    None,
                    None,
                    None,
                    None,
                    Some(CT_TcPr {
                        shading: Some(CT_Shd {
                            val: "clear".to_owned(),
                            color: None,
                            fill: Some("D9EAF7".to_owned()),
                            ..Default::default()
                        }),
                        ..CT_TcPr::default()
                    }),
                ),
        )
        .unwrap();
    document
        .set_default_style(rdocx::StyleType::Paragraph, "CorpusBody")
        .unwrap();
    document
        .set_default_style(rdocx::StyleType::Table, "CorpusTable")
        .unwrap();
    document
        .add_paragraph("")
        .style("CorpusBody")
        .add_run("Corpus style graph")
        .style("CorpusBody");
    let mut table = document.add_table(2, 1);
    table.row(0).unwrap().cell(0).unwrap().set_text("Band one");
    table.row(1).unwrap().cell(0).unwrap().set_text("Band two");
    document
}

#[test]
fn style_inheritance_resolution() {
    let doc = Document::new();

    // Heading1's rpr should have bold from the style definition
    let rpr = doc.resolve_run_properties(Some("Heading1"), None);
    assert_eq!(rpr.bold, Some(true));
    // Font inherited from docDefaults
    assert_eq!(rpr.font_ascii, Some("Calibri".to_string()));

    // Normal paragraph should get docDefaults spacing
    let ppr = doc.resolve_paragraph_properties(Some("Normal"));
    assert!(ppr.space_after.is_some());
}

#[test]
fn section_landscape_round_trip() {
    let mut doc = Document::new();
    doc.set_landscape();

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let sect = doc2.section_properties().unwrap();
    assert!(sect.page_width.unwrap().0 > sect.page_height.unwrap().0);
}

#[test]
fn section_margins_round_trip() {
    let mut doc = Document::new();
    doc.set_margins(
        Length::inches(0.5),
        Length::inches(0.75),
        Length::inches(0.5),
        Length::inches(0.75),
    );

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let sect = doc2.section_properties().unwrap();
    assert_eq!(sect.margin_top.unwrap().0, 720);
    assert_eq!(sect.margin_right.unwrap().0, 1080);
    assert_eq!(sect.margin_bottom.unwrap().0, 720);
    assert_eq!(sect.margin_left.unwrap().0, 1080);
}

#[test]
fn section_columns_round_trip() {
    let mut doc = Document::new();
    doc.set_columns(3, Length::inches(0.25));

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let sect = doc2.section_properties().unwrap();
    let cols = sect.columns.as_ref().unwrap();
    assert_eq!(cols.num, Some(3));
    assert_eq!(cols.equal_width, Some(true));
}

#[test]
fn section_a4_page_size() {
    let mut doc = Document::new();
    doc.set_page_size(Length::cm(21.0), Length::cm(29.7));

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let sect = doc2.section_properties().unwrap();
    let w = sect.page_width.unwrap().0;
    let h = sect.page_height.unwrap().0;
    // A4 dimensions: 11906tw x 16838tw (allow small rounding)
    assert!((w - 11906).abs() < 5, "Expected ~11906, got {w}");
    assert!((h - 16838).abs() < 5, "Expected ~16838, got {h}");
}

#[test]
fn comprehensive_document_round_trip() {
    // Create a document with many Phase 2 features combined
    let mut doc = Document::new();

    // Custom style
    doc.add_style(StyleBuilder::paragraph("BlockQuote", "Block Quote").based_on("Normal"))
        .unwrap();

    // Page setup
    doc.set_margins(
        Length::inches(1.0),
        Length::inches(1.25),
        Length::inches(1.0),
        Length::inches(1.25),
    );

    // Title
    doc.add_paragraph("My Document")
        .style("Heading1")
        .alignment(Alignment::Center)
        .space_after(Length::pt(24.0));

    // Body paragraph with formatting
    let mut para = doc.add_paragraph("");
    para.add_run("This is ").font("Calibri").size(11.0);
    para.add_run("important")
        .bold(true)
        .color("FF0000")
        .font("Calibri")
        .size(11.0);
    para.add_run(" text with ").font("Calibri").size(11.0);
    para.add_run("underline")
        .underline(true)
        .font("Calibri")
        .size(11.0);
    para.add_run(".").font("Calibri").size(11.0);

    // Block quote with indentation and shading
    doc.add_paragraph("This is a block quote.")
        .style("BlockQuote")
        .indent_left(Length::inches(0.5))
        .indent_right(Length::inches(0.5))
        .shading("F2F2F2")
        .space_before(Length::pt(6.0))
        .space_after(Length::pt(6.0));

    // Bordered paragraph
    doc.add_paragraph("Important note")
        .border_all(BorderStyle::Single, 4, "000000")
        .shading("FFFFCC");

    // Save and reload
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 4);
    assert!(doc2.style("BlockQuote").is_some());

    let paras = doc2.paragraphs();
    assert_eq!(paras[0].style_id(), Some("Heading1"));
    assert_eq!(paras[0].alignment(), Some(Alignment::Center));

    let runs: Vec<_> = paras[1].runs().collect();
    assert_eq!(runs.len(), 5);
    assert!(runs[1].is_bold());
    assert_eq!(runs[1].color(), Some("FF0000"));

    assert_eq!(paras[2].shading_fill(), Some("F2F2F2"));
    assert!(paras[3].has_borders());
    assert_eq!(paras[3].shading_fill(), Some("FFFFCC"));
}

// ---- Phase 3 Integration Tests ----

#[test]
fn table_basic_creation_round_trip() {
    let mut doc = Document::new();
    let mut table = doc.add_table(3, 4);
    table.cell(0, 0).unwrap().set_text("A1");
    table.cell(0, 1).unwrap().set_text("B1");
    table.cell(1, 0).unwrap().set_text("A2");
    table.cell(2, 3).unwrap().set_text("D3");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.table_count(), 1);
    let tables = doc2.tables();
    let t = &tables[0];
    assert_eq!(t.row_count(), 3);
    assert_eq!(t.column_count(), 4);
    assert_eq!(t.cell(0, 0).unwrap().text(), "A1");
    assert_eq!(t.cell(0, 1).unwrap().text(), "B1");
    assert_eq!(t.cell(1, 0).unwrap().text(), "A2");
    assert_eq!(t.cell(2, 3).unwrap().text(), "D3");
}

#[test]
fn table_column_width_updates_grid_and_cells() {
    let mut doc = Document::new();
    let mut table = doc.add_table(2, 2);

    assert!(table.set_column_width(0, Length::twips(2_000)));
    assert!(table.set_column_width(1, Length::twips(3_000)));
    assert!(!table.set_column_width(2, Length::twips(1_000)));

    let xml = String::from_utf8(document_xml(&mut doc)).unwrap();
    assert_eq!(
        xml.matches("<w:tblW w:w=\"5000\" w:type=\"dxa\"").count(),
        1,
        "the table width must equal the synchronized grid width"
    );
    assert_eq!(xml.matches("<w:gridCol w:w=\"2000\"").count(), 1);
    assert_eq!(xml.matches("<w:gridCol w:w=\"3000\"").count(), 1);
    assert_eq!(xml.matches("<w:tcW w:w=\"2000\" w:type=\"dxa\"").count(), 2);
    assert_eq!(xml.matches("<w:tcW w:w=\"3000\" w:type=\"dxa\"").count(), 2);
}

#[test]
fn table_with_formatting_round_trip() {
    let mut doc = Document::new();
    doc.add_table(2, 2)
        .borders(BorderStyle::Single, 4, "000000")
        .alignment(Alignment::Center)
        .layout_fixed();

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.table_count(), 1);
    let tables = doc2.tables();
    assert_eq!(tables[0].row_count(), 2);
    assert_eq!(tables[0].column_count(), 2);
}

#[test]
fn table_cell_shading_and_alignment() {
    let mut doc = Document::new();
    let mut table = doc.add_table(2, 2);

    table.cell(0, 0).unwrap().set_text("Header");
    table
        .cell(0, 0)
        .unwrap()
        .shading("4472C4")
        .vertical_alignment(VerticalAlignment::Center);

    table
        .cell(1, 0)
        .unwrap()
        .shading("D9E2F3")
        .vertical_alignment(VerticalAlignment::Bottom);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let tables = doc2.tables();
    let cell_00 = tables[0].cell(0, 0).unwrap();
    assert_eq!(cell_00.shading_fill(), Some("4472C4"));
    assert_eq!(
        cell_00.vertical_alignment(),
        Some(VerticalAlignment::Center)
    );

    let cell_10 = tables[0].cell(1, 0).unwrap();
    assert_eq!(cell_10.shading_fill(), Some("D9E2F3"));
    assert_eq!(
        cell_10.vertical_alignment(),
        Some(VerticalAlignment::Bottom)
    );
}

#[test]
fn table_header_row_round_trip() {
    let mut doc = Document::new();
    let mut table = doc.add_table(3, 2);
    table.row(0).unwrap().header();
    table.cell(0, 0).unwrap().set_text("Col A");
    table.cell(0, 1).unwrap().set_text("Col B");
    table.cell(1, 0).unwrap().set_text("Data 1");
    table.cell(2, 0).unwrap().set_text("Data 2");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let tables = doc2.tables();
    assert!(tables[0].row(0).unwrap().is_header());
    assert!(!tables[0].row(1).unwrap().is_header());
}

#[test]
fn table_rows_clone_remove_and_clear_through_native_and_python() {
    let mut doc = Document::new();
    let mut table = doc.add_table(2, 2);
    table.cell(0, 0).unwrap().set_text("header");
    table.cell(1, 0).unwrap().set_text("entry");
    {
        let mut row = table.row(1).unwrap();
        row.set_height_checked(RowHeight::Exact(Length::twips(480)))
            .unwrap();
        row.set_header_value(Some(true));
        row.set_cant_split_value(Some(true));
    }
    {
        let mut cell = table.cell(1, 1).unwrap();
        let mut nested = cell.add_table_checked(1, 1).unwrap();
        nested.cell(0, 0).unwrap().set_text("nested");
    }

    assert_eq!(doc.clone_table_row(0, 1, 2).unwrap(), 2);
    {
        let mut table = doc.table_mut(0).unwrap();
        let mut copied = table.row(2).unwrap();
        copied.set_header_value(None);
        copied.set_cant_split_value(None);
        copied.cell(0).unwrap().set_text("copied entry");
    }
    assert!(doc.remove_table_row(0, 1).unwrap());

    let reopened = Document::from_bytes(&doc.to_bytes().unwrap()).unwrap();
    let table = reopened.table(0).unwrap();
    assert_eq!(table.row_count(), 2);
    assert_eq!(table.cell(1, 0).unwrap().text(), "copied entry");
    assert_eq!(
        table.row(1).unwrap().height(),
        Some(RowHeight::Exact(Length::twips(480)))
    );
    assert_eq!(table.row(1).unwrap().header_value(), None);
    assert_eq!(table.row(1).unwrap().cant_split_value(), None);
    assert!(table.cell(1, 1).unwrap().items().any(|item| matches!(
        item,
        rdocx::table::CellItemRef::Table(nested)
            if nested.cell(0, 0).unwrap().text() == "nested"
    )));
    assert_eq!(
        reopened.render_page_to_png_deterministic(0, 72.0).unwrap(),
        reopened.render_page_to_png_deterministic(0, 72.0).unwrap()
    );
}

#[test]
fn cloned_table_row_freshens_identities_and_normalizes_root_namespaces() {
    let mut seed = Document::new();
    seed.add_paragraph("entry");
    let entry = RunRange {
        start: RunPosition {
            body_index: 0,
            run_index: 0,
        },
        end: RunPosition {
            body_index: 0,
            run_index: 1,
        },
    };
    seed.add_bookmark("entry", entry).unwrap();
    seed.add_comment(entry, "Ada", None, "Check this entry")
        .unwrap();
    seed.add_picture(
        PNG_2_BY_3,
        "entry.png",
        Length::twips(240),
        Length::twips(360),
    );
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap()))
        .expect("open seed package");
    let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap())
        .unwrap()
        .to_owned();
    let body_start = xml.find("<w:body>").unwrap() + "<w:body>".len();
    let section_start = body_start + xml[body_start..].find("<w:sectPr").unwrap();
    let table = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:sdt><w:sdtPr><w:id w:val="42"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt><x:keep/>{}</w:tc></w:tr></w:tbl>"#,
        &xml[body_start..section_start]
    );
    let xml = format!(
        "{}{table}{}",
        &xml[..body_start],
        &xml[section_start..]
    )
    .replacen(
        "<w:document ",
        r#"<w:document xmlns="http://schemas.microsoft.com/office/tasks/2019/documenttasks" xmlns:x="urn:producer" "#,
        1,
    );
    package.set_part("/word/document.xml", xml.into_bytes());
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(&bytes.into_inner()).unwrap();

    assert_eq!(document.clone_table_row(0, 0, 1).unwrap(), 1);

    let saved = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(reopened.table(0).unwrap().row_count(), 2);
    assert_eq!(reopened.comments().len(), 1);
    let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(!xml.contains("xmlns:xmlns"), "{xml}");
    assert_eq!(xml.matches("commentRangeStart").count(), 1, "{xml}");
    assert_eq!(xml.matches("commentReference").count(), 1, "{xml}");
    assert_eq!(xml.matches("<x:keep").count(), 2, "{xml}");

    let bookmark_starts = xml
        .match_indices("<w:bookmarkStart ")
        .map(|(start, _)| &xml[start..start + xml[start..].find('>').unwrap()])
        .collect::<Vec<_>>();
    assert_eq!(bookmark_starts.len(), 2, "{xml}");
    assert_ne!(bookmark_starts[0], bookmark_starts[1], "{xml}");
    let control_ids = xml
        .match_indices("<w:id w:val=\"")
        .map(|(start, _)| &xml[start..start + xml[start..].find('>').unwrap()])
        .collect::<Vec<_>>();
    assert_eq!(control_ids.len(), 2, "{xml}");
    assert_ne!(control_ids[0], control_ids[1], "{xml}");
    let drawing_ids = xml
        .match_indices("<wp:docPr ")
        .map(|(start, _)| &xml[start..start + xml[start..].find('>').unwrap()])
        .collect::<Vec<_>>();
    assert_eq!(drawing_ids.len(), 2, "{xml}");
    assert_ne!(drawing_ids[0], drawing_ids[1], "{xml}");
    let image_relationships = xml
        .match_indices("r:embed=\"")
        .map(|(start, _)| {
            let value = &xml[start + "r:embed=\"".len()..];
            &value[..value.find('"').unwrap()]
        })
        .collect::<Vec<_>>();
    assert_eq!(image_relationships.len(), 2, "{xml}");
    assert_eq!(image_relationships[0], image_relationships[1], "{xml}");
}

#[test]
fn removing_table_rows_restarts_vertical_merges_and_keeps_one_row() {
    let mut document = Document::new();
    let mut table = document.add_table(3, 1);
    for (row, text) in ["first", "second", "third"].into_iter().enumerate() {
        table.cell(row, 0).unwrap().set_text(text);
    }
    table.cell(0, 0).unwrap().set_v_merge_restart();
    table.cell(1, 0).unwrap().set_v_merge_continue();
    table.cell(2, 0).unwrap().set_v_merge_continue();

    assert!(document.remove_table_row(0, 0).unwrap());
    let saved = document.to_bytes().unwrap();
    let mut reopened = Document::from_bytes(&saved).unwrap();
    let table = reopened.table(0).unwrap();
    assert_eq!(table.row_count(), 2);
    assert_eq!(table.cell(0, 0).unwrap().text(), "second");
    assert_eq!(
        table.cell(0, 0).unwrap().v_merge(),
        Some(&rdocx_oxml::table::VMerge::Restart)
    );
    assert_eq!(
        table.cell(1, 0).unwrap().v_merge(),
        Some(&rdocx_oxml::table::VMerge::Continue)
    );

    assert!(reopened.remove_table_row(0, 1).unwrap());
    let before = reopened.to_bytes().unwrap();
    let error = reopened.remove_table_row(0, 0).unwrap_err();
    assert!(error.to_string().contains("at least one row"), "{error}");
    assert_eq!(reopened.to_bytes().unwrap(), before);
}

#[test]
fn table_row_mutations_preserve_raw_boundaries_and_fail_atomically() {
    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let seed = Document::new().to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(seed)).unwrap();
    package.set_part(
        "/word/document.xml",
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:x="urn:producer"><w:body><w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><x:a/><w:tr><w:tc><w:p><w:r><w:t>first</w:t></w:r></w:p></w:tc></w:tr><x:b/><w:sdt><w:sdtPr><w:id w:val="11"/></w:sdtPr><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>controlled</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt><x:c/><w:tr><w:tc><w:p><w:r><w:t>second</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#
        )
        .into_bytes(),
    );
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(&bytes.into_inner()).unwrap();

    assert_eq!(document.clone_table_row(0, 0, 1).unwrap(), 1);
    assert!(document.remove_table_row(0, 1).unwrap());
    let saved = document.to_bytes().unwrap();
    let xml = String::from_utf8(
        OpcPackage::from_reader(std::io::Cursor::new(&saved))
            .unwrap()
            .get_part("/word/document.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    let positions =
        ["<x:a", "<x:b", "<w:sdt>", "<x:c", ">second<"].map(|needle| xml.find(needle).unwrap());
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]), "{xml}");
    let before = document.to_bytes().unwrap();
    assert!(document.clone_table_row(0, 5, 0).is_err());
    assert!(document.clone_table_row(0, 0, 5).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);

    let mut invalid_package = OpcPackage::from_reader(std::io::Cursor::new(before)).unwrap();
    let invalid = xml.replacen(
        "<w:tc>",
        r#"<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>"#,
        1,
    );
    invalid_package.set_part("/word/document.xml", invalid.into_bytes());
    let mut invalid_bytes = std::io::Cursor::new(Vec::new());
    invalid_package.write_to(&mut invalid_bytes).unwrap();
    let mut invalid = Document::from_bytes(&invalid_bytes.into_inner()).unwrap();
    let before = invalid.to_bytes().unwrap();
    assert!(invalid.clone_table_row(0, 0, 1).is_err());
    assert_eq!(invalid.to_bytes().unwrap(), before);
}

#[test]
fn table_cell_grid_span_round_trip() {
    let mut doc = Document::new();
    let mut table = doc.add_table(2, 3);
    // First row: cell 0 spans 2 columns
    table.cell(0, 0).unwrap().set_text("Merged");
    table.cell(0, 0).unwrap().grid_span(2);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let tables = doc2.tables();
    assert_eq!(tables[0].cell(0, 0).unwrap().grid_span(), Some(2));
}

#[test]
fn table_mixed_with_paragraphs() {
    let mut doc = Document::new();
    doc.add_paragraph("Before the table");
    let mut table = doc.add_table(2, 2);
    table.cell(0, 0).unwrap().set_text("Cell");
    doc.add_paragraph("After the table");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 2);
    assert_eq!(doc2.table_count(), 1);
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Before the table");
    assert_eq!(paras[1].text(), "After the table");
    let tables = doc2.tables();
    assert_eq!(tables[0].cell(0, 0).unwrap().text(), "Cell");
}

#[test]
fn table_cell_multiple_paragraphs() {
    let mut doc = Document::new();
    let mut table = doc.add_table(1, 1);
    let mut cell = table.cell(0, 0).unwrap();
    cell.set_text("First line");
    let mut para = cell.add_paragraph("");
    para.add_run("Second line").bold(true);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let tables = doc2.tables();
    let cell = tables[0].cell(0, 0).unwrap();
    let paras: Vec<_> = cell.paragraphs().collect();
    assert_eq!(paras.len(), 2);
    assert_eq!(paras[0].text(), "First line");
    assert_eq!(paras[1].text(), "Second line");
}

#[test]
fn inline_image_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Before image");

    // Create a minimal 1x1 white PNG
    let png_data: Vec<u8> = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, // 8-bit RGB
        0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
        0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
        0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND chunk
        0xAE, 0x42, 0x60, 0x82,
    ];

    doc.add_picture(
        &png_data,
        "test.png",
        Length::inches(2.0),
        Length::inches(1.5),
    );
    doc.add_paragraph("After image");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Image paragraph is counted as a paragraph too
    assert_eq!(doc2.paragraph_count(), 3);
    assert_eq!(doc2.paragraphs()[0].text(), "Before image");
    assert_eq!(doc2.paragraphs()[2].text(), "After image");
}

#[test]
fn add_picture_auto_uses_native_size_at_72_dpi() {
    let mut document = Document::new();
    document
        .add_picture_auto(PNG_2_BY_3, "two-by-three.png")
        .expect("valid image dimensions");

    let bytes = document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(&bytes)).unwrap();
    let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(xml.contains(r#"<wp:extent cx="25400" cy="38100"/>"#));

    let mut reopened = Document::from_bytes(&bytes).unwrap();
    let round_trip_xml = String::from_utf8(document_xml(&mut reopened)).unwrap();
    assert!(round_trip_xml.contains(r#"<wp:extent cx="25400" cy="38100"/>"#));
}

#[test]
fn add_picture_auto_rejects_unavailable_dimensions_without_mutation() {
    let mut document = Document::new();
    document.add_paragraph("unchanged");
    let before = document.to_bytes().unwrap();

    let error = match document.add_picture_auto(b"not an image", "broken.bin") {
        Ok(_) => panic!("malformed image should fail"),
        Err(error) => error,
    };
    match error {
        rdocx::Error::UnavailableImageDimensions { filename } => {
            assert_eq!(filename, "broken.bin");
        }
        other => panic!("expected unavailable image dimensions, got {other}"),
    }

    assert_eq!(document.content_count(), 1);
    assert!(document.images().is_empty());
    let after = document.to_bytes().unwrap();
    assert_eq!(after, before);

    let package = OpcPackage::from_reader(std::io::Cursor::new(after)).unwrap();
    assert!(
        package
            .get_part_rels("/word/document.xml")
            .unwrap()
            .get_all_by_type(rel_types::IMAGE)
            .is_empty()
    );
    assert!(
        !package
            .parts
            .keys()
            .any(|name| name.starts_with("/word/media/"))
    );
    let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(!xml.contains("<w:drawing>"));
}

#[test]
fn header_footer_round_trip() {
    let mut doc = Document::new();
    doc.set_header("Page Header");
    doc.set_footer("Page Footer");
    doc.add_paragraph("Body text");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.header_text(), Some("Page Header".to_string()));
    assert_eq!(doc2.footer_text(), Some("Page Footer".to_string()));
    assert_eq!(doc2.paragraph_count(), 1);
}

#[test]
fn first_page_header_footer() {
    let mut doc = Document::new();
    doc.set_header("Default Header");
    doc.set_footer("Default Footer");
    doc.set_first_page_header("First Page Header");
    doc.set_first_page_footer("First Page Footer");
    doc.add_paragraph("Content");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.header_text(), Some("Default Header".to_string()));
    assert_eq!(doc2.footer_text(), Some("Default Footer".to_string()));
    // titlePg should be set
    let sect = doc2.section_properties().unwrap();
    assert_eq!(sect.title_pg, Some(true));
}

#[test]
fn bullet_list_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Items:");
    doc.add_bullet_list_item("First item", 0);
    doc.add_bullet_list_item("Second item", 0);
    doc.add_bullet_list_item("Sub-item", 1);
    doc.add_bullet_list_item("Third item", 0);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 5);
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Items:");
    assert_eq!(paras[1].text(), "First item");
    assert_eq!(paras[2].text(), "Second item");
    assert_eq!(paras[3].text(), "Sub-item");
    assert_eq!(paras[4].text(), "Third item");
}

#[test]
fn numbered_list_round_trip() {
    let mut doc = Document::new();
    doc.add_numbered_list_item("Step one", 0);
    doc.add_numbered_list_item("Step two", 0);
    doc.add_numbered_list_item("Sub-step a", 1);
    doc.add_numbered_list_item("Sub-step b", 1);
    doc.add_numbered_list_item("Step three", 0);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 5);
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Step one");
    assert_eq!(paras[4].text(), "Step three");
}

#[test]
fn mixed_lists_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Introduction");
    doc.add_bullet_list_item("Bullet 1", 0);
    doc.add_bullet_list_item("Bullet 2", 0);
    doc.add_paragraph("Transition");
    doc.add_numbered_list_item("Step 1", 0);
    doc.add_numbered_list_item("Step 2", 0);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 6);
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Introduction");
    assert_eq!(paras[1].text(), "Bullet 1");
    assert_eq!(paras[3].text(), "Transition");
    assert_eq!(paras[4].text(), "Step 1");
}

#[test]
fn paragraph_facade_exposes_exact_bottom_border_facts() {
    let mut document = Document::new();
    document
        .add_paragraph("")
        .border_bottom(BorderStyle::Single, 8, "808080");

    let bytes = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    let paragraph = reopened.paragraph(0).unwrap();
    let border = paragraph.bottom_border().unwrap();

    assert_eq!(paragraph.border_count(), 1);
    assert_eq!(border.style(), "single");
    assert_eq!(border.size_eighths_pt(), Some(8));
    assert_eq!(border.space_points(), Some(1));
    assert_eq!(border.color(), Some("808080"));
}

#[test]
fn custom_list_definitions_restart_numbering_per_list() {
    let mut doc = Document::new();
    let first = doc.add_list_definition(&[ListLevel::decimal()]);
    let second = doc.add_list_definition(&[ListLevel::decimal()]);
    assert_ne!(first, second);

    doc.add_paragraph("List one, item one")
        .set_numbering(first, 0);
    doc.add_paragraph("List two, item one")
        .set_numbering(second, 0);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    // Distinct numId means each list numbers independently from its start.
    assert_eq!(paras[0].numbering(), Some((first, 0)));
    assert_eq!(paras[1].numbering(), Some((second, 0)));
    assert_eq!(doc2.numbering_is_bullet(first), Some(false));
    assert_eq!(doc2.numbering_is_bullet(second), Some(false));
}

#[test]
fn custom_list_definition_mixes_formats_across_levels() {
    let mut doc = Document::new();
    let num_id = doc.add_list_definition(&[ListLevel::bullet(), ListLevel::decimal().start(3)]);

    doc.add_paragraph("bullet item").set_numbering(num_id, 0);
    doc.add_paragraph("third decimal item")
        .set_numbering(num_id, 1);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    assert_eq!(paras[0].numbering(), Some((num_id, 0)));
    assert_eq!(paras[1].numbering(), Some((num_id, 1)));
    // Level 0 is a bullet; the definition still resolves as a bullet list.
    assert_eq!(doc2.numbering_is_bullet(num_id), Some(true));

    // The persisted numbering part survives the round trip: the reopened
    // document still knows the definition, so a level can be redefined.
    let mut doc3 = Document::from_bytes(&bytes).unwrap();
    assert!(doc3.set_list_level(num_id, 2, ListLevel::decimal().start(7)));
}

#[test]
fn set_list_level_upgrades_a_deeper_level_after_the_fact() {
    let mut doc = Document::new();
    let num_id = doc.add_list_definition(&[ListLevel::bullet()]);

    // Level 1 starts as the bullet fill; content later needs it decimal.
    assert!(doc.set_list_level(num_id, 1, ListLevel::decimal().start(3)));
    assert!(!doc.set_list_level(999, 1, ListLevel::decimal()));

    doc.add_paragraph("bullet").set_numbering(num_id, 0);
    doc.add_paragraph("decimal from three")
        .set_numbering(num_id, 1);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraphs()[1].numbering(), Some((num_id, 1)));
}

#[test]
fn paragraph_numbering_value_can_be_cleared() {
    let mut doc = Document::new();
    let num_id = doc.add_list_definition(&[ListLevel::bullet()]);

    let mut para = doc.add_paragraph("was a list item");
    para.set_numbering(num_id, 0);
    para.set_numbering_value(None);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraphs()[0].numbering(), None);
}

#[test]
fn comprehensive_phase3_document() {
    // Create a document using all Phase 3 features together
    let mut doc = Document::new();

    // Page setup
    doc.set_margins(
        Length::inches(1.0),
        Length::inches(1.0),
        Length::inches(1.0),
        Length::inches(1.0),
    );
    doc.set_header("Phase 3 Test Document");
    doc.set_footer("Confidential");

    // Title
    doc.add_paragraph("Project Report")
        .style("Heading1")
        .alignment(Alignment::Center);

    // Intro paragraph
    doc.add_paragraph("This document demonstrates Phase 3 features.");

    // Bulleted list section
    doc.add_paragraph("Key Points:").style("Heading2");
    doc.add_bullet_list_item("Tables with formatting", 0);
    doc.add_bullet_list_item("Inline images", 0);
    doc.add_bullet_list_item("Headers and footers", 0);
    doc.add_bullet_list_item("Numbered and bulleted lists", 0);

    // Table section
    doc.add_paragraph("Data Summary:").style("Heading2");
    let mut table = doc
        .add_table(3, 3)
        .borders(BorderStyle::Single, 4, "000000");

    // Header row
    table.row(0).unwrap().header();
    table.cell(0, 0).unwrap().set_text("Category");
    table
        .cell(0, 0)
        .unwrap()
        .shading("4472C4")
        .vertical_alignment(VerticalAlignment::Center);
    table.cell(0, 1).unwrap().set_text("Q1");
    table.cell(0, 1).unwrap().shading("4472C4");
    table.cell(0, 2).unwrap().set_text("Q2");
    table.cell(0, 2).unwrap().shading("4472C4");

    // Data rows
    table.cell(1, 0).unwrap().set_text("Revenue");
    table.cell(1, 1).unwrap().set_text("$1,200");
    table.cell(1, 2).unwrap().set_text("$1,500");
    table.cell(2, 0).unwrap().set_text("Expenses");
    table.cell(2, 1).unwrap().set_text("$800");
    table.cell(2, 2).unwrap().set_text("$900");

    // Numbered steps
    doc.add_paragraph("Next Steps:").style("Heading2");
    doc.add_numbered_list_item("Review Q2 financials", 0);
    doc.add_numbered_list_item("Prepare Q3 forecast", 0);
    doc.add_numbered_list_item("Gather input from teams", 1);
    doc.add_numbered_list_item("Consolidate data", 1);
    doc.add_numbered_list_item("Submit report", 0);

    // Save and reload
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Verify structure
    assert_eq!(doc2.table_count(), 1);
    let tables = doc2.tables();
    assert_eq!(tables[0].row_count(), 3);
    assert_eq!(tables[0].column_count(), 3);
    assert!(tables[0].row(0).unwrap().is_header());
    assert_eq!(tables[0].cell(1, 1).unwrap().text(), "$1,200");

    // Verify header/footer
    assert_eq!(
        doc2.header_text(),
        Some("Phase 3 Test Document".to_string())
    );
    assert_eq!(doc2.footer_text(), Some("Confidential".to_string()));

    // Count paragraphs (heading + intro + 3 heading2 + 4 bullets + 5 numbered = 15 total paragraphs)
    assert!(doc2.paragraph_count() > 10);
}

#[test]
fn metadata_round_trip() {
    let mut doc = Document::new();
    doc.set_title("Test Title");
    doc.set_author("Test Author");
    doc.set_subject("Test Subject");
    doc.set_keywords("rust, docx, test");

    assert_eq!(doc.title(), Some("Test Title"));
    assert_eq!(doc.author(), Some("Test Author"));
    assert_eq!(doc.subject(), Some("Test Subject"));
    assert_eq!(doc.keywords(), Some("rust, docx, test"));

    // Round-trip through DOCX bytes
    let bytes = doc.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(&bytes)).unwrap();
    let relationship = package
        .package_rels
        .get_by_type(rel_types::CORE_PROPERTIES)
        .unwrap();
    assert_eq!(relationship.target, "docProps/core.xml");

    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.title(), Some("Test Title"));
    assert_eq!(doc2.author(), Some("Test Author"));
    assert_eq!(doc2.subject(), Some("Test Subject"));
    assert_eq!(doc2.keywords(), Some("rust, docx, test"));
}

#[test]
fn core_properties_at_relationship_target_round_trip_in_place() {
    let mut source = Document::new();
    source.set_title("Original title");
    source.set_author("Original author");

    let bytes = source.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let core_xml = package.parts.remove("/docProps/core.xml").unwrap();
    package.set_part("/custom/metadata.xml", core_xml);
    package.content_types.remove_override("/docProps/core.xml");
    package.content_types.add_override(
        "/custom/metadata.xml",
        "application/vnd.openxmlformats-package.core-properties+xml",
    );
    package
        .package_rels
        .items
        .retain(|rel| rel.rel_type != rel_types::CORE_PROPERTIES);
    package
        .package_rels
        .add(rel_types::CORE_PROPERTIES, "custom/metadata.xml");

    let mut custom_package = std::io::Cursor::new(Vec::new());
    package.write_to(&mut custom_package).unwrap();
    let mut document = Document::from_bytes(custom_package.get_ref()).unwrap();
    assert_eq!(document.title(), Some("Original title"));
    assert_eq!(document.author(), Some("Original author"));

    document.set_title("Updated title");
    let saved = document.to_bytes().unwrap();
    let saved_package = OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    assert!(saved_package.get_part("/docProps/core.xml").is_none());
    assert!(saved_package.get_part("/custom/metadata.xml").is_some());
    let relationship = saved_package
        .package_rels
        .get_by_type(rel_types::CORE_PROPERTIES)
        .unwrap();
    assert_eq!(relationship.target, "custom/metadata.xml");

    let reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(reopened.title(), Some("Updated title"));
    assert_eq!(reopened.author(), Some("Original author"));
}

#[test]
fn split_run_lets_a_bookmark_and_a_comment_cover_part_of_a_run() {
    let mut doc = Document::new();
    doc.add_paragraph("Hello brave world");
    let mut paragraph = doc.paragraph_mut(0).unwrap();
    paragraph.split_run(0, 6).unwrap();
    paragraph.split_run(1, 5).unwrap();
    assert_eq!(paragraph.split_run(1, 5).unwrap(), 2);
    let error = paragraph.split_run(1, 6).unwrap_err();
    assert!(error.to_string().contains("literal text length"), "{error}");
    let texts: Vec<String> = doc
        .paragraph(0)
        .unwrap()
        .runs()
        .map(|run| run.text())
        .collect();
    assert_eq!(texts, ["Hello ", "brave", " world"]);

    let brave = RunRange {
        start: RunPosition {
            body_index: 0,
            run_index: 1,
        },
        end: RunPosition {
            body_index: 0,
            run_index: 2,
        },
    };
    doc.add_bookmark("brave", brave).unwrap();
    doc.add_comment(brave, "Ada", None, "Which one?").unwrap();

    let reopened = Document::from_bytes(&doc.to_bytes().unwrap()).unwrap();
    let bookmarks = reopened.bookmarks();
    assert_eq!(bookmarks.len(), 1);
    assert_eq!(bookmarks[0].text(), "brave");
    assert_eq!(reopened.comments().len(), 1);
}

#[test]
fn three_comments_and_cross_paragraph_anchors_round_trip_byte_identically() {
    let mut source = Document::new();
    source.add_paragraph("first");
    source.add_paragraph("second");
    let bytes = source.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();

    let document_xml = package.get_part("/word/document.xml").unwrap();
    let document_xml = String::from_utf8(document_xml.to_vec())
        .unwrap()
        .replace(
            "<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>",
            "<w:p><w:commentRangeStart w:id=\"1\"/><w:r><w:t>first</w:t></w:r><w:commentRangeStart w:id=\"2\"/><w:commentRangeEnd w:id=\"2\"/><w:r><w:commentReference w:id=\"2\"/></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r><w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r><w:commentRangeStart w:id=\"3\"/><w:commentRangeEnd w:id=\"3\"/><w:r><w:commentReference w:id=\"3\"/></w:r></w:p>",
        );
    package.set_part("/word/document.xml", document_xml.as_bytes().to_vec());

    let comments_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:author="Ada" w:id="1"><w:p><w:r><w:t>one</w:t></w:r></w:p></w:comment><w:comment w:author="Ben" w:id="2"><w:p><w:r><w:t>two</w:t></w:r></w:p></w:comment><w:comment w:author="Cy" w:id="3"><w:p><w:r><w:t>three</w:t></w:r></w:p></w:comment></w:comments>"#;
    package.set_part("/word/comments.xml", comments_xml.to_vec());
    package.content_types.add_override(
        "/word/comments.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    );
    package
        .get_or_create_part_rels("/word/document.xml")
        .add(rel_types::COMMENTS, "comments.xml");

    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let mut document = Document::from_bytes(input.get_ref()).unwrap();
    let saved = document.to_bytes().unwrap();
    let saved_package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();

    assert_eq!(
        saved_package.get_part("/word/comments.xml"),
        Some(comments_xml.as_slice())
    );
    assert_eq!(
        saved_package.get_part("/word/document.xml"),
        Some(document_xml.as_bytes())
    );
}

#[test]
fn encoded_comment_ids_reopen_as_one_complete_comment_anchor() {
    let mut source = Document::new();
    source.add_paragraph("commented");
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(source.to_bytes().unwrap()))
        .expect("source package");
    package.set_part(
        "/word/document.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:commentRangeStart w:id="&#55;"/><w:r><w:t>commented</w:t></w:r><w:commentRangeEnd w:id="&#x37;"/><w:r><w:commentReference w:id="&#55;"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#.to_vec(),
    );
    package.set_part(
        "/word/comments.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:author="Ada" w:id="7"><w:p><w:r><w:t>encoded anchor</w:t></w:r></w:p></w:comment></w:comments>"#.to_vec(),
    );
    package.content_types.add_override(
        "/word/comments.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    );
    package
        .get_or_create_part_rels("/word/document.xml")
        .add(rel_types::COMMENTS, "comments.xml");

    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(bytes.get_ref()).expect("encoded comment package");
    assert_eq!(document.comments()[0].id(), 7);
    let paragraph = document.paragraph(0).expect("comment paragraph");
    let range_ids = paragraph
        .items()
        .filter_map(|item| match item {
            rdocx::paragraph::ParagraphItemRef::CommentRangeStart { id, .. } => Some((true, id)),
            rdocx::paragraph::ParagraphItemRef::CommentRangeEnd { id, .. } => Some((false, id)),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(range_ids, [(true, 7), (false, 7)]);
    let reference_ids = paragraph
        .runs()
        .filter_map(|run| {
            run.items().find_map(|item| match item {
                rdocx::run::RunItemRef::CommentReference(id) => Some(id),
                _ => None,
            })
        })
        .collect::<Vec<_>>();
    assert_eq!(reference_ids, [7]);

    let saved = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&saved).expect("saved comment package reopens");
    assert_eq!(reopened.comments()[0].id(), 7);
    let saved = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    let xml = std::str::from_utf8(saved.get_part("/word/document.xml").unwrap()).unwrap();
    assert!(xml.contains(r#"<w:commentRangeStart w:id="7"/>"#));
    assert!(xml.contains(r#"<w:commentRangeEnd w:id="7"/>"#));
    assert!(xml.contains(r#"<w:commentReference w:id="7"/>"#));
}

#[test]
fn comments_part_uses_its_existing_relationship_target() {
    let mut source = Document::new();
    source.add_paragraph("body");
    let bytes = source.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let comments_xml = br#"<x:comments xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:comment x:id="9"><x:p><x:r><x:t>custom target</x:t></x:r></x:p><raw:commentChild xmlns:raw="urn:rdocx:comments-raw" raw:kept="exact"/></x:comment><raw:rootChild xmlns:raw="urn:rdocx:comments-raw" raw:kept="exact"/></x:comments>"#;
    package.set_part("/custom/comments-data.xml", comments_xml.to_vec());
    package.content_types.add_override(
        "/custom/comments-data.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    );
    package
        .get_or_create_part_rels("/word/document.xml")
        .add(rel_types::COMMENTS, "../custom/comments-data.xml");
    flat_opc_package_class_tests::add_package_signature_graph(&mut package);

    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let mut document = Document::from_bytes(input.get_ref()).unwrap();
    let assert_canonical_package = |saved_package: &OpcPackage| {
        assert!(saved_package.get_part("/word/comments.xml").is_none());
        assert_eq!(
            saved_package
                .get_part_rels("/word/document.xml")
                .unwrap()
                .get_by_type(rel_types::COMMENTS)
                .unwrap()
                .target,
            "../custom/comments-data.xml"
        );
        let output =
            std::str::from_utf8(saved_package.get_part("/custom/comments-data.xml").unwrap())
                .unwrap();
        assert!(output.contains("<w:comments"), "{output}");
        assert!(output.contains("custom target"), "{output}");
        assert!(
            output.contains(
                r#"<raw:commentChild xmlns:raw="urn:rdocx:comments-raw" raw:kept="exact"/>"#
            ),
            "{output}"
        );
        assert!(
            output.contains(
                r#"<raw:rootChild xmlns:raw="urn:rdocx:comments-raw" raw:kept="exact"/>"#
            ),
            "{output}"
        );
        assert!(
            flat_opc_package_class_tests::has_package_signature_invalidation_marker(saved_package)
        );
        assert!(saved_package.parts.contains_key("/_xmlsignatures/sig1.xml"));
    };

    let saved = document.to_bytes().unwrap();
    assert_eq!(document.to_bytes().unwrap(), saved);
    let saved_package = OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
    assert_canonical_package(&saved_package);
    let mut reopened = Document::from_bytes(&saved).unwrap();
    assert_eq!(reopened.comments()[0].text(), "custom target");
    assert_eq!(reopened.to_bytes().unwrap(), saved);

    let flat = document.to_flat_opc_bytes().unwrap();
    assert_eq!(document.to_flat_opc_bytes().unwrap(), flat);
    let flat_xml = std::str::from_utf8(&flat).unwrap();
    assert!(flat_xml.contains(r#"pkg:name="/custom/comments-data.xml""#));
    assert!(flat_xml.contains("<w:comments"), "{flat_xml}");
    assert!(flat_xml.contains("urn:rdocx:relationships/invalidated-package-signature"));
    assert!(flat_xml.contains(r#"pkg:name="/_xmlsignatures/sig1.xml""#));
    let mut reopened = Document::from_flat_opc_bytes(&flat).unwrap();
    assert_eq!(reopened.comments()[0].text(), "custom target");
    assert_eq!(reopened.to_flat_opc_bytes().unwrap(), flat);
    let reopened_package =
        OpcPackage::from_reader(std::io::Cursor::new(reopened.to_bytes().unwrap())).unwrap();
    assert_canonical_package(&reopened_package);
}

#[test]
fn saving_without_comments_does_not_manufacture_a_comments_part() {
    let mut document = Document::new();
    document.add_paragraph("No review thread");
    let saved = document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();

    assert!(package.get_part("/word/comments.xml").is_none());
    assert!(
        !package
            .content_types
            .overrides
            .contains_key("/word/comments.xml")
    );
    assert!(
        package
            .get_part_rels("/word/document.xml")
            .and_then(|rels| rels.get_by_type(rel_types::COMMENTS))
            .is_none()
    );
}

#[test]
fn comment_parts_relationships_and_content_types_are_word_compatible() {
    let mut document = Document::new();
    document.add_paragraph("review this");
    let id = document
        .add_comment(
            RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 0,
                    run_index: 1,
                },
            },
            "Ada",
            Some("AL"),
            "Please review",
        )
        .unwrap();
    document.reply_to(id, "Ben", "Done").unwrap();
    document.resolve_comment(id, true).unwrap();

    let bytes = document.to_bytes().unwrap();
    let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let relationships = package.get_part_rels("/word/document.xml").unwrap();
    let comments = relationships.get_by_type(rel_types::COMMENTS).unwrap();
    let extended = relationships
        .get_by_type("http://schemas.microsoft.com/office/2011/relationships/commentsExtended")
        .unwrap();
    let comments_part = OpcPackage::resolve_rel_target("/word/document.xml", &comments.target);
    let extended_part = OpcPackage::resolve_rel_target("/word/document.xml", &extended.target);

    assert_eq!(
        package.content_types.content_type_for(&comments_part),
        Some("application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    );
    assert_eq!(
        package.content_types.content_type_for(&extended_part),
        Some("application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml")
    );
    let comments_xml =
        String::from_utf8(package.get_part(&comments_part).unwrap().to_vec()).unwrap();
    let extended_xml =
        String::from_utf8(package.get_part(&extended_part).unwrap().to_vec()).unwrap();
    assert!(
        comments_xml.contains("xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"")
    );
    assert!(comments_xml.contains("w14:paraId="));
    assert!(
        extended_xml.contains("xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"")
    );
    assert!(extended_xml.contains("w15:paraIdParent="));
    assert!(extended_xml.contains("w15:done=\"1\""));
}

#[test]
fn comments_extended_part_uses_its_existing_relationship_target() {
    const COMMENTS_EXTENDED_REL_TYPE: &str =
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
    const COMMENTS_EXTENDED_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";

    let mut document = Document::new();
    document.add_paragraph("review this");
    let id = document
        .add_comment(
            RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 0,
                    run_index: 1,
                },
            },
            "Ada",
            None,
            "Review",
        )
        .unwrap();
    let bytes = document.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let extension = package.parts.remove("/word/commentsExtended.xml").unwrap();
    package.set_part("/custom/thread-data.xml", extension);
    package
        .content_types
        .overrides
        .remove("/word/commentsExtended.xml");
    package
        .content_types
        .add_override("/custom/thread-data.xml", COMMENTS_EXTENDED_CONTENT_TYPE);
    package
        .part_rels
        .get_mut("/word/document.xml")
        .unwrap()
        .items
        .iter_mut()
        .find(|relationship| relationship.rel_type == COMMENTS_EXTENDED_REL_TYPE)
        .unwrap()
        .target = "../custom/thread-data.xml".to_owned();
    let mut seeded = std::io::Cursor::new(Vec::new());
    package.write_to(&mut seeded).unwrap();

    let mut reopened = Document::from_bytes(seeded.get_ref()).unwrap();
    assert!(reopened.resolve_comment(id, true).unwrap());
    let saved = reopened.to_bytes().unwrap();
    let saved_package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
    assert!(
        saved_package
            .get_part("/word/commentsExtended.xml")
            .is_none()
    );
    assert!(saved_package.get_part("/custom/thread-data.xml").is_some());
    let relationship = saved_package
        .get_part_rels("/word/document.xml")
        .and_then(|relationships| relationships.get_by_type(COMMENTS_EXTENDED_REL_TYPE))
        .unwrap();
    assert_eq!(relationship.target, "../custom/thread-data.xml");
}

#[test]
fn nested_table_round_trip() {
    let mut doc = Document::new();

    // Create outer 2x2 table
    let mut tbl = doc.add_table(2, 2);
    tbl.cell(0, 0).unwrap().set_text("Outer A1");
    tbl.cell(0, 1).unwrap().set_text("Outer A2");
    tbl.cell(1, 1).unwrap().set_text("Outer B2");

    // Add a nested 2x1 table inside cell (1, 0)
    let mut cell_b1 = tbl.cell(1, 0).unwrap();
    cell_b1.set_text("Before nested");
    let mut nested = cell_b1.add_table(2, 1);
    nested.cell(0, 0).unwrap().set_text("Inner R1");
    nested.cell(1, 0).unwrap().set_text("Inner R2");

    // Serialize and reload
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Verify outer table structure
    assert_eq!(doc2.table_count(), 1);
    let tables = doc2.tables();
    assert!(!tables.is_empty());
    let tbl2 = &tables[0];
    assert_eq!(tbl2.row_count(), 2);
    assert_eq!(tbl2.cell(0, 0).unwrap().text(), "Outer A1");
    assert_eq!(tbl2.cell(0, 1).unwrap().text(), "Outer A2");
    assert_eq!(tbl2.cell(1, 1).unwrap().text(), "Outer B2");

    // Cell (1,0) should have paragraph text (nested table text excluded from text())
    let cell_b1_ref = tbl2.cell(1, 0).unwrap();
    assert_eq!(cell_b1_ref.text(), "Before nested");
}

#[test]
fn comprehensive_document_round_trip_with_nested() {
    use rdocx::paragraph::Alignment;

    let mut doc = Document::new();

    // Metadata
    doc.set_title("Comprehensive Test Document");
    doc.set_author("rdocx Test Suite");

    // Heading 1
    doc.add_paragraph("Chapter 1: Introduction")
        .style("Heading1");

    // Normal paragraphs
    doc.add_paragraph("This is a normal paragraph with some body text.")
        .alignment(Alignment::Left);

    doc.add_paragraph("This paragraph is centered for emphasis.")
        .alignment(Alignment::Center);

    doc.add_paragraph("This paragraph is justified for a clean look.")
        .alignment(Alignment::Justify);

    // Heading 2
    doc.add_paragraph("Section 1.1: Data Table")
        .style("Heading2");

    // Table with formatting
    let mut tbl = doc.add_table(3, 3);
    tbl = tbl.borders(rdocx::BorderStyle::Single, 4, "000000");

    // Header row
    tbl.row(0).unwrap().header();
    tbl.cell(0, 0).unwrap().set_text("Name");
    tbl.cell(0, 1).unwrap().set_text("Value");
    tbl.cell(0, 2).unwrap().set_text("Status");

    // Data rows
    tbl.cell(1, 0).unwrap().set_text("Alpha");
    tbl.cell(1, 1).unwrap().set_text("100");
    tbl.cell(1, 2).unwrap().set_text("Active");

    tbl.cell(2, 0).unwrap().set_text("Beta");
    tbl.cell(2, 1).unwrap().set_text("200");
    tbl.cell(2, 2).unwrap().set_text("Pending");

    // Another heading
    doc.add_paragraph("Chapter 2: Nested Content")
        .style("Heading1");

    // Table with nested table
    let mut tbl2 = doc.add_table(2, 2);
    tbl2.cell(0, 0).unwrap().set_text("Outer cell");
    tbl2.cell(0, 1).unwrap().set_text("Another outer cell");
    tbl2.cell(1, 0).unwrap().set_text("Simple cell");

    let mut nested_cell = tbl2.cell(1, 1).unwrap();
    nested_cell.set_text("Contains nested table:");
    let mut nested = nested_cell.add_table(2, 2);
    nested.cell(0, 0).unwrap().set_text("N1");
    nested.cell(0, 1).unwrap().set_text("N2");
    nested.cell(1, 0).unwrap().set_text("N3");
    nested.cell(1, 1).unwrap().set_text("N4");

    // Bullet list
    doc.add_paragraph("Chapter 3: Lists").style("Heading1");
    doc.add_paragraph("First bullet point").style("ListBullet");
    doc.add_paragraph("Second bullet point").style("ListBullet");
    doc.add_paragraph("Third bullet point").style("ListBullet");

    // Final paragraph
    doc.add_paragraph("End of document.");

    // Round-trip through DOCX
    let docx_bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&docx_bytes).unwrap();

    // Verify structure
    assert_eq!(doc2.title(), Some("Comprehensive Test Document"));
    assert_eq!(doc2.author(), Some("rdocx Test Suite"));
    assert!(doc2.table_count() >= 2);
}

// ---- Phase 7: Template Engine, Placeholder Replacement & Background Images ----

#[test]
fn template_open_replace_save() {
    // Create a "template" document
    let mut template = Document::new();
    template.set_header("Company: {{company}}");
    template.set_footer("Date: {{date}}");
    template.add_paragraph("Dear {{name}},").style("Heading1");
    template.add_paragraph("Welcome to {{company}}. Your role starts on {{date}}.");
    template.add_paragraph("{{INSERT_CONTENT}}");
    template.add_paragraph("Best regards,");
    template.add_paragraph("HR Department");

    let template_bytes = template.to_bytes().unwrap();

    // Open template and do replacements
    let mut doc = Document::from_bytes(&template_bytes).unwrap();
    doc.replace_text("{{company}}", "Acme Corp");
    doc.replace_text("{{name}}", "Alice");
    doc.replace_text("{{date}}", "2026-03-01");

    // Find and replace content placeholder
    if let Some(idx) = doc.find_content_index("{{INSERT_CONTENT}}") {
        doc.remove_content(idx);
        doc.insert_paragraph(idx, "Your onboarding schedule is attached.");
        doc.insert_paragraph(idx + 1, "Please review and confirm.");
    }

    // Verify
    let paras = doc.paragraphs();
    assert_eq!(paras[0].text(), "Dear Alice,");
    assert_eq!(
        paras[1].text(),
        "Welcome to Acme Corp. Your role starts on 2026-03-01."
    );
    assert_eq!(paras[2].text(), "Your onboarding schedule is attached.");
    assert_eq!(paras[3].text(), "Please review and confirm.");
    assert_eq!(paras[4].text(), "Best regards,");

    assert_eq!(doc.header_text().unwrap(), "Company: Acme Corp");
    assert_eq!(doc.footer_text().unwrap(), "Date: 2026-03-01");

    // Round-trip
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraphs()[0].text(), "Dear Alice,");
}

#[test]
fn replace_all_batch_workflow() {
    let mut doc = Document::new();
    doc.add_paragraph("{{a}} {{b}} {{c}}");

    let mut map = std::collections::HashMap::new();
    map.insert("{{a}}", "X");
    map.insert("{{b}}", "Y");
    map.insert("{{c}}", "Z");
    let count = doc.replace_all(&map);
    assert_eq!(count, 3);
    assert_eq!(doc.paragraphs()[0].text(), "X Y Z");
}

#[test]
fn background_image_end_to_end() {
    // Minimal 1x1 PNG
    let png_data: Vec<u8> = vec![
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
        0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];

    let mut doc = Document::new();
    doc.add_paragraph("Page content here");
    doc.add_background_image(&png_data, "background.png");

    // Background paragraph inserted at index 0
    assert_eq!(doc.content_count(), 2);

    // Round-trip DOCX
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.content_count(), 2);
}

#[test]
fn full_phase7_workflow() {
    // Minimal PNG
    let png_data: Vec<u8> = vec![
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
        0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];

    let mut doc = Document::new();
    doc.set_title("Template Output");
    doc.set_header("{{company}} - Confidential");

    doc.add_paragraph("Report for {{company}}")
        .style("Heading1");
    doc.add_paragraph("Date: {{date}}");
    doc.add_paragraph("{{INSERT_HERE}}");
    doc.add_paragraph("Summary: {{company}} performed well in {{date}}.");

    // Add background image
    doc.add_background_image(&png_data, "bg.png");

    // Replace placeholders
    doc.replace_text("{{company}}", "Acme Corp");
    doc.replace_text("{{date}}", "2026-02-22");

    // Insert content at placeholder position
    if let Some(idx) = doc.find_content_index("{{INSERT_HERE}}") {
        doc.remove_content(idx);
        doc.insert_paragraph(idx, "Revenue increased by 15%.");
    }

    // Verify final state
    let paras = doc.paragraphs();
    // First paragraph is the background image paragraph, skip it
    let text_paras: Vec<_> = paras.iter().filter(|p| !p.text().is_empty()).collect();
    assert!(
        text_paras
            .iter()
            .any(|p| p.text() == "Report for Acme Corp")
    );
    assert!(text_paras.iter().any(|p| p.text() == "Date: 2026-02-22"));
    assert!(
        text_paras
            .iter()
            .any(|p| p.text() == "Revenue increased by 15%.")
    );
    assert!(
        text_paras
            .iter()
            .any(|p| p.text() == "Summary: Acme Corp performed well in 2026-02-22.")
    );

    assert_eq!(doc.header_text().unwrap(), "Acme Corp - Confidential");

    // Save as DOCX
    let bytes = doc.to_bytes().unwrap();
    assert!(!bytes.is_empty());
}

// ---- Phase: Code Quality & Modernization — New Tests ----

#[test]
fn section_break_builders_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("First section");

    // Create a section break to landscape
    doc.add_paragraph("Landscape section break")
        .section_break(SectionBreak::NextPage)
        .section_landscape();

    doc.add_paragraph("In landscape section");

    // Switch back to portrait
    doc.add_paragraph("Portrait section break")
        .section_break(SectionBreak::NextPage)
        .section_portrait();

    doc.add_paragraph("Back to portrait");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraph_count(), 5);
}

#[test]
fn section_page_size_custom() {
    let mut doc = Document::new();
    // Custom page size: 6" x 9" (book size)
    doc.add_paragraph("Small page")
        .section_break(SectionBreak::NextPage)
        .section_page_size(Length::inches(6.0), Length::inches(9.0));

    doc.add_paragraph("On the next section");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraph_count(), 2);
}

#[test]
fn section_break_continuous() {
    let mut doc = Document::new();
    doc.add_paragraph("Before continuous break")
        .section_break(SectionBreak::Continuous);
    doc.add_paragraph("After continuous break");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraph_count(), 2);
}

#[test]
fn tab_stops_all_leader_styles() {
    let mut doc = Document::new();

    // Tab stop with no leader
    doc.add_paragraph("Left\tRight")
        .add_tab_stop(TabAlignment::Right, Length::inches(6.0));

    // Tab stop with dot leader
    doc.add_paragraph("Item\t100").add_tab_stop_with_leader(
        TabAlignment::Right,
        Length::inches(6.0),
        TabLeader::Dot,
    );

    // Tab stop with hyphen leader
    doc.add_paragraph("Section\tPage 5")
        .add_tab_stop_with_leader(TabAlignment::Right, Length::inches(6.0), TabLeader::Hyphen);

    // Tab stop with underscore leader
    doc.add_paragraph("Name\t").add_tab_stop_with_leader(
        TabAlignment::Right,
        Length::inches(6.0),
        TabLeader::Underscore,
    );

    // Multiple alignments
    doc.add_paragraph("A\tB\tC")
        .add_tab_stop(TabAlignment::Center, Length::inches(3.0))
        .add_tab_stop(TabAlignment::Right, Length::inches(6.0))
        .add_tab_stop(TabAlignment::Decimal, Length::inches(4.5));

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    let paras = doc2.paragraphs();

    assert_eq!(paras[0].tab_stop_count(), 1);
    assert_eq!(paras[1].tab_stop_count(), 1);
    assert_eq!(paras[2].tab_stop_count(), 1);
    assert_eq!(paras[3].tab_stop_count(), 1);
    assert_eq!(paras[4].tab_stop_count(), 3);
}

#[test]
fn run_formatting_all_caps_small_caps() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("UPPERCASE").all_caps(true);
    para.add_run("SmallCaps").small_caps(true);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "UPPERCASESmallCaps");
}

#[test]
fn run_formatting_double_strike_and_spacing() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("DStrike").double_strike(true);
    para.add_run("Spaced").character_spacing(Length::pt(3.0));
    para.add_run("Super").superscript();
    para.add_run("Sub").subscript();

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    let paras = doc2.paragraphs();
    let runs: Vec<_> = paras[0].runs().collect();
    assert_eq!(runs.len(), 4);
    assert_eq!(runs[2].vert_align(), Some("superscript"));
    assert_eq!(runs[3].vert_align(), Some("subscript"));
    assert!(runs[1].character_spacing().is_some());
}

#[test]
fn paragraph_border_bottom_only() {
    let mut doc = Document::new();
    doc.add_paragraph("Bottom bordered")
        .border_bottom(BorderStyle::Single, 4, "000000");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    let paras = doc2.paragraphs();
    assert!(paras[0].has_borders());
}

#[test]
fn paragraph_shading_and_indent_combined() {
    let mut doc = Document::new();
    doc.add_paragraph("Shaded and indented")
        .shading("E0E0E0")
        .indent_left(Length::inches(0.75))
        .indent_right(Length::inches(0.5))
        .hanging_indent(Length::inches(0.25));

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].shading_fill(), Some("E0E0E0"));
    assert_eq!(paras[0].text(), "Shaded and indented");
}

#[test]
fn document_header_footer_first_page() {
    let mut doc = Document::new();
    doc.set_header("Default Header");
    doc.set_footer("Default Footer");
    doc.set_first_page_header("First Page Header");
    doc.set_first_page_footer("First Page Footer");
    doc.add_paragraph("Body content");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.header_text(), Some("Default Header".to_string()));
    assert_eq!(doc2.footer_text(), Some("Default Footer".to_string()));
    let sect = doc2.section_properties().unwrap();
    assert_eq!(sect.title_pg, Some(true));
}

#[test]
fn insert_paragraph_at_beginning_and_end() {
    let mut doc = Document::new();
    doc.add_paragraph("Middle");

    // Insert at beginning
    doc.insert_paragraph(0, "First");
    // Insert at end
    let count = doc.content_count();
    doc.insert_paragraph(count, "Last");

    assert_eq!(doc.content_count(), 3);
    let paras = doc.paragraphs();
    assert_eq!(paras[0].text(), "First");
    assert_eq!(paras[1].text(), "Middle");
    assert_eq!(paras[2].text(), "Last");
}

#[test]
fn insert_table_at_index() {
    let mut doc = Document::new();
    doc.add_paragraph("Before");
    doc.add_paragraph("After");

    // Insert table between the two paragraphs
    let mut table = doc.insert_table(1, 2, 2);
    table.cell(0, 0).unwrap().set_text("Cell");

    assert_eq!(doc.content_count(), 3);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.table_count(), 1);
    assert_eq!(doc2.paragraph_count(), 2);
}

#[test]
fn remove_content_basic() {
    let mut doc = Document::new();
    doc.add_paragraph("Keep");
    doc.add_paragraph("Remove");
    doc.add_paragraph("Keep too");

    assert_eq!(doc.content_count(), 3);
    assert!(doc.remove_content(1));
    assert_eq!(doc.content_count(), 2);
    assert_eq!(doc.paragraphs()[0].text(), "Keep");
    assert_eq!(doc.paragraphs()[1].text(), "Keep too");
}

#[test]
fn remove_content_out_of_bounds() {
    let mut doc = Document::new();
    doc.add_paragraph("Only");
    assert!(!doc.remove_content(5));
    assert_eq!(doc.content_count(), 1);
}

#[test]
fn find_content_index_returns_none_for_missing() {
    let mut doc = Document::new();
    doc.add_paragraph("Hello");
    assert_eq!(doc.find_content_index("nonexistent"), None);
}

#[test]
fn section_break_round_trip_preserves() {
    let mut doc = Document::new();
    doc.add_paragraph("Section 1")
        .section_break(SectionBreak::NextPage)
        .section_landscape();

    doc.add_paragraph("Section 2 (landscape)");

    doc.add_paragraph("Section 2 end")
        .section_break(SectionBreak::NextPage)
        .section_portrait();

    doc.add_paragraph("Section 3 (portrait)");

    // Round-trip
    let bytes = doc.to_bytes().unwrap();
    let mut doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraph_count(), 4);

    // The document should still be valid and have the paragraphs
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Section 1");
    assert_eq!(paras[1].text(), "Section 2 (landscape)");
    assert_eq!(paras[2].text(), "Section 2 end");
    assert_eq!(paras[3].text(), "Section 3 (portrait)");

    // Re-round-trip to verify stability
    let bytes2 = doc2.to_bytes().unwrap();
    let doc3 = Document::from_bytes(&bytes2).unwrap();
    assert_eq!(doc3.paragraph_count(), 4);
}

#[test]
fn empty_document_insert_and_remove() {
    let mut doc = Document::new();
    assert_eq!(doc.content_count(), 0);

    doc.insert_paragraph(0, "Inserted");
    assert_eq!(doc.content_count(), 1);
    assert_eq!(doc.paragraphs()[0].text(), "Inserted");

    assert!(doc.remove_content(0));
    assert_eq!(doc.content_count(), 0);
}

#[test]
fn direct_content_indices_map_across_block_content_control_paragraphs() {
    let mut seed = Document::new();
    seed.add_paragraph("first");
    seed.add_paragraph("last");
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap()))
        .expect("open seed package");
    let body = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
    let body = body.replacen(
        "<w:body>",
        "<w:body><w:sdt><w:sdtContent><w:p><w:r><w:t>control one</w:t></w:r></w:p><w:p><w:r><w:t>control two</w:t></w:r></w:p></w:sdtContent></w:sdt><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        1,
    );
    package.set_part("/word/document.xml", body.into_bytes());
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut document = Document::from_bytes(&bytes.into_inner()).unwrap();

    assert_eq!(document.content_count(), 4);
    assert_eq!(document.paragraph_index_of_content(0), None);
    assert_eq!(document.paragraph_index_of_content(1), None);
    assert_eq!(document.paragraph_index_of_content(2), Some(2));
    assert_eq!(document.content_index_of_paragraph(0), None);
    assert_eq!(document.content_index_of_paragraph(1), None);
    assert_eq!(document.content_index_of_paragraph(2), Some(2));
    assert_eq!(document.content_index_of_table(0), Some(1));

    document.insert_paragraph(3, "middle");
    let paragraph_index = document.paragraph_index_of_content(3).unwrap();
    assert_eq!(paragraph_index, 3);
    assert_eq!(
        document.content_index_of_paragraph(paragraph_index),
        Some(3)
    );
    document
        .paragraph_mut(paragraph_index)
        .unwrap()
        .add_run("!");
    assert_eq!(
        document.paragraph(paragraph_index).unwrap().text(),
        "middle!"
    );
}

// ---- PDF rendering tests ----

#[test]
fn to_pdf_simple_document() {
    let mut doc = Document::new();
    doc.add_paragraph("Hello, World!");
    doc.add_paragraph("This is a test document.");

    let result = doc.to_pdf();
    // On systems without fonts, layout may fail — that's OK for CI
    if let Ok(pdf_bytes) = result {
        // Verify it starts with PDF header
        assert!(pdf_bytes.starts_with(b"%PDF"));
        // Verify it's not trivially small
        assert!(pdf_bytes.len() > 100);
        // Verify it ends with %%EOF
        let tail = String::from_utf8_lossy(&pdf_bytes[pdf_bytes.len().saturating_sub(10)..]);
        assert!(tail.contains("%%EOF"));
    }
}

#[test]
fn to_pdf_with_formatting() {
    let mut doc = Document::new();
    doc.add_paragraph("Title")
        .style("Heading1")
        .alignment(Alignment::Center);
    doc.add_paragraph("Normal text with ")
        .add_run("bold")
        .bold(true);
    doc.add_paragraph("Another paragraph");

    let result = doc.to_pdf();
    if let Ok(pdf_bytes) = result {
        assert!(pdf_bytes.starts_with(b"%PDF"));
        assert!(pdf_bytes.len() > 200);
    }
}

#[test]
fn to_pdf_with_table() {
    let mut doc = Document::new();
    doc.add_paragraph("Table test");
    {
        let mut table = doc.add_table(2, 3);
        table.cell(0, 0).unwrap().set_text("A1");
        table.cell(0, 1).unwrap().set_text("B1");
        table.cell(0, 2).unwrap().set_text("C1");
        table.cell(1, 0).unwrap().set_text("A2");
        table.cell(1, 1).unwrap().set_text("B2");
        table.cell(1, 2).unwrap().set_text("C2");
    }
    doc.add_paragraph("After table");

    let result = doc.to_pdf();
    if let Ok(pdf_bytes) = result {
        assert!(pdf_bytes.starts_with(b"%PDF"));
    }
}

#[test]
fn to_pdf_with_metadata() {
    let mut doc = Document::new();
    doc.set_title("Test Document");
    doc.set_author("rdocx");
    doc.add_paragraph("Content");

    let result = doc.to_pdf();
    if let Ok(pdf_bytes) = result {
        assert!(pdf_bytes.starts_with(b"%PDF"));
        // Metadata should be embedded in the PDF
        let pdf_str = String::from_utf8_lossy(&pdf_bytes);
        assert!(pdf_str.contains("Test Document") || pdf_str.contains("rdocx-pdf"));
    }
}

#[test]
fn save_pdf_to_file() {
    let mut doc = Document::new();
    doc.add_paragraph("PDF file test");

    let path = std::env::temp_dir().join(format!("rdocx_test_output_{}.pdf", std::process::id()));
    let result = doc.save_pdf(&path);
    if result.is_ok() {
        let bytes = std::fs::read(&path).unwrap();
        assert!(bytes.starts_with(b"%PDF"));
        std::fs::remove_file(&path).ok();
    }

    let tracked_path = std::env::temp_dir().join(format!(
        "rdocx_test_tracked_output_{}.pdf",
        std::process::id()
    ));
    let result = doc.save_pdf_with_options(
        &tracked_path,
        rdocx::RenderOptions {
            revision_view: rdocx::RevisionView::Tracked,
        },
    );
    if result.is_ok() {
        let bytes = std::fs::read(&tracked_path).unwrap();
        assert!(bytes.starts_with(b"%PDF"));
        std::fs::remove_file(&tracked_path).ok();
    }
}

#[test]
fn line_spacing_multiple_round_trips() {
    let mut doc = Document::new();
    doc.add_paragraph("double spaced")
        .line_spacing_multiple(2.0);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    let p = paras.first().unwrap();
    assert_eq!(p.line_spacing_multiple(), Some(2.0));
}

#[test]
fn non_consuming_setters_mutate_borrowed_wrappers() {
    let mut doc = Document::new();
    doc.add_paragraph("Borrowed paragraph");

    doc.paragraph_mut(0)
        .unwrap()
        .set_alignment(Alignment::Right);
    doc.paragraph_mut(0)
        .unwrap()
        .add_run(" borrowed run")
        .set_bold(true);

    {
        let mut table = doc.add_table(1, 1);
        table.set_layout_fixed();
        table.row(0).unwrap().set_header();
        table
            .cell(0, 0)
            .unwrap()
            .set_vertical_alignment(VerticalAlignment::Center);
    }

    let bytes = doc.to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    let paragraphs = reopened.paragraphs();
    assert_eq!(paragraphs[0].alignment(), Some(Alignment::Right));
    assert!(paragraphs[0].runs().last().unwrap().is_bold());
    let tables = reopened.tables();
    assert!(tables[0].row(0).unwrap().is_header());
    assert_eq!(
        tables[0].cell(0, 0).unwrap().vertical_alignment(),
        Some(VerticalAlignment::Center)
    );
}

#[test]
fn non_consuming_setters_match_consuming_builders() {
    let mut builders = Document::new();
    {
        let mut paragraph = builders
            .add_paragraph("Paragraph")
            .alignment(Alignment::Center)
            .space_after(Length::pt(6.0));
        paragraph
            .add_run(" run")
            .bold(true)
            .font("Arial")
            .color("123456");
    }
    {
        let mut table = builders
            .add_table(1, 1)
            .style("TableGrid")
            .alignment(Alignment::Center)
            .layout_fixed();
        table.row(0).unwrap().height(Length::pt(18.0)).header();
        table
            .cell(0, 0)
            .unwrap()
            .width(Length::inches(2.0))
            .shading("D9EAF7")
            .vertical_alignment(VerticalAlignment::Center);
    }

    let mut setters = Document::new();
    {
        let mut paragraph = setters.add_paragraph("Paragraph");
        paragraph.set_alignment(Alignment::Center);
        paragraph.set_space_after(Length::pt(6.0));
        let mut run = paragraph.add_run(" run");
        run.set_bold(true);
        run.set_font("Arial");
        run.set_color("123456");
    }
    {
        let mut table = setters.add_table(1, 1);
        table.set_style("TableGrid");
        table.set_alignment(Alignment::Center);
        table.set_layout_fixed();
        let mut row = table.row(0).unwrap();
        row.set_height(Length::pt(18.0));
        row.set_header();
        let mut cell = row.cell(0).unwrap();
        cell.set_width(Length::inches(2.0));
        cell.set_shading("D9EAF7");
        cell.set_vertical_alignment(VerticalAlignment::Center);
    }

    assert_eq!(document_xml(&mut setters), document_xml(&mut builders));
}

#[test]
fn direct_facade_accessors_are_total() {
    let mut doc = Document::new();
    doc.add_paragraph("first").add_run(" run");

    assert_eq!(doc.paragraph(0).unwrap().text(), "first run");
    assert!(doc.paragraph(1).is_none());

    let mut paragraph = doc.paragraph_mut(0).unwrap();
    assert_eq!(paragraph.run_count(), 2);
    assert_eq!(paragraph.run(1).unwrap().text(), " run");
    assert!(paragraph.run(2).is_none());
    paragraph.run_mut(1).unwrap().set_text(" changed");
    assert!(paragraph.run_mut(2).is_none());

    assert_eq!(doc.paragraph(0).unwrap().text(), "first changed");
}

#[test]
fn immutable_run_accessors_are_total() {
    let mut doc = Document::new();
    doc.add_paragraph("first").add_run(" run");

    let paragraph = doc.paragraph(0).unwrap();
    assert_eq!(paragraph.run_count(), 2);
    assert_eq!(paragraph.run(1).unwrap().text(), " run");
    assert!(paragraph.run(2).is_none());
}

mod header_footer_pdf {
    use std::collections::{BTreeMap, BTreeSet, HashMap};

    use super::*;
    use oxml_layout::{PageFrame, PositionedElement, walk};
    use rdocx_oxml::header_footer::HdrFtrType;

    const DEFAULT_HEADER: &str = "DEFAULTHEADER";
    const DEFAULT_FOOTER: &str = "DEFAULTFOOTER";
    const FIRST_HEADER: &str = "FIRSTHEADER";
    const FIRST_FOOTER: &str = "FIRSTFOOTER";
    const EVEN_HEADER: &str = "EVENHEADER";
    const EVEN_FOOTER: &str = "EVENFOOTER";
    const STORY_MARKERS: [&str; 6] = [
        DEFAULT_HEADER,
        DEFAULT_FOOTER,
        FIRST_HEADER,
        FIRST_FOOTER,
        EVEN_HEADER,
        EVEN_FOOTER,
    ];
    const PRODUCER_PART: &[u8] = b"producer-private-bytes\x00\xff";
    const PRODUCER_REL_TYPE: &str = "urn:producer:relationships/private-state";
    const PRODUCER_XML: &[u8] = b"<x:opaque x:flag=\"keep\"><x:child/></x:opaque>";

    struct Fixture {
        document: Document,
        saved: Vec<u8>,
    }

    struct PdfObject {
        body: Vec<u8>,
        stream: Option<Vec<u8>>,
    }

    fn find_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
        haystack
            .windows(needle.len())
            .position(|window| window == needle)
    }

    fn parse_decimal(bytes: &[u8], start: usize) -> (usize, usize) {
        let mut end = start;
        while end < bytes.len() && bytes[end].is_ascii_digit() {
            end += 1;
        }
        assert!(end > start, "expected decimal at byte {start}");
        let value = std::str::from_utf8(&bytes[start..end])
            .unwrap()
            .parse()
            .unwrap();
        (value, end)
    }

    fn reference_after(bytes: &[u8], key: &[u8]) -> u32 {
        let start = find_bytes(bytes, key)
            .unwrap_or_else(|| panic!("missing PDF key {}", String::from_utf8_lossy(key)))
            + key.len();
        let mut number_start = start;
        while number_start < bytes.len() && bytes[number_start].is_ascii_whitespace() {
            number_start += 1;
        }
        let (number, end) = parse_decimal(bytes, number_start);
        assert!(bytes[end..].starts_with(b" 0 R"));
        number as u32
    }

    fn references(bytes: &[u8]) -> Vec<u32> {
        let mut refs = Vec::new();
        let mut cursor = 0;
        while cursor < bytes.len() {
            let Some(relative) = find_bytes(&bytes[cursor..], b" 0 R") else {
                break;
            };
            let marker = cursor + relative;
            let mut start = marker;
            while start > 0 && bytes[start - 1].is_ascii_digit() {
                start -= 1;
            }
            if start < marker {
                refs.push(parse_decimal(bytes, start).0 as u32);
            }
            cursor = marker + 4;
        }
        refs
    }

    fn structure_role(bytes: &[u8]) -> Option<&[u8]> {
        let start = find_bytes(bytes, b"/S /")? + b"/S /".len();
        let end = bytes[start..]
            .iter()
            .position(|byte| byte.is_ascii_whitespace())?;
        Some(&bytes[start..start + end])
    }

    fn structure_children(bytes: &[u8]) -> Vec<u32> {
        let Some(start) = find_bytes(bytes, b"/K [").map(|start| start + b"/K [".len()) else {
            return Vec::new();
        };
        let end = start + find_bytes(&bytes[start..], b"]").expect("structure child terminator");
        references(&bytes[start..end])
    }

    fn has_nested_list_depth(
        objects: &BTreeMap<u32, PdfObject>,
        list_ref: u32,
        depth: usize,
    ) -> bool {
        let list = &objects[&list_ref].body;
        if structure_role(list) != Some(b"L") {
            return false;
        }
        if depth == 1 {
            return true;
        }
        structure_children(list)
            .into_iter()
            .filter(|child| structure_role(&objects[child].body) == Some(b"LI"))
            .flat_map(|item| structure_children(&objects[&item].body))
            .filter(|child| structure_role(&objects[child].body) == Some(b"LBody"))
            .flat_map(|body| structure_children(&objects[&body].body))
            .any(|child| has_nested_list_depth(objects, child, depth - 1))
    }

    fn number_tree_entries(bytes: &[u8]) -> Vec<(usize, u32)> {
        let start = find_bytes(bytes, b"/Nums [").expect("parent number tree") + b"/Nums [".len();
        let end = start + find_bytes(&bytes[start..], b"]").expect("parent number tree end");
        let entries = &bytes[start..end];
        let mut cursor = 0;
        let mut result = Vec::new();
        while cursor < entries.len() {
            while cursor < entries.len() && entries[cursor].is_ascii_whitespace() {
                cursor += 1;
            }
            if cursor == entries.len() {
                break;
            }
            let (key, key_end) = parse_decimal(entries, cursor);
            cursor = key_end;
            while cursor < entries.len() && entries[cursor].is_ascii_whitespace() {
                cursor += 1;
            }
            let (reference, reference_end) = parse_decimal(entries, cursor);
            assert!(entries[reference_end..].starts_with(b" 0 R"));
            result.push((key, reference as u32));
            cursor = reference_end + b" 0 R".len();
        }
        result
    }

    fn marked_content_ids(bytes: &[u8]) -> Vec<usize> {
        let mut result = Vec::new();
        let mut cursor = 0;
        while let Some(relative) = find_bytes(&bytes[cursor..], b"/MCID ") {
            let start = cursor + relative + b"/MCID ".len();
            let (mcid, end) = parse_decimal(bytes, start);
            result.push(mcid);
            cursor = end;
        }
        result
    }

    fn marked_content_references(bytes: &[u8]) -> Vec<(u32, usize)> {
        let mut result = Vec::new();
        let mut cursor = 0;
        while let Some(relative) = find_bytes(&bytes[cursor..], b"/Type /MCR") {
            let start = cursor + relative;
            let body = &bytes[start..];
            let page = reference_after(body, b"/Pg");
            let mcid_start = find_bytes(body, b"/MCID ").expect("MCR MCID") + b"/MCID ".len();
            let (mcid, _) = parse_decimal(body, mcid_start);
            result.push((page, mcid));
            cursor = start + b"/Type /MCR".len();
        }
        result
    }

    fn pdf_objects(pdf: &[u8]) -> BTreeMap<u32, PdfObject> {
        let mut objects = BTreeMap::new();
        let mut cursor = 0;
        while let Some(relative_marker) = find_bytes(&pdf[cursor..], b" 0 obj") {
            let marker = cursor + relative_marker;
            let mut number_start = marker;
            while number_start > 0 && pdf[number_start - 1].is_ascii_digit() {
                number_start -= 1;
            }
            let number = parse_decimal(pdf, number_start).0 as u32;
            let content_start = marker + b" 0 obj".len();
            let endobj = content_start
                + find_bytes(&pdf[content_start..], b"endobj").expect("PDF object terminator");
            let stream_marker = find_bytes(&pdf[content_start..endobj], b"\nstream\n")
                .map(|relative| content_start + relative);

            let (body, stream, object_end) = if let Some(stream_marker) = stream_marker {
                let body = pdf[content_start..stream_marker].to_vec();
                let length_start =
                    find_bytes(&body, b"/Length ").expect("PDF stream length") + b"/Length ".len();
                let length = parse_decimal(&body, length_start).0;
                let data_start = stream_marker + b"\nstream\n".len();
                let data_end = data_start + length;
                assert!(pdf[data_end..].starts_with(b"\nendstream"));
                let object_end = data_end
                    + find_bytes(&pdf[data_end..], b"endobj").expect("stream object terminator")
                    + b"endobj".len();
                (body, Some(pdf[data_start..data_end].to_vec()), object_end)
            } else {
                (
                    pdf[content_start..endobj].to_vec(),
                    None,
                    endobj + b"endobj".len(),
                )
            };

            assert!(objects.insert(number, PdfObject { body, stream }).is_none());
            cursor = object_end;
        }
        assert!(
            !objects.is_empty(),
            "deterministic renderer returned no PDF objects"
        );
        objects
    }

    fn decoded_stream(object: &PdfObject) -> Vec<u8> {
        let stream = object.stream.as_ref().expect("expected a PDF stream");
        if find_bytes(&object.body, b"/Filter /FlateDecode").is_some() {
            miniz_oxide::inflate::decompress_to_vec_zlib(stream)
                .expect("deterministic PDF stream should inflate")
        } else {
            stream.clone()
        }
    }

    fn parse_hex(value: &[u8]) -> Vec<u8> {
        assert!(value.len().is_multiple_of(2));
        value
            .chunks_exact(2)
            .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
            .collect()
    }

    fn pdf_literal(line: &[u8], open: usize) -> (Vec<u8>, usize) {
        let mut bytes = Vec::new();
        let mut cursor = open + 1;
        while cursor < line.len() {
            match line[cursor] {
                b')' => return (bytes, cursor + 1),
                b'\\' => {
                    cursor += 1;
                    assert!(cursor < line.len(), "truncated PDF string escape");
                    let escaped = line[cursor];
                    if escaped.is_ascii_digit() && escaped < b'8' {
                        let mut value = 0u16;
                        let mut digits = 0;
                        while cursor < line.len()
                            && digits < 3
                            && line[cursor].is_ascii_digit()
                            && line[cursor] < b'8'
                        {
                            value = value * 8 + u16::from(line[cursor] - b'0');
                            cursor += 1;
                            digits += 1;
                        }
                        bytes.push(value as u8);
                        continue;
                    }
                    bytes.push(match escaped {
                        b'n' => b'\n',
                        b'r' => b'\r',
                        b't' => b'\t',
                        b'b' => 8,
                        b'f' => 12,
                        other => other,
                    });
                    cursor += 1;
                }
                byte => {
                    bytes.push(byte);
                    cursor += 1;
                }
            }
        }
        panic!("unterminated PDF literal string")
    }

    fn cmap(object: &PdfObject) -> BTreeMap<u16, String> {
        let bytes = decoded_stream(object);
        let mut mappings = BTreeMap::new();
        let mut in_bfchar = false;
        for line in bytes.split(|byte| *byte == b'\n') {
            if find_bytes(line, b"beginbfchar").is_some() {
                in_bfchar = true;
                continue;
            }
            if find_bytes(line, b"endbfchar").is_some() {
                in_bfchar = false;
                continue;
            }
            if !in_bfchar {
                continue;
            }
            let fields = line
                .split(|byte| byte.is_ascii_whitespace())
                .collect::<Vec<_>>();
            if fields.len() != 2 {
                continue;
            }
            let glyph_bytes = parse_hex(&fields[0][1..fields[0].len() - 1]);
            let unicode_bytes = parse_hex(&fields[1][1..fields[1].len() - 1]);
            let glyph = u16::from_be_bytes([glyph_bytes[0], glyph_bytes[1]]);
            let utf16 = unicode_bytes
                .chunks_exact(2)
                .map(|pair| u16::from_be_bytes([pair[0], pair[1]]))
                .collect::<Vec<_>>();
            let text = char::decode_utf16(utf16)
                .collect::<Result<String, _>>()
                .expect("valid ToUnicode mapping");
            mappings.insert(glyph, text);
        }
        assert!(!mappings.is_empty(), "missing ToUnicode mappings");
        mappings
    }

    fn font_references(page: &[u8]) -> HashMap<String, u32> {
        let start = find_bytes(page, b"/Font <<").expect("page font resources") + b"/Font <<".len();
        let end = start + find_bytes(&page[start..], b">>").expect("page font dictionary");
        let resources = &page[start..end];
        let mut fonts = HashMap::new();
        let mut cursor = 0;
        while let Some(relative) = find_bytes(&resources[cursor..], b"/F") {
            let name_start = cursor + relative + 1;
            let mut name_end = name_start + 1;
            while name_end < resources.len() && resources[name_end].is_ascii_digit() {
                name_end += 1;
            }
            if name_end == name_start + 1 {
                cursor = name_end;
                continue;
            }
            let name = std::str::from_utf8(&resources[name_start..name_end])
                .unwrap()
                .to_owned();
            let mut ref_start = name_end;
            while ref_start < resources.len() && resources[ref_start].is_ascii_whitespace() {
                ref_start += 1;
            }
            fonts.insert(name, parse_decimal(resources, ref_start).0 as u32);
            cursor = name_end;
        }
        assert!(!fonts.is_empty(), "page has no font resources");
        fonts
    }

    fn content_text(content: &[u8], cmaps: &HashMap<String, BTreeMap<u16, String>>) -> String {
        let mut current_font = None;
        let mut text = String::new();
        let mut inside_actual_text = false;
        for line in content.split(|byte| *byte == b'\n') {
            let line = line
                .iter()
                .copied()
                .skip_while(|byte| byte.is_ascii_whitespace())
                .collect::<Vec<_>>();
            if let Some(actual_text) = find_bytes(&line, b"/ActualText <") {
                let open = actual_text + b"/ActualText <".len();
                let close = open
                    + line[open..]
                        .iter()
                        .position(|byte| *byte == b'>')
                        .expect("ActualText hex string");
                let bytes = parse_hex(&line[open..close]);
                let utf16 = bytes
                    .chunks_exact(2)
                    .map(|pair| u16::from_be_bytes([pair[0], pair[1]]))
                    .skip_while(|unit| *unit == 0xfeff)
                    .collect::<Vec<_>>();
                text.push_str(
                    &char::decode_utf16(utf16)
                        .collect::<Result<String, _>>()
                        .expect("valid ActualText"),
                );
                inside_actual_text = true;
                continue;
            }
            if inside_actual_text {
                if line == b"EMC" {
                    inside_actual_text = false;
                }
                continue;
            }
            if line.ends_with(b" Tf") && line.starts_with(b"/F") {
                let end = line
                    .iter()
                    .position(|byte| byte.is_ascii_whitespace())
                    .expect("font operand");
                current_font = Some(String::from_utf8(line[1..end].to_vec()).unwrap());
                continue;
            }
            if !line.ends_with(b" TJ") {
                continue;
            }
            let font = current_font.as_ref().expect("TJ without selected font");
            let mapping = cmaps.get(font).expect("ToUnicode map for selected font");
            let mut cursor = 0;
            while cursor < line.len() {
                let hex = line[cursor..].iter().position(|byte| *byte == b'<');
                let literal = line[cursor..].iter().position(|byte| *byte == b'(');
                let (glyph_bytes, next) = match (hex, literal) {
                    (None, None) => break,
                    (Some(hex), Some(literal)) if literal < hex => {
                        pdf_literal(&line, cursor + literal)
                    }
                    (_, Some(literal)) if hex.is_none() => pdf_literal(&line, cursor + literal),
                    (Some(hex), _) => {
                        let open = cursor + hex;
                        let close = open
                            + line[open..]
                                .iter()
                                .position(|byte| *byte == b'>')
                                .expect("hex glyph string");
                        (parse_hex(&line[open + 1..close]), close + 1)
                    }
                    _ => unreachable!(),
                };
                for pair in glyph_bytes.chunks_exact(2) {
                    let glyph = u16::from_be_bytes([pair[0], pair[1]]);
                    text.push_str(mapping.get(&glyph).expect("mapped PDF glyph"));
                }
                cursor = next;
            }
        }
        text
    }

    fn pdf_page_text(pdf: &[u8]) -> Vec<String> {
        let objects = pdf_objects(pdf);
        let pages_object = objects
            .values()
            .find(|object| {
                find_bytes(&object.body, b"/Type /Pages").is_some()
                    && find_bytes(&object.body, b"/Kids [").is_some()
            })
            .expect("PDF page tree");
        let kids_start = find_bytes(&pages_object.body, b"/Kids [").unwrap() + b"/Kids [".len();
        let kids_end = kids_start
            + find_bytes(&pages_object.body[kids_start..], b"]").expect("page tree kids");
        let page_refs = references(&pages_object.body[kids_start..kids_end]);
        let first_page = &objects[&page_refs[0]].body;
        let font_refs = font_references(first_page);
        let cmaps = font_refs
            .into_iter()
            .map(|(name, font_ref)| {
                let cmap_ref = reference_after(&objects[&font_ref].body, b"/ToUnicode");
                (name, cmap(&objects[&cmap_ref]))
            })
            .collect::<HashMap<_, _>>();

        page_refs
            .into_iter()
            .map(|page_ref| {
                let content_ref = reference_after(&objects[&page_ref].body, b"/Contents");
                content_text(&decoded_stream(&objects[&content_ref]), &cmaps)
            })
            .collect()
    }

    #[test]
    fn hybrid_rtl_hyphenation_keeps_pdf_and_svg_logical_text_with_visual_origins() {
        let mut seed = Document::new();
        seed.add_paragraph("seed");
        let mut package = OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap()))
            .expect("seed package opens");
        package.set_part(
            "/word/document.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/><w:lang w:bidi="ar-SA"/></w:rPr><w:t>العربية</w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:rtl/></w:rPr><w:t>99</w:t></w:r></w:fldSimple><w:r><w:rPr><w:rtl w:val="0"/><w:lang w:val="en-US"/><w:sz w:val="72"/></w:rPr><w:t>representation</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="3600" w:h="15840"/><w:pgMar w:top="720" w:right="360" w:bottom="720" w:left="360"/></w:sectPr></w:body></w:document>"#.as_bytes().to_vec(),
        );
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let mut document = Document::from_bytes(bytes.get_ref()).expect("hybrid document opens");
        document.set_auto_hyphenation(true).unwrap();

        let layout = document.layout_deterministic().unwrap();
        let mut arabic_x = None;
        let mut english = None::<(String, f64)>;
        let mut has_conditional_hyphen = false;
        for page in &layout.layout.pages {
            walk(&page.elements, &mut |element, _| match element {
                PositionedElement::MultilingualText(run) if run.logical_text == "العربية" => {
                    arabic_x = Some(run.origin.x)
                }
                PositionedElement::Text(run)
                    if run.source.is_some() && "representation".starts_with(&run.text) =>
                {
                    english = Some((run.text.clone(), run.origin.x))
                }
                PositionedElement::Text(run)
                    if run.source.is_none() && run.field_kind.is_none() && run.text == "-" =>
                {
                    has_conditional_hyphen = true
                }
                _ => {}
            });
        }
        let (english_text, english_x) = english.expect("hyphenated English prefix");
        assert!(
            has_conditional_hyphen,
            "the line selects a conditional hyphen"
        );
        assert!(english_x < arabic_x.unwrap(), "paint remains visual RTL");

        let pdf_text = pdf_page_text(&document.to_pdf_deterministic().unwrap());
        let page_text = &pdf_text[0];
        let arabic = page_text
            .find("العربية")
            .unwrap_or_else(|| panic!("missing Arabic PDF extraction in {page_text:?}"));
        let field = page_text
            .find('1')
            .unwrap_or_else(|| panic!("missing PAGE PDF extraction in {page_text:?}"));
        let english = page_text
            .find(&english_text)
            .unwrap_or_else(|| panic!("missing English PDF extraction in {page_text:?}"));
        let hyphen = english
            + page_text[english..]
                .find('-')
                .unwrap_or_else(|| panic!("missing conditional hyphen in {page_text:?}"));
        assert!(
            arabic < field && field < english && english < hyphen,
            "PDF extraction remains logical: {page_text:?}"
        );

        let svg = document
            .render_page_to_svg_deterministic(0)
            .unwrap()
            .expect("first SVG page");
        let arabic = svg.svg.find("العربية</text>").unwrap();
        let field = svg.svg.find(">1</text>").unwrap();
        let english = svg.svg.find(&format!(">{english_text}</text>")).unwrap();
        let hyphen = svg.svg[english..].find(">-</text>").unwrap() + english;
        assert!(
            arabic < field && field < english && english < hyphen,
            "SVG DOM text remains logical"
        );
    }

    #[test]
    fn source_less_stored_field_keeps_logical_pdf_and_svg_order() {
        let mut seed = Document::new();
        seed.add_paragraph("seed");
        let mut package = OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap()))
            .expect("seed package opens");
        package.set_part(
            "/word/document.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/><w:lang w:bidi="he-IL"/></w:rPr><w:t>אבג</w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:rtl/></w:rPr><w:t>99</w:t></w:r></w:fldSimple><w:r><w:rPr><w:rtl w:val="0"/><w:lang w:val="en-US"/></w:rPr><w:t>ABC</w:t></w:r></w:p></w:body></w:document>"#.as_bytes().to_vec(),
        );
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let document = Document::from_bytes(bytes.get_ref()).expect("field document opens");

        let pdf_text = pdf_page_text(&document.to_pdf_deterministic().unwrap());
        let page_text = &pdf_text[0];
        let hebrew = page_text
            .find("אבג")
            .unwrap_or_else(|| panic!("missing Hebrew PDF extraction in {page_text:?}"));
        let page = page_text
            .find('1')
            .unwrap_or_else(|| panic!("missing PAGE PDF extraction in {page_text:?}"));
        let english = page_text
            .find("ABC")
            .unwrap_or_else(|| panic!("missing English PDF extraction in {page_text:?}"));
        assert!(
            hebrew < page && page < english,
            "PDF extraction keeps the logical field position: {page_text:?}"
        );

        let svg = document
            .render_page_to_svg_deterministic(0)
            .unwrap()
            .expect("first SVG page");
        let hebrew = svg.svg.find("אבג</text>").unwrap();
        let page = svg.svg.find(">1</text>").unwrap();
        let english = svg.svg.find("ABC</text>").unwrap();
        assert!(
            hebrew < page && page < english,
            "SVG DOM keeps the logical field position"
        );
    }

    #[test]
    fn mixed_bidi_footnote_and_endnote_keep_logical_pdf_and_svg_order() {
        let mut seed = Document::new();
        seed.add_paragraph("seed");
        let mut package = OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap()))
            .expect("seed package opens");
        let (footnote_rel, endnote_rel) = {
            let relationships = package.get_or_create_part_rels("/word/document.xml");
            (
                relationships.add(
                    oxml_opc::relationship::rel_types::FOOTNOTES,
                    "footnotes.xml",
                ),
                relationships.add(oxml_opc::relationship::rel_types::ENDNOTES, "endnotes.xml"),
            )
        };
        assert!(!footnote_rel.is_empty() && !endnote_rel.is_empty());
        package.set_part(
            "/word/document.xml",
            br#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>BODY</w:t><w:footnoteReference w:id="2"/><w:endnoteReference w:id="3"/></w:r></w:p></w:body></w:document>"#.to_vec(),
        );
        package.set_part(
            "/word/footnotes.xml",
            r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="2"><w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/><w:lang w:bidi="ar-SA"/></w:rPr><w:t>قدم</w:t></w:r><w:fldSimple w:instr="MERGEFIELD Value"><w:r><w:rPr><w:rtl/></w:rPr><w:t>FOOTFIELD</w:t></w:r></w:fldSimple><w:r><w:rPr><w:rtl w:val="0"/><w:lang w:val="en-US"/></w:rPr><w:t>FOOTTAIL</w:t></w:r></w:p></w:footnote></w:footnotes>"#.as_bytes().to_vec(),
        );
        package.set_part(
            "/word/endnotes.xml",
            r#"<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="3"><w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/><w:lang w:bidi="he-IL"/></w:rPr><w:t>סוף</w:t></w:r><w:fldSimple w:instr="MERGEFIELD Value"><w:r><w:rPr><w:rtl/></w:rPr><w:t>ENDFIELD</w:t></w:r></w:fldSimple><w:r><w:rPr><w:rtl w:val="0"/><w:lang w:val="en-US"/></w:rPr><w:t>ENDTAIL</w:t></w:r></w:p></w:endnote></w:endnotes>"#.as_bytes().to_vec(),
        );
        package.content_types.add_override(
            "/word/footnotes.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        );
        package.content_types.add_override(
            "/word/endnotes.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
        );
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let document = Document::from_bytes(bytes.get_ref()).expect("note document opens");

        let pdf_text = pdf_page_text(&document.to_pdf_deterministic().unwrap()).join("\n");
        for expected in [
            ("قدم", "FOOTFIELD", "FOOTTAIL"),
            ("סוף", "ENDFIELD", "ENDTAIL"),
        ] {
            let first = pdf_text.find(expected.0).unwrap();
            let second = pdf_text.find(expected.1).unwrap();
            let third = pdf_text.find(expected.2).unwrap();
            assert!(
                first < second && second < third,
                "PDF note extraction stays logical: {pdf_text:?}"
            );
        }

        let svg = (0..document.layout_deterministic().unwrap().layout.pages.len())
            .filter_map(|page| document.render_page_to_svg_deterministic(page).unwrap())
            .map(|page| page.svg)
            .collect::<String>();
        for expected in [
            ("قدم", "FOOTFIELD", "FOOTTAIL"),
            ("סוף", "ENDFIELD", "ENDTAIL"),
        ] {
            let first = svg.find(expected.0).unwrap();
            let second = svg.find(expected.1).unwrap();
            let third = svg.find(expected.2).unwrap();
            assert!(
                first < second && second < third,
                "SVG note extraction stays logical"
            );
        }
    }

    #[test]
    fn source_less_numbering_marker_keeps_logical_pdf_and_svg_order() {
        let mut seed = Document::new();
        seed.add_numbered_list_item("seed", 0);
        let mut package = OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap()))
            .expect("seed package opens");
        package.set_part(
            "/word/document.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl w:val="0"/><w:lang w:val="en-US"/></w:rPr><w:t xml:space="preserve">123 </w:t></w:r><w:r><w:rPr><w:rtl/><w:lang w:bidi="ar-SA"/></w:rPr><w:t>العربية</w:t></w:r></w:p></w:body></w:document>"#.as_bytes().to_vec(),
        );
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let document = Document::from_bytes(bytes.get_ref()).expect("numbered document opens");

        let pdf_text = pdf_page_text(&document.to_pdf_deterministic().unwrap());
        let page_text = &pdf_text[0];
        let marker = page_text
            .find("1.")
            .unwrap_or_else(|| panic!("missing numbering marker in {page_text:?}"));
        let number = page_text
            .find("123")
            .unwrap_or_else(|| panic!("missing number in {page_text:?}"));
        let arabic = page_text
            .find("العربية")
            .unwrap_or_else(|| panic!("missing Arabic in {page_text:?}"));
        assert!(
            marker < number && number < arabic,
            "PDF extraction keeps the logical marker prefix: {page_text:?}"
        );

        let svg = document
            .render_page_to_svg_deterministic(0)
            .unwrap()
            .expect("first SVG page");
        let marker = svg.svg.find(">1.</text>").unwrap();
        let number = svg.svg.find("123").unwrap();
        let arabic = svg.svg.find("العربية</text>").unwrap();
        assert!(
            marker < number && number < arabic,
            "SVG DOM keeps the logical marker prefix"
        );
    }

    #[test]
    fn word_semantics_reach_owned_multi_page_pdf_structure() {
        fn count_bytes(haystack: &[u8], needle: &[u8]) -> usize {
            haystack
                .windows(needle.len())
                .filter(|window| *window == needle)
                .count()
        }

        let mut document = Document::new();
        for level in 1..=6 {
            document
                .add_paragraph(&format!("Heading level {level}"))
                .style(&format!("Heading{level}"));
        }
        document.add_bullet_list_item("Outer list item", 0);
        document.add_bullet_list_item("Nested list item", 1);
        document.add_bullet_list_item("Deeply nested list item", 2);
        document.add_bullet_list_item("Second outer list item", 0);
        {
            let mut table = document.add_table(120, 2);
            table.row(0).unwrap().header();
            table.cell(0, 0).unwrap().set_text("Name");
            table.cell(0, 1).unwrap().set_text("Value");
            for row in 1..120 {
                table.cell(row, 0).unwrap().set_text(&format!("Item {row}"));
                table
                    .cell(row, 1)
                    .unwrap()
                    .set_text(&format!("Value {row}"));
            }
        }

        let pdf = document
            .to_pdf_deterministic()
            .expect("Word document should render deterministically");
        let objects = pdf_objects(&pdf);
        let pages_object = objects
            .values()
            .find(|object| {
                find_bytes(&object.body, b"/Type /Pages").is_some()
                    && find_bytes(&object.body, b"/Kids [").is_some()
            })
            .expect("PDF page tree");
        let kids_start = find_bytes(&pages_object.body, b"/Kids [").unwrap() + b"/Kids [".len();
        let kids_end = kids_start
            + find_bytes(&pages_object.body[kids_start..], b"]").expect("page tree kids");
        let page_refs = references(&pages_object.body[kids_start..kids_end]);
        assert!(
            page_refs.len() > 1,
            "fixture must exercise repeated headers"
        );

        let structure_root = objects
            .values()
            .find(|object| find_bytes(&object.body, b"/Type /StructTreeRoot").is_some())
            .expect("PDF structure root");
        let parent_tree = number_tree_entries(&structure_root.body);
        assert_eq!(parent_tree.len(), page_refs.len());
        let all_mcrs = objects
            .iter()
            .flat_map(|(owner, object)| {
                marked_content_references(&object.body)
                    .into_iter()
                    .map(|(page, mcid)| (*owner, page, mcid))
            })
            .collect::<Vec<_>>();
        let mut expected_mcrs = BTreeSet::new();
        let mut struct_parent_keys = BTreeSet::new();
        for (page_index, page_ref) in page_refs.iter().copied().enumerate() {
            let page = &objects[&page_ref].body;
            let key_start = find_bytes(page, b"/StructParents ")
                .expect("every content-owning page needs a parent-tree key")
                + b"/StructParents ".len();
            let struct_parent = parse_decimal(page, key_start).0;
            assert!(
                struct_parent_keys.insert(struct_parent),
                "StructParents keys must be page-unique"
            );
            assert_eq!(
                struct_parent, page_index,
                "deterministic StructParents keys follow page order"
            );
            let parent_array_ref = parent_tree
                .iter()
                .find_map(|(key, reference)| (*key == struct_parent).then_some(*reference))
                .expect("StructParents key must resolve through ParentTree");
            let parent_refs = references(&objects[&parent_array_ref].body);
            let content_ref = reference_after(page, b"/Contents");
            let content = decoded_stream(&objects[&content_ref]);
            let content_mcids = marked_content_ids(&content);
            assert_eq!(
                content_mcids,
                (0..content_mcids.len()).collect::<Vec<_>>(),
                "MCIDs must be page-local and contiguous"
            );
            assert_eq!(
                parent_refs.len(),
                content_mcids.len(),
                "ParentTree array slots must cover every page MCID"
            );
            for (mcid, owner) in parent_refs.into_iter().enumerate() {
                assert!(expected_mcrs.insert((owner, page_ref, mcid)));
                assert_eq!(
                    all_mcrs
                        .iter()
                        .filter(|candidate| **candidate == (owner, page_ref, mcid))
                        .count(),
                    1,
                    "each ParentTree slot must have one matching owner MCR"
                );
            }
        }
        assert_eq!(struct_parent_keys.len(), page_refs.len());
        assert_eq!(
            all_mcrs.len(),
            expected_mcrs.len(),
            "MCR dictionaries and ParentTree slots must have exact ownership"
        );
        assert!(
            all_mcrs.iter().all(|entry| expected_mcrs.contains(entry)),
            "every MCR must point to the matching page-local ParentTree slot"
        );

        let heading_positions = (1..=6)
            .map(|level| {
                let tag = format!("/S /H{level}\n");
                find_bytes(&pdf, tag.as_bytes()).expect("heading structure role")
            })
            .collect::<Vec<_>>();
        assert!(
            heading_positions.windows(2).all(|pair| pair[0] < pair[1]),
            "heading roles must preserve Word source order"
        );
        let list_refs = objects
            .iter()
            .filter_map(|(reference, object)| {
                (structure_role(&object.body) == Some(b"L")).then_some(*reference)
            })
            .collect::<Vec<_>>();
        assert!(
            list_refs
                .iter()
                .any(|reference| has_nested_list_depth(&objects, *reference, 3)),
            "Word levels 0, 1, and 2 must form one three-level list tree"
        );
        assert!(count_bytes(&pdf, b"/S /LI\n") >= 4, "list item nodes");
        assert_eq!(count_bytes(&pdf, b"/S /TH\n"), 2, "logical header cells");

        let header_owned_ranges = objects
            .values()
            .filter(|object| find_bytes(&object.body, b"/S /TH\n").is_some())
            .map(|header| {
                let kids_start =
                    find_bytes(&header.body, b"/K [").expect("TH child array") + b"/K [".len();
                let kids_end = kids_start
                    + find_bytes(&header.body[kids_start..], b"]").expect("TH child terminator");
                let kids = references(&header.body[kids_start..kids_end]);
                assert_eq!(kids.len(), 1, "TH should own one paragraph");
                count_bytes(&objects[&kids[0]].body, b"/Type /MCR")
            })
            .sum::<usize>();
        assert!(
            header_owned_ranges > 2,
            "repeated header paragraph paint should remain below its logical TH"
        );
        assert_eq!(
            expected_mcrs.len(),
            count_bytes(&pdf, b"/Type /MCR"),
            "each semantic content range belongs to exactly one structure node"
        );
    }

    fn header_footer_xml(root: &str, text: &str) -> Vec<u8> {
        format!(
            r#"<w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:{root}>"#
        )
        .into_bytes()
    }

    fn relationship_id(package: &OpcPackage, rel_type: &str, part_name: &str) -> String {
        package
            .get_part_rels("/word/document.xml")
            .unwrap()
            .items
            .iter()
            .find(|relationship| {
                relationship.rel_type == rel_type
                    && OpcPackage::resolve_rel_target("/word/document.xml", &relationship.target)
                        == part_name
            })
            .unwrap_or_else(|| panic!("missing relationship for {part_name}"))
            .id
            .clone()
    }

    fn section_references(package: &OpcPackage) -> String {
        let default_header = relationship_id(package, rel_types::HEADER, "/word/header1.xml");
        let first_header = relationship_id(package, rel_types::HEADER, "/word/headerFirst1.xml");
        let even_header = relationship_id(package, rel_types::HEADER, "/word/headerEven1.xml");
        let default_footer = relationship_id(package, rel_types::FOOTER, "/word/footer1.xml");
        let first_footer = relationship_id(package, rel_types::FOOTER, "/word/footerFirst1.xml");
        let even_footer = relationship_id(package, rel_types::FOOTER, "/word/footerEven1.xml");
        format!(
            r#"<w:headerReference w:type="default" r:id="{default_header}"/><w:headerReference w:type="first" r:id="{first_header}"/><w:headerReference w:type="even" r:id="{even_header}"/><w:footerReference w:type="default" r:id="{default_footer}"/><w:footerReference w:type="first" r:id="{first_footer}"/><w:footerReference w:type="even" r:id="{even_footer}"/>"#
        )
    }

    fn page_paragraph(text: &str, page_break_before: bool, section: Option<&str>) -> String {
        let page_break = if page_break_before {
            "<w:pageBreakBefore/>"
        } else {
            ""
        };
        let section = section.unwrap_or("");
        format!("<w:p><w:pPr>{page_break}{section}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>")
    }

    fn enable_even_headers(package: &mut OpcPackage) {
        package.set_part(
            "/word/settings.xml",
            br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:evenAndOddHeaders/></w:settings>"#.to_vec(),
        );
        package.content_types.add_override(
            "/word/settings.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        );
        let relationships = package.get_or_create_part_rels("/word/document.xml");
        if relationships.get_by_type(rel_types::SETTINGS).is_none() {
            relationships.add(rel_types::SETTINGS, "settings.xml");
        }
    }

    fn save_and_reopen(package: OpcPackage) -> Fixture {
        let mut staged = std::io::Cursor::new(Vec::new());
        package.write_to(&mut staged).unwrap();
        let mut document = Document::from_bytes(staged.get_ref()).unwrap();
        let saved = document.to_bytes().unwrap();
        let document = Document::from_bytes(&saved).unwrap();
        Fixture { document, saved }
    }

    fn full_fixture() -> Fixture {
        let mut authored = Document::new();
        authored.set_header(DEFAULT_HEADER);
        authored.set_footer(DEFAULT_FOOTER);
        authored.set_first_page_header(FIRST_HEADER);
        authored.set_first_page_footer(FIRST_FOOTER);
        authored.set_raw_header_with_images(
            header_footer_xml("hdr", EVEN_HEADER),
            &[],
            HdrFtrType::Even,
        );
        authored.set_raw_footer_with_images(
            header_footer_xml("ftr", EVEN_FOOTER),
            &[],
            HdrFtrType::Even,
        );
        authored.add_paragraph("authoring seed");

        let bytes = authored.to_bytes().unwrap();
        let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        enable_even_headers(&mut package);
        let references = section_references(&package);
        let first_section = format!(
            r#"<w:sectPr>{references}<w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/><w:titlePg/></w:sectPr>"#
        );
        let final_section = r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/><w:titlePg/></w:sectPr>"#;
        let document_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:producer"><w:body>{}{}{}{}{}{}{}{}</w:body></w:document>"#,
            page_paragraph("SECTIONONEFIRSTBODY", false, None),
            page_paragraph("SECTIONONEEVENBODY", true, None),
            page_paragraph("SECTIONONEDEFAULTBODY", true, Some(&first_section)),
            page_paragraph("SECTIONTWOFIRSTBODY", false, None),
            PRODUCER_XML
                .iter()
                .map(|byte| *byte as char)
                .collect::<String>(),
            page_paragraph("SECTIONTWODEFAULTBODY", true, None),
            page_paragraph("SECTIONTWOEVENBODY", true, None),
            final_section,
        );
        package.set_part("/word/document.xml", document_xml.into_bytes());
        package.set_part("/custom/producer.bin", PRODUCER_PART.to_vec());
        package
            .content_types
            .add_override("/custom/producer.bin", "application/x-producer-private");
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(PRODUCER_REL_TYPE, "../custom/producer.bin");
        save_and_reopen(package)
    }

    fn blank_fixture() -> Fixture {
        let mut authored = Document::new();
        authored.set_header(DEFAULT_HEADER);
        authored.set_footer(DEFAULT_FOOTER);
        authored.set_first_page_header("");
        authored.set_first_page_footer("");
        authored.set_raw_header_with_images(header_footer_xml("hdr", ""), &[], HdrFtrType::Even);
        authored.set_raw_footer_with_images(header_footer_xml("ftr", ""), &[], HdrFtrType::Even);

        let bytes = authored.to_bytes().unwrap();
        let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        enable_even_headers(&mut package);
        let references = section_references(&package);
        let final_section = format!(
            r#"<w:sectPr>{references}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/><w:titlePg/></w:sectPr>"#
        );
        let document_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{}{}{}</w:body></w:document>"#,
            page_paragraph("BLANKFIRSTBODY", false, None),
            page_paragraph("BLANKEVENBODY", true, None),
            final_section,
        );
        package.set_part("/word/document.xml", document_xml.into_bytes());
        save_and_reopen(package)
    }

    fn page_text(page: &PageFrame) -> String {
        let mut text = String::new();
        walk(&page.elements, &mut |element, _| {
            if let PositionedElement::Text(run) = element {
                text.push_str(&run.text);
            }
        });
        text
    }

    fn marker_y(page: &PageFrame, marker: &str) -> f64 {
        let mut ys = Vec::new();
        walk(&page.elements, &mut |element, _| {
            if let PositionedElement::Text(run) = element
                && run.text == marker
            {
                ys.push(run.origin.y);
            }
        });
        assert_eq!(ys.len(), 1, "expected one positioned {marker}");
        ys[0]
    }

    fn assert_story_selection(text: &str, header: &str, body: &str, footer: &str) {
        assert!(text.contains(header), "missing {header} in {text:?}");
        assert!(text.contains(body), "missing {body} in {text:?}");
        assert!(text.contains(footer), "missing {footer} in {text:?}");
        for marker in STORY_MARKERS {
            assert_eq!(
                text.matches(marker).count(),
                usize::from(marker == header || marker == footer),
                "unexpected {marker} in {text:?}"
            );
        }
    }

    #[test]
    fn authored_reopened_headers_and_footers_reach_pdf() {
        let fixture = full_fixture();
        let layout = fixture.document.layout().unwrap();
        let expected = [
            (FIRST_HEADER, "SECTIONONEFIRSTBODY", FIRST_FOOTER),
            (EVEN_HEADER, "SECTIONONEEVENBODY", EVEN_FOOTER),
            (DEFAULT_HEADER, "SECTIONONEDEFAULTBODY", DEFAULT_FOOTER),
            (FIRST_HEADER, "SECTIONTWOFIRSTBODY", FIRST_FOOTER),
            (DEFAULT_HEADER, "SECTIONTWODEFAULTBODY", DEFAULT_FOOTER),
            (EVEN_HEADER, "SECTIONTWOEVENBODY", EVEN_FOOTER),
        ];
        assert_eq!(
            layout.layout.pages.len(),
            expected.len(),
            "page text: {:?}",
            layout
                .layout
                .pages
                .iter()
                .map(|page| page_text(page))
                .collect::<Vec<_>>()
        );
        for (page, (header, body, footer)) in layout.layout.pages.iter().zip(expected) {
            assert_story_selection(&page_text(page), header, body, footer);
            assert!(marker_y(page, header) < marker_y(page, body));
            assert!(marker_y(page, body) < marker_y(page, footer));
        }

        let pdf_text = pdf_page_text(&fixture.document.to_pdf_deterministic().unwrap());
        assert_eq!(pdf_text.len(), expected.len());
        for (text, (header, body, footer)) in pdf_text.iter().zip(expected) {
            assert_story_selection(text, header, body, footer);
        }
    }

    #[test]
    fn blank_first_and_even_variants_do_not_borrow_defaults() {
        let fixture = blank_fixture();
        let layout = fixture.document.layout().unwrap();
        assert_eq!(
            layout.layout.pages.len(),
            2,
            "page text: {:?}",
            layout
                .layout
                .pages
                .iter()
                .map(|page| page_text(page))
                .collect::<Vec<_>>()
        );
        assert_eq!(page_text(&layout.layout.pages[0]), "BLANKFIRSTBODY");
        assert_eq!(page_text(&layout.layout.pages[1]), "BLANKEVENBODY");

        let pdf_text = pdf_page_text(&fixture.document.to_pdf_deterministic().unwrap());
        assert_eq!(pdf_text, ["BLANKFIRSTBODY", "BLANKEVENBODY"]);
    }

    #[test]
    fn header_footer_pdf_fixture_preserves_unrelated_package_state() {
        let fixture = full_fixture();
        let package = OpcPackage::from_reader(std::io::Cursor::new(fixture.saved)).unwrap();
        assert_eq!(
            package.get_part("/custom/producer.bin").unwrap(),
            PRODUCER_PART
        );
        assert_eq!(
            package.content_types.override_for("/custom/producer.bin"),
            Some("application/x-producer-private")
        );
        let relationship = package
            .get_part_rels("/word/document.xml")
            .unwrap()
            .items
            .iter()
            .find(|relationship| relationship.rel_type == PRODUCER_REL_TYPE)
            .expect("preserved producer relationship");
        assert_eq!(relationship.target, "../custom/producer.bin");
        assert_eq!(relationship.target_mode, None);
        let document_xml = package.get_part("/word/document.xml").unwrap();
        assert!(
            document_xml
                .windows(PRODUCER_XML.len())
                .any(|window| window == PRODUCER_XML),
            "unmodelled producer subtree was not preserved byte for byte"
        );
    }
}

mod legacy_forms_and_building_blocks {
    use oxml_opc::OpcPackage;
    use oxml_opc::relationship::rel_types;
    use rdocx::{BuildingBlockKind, Document, LegacyFormFieldKind, LegacyFormFieldValue};

    const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const GLOSSARY_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";
    const HEADER_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    const FOOTER_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    const FOOTNOTES_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
    const ENDNOTES_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";

    fn package_bytes(package: OpcPackage) -> Vec<u8> {
        let mut output = std::io::Cursor::new(Vec::new());
        package.write_to(&mut output).expect("package writes");
        output.into_inner()
    }

    fn base_package() -> OpcPackage {
        let mut document = Document::new();
        let bytes = document.to_bytes().expect("base document");
        OpcPackage::from_reader(std::io::Cursor::new(bytes)).expect("base package")
    }

    fn story_content_type(relationship_type: &str) -> &'static str {
        match relationship_type {
            rel_types::HEADER => HEADER_CONTENT_TYPE,
            rel_types::FOOTER => FOOTER_CONTENT_TYPE,
            rel_types::FOOTNOTES => FOOTNOTES_CONTENT_TYPE,
            rel_types::ENDNOTES => ENDNOTES_CONTENT_TYPE,
            _ => panic!("unsupported story relationship type"),
        }
    }

    fn add_story_content_type(package: &mut OpcPackage, part_name: &str, relationship_type: &str) {
        package
            .content_types
            .add_override(part_name, story_content_type(relationship_type));
    }

    #[test]
    fn legacy_form_fields_round_trip_typed_values_and_preserve_unmodelled_ffdata() {
        let mut package = base_package();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{WORD_NS}" xmlns:q="{WORD_NS}" xmlns:x="urn:producer"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="TextField"/><w:enabled/><w:calcOnExit w:val="0"/><x:producer keep="yes"><x:nested/></x:producer><w:textInput><w:default w:val="old"/><w:maxLength w:val="12"/></w:textInput></w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>old</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><q:p><q:r><q:fldChar q:fldCharType="begin"><q:ffData><q:name q:val="Check"/><q:checkBox><q:sizeAuto/><q:default q:val="0"/><q:checked q:val="1"/></q:checkBox></q:ffData></q:fldChar></q:r><q:r><q:instrText> FORMCHECKBOX </q:instrText></q:r><q:r><q:fldChar q:fldCharType="end"/></q:r></q:p><q:p><q:r><q:fldChar q:fldCharType="begin"><q:ffData><q:name q:val="Choice"/><q:ddList><q:result q:val="0"/><q:listEntry q:val="one"/><q:listEntry q:val="two"/></q:ddList></q:ffData></q:fldChar></q:r><q:r><q:instrText> FORMDROPDOWN </q:instrText></q:r><q:r><q:fldChar q:fldCharType="separate"/></q:r><q:r><q:t>one</q:t></q:r><q:r><q:fldChar q:fldCharType="end"/></q:r></q:p><w:sectPr/></w:body></w:document>"#
        );
        package.set_part("/word/document.xml", xml.into_bytes());

        let mut document = Document::from_bytes(&package_bytes(package)).expect("form document");
        let fields = document.legacy_form_fields().expect("form inventory");
        assert_eq!(fields.len(), 3);
        assert_eq!(fields[0].source_part, "/word/document.xml");
        assert_eq!(fields[0].ordinal, 0);
        assert_eq!(fields[0].kind, LegacyFormFieldKind::TextInput);
        assert_eq!(fields[0].value, LegacyFormFieldValue::Text("old".into()));
        assert_eq!(fields[1].kind, LegacyFormFieldKind::CheckBox);
        assert_eq!(fields[1].value, LegacyFormFieldValue::Checked(true));
        assert_eq!(fields[2].kind, LegacyFormFieldKind::DropDownList);
        assert_eq!(fields[2].choices, ["one", "two"]);

        document
            .set_legacy_form_field_value(
                "/word/document.xml",
                0,
                LegacyFormFieldValue::Text("new".into()),
            )
            .expect("valid text replacement");
        document
            .set_legacy_form_field_value(
                "/word/document.xml",
                2,
                LegacyFormFieldValue::SelectedIndex(1),
            )
            .expect("valid drop-down replacement");
        let saved = document.to_bytes().expect("save form document");
        let reopened = Document::from_bytes(&saved).expect("reopen form document");
        assert_eq!(
            reopened.legacy_form_fields().unwrap()[0].value,
            LegacyFormFieldValue::Text("new".into())
        );
        assert_eq!(
            reopened.legacy_form_fields().unwrap()[2].value,
            LegacyFormFieldValue::SelectedIndex(1)
        );
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(xml.contains(r#"<x:producer keep="yes"><x:nested/></x:producer>"#));
    }

    #[test]
    fn glossary_entries_autotext_and_building_blocks_round_trip() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::GLOSSARY_DOCUMENT, "../custom/glossary.xml");
        package
            .content_types
            .add_override("/custom/glossary.xml", GLOSSARY_CONTENT_TYPE);
        package.set_part(
            "/custom/glossary.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><q:glossaryDocument xmlns:q="{WORD_NS}" xmlns:x="urn:producer"><q:docParts><q:docPart><q:docPartPr><q:name q:val="Greeting"/><q:types><q:type q:val="autoExp"/></q:types><q:description q:val="old description"/><x:property keep="yes"/></q:docPartPr><q:docPartBody><q:p><q:r><q:t>Hello</q:t></q:r></q:p><x:body keep="yes"/></q:docPartBody></q:docPart><q:docPart><q:docPartPr><q:name q:val="Clause"/><q:types><q:type q:val="bbPlcHdr"/></q:types></q:docPartPr><q:docPartBody><q:p><q:r><q:t>Clause body</q:t></q:r></q:p></q:docPartBody></q:docPart></q:docParts></q:glossaryDocument>"#
            )
            .into_bytes(),
        );

        let mut document =
            Document::from_bytes(&package_bytes(package)).expect("glossary document");
        let entries = document
            .building_blocks()
            .expect("building block inventory");
        assert_eq!(entries.len(), 2);
        assert_eq!(entries[0].glossary_part, "/custom/glossary.xml");
        assert_eq!(entries[0].ordinal, 0);
        assert_eq!(entries[0].block.kind, BuildingBlockKind::AutoText);
        assert_eq!(entries[0].block.name, "Greeting");

        let mut replacement = entries[0].block.clone();
        replacement.description = Some("new description".into());
        document
            .replace_building_block("/custom/glossary.xml", 0, replacement)
            .expect("building block replacement");
        let saved = document.to_bytes().expect("save glossary document");
        let reopened = Document::from_bytes(&saved).expect("reopen glossary document");
        assert_eq!(
            reopened.building_blocks().unwrap()[0]
                .block
                .description
                .as_deref(),
            Some("new description")
        );
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        let xml = std::str::from_utf8(package.get_part("/custom/glossary.xml").unwrap()).unwrap();
        assert!(xml.contains(r#"<x:property keep="yes"/>"#));
        assert!(xml.contains(r#"<x:body keep="yes"/>"#));
    }

    #[test]
    fn explicit_internal_glossary_relationship_is_supported() {
        let mut package = base_package();
        let relationships = package.get_or_create_part_rels("/word/document.xml");
        let id = relationships.add(rel_types::GLOSSARY_DOCUMENT, "glossary/document.xml");
        relationships
            .items
            .iter_mut()
            .find(|relationship| relationship.id == id)
            .unwrap()
            .target_mode = Some("Internal".into());
        package
            .content_types
            .add_override("/word/glossary/document.xml", GLOSSARY_CONTENT_TYPE);
        package.set_part(
            "/word/glossary/document.xml",
            format!(r#"<w:glossaryDocument xmlns:w="{WORD_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#).into_bytes(),
        );
        let document = Document::from_bytes(&package_bytes(package)).unwrap();
        assert_eq!(document.building_blocks().unwrap().len(), 1);
    }

    fn text_form_runs(name: &str, value: &str) -> String {
        format!(
            r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="{name}"/><w:textInput><w:default w:val="{value}"/></w:textInput></w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{value}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
        )
    }

    fn text_form(name: &str, value: &str) -> String {
        format!(r#"<w:p>{}</w:p>"#, text_form_runs(name, value))
    }

    fn inline_text_form(name: &str, value: &str) -> String {
        format!(
            r#"<w:p><w:sdt><w:sdtContent>{}</w:sdtContent></w:sdt></w:p>"#,
            text_form_runs(name, value)
        )
    }

    fn nested_inline_form_runs(
        outer_name: Option<&str>,
        nested_name: &str,
        nested_value: &str,
    ) -> String {
        let (outer_data, instruction, tail, outer_result) = if let Some(name) = outer_name {
            (
                format!(
                    r#"<w:ffData><w:name w:val="{name}"/><w:textInput><w:default w:val="outer-old"/></w:textInput></w:ffData>"#
                ),
                " FORMTEXT ",
                " ",
                "outer-old",
            )
        } else {
            (
                String::new(),
                " IF ",
                r#" = &quot;x&quot; &quot;yes&quot; &quot;no&quot; "#,
                "yes",
            )
        };
        format!(
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="begin">{outer_data}</w:fldChar></w:r>"#,
                r#"<w:r><w:instrText xml:space="preserve">{instruction}</w:instrText></w:r>"#,
                "{}",
                r#"<w:r><w:instrText xml:space="preserve">{tail}</w:instrText></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                r#"<w:r><w:t>{outer_result}</w:t></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            ),
            text_form_runs(nested_name, nested_value),
            outer_data = outer_data,
            instruction = instruction,
            tail = tail,
            outer_result = outer_result,
        )
    }

    fn nested_instruction_form_runs() -> String {
        format!(
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                r#"<w:r><w:instrText xml:space="preserve"> TOC \t </w:instrText></w:r>"#,
                "{}",
                r#"<w:r><w:instrText xml:space="preserve"> </w:instrText></w:r>"#,
                "{}",
                r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                r#"<w:r><w:t>result</w:t></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            ),
            text_form_runs("switch-first", "switch-old"),
            text_form_runs("positional-second", "positional-old"),
        )
    }

    fn drop_down_form(name: &str) -> String {
        format!(
            r#"<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="{name}"/><w:ddList><w:result w:val="0"/><w:listEntry w:val="one"/><w:listEntry w:val="two"/></w:ddList></w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>one</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#
        )
    }

    #[test]
    fn legacy_form_field_identity_is_story_part_and_source_ordinal() {
        let mut package = base_package();
        let document_xml = format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="{WORD_NS}"><w:body>{}<w:sectPr/></w:body></w:document>"#,
            text_form("duplicate", "main")
        );
        package.set_part("/word/document.xml", document_xml.into_bytes());
        let stories = [
            (rel_types::HEADER, "/word/header-one.xml", "header", "hdr"),
            (rel_types::FOOTER, "/word/footer-one.xml", "footer", "ftr"),
            (
                rel_types::FOOTNOTES,
                "/word/footnotes-one.xml",
                "footnote",
                "footnotes",
            ),
            (
                rel_types::ENDNOTES,
                "/word/endnotes-one.xml",
                "endnote",
                "endnotes",
            ),
        ];
        for (relationship_type, part_name, value, root) in stories {
            let target = part_name.strip_prefix("/word/").unwrap();
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(relationship_type, target);
            add_story_content_type(&mut package, part_name, relationship_type);
            let content = if matches!(root, "hdr" | "ftr") {
                format!(
                    r#"<?xml version="1.0"?><w:{root} xmlns:w="{WORD_NS}">{}{}</w:{root}>"#,
                    text_form("duplicate", value),
                    if root == "hdr" {
                        text_form("duplicate", "header-2")
                    } else {
                        String::new()
                    },
                )
            } else {
                let item = if root == "footnotes" {
                    "footnote"
                } else {
                    "endnote"
                };
                format!(
                    r#"<?xml version="1.0"?><w:{root} xmlns:w="{WORD_NS}"><w:{item} w:id="2">{}</w:{item}></w:{root}>"#,
                    text_form("duplicate", value),
                )
            };
            package.set_part(part_name, content.into_bytes());
        }
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let fields = document.legacy_form_fields().unwrap();
        assert_eq!(fields.len(), 6);
        assert_eq!(fields[0].source_part, "/word/document.xml");
        let header = fields
            .iter()
            .find(|field| field.source_part == "/word/header-one.xml" && field.ordinal == 1)
            .expect("second duplicate-named header field");
        assert_eq!(header.value, LegacyFormFieldValue::Text("header-2".into()));

        document
            .set_legacy_form_field_value(
                "/word/header-one.xml",
                1,
                LegacyFormFieldValue::Text("changed".into()),
            )
            .unwrap();
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let changed = reopened
            .legacy_form_fields()
            .unwrap()
            .into_iter()
            .find(|field| field.source_part == "/word/header-one.xml" && field.ordinal == 1)
            .unwrap();
        assert_eq!(changed.value, LegacyFormFieldValue::Text("changed".into()));
    }

    #[test]
    fn unsafe_legacy_form_story_relationships_fail_closed() {
        for case in ["unknown-mode", "traversal"] {
            let mut package = base_package();
            let relationships = package.get_or_create_part_rels("/word/document.xml");
            let target = if case == "traversal" {
                "../../outside.xml"
            } else {
                "header.xml"
            };
            let id = relationships.add(rel_types::HEADER, target);
            if case == "unknown-mode" {
                relationships
                    .items
                    .iter_mut()
                    .find(|relationship| relationship.id == id)
                    .unwrap()
                    .target_mode = Some("ProducerDefined".into());
            }
            let part_name = if case == "traversal" {
                "/outside.xml"
            } else {
                "/word/header.xml"
            };
            package.set_part(
                part_name,
                format!(
                    r#"<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr>"#,
                    text_form("field", "old")
                )
                .into_bytes(),
            );
            let document = Document::from_bytes(&package_bytes(package)).unwrap();
            if case == "traversal" {
                assert!(document.legacy_form_fields().is_err(), "{case}");
            } else {
                assert!(document.legacy_form_fields().unwrap().is_empty(), "{case}");
            }
        }

        let mut package = base_package();
        let relationships = package.get_or_create_part_rels("/word/document.xml");
        let malformed = relationships.add(rel_types::HEADER, "ignored.xml");
        relationships
            .items
            .iter_mut()
            .find(|relationship| relationship.id == malformed)
            .unwrap()
            .target_mode = Some("internal".into());
        relationships.add(rel_types::HEADER, "valid.xml");
        package.set_part(
            "/word/valid.xml",
            format!(
                r#"<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr>"#,
                text_form("valid", "value")
            )
            .into_bytes(),
        );
        package.content_types.add_override(
            "/word/valid.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        );
        let document = Document::from_bytes(&package_bytes(package)).unwrap();
        assert_eq!(document.legacy_form_fields().unwrap().len(), 1);
    }

    #[test]
    fn forms_in_table_and_row_owned_content_controls_keep_their_ordinals() {
        let mut package = base_package();
        let table_control = format!(
            r#"<w:sdt><w:sdtContent><w:tr><w:tc>{}</w:tc></w:tr></w:sdtContent></w:sdt>"#,
            text_form("table-control", "first")
        );
        let row_control = format!(
            r#"<w:sdt><w:sdtContent><w:tc>{}</w:tc></w:sdtContent></w:sdt>"#,
            text_form("row-control", "second")
        );
        package.set_part(
            "/word/document.xml",
            format!(
                r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:tbl>{table_control}<w:tr>{row_control}<w:tc>{}</w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#,
                text_form("later", "third")
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let fields = document.legacy_form_fields().unwrap();
        assert_eq!(fields.len(), 3);
        assert_eq!(fields[0].name.as_deref(), Some("table-control"));
        assert_eq!(fields[1].name.as_deref(), Some("row-control"));
        document
            .set_legacy_form_field_value(
                "/word/document.xml",
                0,
                LegacyFormFieldValue::Text("first changed".into()),
            )
            .unwrap();
        document
            .set_legacy_form_field_value(
                "/word/document.xml",
                1,
                LegacyFormFieldValue::Text("second changed".into()),
            )
            .unwrap();
        let fields = Document::from_bytes(&document.to_bytes().unwrap())
            .unwrap()
            .legacy_form_fields()
            .unwrap();
        assert_eq!(
            fields[0].value,
            LegacyFormFieldValue::Text("first changed".into())
        );
        assert_eq!(
            fields[1].value,
            LegacyFormFieldValue::Text("second changed".into())
        );
        assert_eq!(fields[2].value, LegacyFormFieldValue::Text("third".into()));
    }

    #[test]
    fn invalid_legacy_form_mutations_are_atomic() {
        let mut package = base_package();
        let xml = format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="{WORD_NS}"><w:body>{}{}<w:sectPr/></w:body></w:document>"#,
            text_form("text", "old"),
            drop_down_form("choice")
        );
        package.set_part("/word/document.xml", xml.into_bytes());
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let before = document.to_bytes().unwrap();
        assert!(
            document
                .set_legacy_form_field_value(
                    "/word/document.xml",
                    0,
                    LegacyFormFieldValue::Checked(true),
                )
                .is_err()
        );
        assert!(
            document
                .set_legacy_form_field_value(
                    "/word/document.xml",
                    1,
                    LegacyFormFieldValue::SelectedIndex(9),
                )
                .is_err()
        );
        assert!(
            document
                .set_legacy_form_field_value(
                    "/word/document.xml",
                    9,
                    LegacyFormFieldValue::Text("missing".into()),
                )
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), before);
    }

    #[test]
    fn unsafe_or_malformed_glossary_graphs_fail_closed() {
        let glossary_xml = format!(
            r#"<?xml version="1.0"?><w:glossaryDocument xmlns:w="{WORD_NS}"><w:docParts/></w:glossaryDocument>"#
        );
        for case in [
            "duplicate",
            "traversal",
            "missing",
            "wrong-type",
            "wrong-root",
        ] {
            let mut package = base_package();
            let relationships = package.get_or_create_part_rels("/word/document.xml");
            let target = if case == "traversal" {
                "../../outside.xml"
            } else {
                "glossary/document.xml"
            };
            relationships.add(rel_types::GLOSSARY_DOCUMENT, target);
            if case == "duplicate" {
                relationships.add(rel_types::GLOSSARY_DOCUMENT, "glossary/other.xml");
            }
            let part_name = "/word/glossary/document.xml";
            if case != "missing" && case != "traversal" {
                package.set_part(
                    part_name,
                    if case == "wrong-root" {
                        format!(r#"<w:document xmlns:w="{WORD_NS}"/>"#).into_bytes()
                    } else {
                        glossary_xml.as_bytes().to_vec()
                    },
                );
                package.content_types.add_override(
                    part_name,
                    if case == "wrong-type" {
                        "application/xml"
                    } else {
                        GLOSSARY_CONTENT_TYPE
                    },
                );
            }
            assert!(
                Document::from_bytes(&package_bytes(package)).is_err(),
                "{case} glossary graph must fail closed"
            );
        }

        for mode in ["External", "internal", "ProducerDefined"] {
            let mut package = base_package();
            let relationships = package.get_or_create_part_rels("/word/document.xml");
            let ignored = relationships.add(rel_types::GLOSSARY_DOCUMENT, "glossary/ignored.xml");
            relationships
                .items
                .iter_mut()
                .find(|relationship| relationship.id == ignored)
                .unwrap()
                .target_mode = Some(mode.into());
            let document = Document::from_bytes(&package_bytes(package)).unwrap();
            assert!(document.building_blocks().unwrap().is_empty(), "{mode}");
        }

        let mut package = base_package();
        let relationships = package.get_or_create_part_rels("/word/document.xml");
        let ignored = relationships.add(rel_types::GLOSSARY_DOCUMENT, "glossary/ignored.xml");
        relationships
            .items
            .iter_mut()
            .find(|relationship| relationship.id == ignored)
            .unwrap()
            .target_mode = Some("internal".into());
        relationships.add(rel_types::GLOSSARY_DOCUMENT, "glossary/document.xml");
        package.set_part(
            "/word/glossary/document.xml",
            format!(r#"<w:glossaryDocument xmlns:w="{WORD_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="valid"/></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#).into_bytes(),
        );
        package
            .content_types
            .add_override("/word/glossary/document.xml", GLOSSARY_CONTENT_TYPE);
        let document = Document::from_bytes(&package_bytes(package)).unwrap();
        assert_eq!(document.building_blocks().unwrap().len(), 1);
    }

    #[test]
    fn glossary_requires_an_explicit_content_type_override() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::GLOSSARY_DOCUMENT, "glossary/document.xml");
        package
            .content_types
            .defaults
            .insert("xml".to_owned(), GLOSSARY_CONTENT_TYPE.to_owned());
        package.set_part(
            "/word/glossary/document.xml",
            format!(
                r#"<w:glossaryDocument xmlns:w="{WORD_NS}"><w:docParts/></w:glossaryDocument>"#
            )
            .into_bytes(),
        );
        assert!(Document::from_bytes(&package_bytes(package)).is_err());
    }

    #[test]
    fn inline_run_content_control_forms_are_inventoried_and_mutated() {
        let mut package = base_package();
        let inline_runs = text_form_runs("inline-control", "old").replacen(
            "</w:r>",
            r#"</w:r><w:proofErr w:type="spellStart"/>"#,
            1,
        );
        package.set_part(
            "/word/document.xml",
            format!(
                r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p><w:sdt><w:sdtContent>{}</w:sdtContent></w:sdt></w:p><w:sectPr/></w:body></w:document>"#,
                inline_runs
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let fields = document.legacy_form_fields().unwrap();
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].name.as_deref(), Some("inline-control"));
        document
            .set_legacy_form_field_value(
                "/word/document.xml",
                0,
                LegacyFormFieldValue::Text("changed".into()),
            )
            .unwrap();
        let mut reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.legacy_form_fields().unwrap()[0].value,
            LegacyFormFieldValue::Text("changed".into())
        );
        let saved = reopened.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        let begin = xml.find(r#"w:fldCharType="begin""#).unwrap();
        let proofing = xml.find(r#"<w:proofErr w:type="spellStart"/>"#).unwrap();
        let instruction = xml.find("<w:instrText> FORMTEXT </w:instrText>").unwrap();
        assert!(
            begin < proofing && proofing < instruction,
            "saved XML moved the interleaved proofing marker: {xml}"
        );
    }

    #[test]
    fn inline_form_mutation_preserves_run_root_namespace_context() {
        let mut package = base_package();
        let inline_runs = text_form_runs("context", "old")
            .replacen("<w:r>", r#"<w:r xmlns:x="urn:producer" x:run="keep">"#, 1)
            .replacen("</w:ffData>", "<x:retained/></w:ffData>", 1);
        package.set_part(
            "/word/document.xml",
            format!(
                r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p><w:sdt><w:sdtContent>{inline_runs}</w:sdtContent></w:sdt></w:p><w:sectPr/></w:body></w:document>"#
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        document
            .set_legacy_form_field_value(
                "/word/document.xml",
                0,
                LegacyFormFieldValue::Text("changed".into()),
            )
            .unwrap();
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(xml.contains(r#"xmlns:x="urn:producer" x:run="keep""#));
        assert!(xml.contains("<x:retained/>"));
        let reopened = Document::from_bytes(&package_bytes(package)).unwrap();
        assert_eq!(
            reopened.legacy_form_fields().unwrap()[0].value,
            LegacyFormFieldValue::Text("changed".into())
        );
    }

    #[test]
    fn package_story_inline_forms_persist_in_every_supported_story() {
        let mut package = base_package();
        let stories = [
            (rel_types::HEADER, "/word/header-inline.xml", "hdr"),
            (rel_types::FOOTER, "/word/footer-inline.xml", "ftr"),
            (
                rel_types::FOOTNOTES,
                "/word/footnotes-inline.xml",
                "footnotes",
            ),
            (rel_types::ENDNOTES, "/word/endnotes-inline.xml", "endnotes"),
        ];
        for (relationship_type, part_name, root) in stories {
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(relationship_type, part_name.strip_prefix("/word/").unwrap());
            add_story_content_type(&mut package, part_name, relationship_type);
            let name = root.to_owned();
            let content = if matches!(root, "hdr" | "ftr") {
                format!(
                    r#"<w:{root} xmlns:w="{WORD_NS}">{}</w:{root}>"#,
                    inline_text_form(&name, "old")
                )
            } else {
                let item = if root == "footnotes" {
                    "footnote"
                } else {
                    "endnote"
                };
                format!(
                    r#"<w:{root} xmlns:w="{WORD_NS}"><w:{item} w:id="2">{}</w:{item}></w:{root}>"#,
                    inline_text_form(&name, "old")
                )
            };
            package.set_part(part_name, content.into_bytes());
        }
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        for (_, part_name, root) in stories {
            document
                .set_legacy_form_field_value(
                    part_name,
                    0,
                    LegacyFormFieldValue::Text(format!("{root}-changed")),
                )
                .unwrap();
        }
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let fields = reopened.legacy_form_fields().unwrap();
        for (_, part_name, root) in stories {
            assert!(fields.iter().any(|field| {
                field.source_part == part_name
                    && field.value == LegacyFormFieldValue::Text(format!("{root}-changed"))
            }));
        }
    }

    #[test]
    fn package_form_stories_require_one_relationship_appropriate_root() {
        for (relationship_type, root, suffix) in [
            (rel_types::HEADER, "ftr", "wrong-header"),
            (rel_types::FOOTER, "hdr", "wrong-footer"),
        ] {
            let mut package = base_package();
            let part = format!("/word/{suffix}.xml");
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(relationship_type, &format!("{suffix}.xml"));
            add_story_content_type(&mut package, &part, relationship_type);
            package.set_part(
                &part,
                format!(
                    r#"<w:{root} xmlns:w="{WORD_NS}">{}</w:{root}>"#,
                    text_form("wrong-root", "old")
                )
                .into_bytes(),
            );
            let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
            assert!(document.legacy_form_fields().is_err(), "{suffix}");
            let before = document.to_bytes().unwrap();
            assert!(
                document
                    .set_legacy_form_field_value(
                        &part,
                        0,
                        LegacyFormFieldValue::Text("changed".to_owned()),
                    )
                    .is_err(),
                "{suffix}"
            );
            assert_eq!(document.to_bytes().unwrap(), before, "{suffix}");
        }

        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::HEADER, "two-roots.xml");
        add_story_content_type(&mut package, "/word/two-roots.xml", rel_types::HEADER);
        package.set_part(
            "/word/two-roots.xml",
            format!(
                r#"<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr><w:hdr xmlns:w="{WORD_NS}">{}</w:hdr>"#,
                text_form("first-root", "first"),
                text_form("second-root", "second")
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        assert!(document.legacy_form_fields().is_err());
        let before = document.to_bytes().unwrap();
        assert!(
            document
                .set_legacy_form_field_value(
                    "/word/two-roots.xml",
                    0,
                    LegacyFormFieldValue::Text("changed".to_owned()),
                )
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), before);
    }

    #[test]
    fn package_story_character_references_outside_the_root_are_rejected() {
        for (case, xml) in [
            (
                "leading",
                format!(
                    r#"&#65;<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr>"#,
                    text_form("leading", "old")
                ),
            ),
            (
                "trailing",
                format!(
                    r#"<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr>&#65;"#,
                    text_form("trailing", "old")
                ),
            ),
        ] {
            let mut package = base_package();
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(rel_types::HEADER, &format!("{case}.xml"));
            let part_name = format!("/word/{case}.xml");
            add_story_content_type(&mut package, &part_name, rel_types::HEADER);
            package.set_part(&part_name, xml.into_bytes());
            let document = Document::from_bytes(&package_bytes(package)).unwrap();
            assert!(document.legacy_form_fields().is_err(), "{case}");
        }
    }

    #[test]
    fn note_identity_and_type_require_word_namespace_attributes() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::FOOTNOTES, "footnotes-namespaces.xml");
        add_story_content_type(
            &mut package,
            "/word/footnotes-namespaces.xml",
            rel_types::FOOTNOTES,
        );
        package.set_part(
            "/word/footnotes-namespaces.xml",
            format!(
                r#"<w:footnotes xmlns:w="{WORD_NS}" xmlns:x="urn:foreign"><w:footnote x:id="2">{}</w:footnote><w:footnote w:id="-1" x:id="3">{}</w:footnote><w:footnote w:id="4" w:type="separator" x:type="normal">{}</w:footnote><w:footnote w:id="5" x:id="-1" x:type="separator">{}</w:footnote></w:footnotes>"#,
                text_form("foreign-id", "old"),
                text_form("foreign-override", "old"),
                text_form("foreign-type-override", "old"),
                text_form("valid", "old"),
            )
            .into_bytes(),
        );
        let document = Document::from_bytes(&package_bytes(package)).unwrap();
        let fields = document.legacy_form_fields().unwrap();
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].name.as_deref(), Some("valid"));
    }

    #[test]
    fn explicitly_normal_note_ids_are_supported_regardless_of_sign() {
        let mut package = base_package();
        let relationships = package.get_or_create_part_rels("/word/document.xml");
        relationships.add(rel_types::FOOTNOTES, "footnotes-signed.xml");
        relationships.add(rel_types::ENDNOTES, "endnotes-signed.xml");
        add_story_content_type(
            &mut package,
            "/word/footnotes-signed.xml",
            rel_types::FOOTNOTES,
        );
        add_story_content_type(
            &mut package,
            "/word/endnotes-signed.xml",
            rel_types::ENDNOTES,
        );
        package.set_part(
            "/word/footnotes-signed.xml",
            format!(
                r#"<w:footnotes xmlns:w="{WORD_NS}"><w:footnote w:id="0" w:type="normal">{}</w:footnote><w:footnote w:id="-9" w:type="normal">{}</w:footnote></w:footnotes>"#,
                text_form("zero-footnote", "zero"),
                text_form("negative-footnote", "negative"),
            )
            .into_bytes(),
        );
        package.set_part(
            "/word/endnotes-signed.xml",
            format!(
                r#"<w:endnotes xmlns:w="{WORD_NS}"><w:endnote w:id="0" w:type="normal">{}</w:endnote><w:endnote w:id="-11" w:type="normal">{}</w:endnote></w:endnotes>"#,
                text_form("zero-endnote", "zero"),
                text_form("negative-endnote", "negative"),
            )
            .into_bytes(),
        );
        let document = Document::from_bytes(&package_bytes(package)).unwrap();
        let fields = document.legacy_form_fields().unwrap();
        assert_eq!(
            fields
                .iter()
                .map(|field| field.name.as_deref().unwrap())
                .collect::<Vec<_>>(),
            [
                "zero-endnote",
                "negative-endnote",
                "zero-footnote",
                "negative-footnote",
            ]
        );
    }

    #[test]
    fn duplicate_footnotes_and_endnotes_relationships_fail_closed() {
        for (relationship_type, root, item) in [
            (rel_types::FOOTNOTES, "footnotes", "footnote"),
            (rel_types::ENDNOTES, "endnotes", "endnote"),
        ] {
            let mut package = base_package();
            let relationships = package.get_or_create_part_rels("/word/document.xml");
            relationships.add(relationship_type, &format!("{root}-one.xml"));
            relationships.add(relationship_type, &format!("{root}-two.xml"));
            for suffix in ["one", "two"] {
                package.set_part(
                    &format!("/word/{root}-{suffix}.xml"),
                    format!(
                        r#"<w:{root} xmlns:w="{WORD_NS}"><w:{item} w:id="2">{}</w:{item}></w:{root}>"#,
                        text_form(suffix, "old")
                    )
                    .into_bytes(),
                );
            }
            let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
            let before = document.to_bytes().unwrap();
            assert!(document.legacy_form_fields().is_err(), "{root}");
            assert!(
                document
                    .set_legacy_form_field_value(
                        &format!("/word/{root}-one.xml"),
                        0,
                        LegacyFormFieldValue::Text("changed".to_owned()),
                    )
                    .is_err(),
                "{root}"
            );
            assert_eq!(document.to_bytes().unwrap(), before, "{root}");
        }
    }

    #[test]
    fn package_story_declarations_and_doctypes_require_document_positions() {
        for (case, xml) in [
            (
                "nested-declaration",
                format!(
                    r#"<w:hdr xmlns:w="{WORD_NS}"><?xml version="1.0"?>{}</w:hdr>"#,
                    text_form("nested", "old")
                ),
            ),
            (
                "trailing-declaration",
                format!(
                    r#"<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr><?xml version="1.0"?>"#,
                    text_form("trailing", "old")
                ),
            ),
            (
                "trailing-doctype",
                format!(
                    r#"<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr><!DOCTYPE hdr>"#,
                    text_form("doctype", "old")
                ),
            ),
        ] {
            let mut package = base_package();
            let part_name = format!("/word/{case}.xml");
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(rel_types::HEADER, &format!("{case}.xml"));
            package
                .content_types
                .add_override(&part_name, HEADER_CONTENT_TYPE);
            package.set_part(&part_name, xml.into_bytes());
            assert!(
                Document::from_bytes(&package_bytes(package))
                    .unwrap()
                    .legacy_form_fields()
                    .is_err(),
                "{case}"
            );
        }
    }

    #[test]
    fn package_story_document_type_declarations_fail_closed_atomically() {
        for (case, declaration) in [
            ("uppercase-simple", "<!DOCTYPE w:hdr>"),
            ("lowercase-keyword", "<!doctype w:hdr>"),
            ("invalid-root-name", "<!DOCTYPE 1producer>"),
            (
                "external-system-identifier",
                r#"<!DOCTYPE w:hdr SYSTEM "urn:producer">"#,
            ),
            (
                "external-public-identifier",
                r#"<!DOCTYPE w:hdr PUBLIC "producer" "urn:producer">"#,
            ),
            ("internal-subset", "<!DOCTYPE w:hdr [<!ELEMENT w:hdr ANY>]>"),
            (
                "truncated-internal-subset",
                "<!DOCTYPE w:hdr [<!ELEMENT w:hdr ANY>",
            ),
        ] {
            let mut package = base_package();
            let part_name = format!("/word/doctype-{case}.xml");
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(rel_types::HEADER, &format!("doctype-{case}.xml"));
            package
                .content_types
                .add_override(&part_name, HEADER_CONTENT_TYPE);
            package.set_part(
                &part_name,
                format!(
                    r#"{declaration}<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr>"#,
                    text_form(case, "old")
                )
                .into_bytes(),
            );
            let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
            let before = document.to_bytes().unwrap();
            assert!(document.legacy_form_fields().is_err(), "{case}");
            assert!(
                document
                    .set_legacy_form_field_value(
                        &part_name,
                        0,
                        LegacyFormFieldValue::Text("changed".to_owned()),
                    )
                    .is_err(),
                "{case}"
            );
            assert_eq!(document.to_bytes().unwrap(), before, "{case}");
        }
    }

    #[test]
    fn package_story_xml_declarations_require_valid_pseudo_attributes() {
        for (case, declaration) in [
            ("missing-version", r#"<?xml?>"#),
            (
                "encoding-first",
                r#"<?xml encoding="UTF-8" version="1.0"?>"#,
            ),
            (
                "duplicate-version",
                r#"<?xml version="1.0" version="1.0"?>"#,
            ),
            (
                "invalid-standalone",
                r#"<?xml version="1.0" standalone="maybe"?>"#,
            ),
        ] {
            let mut package = base_package();
            let part_name = format!("/word/{case}.xml");
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(rel_types::HEADER, &format!("{case}.xml"));
            package
                .content_types
                .add_override(&part_name, HEADER_CONTENT_TYPE);
            package.set_part(
                &part_name,
                format!(
                    r#"{declaration}<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr>"#,
                    text_form(case, "old")
                )
                .into_bytes(),
            );
            assert!(
                Document::from_bytes(&package_bytes(package))
                    .unwrap()
                    .legacy_form_fields()
                    .is_err(),
                "{case}"
            );
        }
    }

    #[test]
    fn ooxml_package_stories_reject_xml_1_1_declarations() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::HEADER, "xml-1-1.xml");
        package
            .content_types
            .add_override("/word/xml-1-1.xml", HEADER_CONTENT_TYPE);
        package.set_part(
            "/word/xml-1-1.xml",
            format!(
                r#"<?xml version="1.1"?><w:hdr xmlns:w="{WORD_NS}">{}</w:hdr>"#,
                text_form("xml-1-1", "old")
            )
            .into_bytes(),
        );
        assert!(
            Document::from_bytes(&package_bytes(package))
                .unwrap()
                .legacy_form_fields()
                .is_err()
        );
    }

    #[test]
    fn package_story_relationship_role_requires_exact_content_type_override() {
        for (case, content_type) in [
            ("missing", None),
            ("generic", Some("application/xml")),
            ("wrong-role", Some(FOOTER_CONTENT_TYPE)),
        ] {
            let mut package = base_package();
            let part_name = format!("/word/header-{case}.xml");
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(rel_types::HEADER, &format!("header-{case}.xml"));
            if let Some(content_type) = content_type {
                package.content_types.add_override(&part_name, content_type);
            }
            package.set_part(
                &part_name,
                format!(
                    r#"<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr>"#,
                    text_form(case, "old")
                )
                .into_bytes(),
            );
            let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
            let before = document.to_bytes().unwrap();
            assert!(document.legacy_form_fields().is_err(), "{case}");
            assert!(
                document
                    .set_legacy_form_field_value(
                        &part_name,
                        0,
                        LegacyFormFieldValue::Text("changed".to_owned()),
                    )
                    .is_err(),
                "{case}"
            );
            assert_eq!(document.to_bytes().unwrap(), before, "{case}");
        }
    }

    #[test]
    fn encoded_note_identity_and_type_attributes_are_decoded() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::FOOTNOTES, "footnotes-encoded.xml");
        package
            .content_types
            .add_override("/word/footnotes-encoded.xml", FOOTNOTES_CONTENT_TYPE);
        package.set_part(
            "/word/footnotes-encoded.xml",
            format!(r#"<w:footnotes xmlns:w="{WORD_NS}"><w:footnote w:id="&#49;" w:type="norm&#97;l">{}</w:footnote></w:footnotes>"#, text_form("encoded", "old")).into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let fields = document.legacy_form_fields().unwrap();
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].name.as_deref(), Some("encoded"));
        document
            .set_legacy_form_field_value(
                "/word/footnotes-encoded.xml",
                0,
                LegacyFormFieldValue::Text("changed".to_owned()),
            )
            .unwrap();
    }

    #[test]
    fn duplicate_package_story_relationship_ids_fail_closed() {
        let mut package = base_package();
        let relationships = package.get_or_create_part_rels("/word/document.xml");
        relationships.add(rel_types::HEADER, "header-id-one.xml");
        relationships.add(rel_types::HEADER, "header-id-two.xml");
        let duplicate_id = relationships.items[relationships.items.len() - 2]
            .id
            .clone();
        relationships.items.last_mut().unwrap().id = duplicate_id;
        for suffix in ["one", "two"] {
            let part_name = format!("/word/header-id-{suffix}.xml");
            package
                .content_types
                .add_override(&part_name, HEADER_CONTENT_TYPE);
            package.set_part(
                &part_name,
                format!(
                    r#"<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr>"#,
                    text_form(suffix, "old")
                )
                .into_bytes(),
            );
        }
        let mut output = std::io::Cursor::new(Vec::new());
        assert!(matches!(
            package.write_to(&mut output),
            Err(oxml_opc::OpcError::InvalidRelationship)
        ));
        assert!(output.into_inner().is_empty());
    }

    #[test]
    fn conflicting_package_story_relationship_roles_are_rejected() {
        let mut package = base_package();
        let relationships = package.get_or_create_part_rels("/word/document.xml");
        relationships.add(rel_types::HEADER, "shared-story.xml");
        relationships.add(rel_types::FOOTER, "shared-story.xml");
        package.set_part(
            "/word/shared-story.xml",
            format!(
                r#"<w:hdr xmlns:w="{WORD_NS}">{}</w:hdr>"#,
                text_form("shared", "old")
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        assert!(document.legacy_form_fields().is_err());
        let before = document.to_bytes().unwrap();
        assert!(
            document
                .set_legacy_form_field_value(
                    "/word/shared-story.xml",
                    0,
                    LegacyFormFieldValue::Text("changed".to_owned()),
                )
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), before);
    }

    #[test]
    fn default_namespace_valueless_header_form_persists_false() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::HEADER, "header-default.xml");
        add_story_content_type(&mut package, "/word/header-default.xml", rel_types::HEADER);
        package.set_part(
            "/word/header-default.xml",
            format!(
                concat!(
                    r#"<hdr xmlns="{WORD_NS}" xmlns:q="{WORD_NS}"><p><sdt><sdtContent>"#,
                    r#"<r><fldChar q:fldCharType="begin"><ffData><name q:val="default"/><checkBox><sizeAuto/><checked/></checkBox></ffData></fldChar></r>"#,
                    r#"<r><instrText> FORMCHECKBOX </instrText></r><r><fldChar q:fldCharType="separate"/></r>"#,
                    r#"<r><t>☒</t></r><r><fldChar q:fldCharType="end"/></r>"#,
                    r#"</sdtContent></sdt></p></hdr>"#,
                ),
                WORD_NS = WORD_NS,
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        document
            .set_legacy_form_field_value(
                "/word/header-default.xml",
                0,
                LegacyFormFieldValue::Checked(false),
            )
            .unwrap();
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved.clone())).unwrap();
        let xml =
            std::str::from_utf8(package.get_part("/word/header-default.xml").unwrap()).unwrap();
        assert!(xml.contains(r#"<checked q:val="0"/>"#), "{xml}");
        let reopened = Document::from_bytes(&saved).unwrap();
        assert_eq!(
            reopened.legacy_form_fields().unwrap()[0].value,
            LegacyFormFieldValue::Checked(false)
        );
    }

    #[test]
    fn nested_inline_forms_have_mutable_source_ordinals() {
        let mut package = base_package();
        let nested_only = nested_inline_form_runs(None, "nested-only", "first");
        let outer_and_nested =
            nested_inline_form_runs(Some("outer"), "nested-under-form", "second");
        package.set_part(
            "/word/document.xml",
            format!(
                concat!(
                    r#"<w:document xmlns:w="{WORD_NS}"><w:body>"#,
                    r#"<w:p><w:sdt><w:sdtContent>{nested_only}</w:sdtContent></w:sdt></w:p>"#,
                    r#"<w:p><w:sdt><w:sdtContent>{outer_and_nested}</w:sdtContent></w:sdt></w:p>"#,
                    r#"<w:sectPr/></w:body></w:document>"#,
                ),
                WORD_NS = WORD_NS,
                nested_only = nested_only,
                outer_and_nested = outer_and_nested,
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let fields = document.legacy_form_fields().unwrap();
        assert_eq!(
            fields
                .iter()
                .map(|field| field.name.as_deref())
                .collect::<Vec<_>>(),
            [
                Some("nested-only"),
                Some("outer"),
                Some("nested-under-form")
            ]
        );
        document
            .set_legacy_form_field_value(
                "/word/document.xml",
                0,
                LegacyFormFieldValue::Text("first-changed".into()),
            )
            .unwrap();
        document
            .set_legacy_form_field_value(
                "/word/document.xml",
                2,
                LegacyFormFieldValue::Text("second-changed".into()),
            )
            .unwrap();
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let fields = reopened.legacy_form_fields().unwrap();
        assert_eq!(
            fields[0].value,
            LegacyFormFieldValue::Text("first-changed".into())
        );
        assert_eq!(
            fields[2].value,
            LegacyFormFieldValue::Text("second-changed".into())
        );
    }

    #[test]
    fn nested_instruction_forms_keep_source_order_identity() {
        let mut package = base_package();
        package.set_part(
            "/word/document.xml",
            format!(
                r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p>{}</w:p><w:sectPr/></w:body></w:document>"#,
                nested_instruction_form_runs()
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let fields = document.legacy_form_fields().unwrap();
        assert_eq!(
            fields
                .iter()
                .map(|field| field.name.as_deref())
                .collect::<Vec<_>>(),
            [Some("switch-first"), Some("positional-second")]
        );
        document
            .set_legacy_form_field_value(
                "/word/document.xml",
                0,
                LegacyFormFieldValue::Text("switch-changed".into()),
            )
            .unwrap();
        let fields = Document::from_bytes(&document.to_bytes().unwrap())
            .unwrap()
            .legacy_form_fields()
            .unwrap();
        assert_eq!(
            fields[0].value,
            LegacyFormFieldValue::Text("switch-changed".into())
        );
        assert_eq!(
            fields[1].value,
            LegacyFormFieldValue::Text("positional-old".into())
        );
    }

    #[test]
    fn interleaved_nested_inline_controls_keep_source_order_identity() {
        let mut package = base_package();
        let nested = format!(
            r#"<w:sdt><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
            text_form_runs("middle", "middle-old")
        );
        package.set_part(
            "/word/document.xml",
            format!(
                concat!(
                    r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p><w:sdt><w:sdtContent>"#,
                    "{}{}{}",
                    r#"</w:sdtContent></w:sdt></w:p><w:sectPr/></w:body></w:document>"#,
                ),
                text_form_runs("first", "first-old"),
                nested,
                text_form_runs("last", "last-old"),
                WORD_NS = WORD_NS,
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let fields = document.legacy_form_fields().unwrap();
        assert_eq!(
            fields
                .iter()
                .map(|field| field.name.as_deref())
                .collect::<Vec<_>>(),
            [Some("first"), Some("middle"), Some("last")]
        );
        document
            .set_legacy_form_field_value(
                "/word/document.xml",
                1,
                LegacyFormFieldValue::Text("middle-changed".into()),
            )
            .unwrap();
        let fields = Document::from_bytes(&document.to_bytes().unwrap())
            .unwrap()
            .legacy_form_fields()
            .unwrap();
        assert_eq!(
            fields.iter().map(|field| &field.value).collect::<Vec<_>>(),
            [
                &LegacyFormFieldValue::Text("first-old".into()),
                &LegacyFormFieldValue::Text("middle-changed".into()),
                &LegacyFormFieldValue::Text("last-old".into()),
            ]
        );
    }

    #[test]
    fn duplicate_ffdata_owner_is_rejected_without_mutating_preserved_source() {
        let mut package = base_package();
        let duplicate_owner = concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin">"#,
            r#"<w:ffData><w:name w:val="first"/><w:textInput><w:default w:val="old"/></w:textInput></w:ffData>"#,
            r#"<w:ffData><w:name w:val="duplicate"/><w:textInput><w:default w:val="other"/></w:textInput></w:ffData>"#,
            r#"</w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>old</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        );
        package.set_part(
            "/word/document.xml",
            format!(
                r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p>{duplicate_owner}</w:p><w:sectPr/></w:body></w:document>"#
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let before = document.to_bytes().unwrap();
        assert!(document.legacy_form_fields().is_err());
        assert!(
            document
                .set_legacy_form_field_value(
                    "/word/document.xml",
                    0,
                    LegacyFormFieldValue::Text("changed".into()),
                )
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), before);
    }

    #[test]
    fn empty_doc_part_properties_accept_valid_building_block_replacement() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::GLOSSARY_DOCUMENT, "glossary/document.xml");
        package
            .content_types
            .add_override("/word/glossary/document.xml", GLOSSARY_CONTENT_TYPE);
        package.set_part(
            "/word/glossary/document.xml",
            format!(
                r#"<w:glossaryDocument xmlns:w="{WORD_NS}"><w:docParts><w:docPart><w:docPartPr/><w:docPartBody><w:p/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let mut replacement = document.building_blocks().unwrap()[0].block.clone();
        replacement.name = "named".to_owned();
        document
            .replace_building_block("/word/glossary/document.xml", 0, replacement)
            .unwrap();
        let saved = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&saved).unwrap();
        assert_eq!(reopened.building_blocks().unwrap()[0].block.name, "named");
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        let xml =
            std::str::from_utf8(package.get_part("/word/glossary/document.xml").unwrap()).unwrap();
        assert!(xml.contains("<w:docPartPr>"));
        assert!(xml.contains(r#"<w:name w:val="named"/>"#));
    }

    #[test]
    fn package_story_block_content_forms_keep_source_order_and_mutate() {
        let mut package = base_package();
        for (relationship_type, target) in [
            (rel_types::HEADER, "header-block.xml"),
            (rel_types::FOOTER, "footer-block.xml"),
            (rel_types::FOOTNOTES, "footnotes-block.xml"),
            (rel_types::ENDNOTES, "endnotes-block.xml"),
        ] {
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(relationship_type, target);
            add_story_content_type(&mut package, &format!("/word/{target}"), relationship_type);
        }
        package.set_part(
            "/word/header-block.xml",
            format!(
                r#"<w:hdr xmlns:w="{WORD_NS}">{}<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl></w:hdr>"#,
                text_form("header-direct", "direct-old"),
                text_form("header-table", "header-old")
            )
            .into_bytes(),
        );
        package.set_part(
            "/word/footer-block.xml",
            format!(
                r#"<w:ftr xmlns:w="{WORD_NS}"><w:sdt><w:sdtContent>{}</w:sdtContent></w:sdt></w:ftr>"#,
                text_form("footer-control", "footer-old")
            )
            .into_bytes(),
        );
        package.set_part(
            "/word/footnotes-block.xml",
            format!(
                r#"<w:footnotes xmlns:w="{WORD_NS}"><w:footnote w:id="2"><w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl></w:footnote></w:footnotes>"#,
                text_form("footnote-table", "footnote-old")
            )
            .into_bytes(),
        );
        package.set_part(
            "/word/endnotes-block.xml",
            format!(
                r#"<w:endnotes xmlns:w="{WORD_NS}"><w:endnote w:id="2"><w:sdt><w:sdtContent>{}</w:sdtContent></w:sdt></w:endnote></w:endnotes>"#,
                text_form("endnote-control", "endnote-old")
            )
            .into_bytes(),
        );

        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let fields = document.legacy_form_fields().unwrap();
        let names = |part: &str| {
            fields
                .iter()
                .filter(|field| field.source_part == part)
                .map(|field| field.name.as_deref())
                .collect::<Vec<_>>()
        };
        assert_eq!(
            names("/word/header-block.xml"),
            [Some("header-direct"), Some("header-table")]
        );
        assert_eq!(names("/word/footer-block.xml"), [Some("footer-control")]);
        assert_eq!(names("/word/footnotes-block.xml"), [Some("footnote-table")]);
        assert_eq!(names("/word/endnotes-block.xml"), [Some("endnote-control")]);

        for (part, ordinal, value) in [
            ("/word/header-block.xml", 1, "header-changed"),
            ("/word/footer-block.xml", 0, "footer-changed"),
            ("/word/footnotes-block.xml", 0, "footnote-changed"),
            ("/word/endnotes-block.xml", 0, "endnote-changed"),
        ] {
            document
                .set_legacy_form_field_value(
                    part,
                    ordinal,
                    LegacyFormFieldValue::Text(value.to_owned()),
                )
                .unwrap();
        }
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let fields = reopened.legacy_form_fields().unwrap();
        assert!(fields.iter().any(|field| {
            field.source_part == "/word/header-block.xml"
                && field.ordinal == 0
                && field.value == LegacyFormFieldValue::Text("direct-old".into())
        }));
        for (part, value) in [
            ("/word/header-block.xml", "header-changed"),
            ("/word/footer-block.xml", "footer-changed"),
            ("/word/footnotes-block.xml", "footnote-changed"),
            ("/word/endnotes-block.xml", "endnote-changed"),
        ] {
            assert!(fields.iter().any(|field| {
                field.source_part == part
                    && field.value == LegacyFormFieldValue::Text(value.to_owned())
            }));
        }
    }

    #[test]
    fn unrelated_building_block_edits_preserve_every_unsupported_subtree_byte() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::GLOSSARY_DOCUMENT, "glossary/document.xml");
        package
            .content_types
            .add_override("/word/glossary/document.xml", GLOSSARY_CONTENT_TYPE);
        let untouched = r#"<w:docPart producer="untouched"><w:docPartPr><w:name w:val="second"/><x:property> exact </x:property></w:docPartPr><w:docPartBody><w:p/><mc:AlternateContent><mc:Fallback><x:body/></mc:Fallback></mc:AlternateContent></w:docPartBody></w:docPart>"#;
        package.set_part(
            "/word/glossary/document.xml",
            format!(
                r#"<w:glossaryDocument xmlns:w="{WORD_NS}" xmlns:x="urn:producer" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:docParts><w:docPart><w:docPartPr><w:name w:val="first"/><x:selected keep="exact"/></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart>{untouched}</w:docParts></w:glossaryDocument>"#
            )
            .into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let mut replacement = document.building_blocks().unwrap()[0].block.clone();
        replacement.description = Some("changed".into());
        document
            .replace_building_block("/word/glossary/document.xml", 0, replacement)
            .unwrap();
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        let xml =
            std::str::from_utf8(package.get_part("/word/glossary/document.xml").unwrap()).unwrap();
        assert!(xml.contains(r#"<x:selected keep="exact"/>"#));
        assert!(xml.contains(untouched));
    }

    #[test]
    fn failed_building_block_replacements_leave_document_bytes_unchanged() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::GLOSSARY_DOCUMENT, "glossary/document.xml");
        package
            .content_types
            .add_override("/word/glossary/document.xml", GLOSSARY_CONTENT_TYPE);
        package.set_part(
            "/word/glossary/document.xml",
            format!(r#"<w:glossaryDocument xmlns:w="{WORD_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#).into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let before = document.to_bytes().unwrap();
        let mut invalid = document.building_blocks().unwrap()[0].block.clone();
        invalid.name.clear();
        assert!(
            document
                .replace_building_block("/word/glossary/document.xml", 0, invalid)
                .is_err()
        );
        assert!(
            document
                .replace_building_block(
                    "/word/glossary/document.xml",
                    4,
                    document.building_blocks().unwrap()[0].block.clone(),
                )
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), before);
    }

    #[test]
    fn building_block_replacement_rejects_partial_category_pairs_atomically() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::GLOSSARY_DOCUMENT, "glossary/document.xml");
        package
            .content_types
            .add_override("/word/glossary/document.xml", GLOSSARY_CONTENT_TYPE);
        package.set_part(
            "/word/glossary/document.xml",
            format!(r#"<w:glossaryDocument xmlns:w="{WORD_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:category><w:name w:val="category"/><w:gallery w:val="autoTxt"/></w:category></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#).into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let before = document.to_bytes().unwrap();
        for clear_category in [true, false] {
            let mut invalid = document.building_blocks().unwrap()[0].block.clone();
            if clear_category {
                invalid.category = None;
            } else {
                invalid.gallery = None;
            }
            assert!(
                document
                    .replace_building_block("/word/glossary/document.xml", 0, invalid)
                    .is_err()
            );
            assert_eq!(document.to_bytes().unwrap(), before);
        }
    }

    #[test]
    fn building_block_replacement_rejects_invalid_gallery_and_behavior_enums_atomically() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::GLOSSARY_DOCUMENT, "glossary/document.xml");
        package
            .content_types
            .add_override("/word/glossary/document.xml", GLOSSARY_CONTENT_TYPE);
        package.set_part(
            "/word/glossary/document.xml",
            format!(r#"<w:glossaryDocument xmlns:w="{WORD_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:category><w:name w:val="category"/><w:gallery w:val="autoTxt"/></w:category><w:behaviors><w:behavior w:val="content"/></w:behaviors></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#).into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let before = document.to_bytes().unwrap();
        for invalid_gallery in [true, false] {
            let mut invalid = document.building_blocks().unwrap()[0].block.clone();
            if invalid_gallery {
                invalid.gallery = Some("not-a-gallery".to_owned());
            } else {
                invalid.behaviors = vec!["not-a-behavior".to_owned()];
            }
            assert!(
                document
                    .replace_building_block("/word/glossary/document.xml", 0, invalid)
                    .is_err()
            );
            assert_eq!(document.to_bytes().unwrap(), before);
        }
    }

    #[test]
    fn building_block_replacement_rejects_invalid_guid_atomically() {
        let mut package = base_package();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::GLOSSARY_DOCUMENT, "glossary/document.xml");
        package
            .content_types
            .add_override("/word/glossary/document.xml", GLOSSARY_CONTENT_TYPE);
        package.set_part(
            "/word/glossary/document.xml",
            format!(r#"<w:glossaryDocument xmlns:w="{WORD_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:guid w:val="{{01234567-89AB-CDEF-0123-456789ABCDEF}}"/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#).into_bytes(),
        );
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let before = document.to_bytes().unwrap();
        let mut invalid = document.building_blocks().unwrap()[0].block.clone();
        invalid.guid = Some("not-a-guid".to_owned());
        assert!(
            document
                .replace_building_block("/word/glossary/document.xml", 0, invalid)
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), before);
    }

    #[test]
    fn glossary_and_story_relationship_targets_must_be_normalized() {
        for target in ["glossary//document.xml", "glossary/./document.xml"] {
            let mut package = base_package();
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(rel_types::GLOSSARY_DOCUMENT, target);
            package
                .content_types
                .add_override("/word/glossary/document.xml", GLOSSARY_CONTENT_TYPE);
            package.set_part(
                "/word/glossary/document.xml",
                format!(r#"<w:glossaryDocument xmlns:w="{WORD_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#).into_bytes(),
            );
            if let Ok(document) = Document::from_bytes(&package_bytes(package)) {
                assert!(document.building_blocks().is_err(), "{target}");
            }
        }

        for target in ["stories//header.xml", "stories/./header.xml"] {
            let mut package = base_package();
            package
                .get_or_create_part_rels("/word/document.xml")
                .add(rel_types::HEADER, target);
            add_story_content_type(&mut package, "/word/stories/header.xml", rel_types::HEADER);
            package.set_part(
                "/word/stories/header.xml",
                format!(r#"<w:hdr xmlns:w="{WORD_NS}"><w:p/></w:hdr>"#).into_bytes(),
            );
            if let Ok(document) = Document::from_bytes(&package_bytes(package)) {
                assert!(document.legacy_form_fields().is_err(), "{target}");
            }
        }
    }
}

mod f264_paragraph_property_tests {
    use super::*;
    use rdocx::{
        DropCap, FrameAnchor, FrameWrap, ParagraphBorderEdge, ParagraphFrame,
        ParagraphTextAlignment, ParagraphTextDirection, TextboxTightWrap,
    };

    const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// Children of `w:pPr` that rdocx does not model, at four schema slots
    /// around the ones this feature types.
    const PRODUCER_CHILDREN: [&str; 4] = [
        "<w:kinsoku/>",
        "<w:snapToGrid/>",
        r#"<w:cnfStyle w:val="100000000000"/>"#,
        r#"<ext:marker xmlns:ext="urn:producer" ext:keep="exact"/>"#,
    ];

    fn producer_document(paragraph_properties: &str) -> Vec<u8> {
        let mut seed = Document::new();
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
        package.set_part(
            "/word/document.xml",
            format!(
                concat!(
                    r#"<w:document xmlns:w="{}">"#,
                    r#"<w:body><w:p><w:pPr>{}</w:pPr><w:r><w:t>framed</w:t></w:r></w:p>"#,
                    r#"<w:sectPr/></w:body></w:document>"#,
                ),
                WORD_NS, paragraph_properties
            )
            .into_bytes(),
        );
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        bytes.into_inner()
    }

    fn document_xml(bytes: &[u8]) -> String {
        let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap()
    }

    /// The first `w:pPr` element of the saved body, or the empty string when
    /// the paragraph carries no properties at all.
    fn paragraph_properties_xml(bytes: &[u8]) -> String {
        let xml = document_xml(bytes);
        let Some(start) = xml.find("<w:pPr>") else {
            return String::new();
        };
        let end = xml.find("</w:pPr>").expect("a w:pPr end tag");
        xml[start..end].to_owned()
    }

    fn authored_frame() -> ParagraphFrame {
        ParagraphFrame {
            width: Some(Length::twips(2880)),
            height: Some(Length::twips(1440)),
            horizontal_space: Some(Length::twips(180)),
            vertical_space: Some(Length::twips(120)),
            horizontal_position: Some(Length::twips(720)),
            vertical_position: Some(Length::twips(360)),
            horizontal_anchor: Some(FrameAnchor::Margin),
            vertical_anchor: Some(FrameAnchor::Text),
            wrap: Some(FrameWrap::Around),
            drop_cap: Some(DropCap::Drop),
            drop_cap_lines: Some(3),
            anchor_lock: Some(true),
        }
    }

    #[test]
    fn every_public_paragraph_property_reopens_and_preserves_unrelated_xml() {
        let source = producer_document(&PRODUCER_CHILDREN.concat());
        let mut document = Document::from_bytes(&source).unwrap();
        {
            let mut paragraph = document.paragraph_mut(0).unwrap();
            paragraph.set_indent_start(Length::twips(720));
            paragraph.set_indent_end(Length::twips(360));
            paragraph.set_hanging_indent_value(Some(Length::twips(240)));
            paragraph.set_mirror_indents(true);
            paragraph.set_adjust_right_indent(true);
            paragraph.set_space_before_auto(true);
            paragraph.set_space_after_auto(false);
            paragraph.set_contextual_spacing(true);
            paragraph.set_border(
                ParagraphBorderEdge::Between,
                BorderStyle::Dashed,
                6,
                "FF0000",
            );
            paragraph.set_shading_pattern("pct20", "FFFF00", "auto");
            paragraph.set_add_tab_stop(TabAlignment::Right, Length::twips(8640));
            paragraph.set_suppress_line_numbers(true);
            paragraph.set_suppress_auto_hyphens(true);
            paragraph.set_frame(authored_frame());
            paragraph.set_suppress_overlap(true);
            paragraph.set_textbox_tight_wrap(TextboxTightWrap::FirstAndLastLine);
            assert!(paragraph.set_outline_level_value(Some(9)));
            paragraph.set_right_to_left(true);
            paragraph.set_text_direction(ParagraphTextDirection::TopToBottomRightToLeftVertical);
            paragraph.set_text_alignment(ParagraphTextAlignment::Center);
            paragraph.set_div_id_value(Some(11));
            paragraph.mark().set_bold(true);
            paragraph.mark().set_color("00FF00");
        }

        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        let paragraph = reopened.paragraph(0).unwrap();
        assert_eq!(paragraph.indent_start(), Some(Length::twips(720)));
        assert_eq!(paragraph.indent_end(), Some(Length::twips(360)));
        assert_eq!(paragraph.hanging_indent(), Some(Length::twips(240)));
        assert_eq!(paragraph.mirror_indents_value(), Some(true));
        assert_eq!(paragraph.adjust_right_indent_value(), Some(true));
        assert_eq!(paragraph.space_before_auto_value(), Some(true));
        assert_eq!(paragraph.space_after_auto_value(), Some(false));
        assert_eq!(paragraph.contextual_spacing_value(), Some(true));
        let between = paragraph.border(ParagraphBorderEdge::Between).unwrap();
        assert_eq!(between.style(), "dashed");
        assert_eq!(between.size_eighths_pt(), Some(6));
        assert_eq!(between.color(), Some("FF0000"));
        assert_eq!(paragraph.shading_pattern(), Some("pct20"));
        assert_eq!(paragraph.shading_fill(), Some("FFFF00"));
        assert_eq!(paragraph.shading_color(), Some("auto"));
        let tab = paragraph.tab_stop(0).unwrap();
        assert_eq!(tab.alignment(), Some(TabAlignment::Right));
        assert_eq!(tab.position(), Length::twips(8640));
        assert_eq!(paragraph.suppress_line_numbers_value(), Some(true));
        assert_eq!(paragraph.suppress_auto_hyphens_value(), Some(true));
        assert_eq!(paragraph.frame(), Some(authored_frame()));
        assert_eq!(paragraph.suppress_overlap_value(), Some(true));
        assert_eq!(
            paragraph.textbox_tight_wrap(),
            Some(TextboxTightWrap::FirstAndLastLine)
        );
        assert_eq!(paragraph.outline_level(), Some(9));
        assert_eq!(paragraph.right_to_left_value(), Some(true));
        assert_eq!(
            paragraph.text_direction(),
            Some(ParagraphTextDirection::TopToBottomRightToLeftVertical)
        );
        assert_eq!(
            paragraph.text_alignment(),
            Some(ParagraphTextAlignment::Center)
        );
        assert_eq!(paragraph.div_id(), Some(11));
        assert_eq!(paragraph.mark().bold_value(), Some(true));
        assert_eq!(paragraph.mark().color(), Some("00FF00"));

        let xml = document_xml(&bytes);
        for child in PRODUCER_CHILDREN {
            assert!(xml.contains(child), "{child} was lost: {xml}");
        }
    }

    #[test]
    fn paragraph_border_edges_author_read_and_clear_individually() {
        let source = producer_document(
            r#"<w:pBdr><w:top w:val="single" w:sz="4" w:themeColor="accent1"/></w:pBdr>"#,
        );
        let mut document = Document::from_bytes(&source).unwrap();
        let edges = [
            ParagraphBorderEdge::Top,
            ParagraphBorderEdge::Bottom,
            ParagraphBorderEdge::Left,
            ParagraphBorderEdge::Right,
            ParagraphBorderEdge::Between,
            ParagraphBorderEdge::Bar,
        ];
        {
            let mut paragraph = document.paragraph_mut(0).unwrap();
            for (index, edge) in edges.into_iter().enumerate() {
                paragraph.set_border(edge, BorderStyle::Single, index as u32 + 2, "0000FF");
            }
        }

        let bytes = document.to_bytes().unwrap();
        assert!(
            document_xml(&bytes).contains(r#"w:themeColor="accent1""#),
            "an edge mutation dropped a retained attribute"
        );

        let mut document = Document::from_bytes(&bytes).unwrap();
        for (index, edge) in edges.into_iter().enumerate() {
            let paragraph = document.paragraph(0).unwrap();
            let border = paragraph.border(edge).expect("an authored edge");
            assert_eq!(border.style(), "single");
            assert_eq!(border.size_eighths_pt(), Some(index as u32 + 2));
        }

        document
            .paragraph_mut(0)
            .unwrap()
            .set_border_value(ParagraphBorderEdge::Left, None);
        let paragraph = document.paragraph(0).unwrap();
        assert!(paragraph.border(ParagraphBorderEdge::Left).is_none());
        assert_eq!(paragraph.border_count(), 5);

        document.paragraph_mut(0).unwrap().clear_borders();
        assert_eq!(document.paragraph(0).unwrap().border_count(), 0);
        assert!(!document.paragraph(0).unwrap().has_borders());
    }

    #[test]
    fn tab_stops_read_mutate_and_remove_by_index() {
        let mut document = Document::new();
        {
            let mut paragraph = document.add_paragraph("tabs");
            paragraph.set_add_tab_stop(TabAlignment::Left, Length::twips(720));
            paragraph.set_add_tab_stop_with_leader(
                TabAlignment::Center,
                Length::twips(2880),
                TabLeader::Dot,
            );
            paragraph.set_add_tab_stop(TabAlignment::Right, Length::twips(8640));
            assert!(paragraph.tab_stop(3).is_none());
            assert!(!paragraph.set_tab_stop(3, TabAlignment::Left, Length::twips(100), None));
            assert!(!paragraph.remove_tab_stop(3));
            assert!(paragraph.set_tab_stop(
                1,
                TabAlignment::Decimal,
                Length::twips(4320),
                Some(TabLeader::Hyphen)
            ));
            assert!(paragraph.remove_tab_stop(0));
        }

        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        let paragraph = reopened.paragraph(0).unwrap();
        assert_eq!(paragraph.tab_stop_count(), 2);
        let first = paragraph.tab_stop(0).unwrap();
        assert_eq!(first.alignment(), Some(TabAlignment::Decimal));
        assert_eq!(first.position(), Length::twips(4320));
        assert_eq!(first.leader(), Some(TabLeader::Hyphen));
        let second = paragraph.tab_stop(1).unwrap();
        assert_eq!(second.alignment(), Some(TabAlignment::Right));
        assert_eq!(second.position(), Length::twips(8640));
        assert!(paragraph.tab_stop(2).is_none());

        let mut document = reopened;
        document.paragraph_mut(0).unwrap().clear_tab_stops();
        assert_eq!(document.paragraph(0).unwrap().tab_stop_count(), 0);
        assert!(document.paragraph(0).unwrap().tab_stop(0).is_none());
    }

    #[test]
    fn paragraph_mark_formatting_authors_reads_and_clears() {
        let mut document = Document::new();
        {
            let mut paragraph = document.add_paragraph("marked");
            let mut mark = paragraph.mark();
            mark.set_bold(true);
            mark.set_italic(true);
            mark.set_underline(true);
            mark.set_strike(true);
            mark.set_size(18.0);
            mark.set_font("Georgia");
            mark.set_color("112233");
        }

        let bytes = document.to_bytes().unwrap();
        assert!(
            paragraph_properties_xml(&bytes).contains("<w:rPr>"),
            "the mark properties must live inside w:pPr"
        );

        let mut reopened = Document::from_bytes(&bytes).unwrap();
        {
            let paragraph = reopened.paragraph(0).unwrap();
            let mark = paragraph.mark();
            assert!(mark.is_present());
            assert_eq!(mark.bold_value(), Some(true));
            assert_eq!(mark.italic_value(), Some(true));
            assert_eq!(mark.underline_value(), Some(true));
            assert_eq!(mark.strike_value(), Some(true));
            assert_eq!(mark.size(), Some(18.0));
            assert_eq!(mark.font_name(), Some("Georgia"));
            assert_eq!(mark.color(), Some("112233"));
        }

        reopened.paragraph_mut(0).unwrap().clear_mark();
        assert!(!reopened.paragraph(0).unwrap().mark().is_present());
        let cleared = reopened.to_bytes().unwrap();
        assert!(
            !paragraph_properties_xml(&cleared).contains("<w:rPr>"),
            "clear_mark must remove the element"
        );
        assert!(
            !Document::from_bytes(&cleared)
                .unwrap()
                .paragraph(0)
                .unwrap()
                .mark()
                .is_present()
        );
    }

    #[test]
    fn outline_level_accepts_zero_through_nine_and_rejects_above() {
        let mut document = Document::new();
        document.add_paragraph("outline");
        for level in 0..=9 {
            assert!(
                document
                    .paragraph_mut(0)
                    .unwrap()
                    .set_outline_level_value(Some(level)),
                "{level}"
            );
            assert_eq!(document.paragraph(0).unwrap().outline_level(), Some(level));
        }
        for rejected in [10, u32::MAX] {
            assert!(
                !document
                    .paragraph_mut(0)
                    .unwrap()
                    .set_outline_level_value(Some(rejected)),
                "{rejected}"
            );
            assert_eq!(document.paragraph(0).unwrap().outline_level(), Some(9));
        }
        assert!(
            document
                .paragraph_mut(0)
                .unwrap()
                .set_outline_level_value(None)
        );
        assert_eq!(document.paragraph(0).unwrap().outline_level(), None);
    }
}

mod settings_and_web_settings_authoring_tests {
    use super::*;
    use rdocx::{
        CharacterSpacingControl, CompatibilityOption, CompatibilitySetting, CryptAlgorithmClass,
        CryptAlgorithmType, CryptProviderType, DocumentProofState, DocumentProtection,
        DocumentView, DocumentZoom, MailMerge, MailMergeDestination, MailMergeDocumentType,
        ProofState, ProtectionMode, ThemeFontLanguage, Twips, ZoomKind,
    };

    const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    const WEB_SETTINGS_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml";

    fn mail_merge() -> MailMerge {
        MailMerge {
            main_document_type: Some(MailMergeDocumentType::FormLetters),
            link_to_query: Some(true),
            data_type: Some("native".to_owned()),
            connect_string: Some("DSN=Contacts".to_owned()),
            query: Some("SELECT * FROM People".to_owned()),
            do_not_suppress_blank_lines: Some(true),
            destination: Some(MailMergeDestination::Printer),
            address_field_name: Some("Address".to_owned()),
            mail_subject: Some("Invitation".to_owned()),
            mail_as_attachment: Some(false),
            view_merged_data: Some(true),
            active_record: Some(3),
            check_errors: Some(2),
        }
    }

    /// Author every supported settings and web settings member.
    fn author_every_supported_member(document: &mut Document) {
        document.set_view(DocumentView::Print).unwrap();
        document
            .set_zoom(DocumentZoom {
                kind: Some(ZoomKind::FullPage),
                percent: Some(120),
            })
            .unwrap();
        document.set_remove_personal_information(true).unwrap();
        document.set_remove_date_and_time(false).unwrap();
        document.set_mirror_margins(true).unwrap();
        document.set_gutter_at_top(true).unwrap();
        document
            .set_proof_state(DocumentProofState {
                spelling: Some(ProofState::Clean),
                grammar: Some(ProofState::Dirty),
            })
            .unwrap();
        document.set_link_styles(true).unwrap();
        document.set_mail_merge_settings(mail_merge()).unwrap();
        document.set_track_revisions(true).unwrap();
        document.set_do_not_track_moves(true).unwrap();
        document.set_do_not_track_formatting(false).unwrap();
        document
            .set_document_protection(DocumentProtection {
                mode: ProtectionMode::Forms,
                enforcement: Some(true),
                formatting: Some(false),
                provider_type: Some(CryptProviderType::RsaAes),
                algorithm_class: Some(CryptAlgorithmClass::Hash),
                algorithm_type: Some(CryptAlgorithmType::Any),
                algorithm_sid: Some(4),
                spin_count: Some(100_000),
                hash: Some("CALLER-HASH".to_owned()),
                salt: Some("CALLER-SALT".to_owned()),
            })
            .unwrap();
        document.set_default_tab_stop(Twips(720)).unwrap();
        document.set_auto_hyphenation(true).unwrap();
        document.set_consecutive_hyphen_limit(2).unwrap();
        document.set_hyphenation_zone(Twips(360)).unwrap();
        document.set_do_not_hyphenate_caps(true).unwrap();
        document.set_default_table_style("TableNormal").unwrap();
        document.set_even_and_odd_headers(true).unwrap();
        document.set_book_fold_rev_printing(true).unwrap();
        document.set_book_fold_printing(true).unwrap();
        document.set_book_fold_printing_sheets(4).unwrap();
        document
            .set_character_spacing_control(CharacterSpacingControl::DoNotCompress)
            .unwrap();
        document.set_update_fields_on_open(Some(true)).unwrap();
        document
            .set_compatibility_option(CompatibilityOption::NoTabHangInd, true)
            .unwrap();
        document
            .set_compatibility_option(CompatibilityOption::CachedColBalance, false)
            .unwrap();
        document
            .set_compatibility_setting(
                "compatibilityMode",
                "http://schemas.microsoft.com/office/word",
                "15",
            )
            .unwrap();
        document.set_document_variable("Customer", "Ada").unwrap();
        document
            .set_theme_font_language(ThemeFontLanguage {
                latin: Some("en-US".to_owned()),
                east_asia: None,
                bidi: None,
            })
            .unwrap();
        document.set_decimal_symbol(".").unwrap();
        document.set_list_separator(",").unwrap();

        document.set_web_encoding("utf-8").unwrap();
        document.set_web_optimize_for_browser(true).unwrap();
        document.set_web_rely_on_vml(false).unwrap();
        document.set_web_allow_png(true).unwrap();
        document.set_web_do_not_rely_on_css(true).unwrap();
        document.set_web_do_not_save_as_single_file(true).unwrap();
        document.set_web_do_not_organize_in_folder(true).unwrap();
        document.set_web_do_not_use_long_file_names(true).unwrap();
        document.set_web_pixels_per_inch(96).unwrap();
        document.set_web_target_screen_size("800x600").unwrap();
        document.set_web_save_smart_tags_as_xml(true).unwrap();
    }

    fn assert_every_supported_member(document: &Document) {
        assert_eq!(document.view(), Some(DocumentView::Print));
        assert_eq!(
            document.zoom(),
            Some(DocumentZoom {
                kind: Some(ZoomKind::FullPage),
                percent: Some(120),
            })
        );
        assert_eq!(document.remove_personal_information(), Some(true));
        assert_eq!(document.remove_date_and_time(), Some(false));
        assert_eq!(document.mirror_margins(), Some(true));
        assert_eq!(document.gutter_at_top(), Some(true));
        assert_eq!(
            document.proof_state(),
            Some(DocumentProofState {
                spelling: Some(ProofState::Clean),
                grammar: Some(ProofState::Dirty),
            })
        );
        assert_eq!(document.link_styles(), Some(true));
        assert_eq!(document.mail_merge_settings(), Some(&mail_merge()));
        assert_eq!(document.track_revisions(), Some(true));
        assert_eq!(document.do_not_track_moves(), Some(true));
        assert_eq!(document.do_not_track_formatting(), Some(false));
        let protection = document.document_protection().unwrap();
        assert_eq!(protection.mode, ProtectionMode::Forms);
        assert_eq!(protection.hash.as_deref(), Some("CALLER-HASH"));
        assert_eq!(protection.salt.as_deref(), Some("CALLER-SALT"));
        assert_eq!(protection.spin_count, Some(100_000));
        assert_eq!(document.default_tab_stop(), Some(Twips(720)));
        assert_eq!(document.consecutive_hyphen_limit(), Some(2));
        assert_eq!(document.hyphenation_zone(), Some(Twips(360)));
        assert_eq!(document.do_not_hyphenate_caps(), Some(true));
        assert_eq!(document.default_table_style(), Some("TableNormal"));
        assert!(document.even_and_odd_headers());
        assert_eq!(document.book_fold_rev_printing(), Some(true));
        assert_eq!(document.book_fold_printing(), Some(true));
        assert_eq!(document.book_fold_printing_sheets(), Some(4));
        assert_eq!(
            document.character_spacing_control(),
            Some(CharacterSpacingControl::DoNotCompress)
        );
        assert_eq!(document.update_fields_on_open(), Some(true));
        assert_eq!(
            document.compatibility_options(),
            [
                (CompatibilityOption::NoTabHangInd, true),
                (CompatibilityOption::CachedColBalance, false),
            ]
        );
        assert_eq!(
            document.compatibility_settings(),
            [CompatibilitySetting {
                name: "compatibilityMode".to_owned(),
                uri: "http://schemas.microsoft.com/office/word".to_owned(),
                value: "15".to_owned(),
            }]
        );
        assert_eq!(document.document_variable("Customer"), Some("Ada"));
        assert_eq!(
            document.theme_font_language().unwrap().latin.as_deref(),
            Some("en-US")
        );
        assert_eq!(document.decimal_symbol(), Some("."));
        assert_eq!(document.list_separator(), Some(","));

        assert_eq!(document.web_encoding(), Some("utf-8"));
        assert_eq!(document.web_optimize_for_browser(), Some(true));
        assert_eq!(document.web_rely_on_vml(), Some(false));
        assert_eq!(document.web_allow_png(), Some(true));
        assert_eq!(document.web_do_not_rely_on_css(), Some(true));
        assert_eq!(document.web_do_not_save_as_single_file(), Some(true));
        assert_eq!(document.web_do_not_organize_in_folder(), Some(true));
        assert_eq!(document.web_do_not_use_long_file_names(), Some(true));
        assert_eq!(document.web_pixels_per_inch(), Some(96));
        assert_eq!(document.web_target_screen_size(), Some("800x600"));
        assert_eq!(document.web_save_smart_tags_as_xml(), Some(true));
    }

    fn part_text(bytes: &[u8], part: &str) -> Option<String> {
        let package = OpcPackage::from_reader(std::io::Cursor::new(bytes.to_vec())).unwrap();
        package
            .get_part(part)
            .map(|xml| String::from_utf8(xml.to_vec()).unwrap())
    }

    #[test]
    fn public_authored_settings_package_reports_no_unmodeled_supported_children() {
        let producer_settings = format!(
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<w:settings xmlns:w="{word}" xmlns:x="urn:producer">"#,
                r#"<x:ext x:before="1"/>"#,
                r#"<w:compat><x:ext x:inside-compat="1"/></w:compat>"#,
                r#"<w:mailMerge><x:ext x:inside-mail-merge="1"/></w:mailMerge>"#,
                r#"<x:ext x:after="1"/>"#,
                r#"</w:settings>"#,
            ),
            word = WORD_NS,
        );
        let producer_web_settings = format!(
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<w:webSettings xmlns:w="{word}" xmlns:x="urn:producer">"#,
                r#"<x:ext x:inside-web-settings="1"/>"#,
                r#"</w:webSettings>"#,
            ),
            word = WORD_NS,
        );

        let mut seeded = Document::new_with_profile(WordCreationProfile::WordCompatible(
            WordPackageClass::Document,
        ));
        seeded.add_paragraph("Settings corpus");
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seeded.to_bytes().unwrap())).unwrap();
        package.set_part("/word/settings.xml", producer_settings.clone().into_bytes());
        package.set_part(
            "/word/webSettings.xml",
            producer_web_settings.clone().into_bytes(),
        );
        package
            .content_types
            .add_override("/word/webSettings.xml", WEB_SETTINGS_CONTENT_TYPE);
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::WEB_SETTINGS, "webSettings.xml");
        let mut buffer = std::io::Cursor::new(Vec::new());
        package.write_to(&mut buffer).unwrap();

        let mut document = Document::from_bytes(buffer.get_ref()).unwrap();
        author_every_supported_member(&mut document);
        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();

        assert_every_supported_member(&reopened);
        assert_eq!(reopened.settings_diagnostics(), &[]);
        assert_eq!(reopened.web_settings_diagnostics(), &[]);

        let settings = part_text(&bytes, "/word/settings.xml").unwrap();
        for retained in [
            r#"<x:ext x:before="1"/>"#,
            r#"<x:ext x:inside-compat="1"/>"#,
            r#"<x:ext x:inside-mail-merge="1"/>"#,
            r#"<x:ext x:after="1"/>"#,
        ] {
            assert!(settings.contains(retained), "{settings}");
        }
        let web_settings = part_text(&bytes, "/word/webSettings.xml").unwrap();
        assert!(
            web_settings.contains(r#"<x:ext x:inside-web-settings="1"/>"#),
            "{web_settings}"
        );
    }

    #[test]
    fn settings_children_serialize_in_schema_sequence_order() {
        let mut document = Document::new_with_profile(WordCreationProfile::WordCompatible(
            WordPackageClass::Document,
        ));
        // Reverse schema order: the last member is authored first.
        document.set_list_separator(",").unwrap();
        document.set_decimal_symbol(".").unwrap();
        document.set_book_fold_printing_sheets(4).unwrap();
        document.set_even_and_odd_headers(true).unwrap();
        document.set_default_tab_stop(Twips(720)).unwrap();
        document.set_mail_merge_settings(mail_merge()).unwrap();
        document.set_link_styles(true).unwrap();
        document.set_mirror_margins(true).unwrap();
        document.set_view(DocumentView::Print).unwrap();
        document
            .set_compatibility_option(CompatibilityOption::CachedColBalance, true)
            .unwrap();
        document
            .set_compatibility_option(CompatibilityOption::NoTabHangInd, true)
            .unwrap();
        document.set_web_save_smart_tags_as_xml(true).unwrap();
        document.set_web_encoding("utf-8").unwrap();

        let bytes = document.to_bytes().unwrap();
        let settings = part_text(&bytes, "/word/settings.xml").unwrap();
        let top_level = [
            "<w:view",
            "<w:mirrorMargins",
            "<w:linkStyles",
            "<w:mailMerge",
            "<w:defaultTabStop",
            "<w:evenAndOddHeaders",
            "<w:bookFoldPrintingSheets",
            "<w:compat",
            "<w:decimalSymbol",
            "<w:listSeparator",
        ];
        assert_ordered(&settings, &top_level);
        assert_ordered(&settings, &["<w:noTabHangInd", "<w:cachedColBalance"]);
        assert_ordered(
            &settings,
            &[
                "<w:mainDocumentType",
                "<w:linkToQuery",
                "<w:dataType",
                "<w:connectString",
                "<w:query",
                "<w:doNotSuppressBlankLines",
                "<w:destination",
                "<w:addressFieldName",
                "<w:mailSubject",
                "<w:mailAsAttachment",
                "<w:viewMergedData",
                "<w:activeRecord",
                "<w:checkErrors",
            ],
        );

        let web_settings = part_text(&bytes, "/word/webSettings.xml").unwrap();
        assert_ordered(&web_settings, &["<w:encoding", "<w:saveSmartTagsAsXml"]);
    }

    fn assert_ordered(xml: &str, names: &[&str]) {
        let mut previous = 0usize;
        for name in names {
            let position = xml
                .find(name)
                .unwrap_or_else(|| panic!("{name} is missing from {xml}"));
            assert!(
                position >= previous,
                "{name} is out of schema order in {xml}"
            );
            previous = position;
        }
    }

    #[test]
    fn web_settings_part_is_created_on_demand_and_pruned_when_empty() {
        let mut document = Document::new_with_profile(WordCreationProfile::WordCompatible(
            WordPackageClass::Document,
        ));
        document.add_paragraph("Web settings on demand");

        let untouched = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(untouched.clone())).unwrap();
        assert!(package.get_part("/word/webSettings.xml").is_none());
        assert!(
            !package
                .get_part_rels("/word/document.xml")
                .unwrap()
                .items
                .iter()
                .any(|relationship| relationship.rel_type == rel_types::WEB_SETTINGS)
        );

        document.set_web_allow_png(true).unwrap();
        let authored = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(authored)).unwrap();
        let part_name = package
            .get_part_rels("/word/document.xml")
            .unwrap()
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::WEB_SETTINGS)
            .map(|relationship| {
                OpcPackage::resolve_rel_target("/word/document.xml", &relationship.target)
            })
            .expect("authored web settings relationship");
        assert!(package.get_part(&part_name).is_some());
        assert_eq!(
            package.content_types.override_for(&part_name),
            Some(WEB_SETTINGS_CONTENT_TYPE)
        );

        assert_eq!(document.remove_web_allow_png().unwrap(), Some(true));
        let pruned = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(pruned)).unwrap();
        assert!(package.get_part(&part_name).is_none());
        assert_eq!(package.content_types.override_for(&part_name), None);
        assert!(
            !package
                .get_part_rels("/word/document.xml")
                .unwrap()
                .items
                .iter()
                .any(|relationship| relationship.rel_type == rel_types::WEB_SETTINGS)
        );
    }

    #[test]
    fn fresh_package_profiles_gain_no_web_settings_part() {
        for profile in [
            WordCreationProfile::WordCompatible(WordPackageClass::Document),
            WordCreationProfile::WordCompatible(WordPackageClass::Template),
            WordCreationProfile::Minimal(WordPackageClass::Document),
            WordCreationProfile::Minimal(WordPackageClass::Template),
        ] {
            let mut document = Document::new_with_profile(profile);
            let bytes = document.to_bytes().unwrap();
            let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
            assert!(
                package
                    .parts
                    .keys()
                    .all(|name| !name.to_ascii_lowercase().contains("websettings")),
                "{profile:?} gained a web settings part"
            );
            assert!(document.web_settings_diagnostics().is_empty());
            assert!(document.web_division_ids().is_empty());
        }
    }

    #[test]
    fn web_settings_divisions_are_reported_for_reference_checking() {
        let producer = format!(
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<w:webSettings xmlns:w="{word}">"#,
                r#"<w:divs><w:div w:id="11"><w:divsChild><w:div w:id="12"/></w:divsChild></w:div>"#,
                r#"<w:div w:id="13"/></w:divs>"#,
                r#"<w:allowPNG/>"#,
                r#"</w:webSettings>"#,
            ),
            word = WORD_NS,
        );
        let mut seeded = Document::new_with_profile(WordCreationProfile::WordCompatible(
            WordPackageClass::Document,
        ));
        seeded.add_paragraph("Divisions");
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seeded.to_bytes().unwrap())).unwrap();
        package.set_part("/word/webSettings.xml", producer.clone().into_bytes());
        package
            .content_types
            .add_override("/word/webSettings.xml", WEB_SETTINGS_CONTENT_TYPE);
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::WEB_SETTINGS, "webSettings.xml");
        let mut buffer = std::io::Cursor::new(Vec::new());
        package.write_to(&mut buffer).unwrap();

        let mut document = Document::from_bytes(buffer.get_ref()).unwrap();
        assert_eq!(document.web_division_ids(), vec![11, 12, 13]);
        document.set_web_encoding("utf-8").unwrap();
        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.web_division_ids(), vec![11, 12, 13]);
        let stored = part_text(&bytes, "/word/webSettings.xml").unwrap();
        assert!(
            stored.contains(r#"<w:divs><w:div w:id="11"><w:divsChild><w:div w:id="12"/></w:divsChild></w:div><w:div w:id="13"/></w:divs>"#),
            "{stored}"
        );

        let empty = Document::new_with_profile(WordCreationProfile::WordCompatible(
            WordPackageClass::Document,
        ));
        assert!(empty.web_division_ids().is_empty());
    }

    #[test]
    fn mirror_margins_gutter_at_top_and_book_fold_round_trip() {
        let producer = format!(
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<w:settings xmlns:w="{word}">"#,
                r#"<w:saveFormsData/><w:hideSpellingErrors/>"#,
                r#"<w:evenAndOddHeaders/><w:characterSpacingControl w:val="doNotCompress"/>"#,
                r#"</w:settings>"#,
            ),
            word = WORD_NS,
        );
        let mut seeded = Document::new_with_profile(WordCreationProfile::WordCompatible(
            WordPackageClass::Document,
        ));
        seeded.add_paragraph("Book fold");
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seeded.to_bytes().unwrap())).unwrap();
        package.set_part("/word/settings.xml", producer.into_bytes());
        let mut buffer = std::io::Cursor::new(Vec::new());
        package.write_to(&mut buffer).unwrap();

        let mut document = Document::from_bytes(buffer.get_ref()).unwrap();
        document.set_mirror_margins(true).unwrap();
        document.set_gutter_at_top(false).unwrap();
        document.set_book_fold_rev_printing(true).unwrap();
        document.set_book_fold_printing(true).unwrap();
        document.set_book_fold_printing_sheets(8).unwrap();

        let bytes = document.to_bytes().unwrap();
        let settings = part_text(&bytes, "/word/settings.xml").unwrap();
        assert_ordered(
            &settings,
            &[
                "<w:saveFormsData",
                "<w:mirrorMargins",
                "<w:gutterAtTop",
                "<w:hideSpellingErrors",
            ],
        );
        assert_ordered(
            &settings,
            &[
                "<w:evenAndOddHeaders",
                "<w:bookFoldRevPrinting",
                "<w:bookFoldPrinting",
                "<w:bookFoldPrintingSheets",
                "<w:characterSpacingControl",
            ],
        );

        let mut reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.mirror_margins(), Some(true));
        assert_eq!(reopened.gutter_at_top(), Some(false));
        assert_eq!(reopened.book_fold_rev_printing(), Some(true));
        assert_eq!(reopened.book_fold_printing(), Some(true));
        assert_eq!(reopened.book_fold_printing_sheets(), Some(8));
        assert_eq!(reopened.settings_diagnostics(), &[]);

        assert_eq!(reopened.remove_mirror_margins().unwrap(), Some(true));
        assert_eq!(reopened.remove_gutter_at_top().unwrap(), Some(false));
        assert_eq!(
            reopened.remove_book_fold_rev_printing().unwrap(),
            Some(true)
        );
        assert_eq!(reopened.remove_book_fold_printing().unwrap(), Some(true));
        assert_eq!(
            reopened.remove_book_fold_printing_sheets().unwrap(),
            Some(8)
        );
        let cleared = reopened.to_bytes().unwrap();
        let settings = part_text(&cleared, "/word/settings.xml").unwrap();
        for absent in [
            "mirrorMargins",
            "gutterAtTop",
            "bookFoldRevPrinting",
            "bookFoldPrinting",
        ] {
            assert!(!settings.contains(absent), "{settings}");
        }
        assert!(settings.contains("<w:saveFormsData/>"), "{settings}");
        assert!(settings.contains("<w:hideSpellingErrors/>"), "{settings}");
    }

    /// First glyph origin of each text run on page one, in layout order.
    fn run_origins(document: &Document) -> Vec<(String, f64)> {
        let layout = document.layout_deterministic().unwrap();
        let mut origins = Vec::new();
        oxml_layout::walk(&layout.layout.pages[0].elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Text(run) = element {
                origins.push((run.text.clone(), run.origin.x));
            }
        });
        origins
    }

    #[test]
    fn document_default_tab_stop_drives_implicit_tab_positions() {
        let run_start = |default_tab_stop: Option<Twips>, text: &str| -> f64 {
            let mut document = Document::new_with_profile(WordCreationProfile::WordCompatible(
                WordPackageClass::Document,
            ));
            {
                let mut paragraph = document.add_paragraph("A");
                let mut run = paragraph.add_run("");
                run.add_tab();
                run.add_text("B");
            }
            if let Some(value) = default_tab_stop {
                document.set_default_tab_stop(value).unwrap();
            }
            run_origins(&document)
                .into_iter()
                .find(|(run_text, _)| run_text == text)
                .unwrap_or_else(|| panic!("{text} must be laid out"))
                .1
        };

        // Deterministic bundled fonts, so the positions below are exact. The
        // paragraph declares no tab stop of its own, so the tab resolves
        // against the document interval alone.
        assert!((run_start(None, "A") - 72.0).abs() < 0.01);

        // An absent setting reproduces the 36.0 point fallback exactly, and
        // Word's own half-inch value lands in the same place.
        let fallback = run_start(None, "B");
        assert!((fallback - 108.0).abs() < 0.01, "{fallback}");
        let word_default = run_start(Some(Twips(720)), "B");
        assert!(
            (word_default - fallback).abs() < f64::EPSILON,
            "{word_default}"
        );

        // A two-inch interval moves the implicit stop onto the wider grid.
        let two_inch = run_start(Some(Twips(2880)), "B");
        assert!((two_inch - 180.0).abs() < 0.01, "{two_inch}");
    }
}

/// F-265, complete run property and inline authoring.
///
/// Every test here reads or authors `w:rPr` children and ordered run content
/// that F-265 modeled, and proves the bytes outside the typed projection
/// survive untouched beside them.
mod f265_run_property_and_inline_tests {
    use super::*;
    use rdocx::{
        RunFontSlot, RunItemRef, ST_PTabAlignment, ST_PTabLeader, ST_PTabRelativeTo,
        SpecialCharacter,
    };
    use rdocx_oxml::properties::{CT_EastAsianLayout, CT_FitText, ST_Em, ST_TextEffect};
    use rdocx_oxml::shared::ST_Border;
    use rdocx_oxml::units::Twips;

    const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// The oracle for the run-formatting gate.
    ///
    /// Microsoft Word GUI capture is not available on this machine, so the
    /// reference is the recorded WordprocessingML for these properties rather
    /// than a fresh save. The no-repair confirmation is tracked as a human
    /// action in `docs/hld/14-development-backlog.md`.
    const WORD_RUN_REFERENCE: &str = "recorded WordprocessingML, ECMA-376 EG_RPrBase";

    /// Every `EG_RPrBase` child F-265 typed, in schema order, interleaved with
    /// unmodelled producer siblings at three schema slots.
    const EVERY_NEW_RUN_PROPERTY: &str = concat!(
        r#"<w:rFonts w:hint="eastAsia" w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="MS Mincho""#,
        r#" w:cs="Arial" w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi""#,
        r#" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>"#,
        r#"<w:outline/><w:shadow/><w:emboss/><w:imprint/><w:noProof/><w:snapToGrid/>"#,
        r#"<w:webHidden/>"#,
        r#"<w:color w:val="4472C4" w:themeColor="accent1" w:themeTint="66" w:themeShade="BF"/>"#,
        r#"<w:kern w:val="16"/>"#,
        r#"<w:effect w:val="antsRed"/>"#,
        r#"<w:bdr w:val="single" w:sz="4" w:space="1" w:color="FF0000"/>"#,
        r#"<w:shd w:val="pct20" w:color="4472C4" w:themeColor="accent1" w:themeTint="66""#,
        r#" w:themeShade="BF" w:fill="ED7D31" w:themeFill="accent2" w:themeFillTint="33""#,
        r#" w:themeFillShade="80"/>"#,
        r#"<w:fitText w:val="1440" w:id="3"/>"#,
        r#"<w:rtl/><w:cs/><w:em w:val="dot"/>"#,
        r#"<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>"#,
        r#"<w:eastAsianLayout w:id="7" w:combine="1" w:combineBrackets="round" w:vert="1""#,
        r#" w:vertCompress="1"/>"#,
        r#"<w:specVanish/><w:oMath/>"#,
    );

    /// Run children that rdocx keeps in positioned raw capture, one before the
    /// typed sequence, one in the middle of it, one after.
    const PRODUCER_RUN_PROPERTY_SIBLINGS: [&str; 3] = [
        r#"<ext:before xmlns:ext="urn:producer" ext:keep="exact"/>"#,
        r#"<ext:middle xmlns:ext="urn:producer" ext:keep="exact"/>"#,
        r#"<ext:after xmlns:ext="urn:producer" ext:keep="exact"/>"#,
    ];

    /// The saved part is indented, so element order is compared with the
    /// inter-element whitespace removed. Text content is left alone.
    fn without_layout_whitespace(xml: &str) -> String {
        let mut output = String::with_capacity(xml.len());
        let mut rest = xml;
        while let Some(start) = rest.find('>') {
            output.push_str(&rest[..=start]);
            rest = &rest[start + 1..];
            let trimmed = rest.trim_start_matches([' ', '\n', '\r', '\t']);
            if trimmed.starts_with('<') {
                rest = trimmed;
            }
        }
        output.push_str(rest);
        output
    }

    fn producer_document(run_properties: &str, run_content: &str) -> Vec<u8> {
        let mut seed = Document::new();
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
        package.set_part(
            "/word/document.xml",
            format!(
                concat!(
                    r#"<w:document xmlns:w="{}">"#,
                    r#"<w:body><w:p><w:r><w:rPr>{}</w:rPr>{}</w:r></w:p>"#,
                    r#"<w:sectPr/></w:body></w:document>"#,
                ),
                WORD_NS, run_properties, run_content
            )
            .into_bytes(),
        );
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        bytes.into_inner()
    }

    fn document_xml(bytes: &[u8]) -> String {
        let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap()
    }

    /// The first `w:r` element of the saved body, verbatim.
    fn run_xml(bytes: &[u8]) -> String {
        let xml = document_xml(bytes);
        let start = xml.find("<w:r>").unwrap();
        let end = xml.find("</w:r>").unwrap() + "</w:r>".len();
        without_layout_whitespace(&xml[start..end])
    }

    /// The story-level test gate. Every modeled run property and inline item
    /// read from the pinned Word reference produces the same effective run
    /// formatting and the same inline order after a save and a reopen, and the
    /// deterministic render of the reopened document is stable.
    #[test]
    fn complete_run_formatting_and_inline_order_match_the_pinned_word_reference() {
        assert!(WORD_RUN_REFERENCE.contains("EG_RPrBase"));
        let inline = concat!(
            r#"<w:t xml:space="preserve">a </w:t>"#,
            r#"<w:sym w:font="Wingdings" w:char="F0FC"/>"#,
            r#"<w:cr/><w:noBreakHyphen/><w:tab/><w:softHyphen/>"#,
            r#"<w:ptab w:alignment="right" w:relativeTo="margin" w:leader="dot"/>"#,
            r#"<w:lastRenderedPageBreak/><w:br/><w:t>b</w:t>"#,
        );
        let source = producer_document(EVERY_NEW_RUN_PROPERTY, inline);

        let mut document = Document::from_bytes(&source).unwrap();
        let saved = document.to_bytes().unwrap();
        assert_eq!(run_xml(&saved), run_xml(&source));

        let reopened = Document::from_bytes(&saved).unwrap();
        let paragraphs = reopened.paragraphs();
        let run = paragraphs[0].runs().next().unwrap();
        assert_eq!(run.slot_font(RunFontSlot::EastAsia), Some("MS Mincho"));
        assert_eq!(
            run.slot_theme_font(RunFontSlot::ComplexScript),
            Some("minorBidi")
        );
        assert_eq!(run.font_hint(), Some("eastAsia"));
        assert_eq!(run.outline_value(), Some(true));
        assert_eq!(run.shadow_value(), Some(true));
        assert_eq!(run.emboss_value(), Some(true));
        assert_eq!(run.imprint_value(), Some(true));
        assert_eq!(run.no_proof_value(), Some(true));
        assert_eq!(run.snap_to_grid_value(), Some(true));
        assert_eq!(run.web_hidden_value(), Some(true));
        assert_eq!(run.color(), Some("4472C4"));
        assert_eq!(run.color_theme(), Some("accent1"));
        assert_eq!(run.color_theme_tint(), Some(0x66));
        assert_eq!(run.color_theme_shade(), Some(0xBF));
        assert_eq!(run.kern(), Some(8.0));
        assert_eq!(run.effect(), Some(&ST_TextEffect::AntsRed));
        assert_eq!(
            run.character_border().map(|border| border.val),
            Some(ST_Border::Single)
        );
        assert_eq!(run.fit_text().map(|fit| fit.val), Some(Twips(1440)));
        assert_eq!(run.rtl_value(), Some(true));
        assert_eq!(run.complex_script_value(), Some(true));
        assert_eq!(run.emphasis_mark(), Some(&ST_Em::Dot));
        assert_eq!(run.language_east_asia(), Some("ja-JP"));
        assert_eq!(run.language_bidi(), Some("ar-SA"));
        assert_eq!(
            run.east_asian_layout().and_then(|layout| layout.id),
            Some(7)
        );
        assert_eq!(run.spec_vanish_value(), Some(true));
        assert_eq!(run.office_math_value(), Some(true));

        let items = run.items().collect::<Vec<_>>();
        assert!(matches!(items[0], RunItemRef::Text("a ")), "{:?}", items[0]);
        assert!(
            matches!(
                items[1],
                RunItemRef::Symbol {
                    font: "Wingdings",
                    char_code: 0xF0FC
                }
            ),
            "{:?}",
            items[1]
        );
        assert!(matches!(
            items[2],
            RunItemRef::SpecialCharacter(SpecialCharacter::CarriageReturn)
        ));
        assert!(matches!(
            items[3],
            RunItemRef::SpecialCharacter(SpecialCharacter::NoBreakHyphen)
        ));
        assert!(matches!(items[4], RunItemRef::Tab));
        assert!(matches!(
            items[5],
            RunItemRef::SpecialCharacter(SpecialCharacter::SoftHyphen)
        ));
        assert!(matches!(
            items[6],
            RunItemRef::SpecialCharacter(SpecialCharacter::PositionalTab {
                alignment: ST_PTabAlignment::Right,
                relative_to: ST_PTabRelativeTo::Margin,
                leader: ST_PTabLeader::Dot,
            })
        ));
        assert!(
            matches!(items[7], RunItemRef::LastRenderedPageBreak(_)),
            "{:?}",
            items[7]
        );
        assert!(matches!(
            items[8],
            RunItemRef::Break(rdocx::BreakKind::Line)
        ));
        assert!(matches!(items[9], RunItemRef::Text("b")));
        assert_eq!(items.len(), 10);

        let reopened = Document::from_bytes(&saved).unwrap();
        let rendered = reopened.to_pdf_deterministic().unwrap();
        assert_eq!(rendered, reopened.to_pdf_deterministic().unwrap());
    }

    #[test]
    fn every_new_run_property_survives_a_no_op_save_beside_untouched_raw_siblings() {
        let interleaved = format!(
            "{}{}{}{}{}",
            PRODUCER_RUN_PROPERTY_SIBLINGS[0],
            &EVERY_NEW_RUN_PROPERTY[..EVERY_NEW_RUN_PROPERTY.find("<w:kern").unwrap()],
            PRODUCER_RUN_PROPERTY_SIBLINGS[1],
            &EVERY_NEW_RUN_PROPERTY[EVERY_NEW_RUN_PROPERTY.find("<w:kern").unwrap()..],
            PRODUCER_RUN_PROPERTY_SIBLINGS[2],
        );
        let source = producer_document(&interleaved, "<w:t>x</w:t>");
        let mut document = Document::from_bytes(&source).unwrap();
        let saved = document.to_bytes().unwrap();
        assert_eq!(run_xml(&saved), run_xml(&source));

        // A run whose only property is one newly modeled child keeps it.
        for single in [
            "<w:outline/>",
            "<w:webHidden/>",
            r#"<w:kern w:val="18"/>"#,
            r#"<w:effect w:val="shimmer"/>"#,
            r#"<w:bdr w:val="double"/>"#,
            r#"<w:fitText w:val="720"/>"#,
            "<w:cs/>",
            r#"<w:em w:val="circle"/>"#,
            r#"<w:eastAsianLayout w:id="1"/>"#,
            "<w:specVanish/>",
            "<w:oMath/>",
        ] {
            let source = producer_document(single, "<w:t>x</w:t>");
            let saved = Document::from_bytes(&source).unwrap().to_bytes().unwrap();
            assert_eq!(run_xml(&saved), run_xml(&source), "{single}");
        }
    }

    #[test]
    fn symbols_and_special_characters_are_publicly_authored_and_reopen_in_order() {
        let mut document = Document::new();
        {
            let mut paragraph = document.add_paragraph("");
            let mut run = paragraph.add_run("start");
            run.add_symbol_char("Wingdings", 0xF0FC);
            run.add_special_character(SpecialCharacter::CarriageReturn);
            run.add_special_character(SpecialCharacter::NoBreakHyphen);
            run.add_special_character(SpecialCharacter::SoftHyphen);
            run.add_special_character(SpecialCharacter::PositionalTab {
                alignment: ST_PTabAlignment::Center,
                relative_to: ST_PTabRelativeTo::Indent,
                leader: ST_PTabLeader::Underscore,
            });
            // The shipped F-260 meaning is unchanged: one Unicode scalar as
            // ordinary text, not a `w:sym`.
            run.add_symbol('\u{2713}');
        }

        let saved = document.to_bytes().unwrap();
        let xml = document_xml(&saved);
        let ordered = [
            "<w:t>start</w:t>",
            r#"<w:sym w:font="Wingdings" w:char="F0FC"/>"#,
            "<w:cr/>",
            "<w:noBreakHyphen/>",
            "<w:softHyphen/>",
            r#"<w:ptab w:alignment="center" w:relativeTo="indent" w:leader="underscore"/>"#,
            "<w:t>\u{2713}</w:t>",
        ]
        .map(|needle| {
            xml.find(needle)
                .unwrap_or_else(|| panic!("{needle} in {xml}"))
        });
        assert!(ordered.windows(2).all(|pair| pair[0] < pair[1]), "{xml}");

        let mut reopened = Document::from_bytes(&saved).unwrap();
        assert_eq!(run_xml(&reopened.to_bytes().unwrap()), run_xml(&saved));
    }

    /// The F-266a contract. Every run property F-266a builds its golden
    /// fixture from is publicly authored, read back and reopened, with no raw
    /// XML left in the run.
    #[test]
    fn f266a_prerequisite_run_properties_author_and_reopen() {
        let mut document = Document::new();
        {
            let mut paragraph = document.add_paragraph("");
            let mut run = paragraph.add_run("multiscript");
            run.set_slot_font(RunFontSlot::Ascii, Some("Arial"));
            run.set_slot_font(RunFontSlot::HighAnsi, Some("Arial"));
            run.set_slot_font(RunFontSlot::EastAsia, Some("MS Mincho"));
            run.set_slot_font(RunFontSlot::ComplexScript, Some("Arial"));
            run.set_font_hint(Some("eastAsia"));
            run.set_rtl_value(Some(true));
            run.set_complex_script_value(Some(true));
            run.set_language_value(Some("en-US"));
            run.set_language_east_asia_value(Some("ja-JP"));
            run.set_language_bidi_value(Some("ar-SA"));
            run.set_bold_cs_value(Some(true));
            run.set_italic_cs_value(Some(true));
            run.set_size_cs_value(Some(14.0));
        }

        let saved = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&saved).unwrap();
        let paragraphs = reopened.paragraphs();
        let run = paragraphs[0].runs().next().unwrap();
        assert_eq!(run.slot_font(RunFontSlot::Ascii), Some("Arial"));
        assert_eq!(run.slot_font(RunFontSlot::HighAnsi), Some("Arial"));
        assert_eq!(run.slot_font(RunFontSlot::EastAsia), Some("MS Mincho"));
        assert_eq!(run.slot_font(RunFontSlot::ComplexScript), Some("Arial"));
        assert_eq!(run.font_hint(), Some("eastAsia"));
        assert_eq!(run.rtl_value(), Some(true));
        assert_eq!(run.complex_script_value(), Some(true));
        assert_eq!(run.language(), Some("en-US"));
        assert_eq!(run.language_east_asia(), Some("ja-JP"));
        assert_eq!(run.language_bidi(), Some("ar-SA"));
        assert_eq!(run.bold_cs_value(), Some(true));
        assert_eq!(run.italic_cs_value(), Some(true));
        assert_eq!(run.size_cs(), Some(14.0));
        assert!(
            run.items()
                .all(|item| !matches!(item, RunItemRef::UnsupportedXml(_))),
            "the F-266a prerequisites must not need raw XML"
        );
    }

    #[test]
    fn every_new_run_property_is_publicly_authored_and_reopens() {
        let mut document = Document::new();
        {
            let mut paragraph = document.add_paragraph("");
            let mut run = paragraph.add_run("authored");
            run.set_outline_value(Some(true));
            run.set_shadow_value(Some(true));
            run.set_emboss_value(Some(false));
            run.set_imprint_value(Some(true));
            run.set_no_proof_value(Some(true));
            run.set_snap_to_grid_value(Some(false));
            run.set_web_hidden_value(Some(true));
            run.set_kern_value(Some(9.0));
            run.set_effect_value(Some(ST_TextEffect::Shimmer));
            run.set_character_border_value(Some(CT_BorderEdge::new(ST_Border::Double)));
            run.set_fit_text_value(Some(CT_FitText {
                val: Twips(1200),
                id: Some(5),
                extra_attributes: Vec::new(),
            }));
            run.set_emphasis_mark_value(Some(ST_Em::UnderDot));
            run.set_east_asian_layout_value(Some(CT_EastAsianLayout {
                combine: Some(true),
                ..Default::default()
            }));
            run.set_spec_vanish_value(Some(true));
            run.set_office_math_value(Some(true));
        }

        let saved = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&saved).unwrap();
        let paragraphs = reopened.paragraphs();
        let run = paragraphs[0].runs().next().unwrap();
        assert_eq!(run.outline_value(), Some(true));
        assert_eq!(run.shadow_value(), Some(true));
        assert_eq!(run.emboss_value(), Some(false));
        assert_eq!(run.imprint_value(), Some(true));
        assert_eq!(run.no_proof_value(), Some(true));
        assert_eq!(run.snap_to_grid_value(), Some(false));
        assert_eq!(run.web_hidden_value(), Some(true));
        assert_eq!(run.kern(), Some(9.0));
        assert_eq!(run.effect(), Some(&ST_TextEffect::Shimmer));
        assert_eq!(
            run.character_border().map(|border| border.val),
            Some(ST_Border::Double)
        );
        assert_eq!(run.fit_text().map(|fit| fit.id), Some(Some(5)));
        assert_eq!(run.emphasis_mark(), Some(&ST_Em::UnderDot));
        assert_eq!(
            run.east_asian_layout().and_then(|layout| layout.combine),
            Some(true)
        );
        assert_eq!(run.spec_vanish_value(), Some(true));
        assert_eq!(run.office_math_value(), Some(true));
    }

    /// A symbol is font-encoded rather than Unicode, so the ODT, EPUB, RTF,
    /// HTML and Markdown projections carry no portable spelling for it. Each
    /// exporter that drops one says so, and the special characters that do
    /// have a spelling reach the output instead.
    #[test]
    fn a_dropped_symbol_is_diagnosed_and_the_special_characters_are_exported() {
        let mut document = Document::new();
        {
            let mut paragraph = document.add_paragraph("");
            let mut run = paragraph.add_run("before");
            run.add_symbol_char("Wingdings", 0xF0FC);
            run.add_special_character(SpecialCharacter::NoBreakHyphen);
            run.add_special_character(SpecialCharacter::CarriageReturn);
        }

        let odt = document.to_odt_bytes().unwrap();
        assert!(
            odt.diagnostics.iter().any(|diagnostic| diagnostic
                .message
                .contains("symbol character was dropped during ODT export")),
            "{:?}",
            odt.diagnostics
        );

        let rtf = document.to_rtf_bytes().unwrap();
        assert!(
            rtf.diagnostics.iter().any(|diagnostic| diagnostic
                .message
                .contains("symbol character was dropped during RTF export")),
            "{:?}",
            rtf.diagnostics
        );
        let rtf_text = String::from_utf8_lossy(&rtf.bytes).into_owned();
        assert!(rtf_text.contains("\\_"), "{rtf_text}");
        assert!(rtf_text.contains("\\line "), "{rtf_text}");
    }

    /// `w:themeTint` and `w:themeShade` were dropped on save before F-265, so
    /// `rdocx_oxml::theme::apply_tint_shade` had no production caller. This
    /// proves the call is live on the Word render path, with its own Word
    /// 0-255 arithmetic rather than the spec-correct DrawingML functions.
    #[test]
    fn a_theme_colour_tint_and_shade_reach_the_rendered_colour() {
        fn rendered(colour: &str) -> Vec<u8> {
            let source = producer_document(colour, "<w:t>tinted</w:t>");
            let document = Document::from_bytes(&source).unwrap();
            document
                .render_page_to_png_deterministic(0, 150.0)
                .unwrap()
                .unwrap()
        }

        let plain = rendered(r#"<w:color w:val="000000" w:themeColor="accent1"/>"#);
        let tinted =
            rendered(r#"<w:color w:val="000000" w:themeColor="accent1" w:themeTint="66"/>"#);
        let shaded =
            rendered(r#"<w:color w:val="000000" w:themeColor="accent1" w:themeShade="BF"/>"#);
        assert_ne!(plain, tinted);
        assert_ne!(plain, shaded);
        assert_ne!(tinted, shaded);
    }

    /// `w:effect`, `w:noProof`, `w:webHidden`, `w:specVanish` and `w:oMath`
    /// are modeled and round-tripped with no visible render projection, which
    /// is what Word prints. The classification cannot rot into an oversight
    /// while this holds.
    #[test]
    fn non_rendering_run_properties_change_no_pixels() {
        fn rendered(apply: bool) -> Vec<u8> {
            let mut document = Document::new();
            {
                let mut paragraph = document.add_paragraph("");
                let mut run = paragraph.add_run("unchanged pixels");
                if apply {
                    run.set_effect_value(Some(ST_TextEffect::Shimmer));
                    run.set_no_proof_value(Some(true));
                    run.set_web_hidden_value(Some(true));
                    run.set_spec_vanish_value(Some(true));
                    run.set_office_math_value(Some(true));
                }
            }
            document.to_pdf_deterministic().unwrap()
        }

        assert_eq!(rendered(true), rendered(false));
    }
}

/// F-269, section page semantics.
///
/// Columns, page borders, line numbering, vertical alignment, mirrored margins
/// and the round-trip-only section children.
mod f269_section_page_semantics {
    use super::*;
    use rdocx_oxml::document::{
        CT_LineNumber, CT_NoteProperties, CT_PageBorders, ST_LineNumberRestart,
        ST_PageBorderDisplay, ST_PageBorderOffset, ST_PageBorderZOrder,
    };
    use rdocx_oxml::table::ST_VerticalJc;
    use rdocx_oxml::units::Twips;

    /// The LibreOffice build this story compares its renders against.
    const F269_LIBREOFFICE_ORACLE: &str =
        "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb";
    /// The rasterizer this story compares its renders through.
    const F269_PDFTOPPM_ORACLE: &str = "pdftoppm version 26.01.0";
    /// Rasterization resolution, high enough to separate adjacent column
    /// tracks and low enough to keep the comparison quick.
    const F269_RASTER_DPI: f64 = 150.0;
    /// Structural similarity floor for the page-semantics render.
    ///
    /// Global-window luminance SSIM over a page of 11 point prose measures
    /// glyph rasterization far more than it measures layout, because two
    /// independent shapers and rasterizers never put the same ink in the same
    /// pixel. The measured agreement is 0.21 and the same page against a blank
    /// sheet scores 0.02, so this floor is a collapse guard an order of
    /// magnitude above a blank render. The layout claim is gated by the ink
    /// block comparison below, which is what actually answers whether the
    /// columns, the rule and the numbers landed where LibreOffice put them.
    const F269_SSIM_FLOOR: f64 = 0.15;
    /// Pixel tolerance for each ink block edge, at the raster resolution.
    ///
    /// Six pixels at 150 DPI is 2.9 points, which covers the glyph edge and
    /// border stroke differences between two renderers without admitting a
    /// misplaced column track, whose nearest error is a 36 point gutter.
    const F269_BLOCK_TOLERANCE_PX: i64 = 6;

    /// Open a document whose `/word/document.xml` is exactly `xml`.
    fn document_from_xml(xml: &str) -> Document {
        let mut seed = Document::new();
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
        package.set_part("/word/document.xml", xml.as_bytes().to_vec());
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        Document::from_bytes(bytes.get_ref()).unwrap()
    }

    /// The saved `/word/document.xml` of a document.
    fn saved_document_xml(document: &mut Document) -> String {
        let package =
            OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
        String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap()
    }

    /// The `w:sectPr` element of a saved document, without the indentation the
    /// writer adds between elements.
    ///
    /// Indentation is a serialisation decision this workspace owns, so it is
    /// removed before the comparison rather than baked into the expectation.
    fn saved_sect_pr(document: &mut Document) -> String {
        let xml = saved_document_xml(document);
        let start = xml.find("<w:sectPr").expect("saved section properties");
        let end = xml.find("</w:sectPr>").expect("saved section end") + "</w:sectPr>".len();
        let mut out = String::with_capacity(end - start);
        let mut pending = String::new();
        for character in xml[start..end].chars() {
            if character.is_whitespace() && !pending.is_empty() {
                continue;
            }
            if character == '>' {
                out.push(character);
                pending.push('>');
                continue;
            }
            if !pending.is_empty() {
                pending.clear();
            }
            out.push(character);
        }
        out
    }

    /// A document whose only body child is a paragraph, followed by `sect_pr`.
    fn document_with_sect_pr(sect_pr: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p>{sect_pr}</w:body></w:document>"#
        )
    }

    /// Left edge of every shaped run on a page, in document order.
    fn glyph_origins(page: &oxml_layout::PageFrame) -> Vec<f64> {
        let mut origins = Vec::new();
        oxml_layout::walk(&page.elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Text(run) = element
                && !run.text.trim().is_empty()
            {
                origins.push(run.origin.x);
            }
        });
        origins
    }

    /// Every text run inside marked content that carries no structure.
    fn artifact_texts(elements: &[oxml_layout::PositionedElement]) -> Vec<(String, f64, f64)> {
        let mut found = Vec::new();
        for element in elements {
            match element {
                oxml_layout::PositionedElement::MarkedContent {
                    structure: None,
                    children,
                } => {
                    for child in children {
                        if let oxml_layout::PositionedElement::Text(run) = child {
                            found.push((run.text.clone(), run.origin.x, run.origin.y));
                        }
                    }
                }
                oxml_layout::PositionedElement::MarkedContent {
                    structure: Some(_),
                    children,
                } => found.extend(artifact_texts(children)),
                oxml_layout::PositionedElement::Group(group) => {
                    found.extend(artifact_texts(&group.children));
                }
                _ => {}
            }
        }
        found
    }

    /// Every vertical rule drawn on a page.
    fn vertical_rules(page: &oxml_layout::PageFrame) -> Vec<f64> {
        let mut rules = Vec::new();
        oxml_layout::walk(&page.elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Line { start, end, .. } = element
                && (start.x - end.x).abs() < f64::EPSILON
                && (end.y - start.y).abs() > 1.0
            {
                rules.push(start.x);
            }
        });
        rules
    }

    /// Fifty short paragraphs, which overflow one column of a Letter page.
    fn fill_with_paragraphs(document: &mut Document, count: usize) {
        for index in 0..count {
            document.add_paragraph(&format!("Line {index:02}"));
        }
    }

    /// Positioned elements of every page, as a comparable record.
    fn page_records(document: &Document) -> Vec<String> {
        document
            .layout_deterministic()
            .unwrap()
            .layout
            .pages
            .iter()
            .map(|page| format!("{}x{} {:?}", page.width, page.height, page.elements))
            .collect()
    }

    #[test]
    fn every_section_property_survives_noop_save() {
        let sect_pr = concat!(
            "<w:sectPr>",
            "<w:footnotePr><w:pos w:val=\"pageBottom\"/><w:numFmt w:val=\"lowerRoman\"/><w:numStart w:val=\"3\"/><w:numRestart w:val=\"eachPage\"/></w:footnotePr>",
            "<w:endnotePr><w:pos w:val=\"docEnd\"/><w:numFmt w:val=\"upperLetter\"/><w:numStart w:val=\"2\"/><w:numRestart w:val=\"eachSect\"/></w:endnotePr>",
            "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>",
            "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:gutter=\"0\" w:header=\"720\" w:footer=\"720\"/>",
            "<w:paperSrc w:first=\"15\" w:other=\"7\"/>",
            "<w:pgBorders w:zOrder=\"back\" w:display=\"notFirstPage\" w:offsetFrom=\"page\">",
            "<w:top w:val=\"single\" w:sz=\"12\" w:space=\"24\" w:color=\"336699\"/>",
            "<w:left w:val=\"single\" w:sz=\"12\" w:space=\"24\" w:color=\"336699\"/>",
            "<w:bottom w:val=\"single\" w:sz=\"12\" w:space=\"24\" w:color=\"336699\"/>",
            "<w:right w:val=\"single\" w:sz=\"12\" w:space=\"24\" w:color=\"336699\"/>",
            "</w:pgBorders>",
            "<w:lnNumType w:countBy=\"5\" w:start=\"2\" w:distance=\"360\" w:restart=\"continuous\"/>",
            "<w:cols w:num=\"2\" w:equalWidth=\"0\" w:sep=\"1\"><w:col w:w=\"4000\" w:space=\"360\"/><w:col w:w=\"4880\"/></w:cols>",
            "<w:vAlign w:val=\"center\"/>",
            "<w:titlePg/>",
            "<w:textDirection w:val=\"lrTb\"/>",
            "</w:sectPr>",
        );
        let mut document = document_from_xml(&document_with_sect_pr(sect_pr));

        let section = document.sections().next().expect("one section");
        let footnotes = section.footnote_properties().expect("footnote properties");
        assert_eq!(footnotes.pos.as_deref(), Some("pageBottom"));
        assert_eq!(footnotes.num_fmt.as_deref(), Some("lowerRoman"));
        assert_eq!(footnotes.num_start, Some(3));
        assert_eq!(footnotes.num_restart.as_deref(), Some("eachPage"));
        let endnotes = section.endnote_properties().expect("endnote properties");
        assert_eq!(endnotes.pos.as_deref(), Some("docEnd"));
        assert_eq!(endnotes.num_fmt.as_deref(), Some("upperLetter"));
        assert_eq!(endnotes.num_start, Some(2));
        assert_eq!(endnotes.num_restart.as_deref(), Some("eachSect"));
        assert_eq!(section.paper_source(), Some((Some(15), Some(7))));
        let borders = section.page_borders().expect("page borders");
        assert_eq!(borders.z_order, Some(ST_PageBorderZOrder::Back));
        assert_eq!(borders.display, Some(ST_PageBorderDisplay::NotFirstPage));
        assert_eq!(borders.offset_from, Some(ST_PageBorderOffset::Page));
        assert_eq!(borders.top.as_ref().unwrap().space, Some(24));
        let numbering = section.line_numbers().expect("line numbering");
        assert_eq!(numbering.count_by, Some(5));
        assert_eq!(numbering.start, Some(2));
        assert_eq!(numbering.distance, Some(Twips(360)));
        assert_eq!(numbering.restart, Some(ST_LineNumberRestart::Continuous));
        assert_eq!(section.vertical_alignment(), Some(ST_VerticalJc::Center));
        assert_eq!(section.text_direction(), Some("lrTb"));
        assert_eq!(
            section.column_widths(),
            Some(vec![
                (Length::twips(4000), Length::twips(360)),
                (Length::twips(4880), Length::twips(0)),
            ])
        );
        assert_eq!(section.column_separator(), Some(true));

        // The saved section is byte identical, including the xsd:sequence
        // child order and every retained attribute.
        assert_eq!(saved_sect_pr(&mut document), sect_pr);
    }

    #[test]
    fn unmodeled_section_children_stay_byte_exact() {
        let sect_pr = concat!(
            "<w:sectPr>",
            "<w:footnotePr><w:pos w:val=\"sectEnd\"/></w:footnotePr>",
            "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>",
            "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:gutter=\"0\" w:header=\"720\" w:footer=\"720\"/>",
            "<w:paperSrc w:first=\"1\"/>",
            "<w:lnNumType w:countBy=\"1\"/>",
            "<w:cols w:num=\"1\" w:space=\"720\"/>",
            "<w:formProt w:val=\"0\"/>",
            "<w:vAlign w:val=\"bottom\"/>",
            "<w:noEndnote w:val=\"1\"/>",
            "<w:titlePg/>",
            "<w:textDirection w:val=\"tbRl\"/>",
            "<w:bidi w:val=\"0\"/>",
            "<w:rtlGutter w:val=\"0\"/>",
            "<w:docGrid w:type=\"lines\" w:linePitch=\"360\"/>",
            "<w:printerSettings r:id=\"rIdPrinter\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>",
            "</w:sectPr>",
        );
        let mut document = document_from_xml(&document_with_sect_pr(sect_pr));
        assert_eq!(saved_sect_pr(&mut document), sect_pr);
    }

    #[test]
    fn variable_width_columns_parse_and_write_in_schema_order() {
        let sect_pr = concat!(
            "<w:sectPr>",
            "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>",
            "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:gutter=\"0\" w:header=\"720\" w:footer=\"720\"/>",
            "<w:cols w:num=\"3\" w:equalWidth=\"0\" w:sep=\"1\"><w:col w:w=\"2000\" w:space=\"180\"/><w:col w:w=\"3000\" w:space=\"180\"/><w:col w:w=\"4000\"/></w:cols>",
            "</w:sectPr>",
        );
        let mut document = document_from_xml(&document_with_sect_pr(sect_pr));
        assert_eq!(saved_sect_pr(&mut document), sect_pr);

        // Prefix tolerant on read, fixed `w:` on write.
        let aliased = sect_pr
            .replace("<w:", "<x:")
            .replace("</w:", "</x:")
            .replace(" w:", " x:");
        let aliased = aliased.replacen(
            "<x:sectPr",
            "<x:sectPr xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"",
            1,
        );
        let mut reparsed = document_from_xml(&document_with_sect_pr(&aliased));
        assert_eq!(saved_sect_pr(&mut reparsed), sect_pr);
    }

    #[test]
    fn section_vertical_alignment_retains_its_source_value() {
        for value in ["top", "center", "both", "bottom"] {
            let sect_pr = format!(
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:vAlign w:val=\"{value}\"/></w:sectPr>"
            );
            let mut document = document_from_xml(&document_with_sect_pr(&sect_pr));
            assert_eq!(
                document
                    .sections()
                    .next()
                    .unwrap()
                    .vertical_alignment()
                    .map(ST_VerticalJc::to_str),
                Some(value),
                "{value} must not collapse"
            );
            assert_eq!(saved_sect_pr(&mut document), sect_pr);
        }
    }

    /// F-269 authors and preserves `w:sectPr/w:textDirection`. F-266c owns the
    /// render projection over it, and asserts it in
    /// `f266c_character_grid_and_vertical_text`, so this test states the
    /// authoring contract and the round trip alone.
    #[test]
    fn section_text_direction_round_trips_as_authored() {
        let mut document = Document::new();
        document.add_paragraph("body");
        document
            .section_mut(0)
            .expect("final section")
            .set_text_direction("tbRl");
        assert_eq!(
            document.sections().next().unwrap().text_direction(),
            Some("tbRl")
        );

        let mut reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.sections().next().unwrap().text_direction(),
            Some("tbRl")
        );
        assert!(saved_sect_pr(&mut reopened).contains("<w:textDirection w:val=\"tbRl\"/>"));
    }

    #[test]
    fn vertical_alignment_both_lays_out_as_top_with_a_diagnostic() {
        let build = |alignment: Option<ST_VerticalJc>| {
            let mut document = Document::new();
            document.add_paragraph("body");
            if let Some(alignment) = alignment {
                document
                    .section_mut(0)
                    .expect("final section")
                    .set_vertical_alignment(alignment);
            }
            document
        };

        let top = build(Some(ST_VerticalJc::Top));
        let both = build(Some(ST_VerticalJc::Both));
        assert_eq!(page_records(&both), page_records(&top));

        let laid_out = both.layout_deterministic().unwrap();
        let reported = laid_out
            .layout
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.message.contains("vertical alignment both"))
            .count();
        assert_eq!(reported, 1, "{:?}", laid_out.layout.diagnostics);

        // The source value survives the diagnostic.
        let mut both = both;
        assert_eq!(
            both.sections().next().unwrap().vertical_alignment(),
            Some(ST_VerticalJc::Both)
        );
        assert!(saved_sect_pr(&mut both).contains("<w:vAlign w:val=\"both\"/>"));

        // A centred section does move, so the comparison above is not vacuous.
        let centred = build(Some(ST_VerticalJc::Center));
        assert_ne!(page_records(&centred), page_records(&top));
    }

    #[test]
    fn page_border_offset_and_display_attributes_round_trip() {
        let sect_pr = concat!(
            "<w:sectPr>",
            "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>",
            "<w:pgBorders w:zOrder=\"front\" w:display=\"firstPage\" w:offsetFrom=\"text\">",
            "<w:top w:shadow=\"1\" w:themeColor=\"accent1\" w:val=\"double\" w:sz=\"18\" w:space=\"1\" w:color=\"auto\"/>",
            "<w:left w:frame=\"1\" w:val=\"dashed\" w:sz=\"6\" w:space=\"4\"/>",
            "</w:pgBorders>",
            "</w:sectPr>",
        );
        let mut document = document_from_xml(&document_with_sect_pr(sect_pr));
        let section = document.sections().next().unwrap();
        let borders = section.page_borders().unwrap();
        assert_eq!(borders.z_order, Some(ST_PageBorderZOrder::Front));
        assert_eq!(borders.display, Some(ST_PageBorderDisplay::FirstPage));
        assert_eq!(borders.offset_from, Some(ST_PageBorderOffset::Text));
        assert_eq!(
            borders.top.as_ref().unwrap().extra_attributes,
            vec![
                ("w:shadow".to_owned(), "1".to_owned()),
                ("w:themeColor".to_owned(), "accent1".to_owned()),
            ]
        );
        assert_eq!(
            borders.left.as_ref().unwrap().extra_attributes,
            vec![("w:frame".to_owned(), "1".to_owned())]
        );
        assert_eq!(saved_sect_pr(&mut document), sect_pr);
    }

    #[test]
    fn variable_width_columns_place_text_in_resolved_tracks() {
        let build = |separator: bool| {
            let mut document = Document::new();
            fill_with_paragraphs(&mut document, 60);
            {
                let mut section = document.section_mut(0).expect("final section");
                section
                    .set_column_widths(&[
                        (Length::twips(3600), Length::twips(720)),
                        (Length::twips(4320), Length::twips(0)),
                    ])
                    .unwrap();
                section.set_column_separator(separator);
            }
            document
        };

        let document = build(true);
        let laid_out = document.layout_deterministic().unwrap();
        let page = &laid_out.layout.pages[0];
        let origins = glyph_origins(page);

        // Track zero starts at the left margin and track one 3600 twips plus
        // 720 twips of gutter to its right, which is 288 points.
        let track_zero = 72.0;
        let track_one = 72.0 + 180.0 + 36.0;
        assert!(
            origins.iter().any(|x| (x - track_zero).abs() < 0.01),
            "{origins:?}"
        );
        assert!(
            origins.iter().any(|x| (x - track_one).abs() < 0.01),
            "{origins:?}"
        );

        // Track zero fills before track one, so every run in track one follows
        // every run in track zero.
        let first_in_track_one = origins
            .iter()
            .position(|x| *x >= track_one)
            .expect("track one receives text");
        assert!(
            origins[..first_in_track_one]
                .iter()
                .all(|x| *x >= track_zero && *x < track_one),
            "{origins:?}"
        );
        assert!(
            origins[first_in_track_one..]
                .iter()
                .all(|x| *x >= track_one),
            "{origins:?}"
        );

        // The rule sits at the midpoint of the gutter.
        let rules = vertical_rules(page);
        assert_eq!(rules.len(), 1, "{rules:?}");
        assert!((rules[0] - (72.0 + 180.0 + 18.0)).abs() < 0.01, "{rules:?}");

        // No rule without `w:sep`.
        let plain = build(false);
        let plain = plain.layout_deterministic().unwrap();
        assert!(vertical_rules(&plain.layout.pages[0]).is_empty());
    }

    #[test]
    fn line_numbering_count_by_and_restart_place_margin_numbers() {
        let build = |restart: ST_LineNumberRestart| {
            let mut document = Document::new();
            fill_with_paragraphs(&mut document, 60);
            document
                .section_mut(0)
                .expect("final section")
                .set_line_numbers(CT_LineNumber {
                    count_by: Some(2),
                    start: Some(1),
                    distance: Some(Twips(360)),
                    restart: Some(restart),
                    extra_attributes: Vec::new(),
                });
            document
        };

        let per_page = build(ST_LineNumberRestart::NewPage);
        let per_page = per_page.layout_deterministic().unwrap();
        assert!(per_page.layout.pages.len() >= 2);

        let first = artifact_texts(&per_page.layout.pages[0].elements);
        assert_eq!(
            first
                .iter()
                .map(|(text, _, _)| text.as_str())
                .take(3)
                .collect::<Vec<_>>(),
            ["2", "4", "6"]
        );
        // Right aligned, 360 twips clear of the left margin.
        for (text, x, _) in &first {
            assert!(*x < 72.0 - 18.0, "{text} at {x}");
        }
        // The numbers are artifacts, so they carry no structure id, which is
        // what `artifact_texts` selected them by.
        assert!(!first.is_empty());

        // `newPage` restarts on every page, `continuous` does not.
        let second_page_new = artifact_texts(&per_page.layout.pages[1].elements);
        assert_eq!(
            second_page_new.first().map(|(text, _, _)| text.as_str()),
            Some("2")
        );
        let continuous = build(ST_LineNumberRestart::Continuous);
        let continuous = continuous.layout_deterministic().unwrap();
        let second_page_continuous = artifact_texts(&continuous.layout.pages[1].elements);
        assert_ne!(
            second_page_continuous
                .first()
                .map(|(text, _, _)| text.as_str()),
            Some("2"),
            "continuous numbering must not restart"
        );

        // Continuous numbering runs on without a gap or a repeat across every
        // page boundary, which is what a body line that was placed but never
        // counted would break.
        let sequence = continuous
            .layout
            .pages
            .iter()
            .flat_map(|page| artifact_texts(&page.elements))
            .map(|(text, _, _)| text.parse::<u32>().expect("a line number"))
            .collect::<Vec<_>>();
        assert!(sequence.len() > 3, "{sequence:?}");
        assert_eq!(sequence[0], 2);
        for pair in sequence.windows(2) {
            assert_eq!(pair[1], pair[0] + 2, "{sequence:?}");
        }
    }

    #[test]
    fn page_border_display_and_z_order_select_where_the_frame_is_drawn() {
        let build = |display: ST_PageBorderDisplay, z_order: ST_PageBorderZOrder| {
            let mut document = Document::new();
            fill_with_paragraphs(&mut document, 120);
            let mut edge = CT_BorderEdge::new(ST_Border::Single);
            edge.sz = Some(12);
            edge.space = Some(24);
            document
                .section_mut(0)
                .expect("final section")
                .set_page_borders(CT_PageBorders {
                    display: Some(display),
                    offset_from: Some(ST_PageBorderOffset::Page),
                    z_order: Some(z_order),
                    top: Some(edge.clone()),
                    left: Some(edge.clone()),
                    bottom: Some(edge.clone()),
                    right: Some(edge),
                    ..CT_PageBorders::default()
                });
            document
        };

        // Pagination wraps loose page furniture in marked content, so the
        // frame edges are counted through the element walk rather than at the
        // top level.
        let frame_lines = |page: &oxml_layout::PageFrame| {
            let mut count = 0;
            oxml_layout::walk(&page.elements, &mut |element, _| {
                if matches!(element, oxml_layout::PositionedElement::Line { .. }) {
                    count += 1;
                }
            });
            count
        };

        for (display, first, later) in [
            (ST_PageBorderDisplay::AllPages, 4, 4),
            (ST_PageBorderDisplay::FirstPage, 4, 0),
            (ST_PageBorderDisplay::NotFirstPage, 0, 4),
        ] {
            let document = build(display, ST_PageBorderZOrder::Front);
            let laid_out = document.layout_deterministic().unwrap();
            assert!(laid_out.layout.pages.len() >= 2);
            assert_eq!(
                frame_lines(&laid_out.layout.pages[0]),
                first,
                "{display:?} on the first page"
            );
            assert_eq!(
                frame_lines(&laid_out.layout.pages[1]),
                later,
                "{display:?} on a later page"
            );
        }

        // `front` draws after the body and `back` draws before it.
        let position = |z_order: ST_PageBorderZOrder| {
            let document = build(ST_PageBorderDisplay::AllPages, z_order);
            let laid_out = document.layout_deterministic().unwrap();
            let mut first_line = None;
            let mut first_text = None;
            let mut index = 0usize;
            oxml_layout::walk(&laid_out.layout.pages[0].elements, &mut |element, _| {
                match element {
                    oxml_layout::PositionedElement::Line { .. } if first_line.is_none() => {
                        first_line = Some(index);
                    }
                    oxml_layout::PositionedElement::Text(_) if first_text.is_none() => {
                        first_text = Some(index);
                    }
                    _ => {}
                }
                index += 1;
            });
            (
                first_line.expect("a frame edge"),
                first_text.expect("body content"),
            )
        };
        let (front_line, front_text) = position(ST_PageBorderZOrder::Front);
        assert!(front_line > front_text, "{front_line} against {front_text}");
        let (back_line, back_text) = position(ST_PageBorderZOrder::Back);
        assert!(back_line < back_text, "{back_line} against {back_text}");
    }

    #[test]
    fn mirrored_margins_swap_inside_and_outside_on_even_pages() {
        let build = |mirrored: bool| {
            let mut document = Document::new();
            fill_with_paragraphs(&mut document, 120);
            document
                .section_mut(0)
                .expect("final section")
                .set_margins(
                    Length::twips(1440),
                    Length::twips(1440),
                    Length::twips(1440),
                    Length::twips(2880),
                )
                .unwrap();
            if mirrored {
                document.set_mirror_margins(true).unwrap();
            }
            document
        };

        let mirrored = build(true);
        let mirrored = mirrored.layout_deterministic().unwrap();
        assert!(mirrored.layout.pages.len() >= 2);
        let left_edge = |page: &oxml_layout::PageFrame| {
            glyph_origins(page)
                .into_iter()
                .fold(f64::INFINITY, f64::min)
        };
        assert!(
            (left_edge(&mirrored.layout.pages[0]) - 144.0).abs() < 0.01,
            "odd page did not keep the inside margin"
        );
        assert!(
            (left_edge(&mirrored.layout.pages[1]) - 72.0).abs() < 0.01,
            "even page did not mirror"
        );
        assert_eq!(mirrored.layout.pages[1].displayed_page_number, 2);

        // Without the setting nothing swaps, so the assertion above is not
        // reporting a coincidence.
        let plain = build(false);
        let plain = plain.layout_deterministic().unwrap();
        for page in &plain.layout.pages {
            assert!(
                (left_edge(page) - 144.0).abs() < 0.01,
                "page {} moved",
                page.page_number
            );
        }
    }

    /// The subject document for the render comparison.
    ///
    /// Columns with a rule, a page border, line numbering, a centred body band
    /// and mirrored margins, all authored in code so no binary fixture is
    /// needed.
    fn f269_render_subject() -> Document {
        let mut document = Document::new();
        for index in 0..60 {
            document.add_paragraph(&format!(
                "Section page semantics sample line {index:02} of the column flow."
            ));
        }
        {
            let mut section = document.section_mut(0).expect("final section");
            section.set_columns(2, Length::twips(720)).unwrap();
            section.set_column_separator(true);
            section.set_line_numbers(CT_LineNumber {
                count_by: Some(5),
                start: Some(1),
                distance: Some(Twips(360)),
                restart: Some(ST_LineNumberRestart::NewPage),
                extra_attributes: Vec::new(),
            });
            let mut edge = CT_BorderEdge::new(ST_Border::Single);
            edge.sz = Some(12);
            edge.space = Some(24);
            section.set_page_borders(CT_PageBorders {
                display: Some(ST_PageBorderDisplay::AllPages),
                offset_from: Some(ST_PageBorderOffset::Page),
                top: Some(edge.clone()),
                left: Some(edge.clone()),
                bottom: Some(edge.clone()),
                right: Some(edge),
                ..CT_PageBorders::default()
            });
            section.set_vertical_alignment(ST_VerticalJc::Top);
        }
        document.set_mirror_margins(true).unwrap();
        document
    }

    #[test]
    fn section_page_semantics_match_pinned_libreoffice_render() {
        let version = std::process::Command::new("soffice")
            .arg("--version")
            .output()
            .expect("pinned LibreOffice is installed");
        assert!(version.status.success());
        assert_eq!(
            String::from_utf8_lossy(&version.stdout).trim(),
            F269_LIBREOFFICE_ORACLE
        );
        let rasterizer = std::process::Command::new("pdftoppm")
            .arg("-v")
            .output()
            .expect("pinned rasterizer is installed");
        assert!(rasterizer.status.success());
        assert_eq!(
            String::from_utf8_lossy(&rasterizer.stderr).lines().next(),
            Some(F269_PDFTOPPM_ORACLE)
        );

        let root = std::env::temp_dir().join(format!(
            "rdocx-f269-render-{}-{}",
            std::process::id(),
            std::thread::current().name().unwrap_or("test")
        ));
        let _ = std::fs::remove_dir_all(&root);
        let output = root.join("output");
        let profile = root.join("profile");
        std::fs::create_dir_all(&output).unwrap();
        std::fs::create_dir_all(&profile).unwrap();

        let mut document = f269_render_subject();
        let source = root.join("source.docx");
        std::fs::write(&source, document.to_bytes().unwrap()).unwrap();
        let ours = root.join("ours.pdf");
        std::fs::write(&ours, document.to_pdf_deterministic().unwrap()).unwrap();

        let status = std::process::Command::new("soffice")
            .arg("--headless")
            .arg(format!(
                "-env:UserInstallation=file://{}",
                profile.display()
            ))
            .arg("--convert-to")
            .arg("pdf:writer_pdf_Export")
            .arg("--outdir")
            .arg(&output)
            .arg(&source)
            .status()
            .expect("LibreOffice conversion starts");
        assert!(status.success());
        let oracle = output.join("source.pdf");

        let rasterize = |pdf: &std::path::Path, prefix: &str| {
            let target = root.join(prefix);
            let rendered = std::process::Command::new("pdftoppm")
                .args(["-png", "-f", "1", "-l", "1", "-r"])
                .arg(F269_RASTER_DPI.to_string())
                .arg(pdf)
                .arg(&target)
                .output()
                .expect("rasterize page one");
            assert!(
                rendered.status.success(),
                "{}",
                String::from_utf8_lossy(&rendered.stderr)
            );
            root.join(format!("{prefix}-1.png"))
        };
        let ours_png = rasterize(&ours, "ours");
        let oracle_png = rasterize(&oracle, "oracle");

        // Two records per raster: the global SSIM, then the horizontal ink
        // blocks of the page interior. The blocks are the line-number band,
        // each column track and the rule between them.
        let script = r#"import sys
from pathlib import Path
sys.path.insert(0, sys.argv[3])
from golden_png_harness import decode_png
from pptx_ssim_harness import composite_luminance, structural_similarity

BORDER_INSET = 60
INK = 200
MINIMUM_COLUMN_INK = 2
BLOCK_GAP = 30

first = decode_png(Path(sys.argv[1]))
second = decode_png(Path(sys.argv[2]))
width = min(first[0], second[0])
height = min(first[1], second[1])

def crop(image):
    image_width, _, rgba = image
    rows = []
    for y in range(height):
        start = (y * image_width) * 4
        rows.append(rgba[start : start + width * 4])
    return (width, height, b"".join(rows))

def ink_blocks(image):
    image_width, image_height, rgba = image
    luminance = composite_luminance(rgba)
    inked = []
    for x in range(BORDER_INSET, image_width - BORDER_INSET):
        count = sum(
            1
            for y in range(BORDER_INSET, image_height - BORDER_INSET)
            if luminance[y * image_width + x] < INK
        )
        if count > MINIMUM_COLUMN_INK:
            inked.append(x)
    blocks = []
    for x in inked:
        if blocks and x - blocks[-1][1] <= BLOCK_GAP:
            blocks[-1][1] = x
        else:
            blocks.append([x, x])
    return blocks

print(structural_similarity(crop(first), crop(second)))
for image in (first, second):
    print(" ".join(f"{start},{end}" for start, end in ink_blocks(image)))
"#;
        let scripts = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../scripts")
            .canonicalize()
            .unwrap();
        let measured = std::process::Command::new("python3")
            .arg("-c")
            .arg(script)
            .arg(&ours_png)
            .arg(&oracle_png)
            .arg(&scripts)
            .output()
            .expect("structural similarity runs");
        assert!(
            measured.status.success(),
            "{}",
            String::from_utf8_lossy(&measured.stderr)
        );
        let reported = String::from_utf8_lossy(&measured.stdout);
        let mut lines = reported.lines();
        let ssim: f64 = lines
            .next()
            .expect("structural similarity line")
            .trim()
            .parse()
            .expect("structural similarity is a number");
        let parse_blocks = |line: &str| {
            line.split_whitespace()
                .map(|block| {
                    let (start, end) = block.split_once(',').expect("ink block bounds");
                    (
                        start.parse::<i64>().expect("ink block start"),
                        end.parse::<i64>().expect("ink block end"),
                    )
                })
                .collect::<Vec<_>>()
        };
        let ours_blocks = parse_blocks(lines.next().expect("subject ink blocks"));
        let oracle_blocks = parse_blocks(lines.next().expect("oracle ink blocks"));
        let _ = std::fs::remove_dir_all(&root);

        println!(
            "F-269 section page semantics SSIM {ssim}, subject {ours_blocks:?}, oracle {oracle_blocks:?}"
        );
        assert!(
            ssim >= F269_SSIM_FLOOR,
            "structural similarity {ssim} fell below {F269_SSIM_FLOOR}"
        );

        // The line-number band, both column tracks and the rule between them.
        assert_eq!(
            ours_blocks.len(),
            4,
            "expected a number band, two tracks and a rule: {ours_blocks:?}"
        );
        assert_eq!(
            oracle_blocks.len(),
            ours_blocks.len(),
            "LibreOffice found different page structure: {oracle_blocks:?}"
        );
        for (index, (ours, oracle)) in ours_blocks.iter().zip(&oracle_blocks).enumerate() {
            assert!(
                (ours.0 - oracle.0).abs() <= F269_BLOCK_TOLERANCE_PX
                    && (ours.1 - oracle.1).abs() <= F269_BLOCK_TOLERANCE_PX,
                "ink block {index} disagrees: {ours:?} against {oracle:?}"
            );
        }
    }

    #[test]
    #[ignore = "requires installed Microsoft Word GUI automation, which this machine does not have"]
    fn capture_f269_word_section_evidence() {
        // The mandatory human action recorded as a follow-up in
        // `docs/hld/14-development-backlog.md`. It asserts the Word build
        // before it records anything, and it is never part of the gate.
        let build = std::process::Command::new("plutil")
            .args([
                "-extract",
                "CFBundleShortVersionString",
                "raw",
                "/Applications/Microsoft Word.app/Contents/Info.plist",
            ])
            .output()
            .expect("read the installed Word version");
        assert!(build.status.success());
        let version = String::from_utf8_lossy(&build.stdout).trim().to_owned();
        assert!(!version.is_empty(), "Word reports no version");

        let output = std::env::var("RDOCX_F269_WORD_DOCX")
            .expect("set RDOCX_F269_WORD_DOCX to a temporary output path");
        let mut document = f269_render_subject();
        document.save(&output).unwrap();
        println!("F-269 Word section evidence source: {output} against Word {version}");
    }

    #[test]
    fn section_footnote_and_endnote_properties_are_authorable() {
        let mut document = Document::new();
        document.add_paragraph("body");
        {
            let mut section = document.section_mut(0).expect("final section");
            section.set_footnote_properties(CT_NoteProperties {
                pos: Some("beneathText".to_owned()),
                num_fmt: Some("chicago".to_owned()),
                num_start: Some(4),
                num_restart: Some("eachSect".to_owned()),
                extra_xml: Vec::new(),
            });
            section.set_endnote_properties(CT_NoteProperties {
                pos: Some("sectEnd".to_owned()),
                ..CT_NoteProperties::default()
            });
            section.set_paper_source(Some(4), None);
            section.set_page_borders(CT_PageBorders {
                display: Some(ST_PageBorderDisplay::AllPages),
                top: Some(CT_BorderEdge::new(ST_Border::Single)),
                ..CT_PageBorders::default()
            });
        }
        let mut reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let section = reopened.sections().next().unwrap();
        assert_eq!(
            section
                .footnote_properties()
                .unwrap()
                .num_restart
                .as_deref(),
            Some("eachSect")
        );
        assert_eq!(
            section.endnote_properties().unwrap().pos.as_deref(),
            Some("sectEnd")
        );
        assert_eq!(section.paper_source(), Some((Some(4), None)));
        assert_eq!(
            section.page_borders().unwrap().display,
            Some(ST_PageBorderDisplay::AllPages)
        );
        let saved = saved_sect_pr(&mut reopened);
        let footnote = saved.find("<w:footnotePr>").expect("footnotePr written");
        let endnote = saved.find("<w:endnotePr>").expect("endnotePr written");
        let paper = saved.find("<w:paperSrc").expect("paperSrc written");
        let borders = saved.find("<w:pgBorders").expect("pgBorders written");
        assert!(
            footnote < endnote && endnote < paper && paper < borders,
            "{saved}"
        );
    }
}

/// F-267. Table style and conditional formatting authoring.
mod f267_table_style_conditional_tests {
    use super::*;
    use rdocx::table::TableLook;
    use rdocx_oxml::table::{CT_TblCellMar, CT_TrPr};
    use rdocx_oxml::units::{HalfPoint, Twips};

    /// The LibreOffice build this story compares its renders against.
    const F267_LIBREOFFICE_ORACLE: &str =
        "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb";
    /// The rasterizer this story compares its renders through.
    const F267_PDFTOPPM_ORACLE: &str = "pdftoppm version 26.01.0";
    /// Rasterization resolution, high enough to separate adjacent table rows
    /// and low enough to keep the comparison quick.
    const F267_RASTER_DPI: f64 = 150.0;

    /// Word GUI capture is not available on this machine, so the reference is
    /// the `w:tblStylePr` tree Word writes, pinned here as source XML. The
    /// structural side of the gate asserts against this tree. The confirmation
    /// that Word itself reopens the authored package without offering to
    /// repair it is a tracked human action recorded in
    /// `docs/hld/14-development-backlog.md`, not a gate this machine can run.
    const F267_WORD_REFERENCE_REGIONS: &str = concat!(
        r#"<w:tblStylePr w:type="wholeTable"><w:pPr><w:spacing w:after="10"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="10" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="band1Vert"><w:pPr><w:spacing w:after="11"/></w:pPr><w:rPr><w:sz w:val="21"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="11" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="E1E1F1"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="band2Vert"><w:pPr><w:spacing w:after="12"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="12" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="E2E2F2"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="band1Horz"><w:pPr><w:spacing w:after="13"/></w:pPr><w:rPr><w:sz w:val="23"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="13" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D1F1D1"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="band2Horz"><w:pPr><w:spacing w:after="14"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="14" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D2F2D2"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="firstCol"><w:pPr><w:spacing w:after="15"/></w:pPr><w:rPr><w:sz w:val="25"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="15" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="C1C1F1"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="lastCol"><w:pPr><w:spacing w:after="16"/></w:pPr><w:rPr><w:sz w:val="26"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="16" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="C2C2F2"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="firstRow"><w:pPr><w:spacing w:after="17"/></w:pPr><w:rPr><w:sz w:val="27"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="17" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="B1F1B1"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="lastRow"><w:pPr><w:spacing w:after="18"/></w:pPr><w:rPr><w:sz w:val="28"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="18" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="B2F2B2"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="nwCell"><w:pPr><w:spacing w:after="19"/></w:pPr><w:rPr><w:sz w:val="29"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="19" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="A10000"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="neCell"><w:pPr><w:spacing w:after="20"/></w:pPr><w:rPr><w:sz w:val="30"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="20" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="A20000"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="swCell"><w:pPr><w:spacing w:after="21"/></w:pPr><w:rPr><w:sz w:val="31"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="21" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="A30000"/></w:tcPr></w:tblStylePr>"#,
        r#"<w:tblStylePr w:type="seCell"><w:pPr><w:spacing w:after="22"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr><w:tblPr><w:tblCellMar><w:top w:w="22" w:type="dxa"/></w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="A40000"/></w:tcPr></w:tblStylePr>"#,
    );

    /// One region of the pinned reference, in Word's priority order.
    ///
    /// The tuple is the region, its `w:spacing w:after`, its `w:sz`, its
    /// `w:tblCellMar w:top` and its `w:shd w:fill`, which is what the five
    /// layers of the reference carry.
    const F267_REGIONS: [(TableStyleRegion, i32, u32, i32, &str); 13] = [
        (TableStyleRegion::WholeTable, 10, 20, 10, "EEEEEE"),
        (TableStyleRegion::Band1Vert, 11, 21, 11, "E1E1F1"),
        (TableStyleRegion::Band2Vert, 12, 22, 12, "E2E2F2"),
        (TableStyleRegion::Band1Horz, 13, 23, 13, "D1F1D1"),
        (TableStyleRegion::Band2Horz, 14, 24, 14, "D2F2D2"),
        (TableStyleRegion::FirstCol, 15, 25, 15, "C1C1F1"),
        (TableStyleRegion::LastCol, 16, 26, 16, "C2C2F2"),
        (TableStyleRegion::FirstRow, 17, 27, 17, "B1F1B1"),
        (TableStyleRegion::LastRow, 18, 28, 18, "B2F2B2"),
        (TableStyleRegion::NwCell, 19, 29, 19, "A10000"),
        (TableStyleRegion::NeCell, 20, 30, 20, "A20000"),
        (TableStyleRegion::SwCell, 21, 31, 21, "A30000"),
        (TableStyleRegion::SeCell, 22, 32, 22, "A40000"),
    ];

    fn f267_shading(fill: &str) -> CT_Shd {
        CT_Shd {
            val: "clear".to_owned(),
            color: Some("auto".to_owned()),
            fill: Some(fill.to_owned()),
            ..CT_Shd::default()
        }
    }

    /// Author the pinned reference's thirteen regions through the facade.
    fn f267_authored_style() -> StyleBuilder {
        let mut builder = StyleBuilder::table("RegionGrid", "Region Grid");
        for (region, after, size, margin, fill) in F267_REGIONS {
            builder = builder.conditional_table_style(
                region,
                Some(CT_PPr {
                    space_after: Some(Twips(after)),
                    ..CT_PPr::default()
                }),
                Some(CT_RPr {
                    sz: Some(HalfPoint(size)),
                    ..CT_RPr::default()
                }),
                Some(CT_TblPr {
                    cell_margin: Some(CT_TblCellMar {
                        top: Some(Twips(margin)),
                        ..CT_TblCellMar::default()
                    }),
                    ..CT_TblPr::default()
                }),
                Some(CT_TrPr {
                    cant_split: Some(true),
                    header: (region == TableStyleRegion::FirstRow).then_some(true),
                    ..CT_TrPr::default()
                }),
                Some(CT_TcPr {
                    shading: Some(f267_shading(fill)),
                    ..CT_TcPr::default()
                }),
            );
        }
        builder
    }

    /// Reopen a package whose `RegionGrid` style is the pinned reference.
    fn f267_reference_document() -> Document {
        let mut seed = Document::new();
        seed.add_style(StyleBuilder::table("RegionGrid", "Region Grid"))
            .expect("seed table style");
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
        let styles =
            String::from_utf8(package.get_part("/word/styles.xml").unwrap().to_vec()).unwrap();
        let injected = styles.replace(
            r#"<w:name w:val="Region Grid"/>"#,
            &format!(r#"<w:name w:val="Region Grid"/>{F267_WORD_REFERENCE_REGIONS}"#),
        );
        package.set_part("/word/styles.xml", injected.into_bytes());
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        Document::from_bytes(bytes.get_ref()).expect("the reference reopens")
    }

    /// The five typed layers of one region, as comparable scalars.
    type RegionLayers = (
        Option<TableStyleRegion>,
        Option<Twips>,
        Option<HalfPoint>,
        Option<Twips>,
        Option<bool>,
        Option<bool>,
        Option<String>,
    );

    fn f267_region_layers(style: &rdocx::Style<'_>) -> Vec<RegionLayers> {
        style
            .conditional_table_styles()
            .iter()
            .map(|region| {
                (
                    region.region(),
                    region
                        .paragraph_properties()
                        .and_then(|properties| properties.space_after),
                    region.run_properties().and_then(|properties| properties.sz),
                    region
                        .table_properties()
                        .and_then(|properties| properties.cell_margin.as_ref())
                        .and_then(|margins| margins.top),
                    region
                        .row_properties()
                        .and_then(|properties| properties.cant_split),
                    region
                        .row_properties()
                        .and_then(|properties| properties.header),
                    region
                        .cell_properties()
                        .and_then(|properties| properties.shading.as_ref())
                        .and_then(|shading| shading.fill.clone()),
                )
            })
            .collect()
    }

    /// A four by four table selecting every conditional region.
    fn f267_region_table(style: StyleBuilder) -> Document {
        let mut document = Document::new();
        document.add_style(style).expect("region style is valid");
        let mut table = document.add_table(4, 4);
        table.set_style("RegionGrid");
        table.set_look(TableLook {
            first_row: true,
            last_row: true,
            first_column: true,
            last_column: true,
            horizontal_banding: true,
            vertical_banding: true,
        });
        for row in 0..4 {
            for column in 0..4 {
                table
                    .row(row)
                    .unwrap()
                    .cell(column)
                    .unwrap()
                    .set_text(&format!("r{row}c{column}"));
            }
        }
        document
    }

    /// Flatten grouped and marked page content into drawable elements.
    fn f267_page_elements(
        elements: &[oxml_layout::PositionedElement],
    ) -> Vec<&oxml_layout::PositionedElement> {
        fn collect<'a>(
            elements: &'a [oxml_layout::PositionedElement],
            output: &mut Vec<&'a oxml_layout::PositionedElement>,
        ) {
            for element in elements {
                match element {
                    oxml_layout::PositionedElement::MarkedContent { children, .. } => {
                        collect(children, output)
                    }
                    oxml_layout::PositionedElement::Group(group) => {
                        collect(&group.children, output)
                    }
                    other => output.push(other),
                }
            }
        }
        let mut output = Vec::new();
        collect(elements, &mut output);
        output
    }

    /// The resolved fill of every cell, in row then column order.
    fn f267_cell_fills(document: &Document) -> Vec<String> {
        let layout = document
            .layout_with_fonts_and_bundled_fallback(&[])
            .expect("deterministic region layout");
        let mut fills = f267_page_elements(&layout.layout.pages[0].elements)
            .into_iter()
            .filter_map(|element| match element {
                oxml_layout::PositionedElement::FilledRect { rect, color } => Some((
                    (rect.y * 100.0).round() as i64,
                    (rect.x * 100.0).round() as i64,
                    format!(
                        "{:02X}{:02X}{:02X}",
                        (color.r * 255.0).round() as u8,
                        (color.g * 255.0).round() as u8,
                        (color.b * 255.0).round() as u8
                    ),
                )),
                _ => None,
            })
            .collect::<Vec<_>>();
        fills.sort();
        fills.into_iter().map(|(_, _, fill)| fill).collect()
    }

    #[test]
    fn every_conditional_table_region_matches_word() {
        // 1. The pinned Word-authored tree projects to all five layers.
        let reference = f267_reference_document();
        let style = reference.style("RegionGrid").unwrap();
        let reference_layers = f267_region_layers(&style);
        assert_eq!(reference_layers.len(), 13);
        for (index, (region, after, size, margin, fill)) in F267_REGIONS.into_iter().enumerate() {
            assert_eq!(
                reference_layers[index],
                (
                    Some(region),
                    Some(Twips(after)),
                    Some(HalfPoint(size)),
                    Some(Twips(margin)),
                    Some(true),
                    (region == TableStyleRegion::FirstRow).then_some(true),
                    Some(fill.to_owned()),
                ),
                "{}",
                region.to_str()
            );
        }

        // 2. An untouched reference serialises back as its original bytes.
        let mut reopened = reference;
        let saved = reopened.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        let styles =
            String::from_utf8(package.get_part("/word/styles.xml").unwrap().to_vec()).unwrap();
        assert!(styles.contains(F267_WORD_REFERENCE_REGIONS), "{styles}");

        // 3. The authored tree agrees with the reference tree. The comparison
        //    is over the parsed tree, not the bytes, because attribute order
        //    and prefix choice are ours to decide.
        let mut authored = Document::new();
        authored
            .add_style(f267_authored_style())
            .expect("authored regions are valid");
        let authored = Document::from_bytes(&authored.to_bytes().unwrap()).unwrap();
        let authored_style = authored.style("RegionGrid").unwrap();
        assert_eq!(f267_region_layers(&authored_style), reference_layers);

        // 4. The authored `w:tblStylePr` keeps the schema sequence.
        let mut authored_bytes = authored;
        let package =
            OpcPackage::from_reader(std::io::Cursor::new(authored_bytes.to_bytes().unwrap()))
                .unwrap();
        let styles =
            String::from_utf8(package.get_part("/word/styles.xml").unwrap().to_vec()).unwrap();
        let first_row = styles
            .split(r#"<w:tblStylePr w:type="firstRow">"#)
            .nth(1)
            .and_then(|tail| tail.split("</w:tblStylePr>").next())
            .expect("an authored firstRow region");
        let order = ["<w:pPr>", "<w:rPr>", "<w:tblPr>", "<w:trPr>", "<w:tcPr>"].map(|tag| {
            first_row
                .find(tag)
                .unwrap_or_else(|| panic!("{tag}: {first_row}"))
        });
        assert!(
            order.windows(2).all(|pair| pair[0] < pair[1]),
            "{first_row}"
        );

        // 5. Every region resolves onto the cell Word's priority order picks.
        let resolved = f267_cell_fills(&f267_region_table(f267_authored_style()));
        assert_eq!(
            resolved,
            [
                // Row 0 is the header row, so no horizontal band applies.
                "A10000", "B1F1B1", "B1F1B1", "A20000", //
                // Row 1 is the first banded row, column 0 the first column.
                "C1C1F1", "D1F1D1", "D1F1D1", "C2C2F2", //
                // Row 2 is the second banded row.
                "C1C1F1", "D2F2D2", "D2F2D2", "C2C2F2", //
                // Row 3 is the last row, whose corners outrank it.
                "A30000", "B2F2B2", "B2F2B2", "A40000",
            ]
            .map(str::to_owned)
        );

        // The whole-table region is what a table with no edges and no bands
        // resolves to.
        let mut plain = Document::new();
        plain
            .add_style(f267_authored_style())
            .expect("region style is valid");
        let mut table = plain.add_table(1, 1);
        table.set_style("RegionGrid");
        table.set_look(TableLook {
            first_row: false,
            last_row: false,
            first_column: false,
            last_column: false,
            horizontal_banding: false,
            vertical_banding: false,
        });
        table.row(0).unwrap().cell(0).unwrap().set_text("plain");
        assert_eq!(f267_cell_fills(&plain), vec!["EEEEEE".to_owned()]);

        // 6. The raster side, against the pinned oracle.
        f267_assert_render_matches_oracle(f267_region_table(f267_authored_style()));
    }

    /// Rasterize the styled table through the pinned oracle and compare.
    fn f267_assert_render_matches_oracle(mut document: Document) {
        let version = std::process::Command::new("soffice")
            .arg("--version")
            .output()
            .expect("pinned LibreOffice is installed");
        assert!(version.status.success());
        assert_eq!(
            String::from_utf8_lossy(&version.stdout).trim(),
            F267_LIBREOFFICE_ORACLE
        );
        let rasterizer = std::process::Command::new("pdftoppm")
            .arg("-v")
            .output()
            .expect("pinned rasterizer is installed");
        assert!(rasterizer.status.success());
        assert_eq!(
            String::from_utf8_lossy(&rasterizer.stderr).lines().next(),
            Some(F267_PDFTOPPM_ORACLE)
        );

        let root = std::env::temp_dir().join(format!(
            "rdocx-f267-render-{}-{}",
            std::process::id(),
            std::thread::current().name().unwrap_or("test")
        ));
        let _ = std::fs::remove_dir_all(&root);
        let output = root.join("output");
        let profile = root.join("profile");
        std::fs::create_dir_all(&output).unwrap();
        std::fs::create_dir_all(&profile).unwrap();

        let source = root.join("source.docx");
        std::fs::write(&source, document.to_bytes().unwrap()).unwrap();
        let ours = root.join("ours.pdf");
        std::fs::write(&ours, document.to_pdf_deterministic().unwrap()).unwrap();

        let status = std::process::Command::new("soffice")
            .arg("--headless")
            .arg(format!(
                "-env:UserInstallation=file://{}",
                profile.display()
            ))
            .arg("--convert-to")
            .arg("pdf:writer_pdf_Export")
            .arg("--outdir")
            .arg(&output)
            .arg(&source)
            .status()
            .expect("LibreOffice conversion starts");
        assert!(status.success());
        let oracle = output.join("source.pdf");

        let rasterize = |pdf: &std::path::Path, prefix: &str| {
            let target = root.join(prefix);
            let rendered = std::process::Command::new("pdftoppm")
                .args(["-png", "-f", "1", "-l", "1", "-r"])
                .arg(F267_RASTER_DPI.to_string())
                .arg(pdf)
                .arg(&target)
                .output()
                .expect("rasterize page one");
            assert!(
                rendered.status.success(),
                "{}",
                String::from_utf8_lossy(&rendered.stderr)
            );
            root.join(format!("{prefix}-1.png"))
        };
        let ours_png = rasterize(&ours, "ours");
        let oracle_png = rasterize(&oracle, "oracle");

        // Two records per raster: the global SSIM, then the vertical extent of
        // every band of shaded rows in the page interior. The bands are what
        // the conditional regions paint, so agreeing on where they start and
        // end is the layout claim.
        let script = r#"import sys
from pathlib import Path
sys.path.insert(0, sys.argv[3])
from golden_png_harness import decode_png
from pptx_ssim_harness import composite_luminance, structural_similarity

SHADED = 250
MINIMUM_ROW_INK = 40
BLOCK_GAP = 4

first = decode_png(Path(sys.argv[1]))
second = decode_png(Path(sys.argv[2]))
width = min(first[0], second[0])
height = min(first[1], second[1])

def crop(image):
    image_width, _, rgba = image
    rows = []
    for y in range(height):
        start = (y * image_width) * 4
        rows.append(rgba[start : start + width * 4])
    return (width, height, b"".join(rows))

def shaded_rows(image):
    image_width, image_height, rgba = image
    luminance = composite_luminance(rgba)
    inked = []
    for y in range(image_height):
        count = sum(
            1
            for x in range(image_width)
            if luminance[y * image_width + x] < SHADED
        )
        if count > MINIMUM_ROW_INK:
            inked.append(y)
    blocks = []
    for y in inked:
        if blocks and y - blocks[-1][1] <= BLOCK_GAP:
            blocks[-1][1] = y
        else:
            blocks.append([y, y])
    return blocks

print(structural_similarity(crop(first), crop(second)))
for image in (first, second):
    print(" ".join(f"{start},{end}" for start, end in shaded_rows(image)))
for image in (first, second):
    image_width, image_height, rgba = image
    counts = {}
    for index in range(0, len(rgba), 4):
        pixel = rgba[index : index + 3]
        counts[bytes(pixel)] = counts.get(bytes(pixel), 0) + 1
    fills = sorted(
        (value.hex().upper(), count)
        for value, count in counts.items()
        if count >= 200
    )
    print(" ".join(f"{name}:{count}" for name, count in fills))
"#;
        let scripts = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../scripts")
            .canonicalize()
            .unwrap();
        let measured = std::process::Command::new("python3")
            .arg("-c")
            .arg(script)
            .arg(&ours_png)
            .arg(&oracle_png)
            .arg(&scripts)
            .output()
            .expect("structural similarity runs");
        assert!(
            measured.status.success(),
            "{}",
            String::from_utf8_lossy(&measured.stderr)
        );
        let reported = String::from_utf8_lossy(&measured.stdout);
        let mut lines = reported.lines();
        let ssim: f64 = lines
            .next()
            .expect("structural similarity line")
            .trim()
            .parse()
            .expect("structural similarity is a number");
        let parse_blocks = |line: &str| {
            line.split_whitespace()
                .map(|block| {
                    let (start, end) = block.split_once(',').expect("shaded block bounds");
                    (
                        start.parse::<i64>().expect("shaded block start"),
                        end.parse::<i64>().expect("shaded block end"),
                    )
                })
                .collect::<Vec<_>>()
        };
        let ours_blocks = parse_blocks(lines.next().expect("our shaded rows"));
        let oracle_blocks = parse_blocks(lines.next().expect("oracle shaded rows"));
        let region_fills = F267_REGIONS.map(|(_, _, _, _, fill)| fill);
        let parse_fills = |line: &str| {
            let mut fills = line
                .split_whitespace()
                .filter_map(|entry| entry.split_once(':').map(|(name, _)| name.to_owned()))
                .filter(|name| region_fills.contains(&name.as_str()))
                .collect::<Vec<_>>();
            fills.sort();
            fills
        };
        let ours_fills = parse_fills(lines.next().expect("our painted fills"));
        let oracle_fills = parse_fills(lines.next().expect("oracle painted fills"));

        // The page-wide similarity is a collapse guard, not the layout claim.
        // Two engines never put the same ink in the same pixel, and the
        // declared band divergence below moves a sixth of the painted area.
        // The measured agreement is 0.73.
        assert!(ssim > 0.60, "structural similarity {ssim}");

        // The table paints one contiguous band of shaded rows in both, and
        // both start it on the same scanline. Where it ends differs, because
        // the two engines disagree about row height by about three points a
        // row, which is not what this story changes.
        assert_eq!(ours_blocks.len(), 1, "{ours_blocks:?}");
        assert_eq!(oracle_blocks.len(), 1, "{oracle_blocks:?}");
        assert!(
            (ours_blocks[0].0 - oracle_blocks[0].0).abs() <= 2,
            "shaded table start: {ours_blocks:?} against {oracle_blocks:?}"
        );

        // Ten of the thirteen regions reach a cell in this table. The four
        // corners, the two row edges and the two column edges resolve
        // identically in both engines.
        let shared = [
            "A10000", "A20000", "A30000", "A40000", "B1F1B1", "B2F2B2", "C1C1F1", "C2C2F2",
        ];
        for fill in shared {
            assert!(ours_fills.contains(&fill.to_owned()), "{ours_fills:?}");
            assert!(oracle_fills.contains(&fill.to_owned()), "{oracle_fills:?}");
        }

        // The declared divergence, asserted so a later change cannot drop it
        // silently. ECMA-376 orders `w:tblStylePr` band1Vert and band2Vert
        // before band1Horz and band2Horz, and a later region overrides an
        // earlier one, so the horizontal band outranks the vertical band.
        // Word resolves it that way and this workspace follows Word.
        // LibreOffice 26.2.5.2 resolves it the other way and paints the
        // vertical band. The oracle is wrong here, so the divergence is
        // recorded rather than followed.
        assert_eq!(
            ours_fills,
            {
                let mut expected = shared.to_vec();
                expected.extend(["D1F1D1", "D2F2D2"]);
                expected.sort();
                expected.into_iter().map(str::to_owned).collect::<Vec<_>>()
            },
            "{F267_LIBREOFFICE_ORACLE}"
        );
        assert_eq!(
            oracle_fills,
            {
                let mut expected = shared.to_vec();
                expected.extend(["E1E1F1", "E2E2F2"]);
                expected.sort();
                expected.into_iter().map(str::to_owned).collect::<Vec<_>>()
            },
            "{F267_LIBREOFFICE_ORACLE} no longer inverts the band priority"
        );
    }
    #[test]
    fn conditional_run_properties_reach_resolved_runs() {
        let mut document = Document::new();
        document
            .add_style(
                StyleBuilder::table("HeaderRun", "Header Run").conditional_table_style(
                    TableStyleRegion::FirstRow,
                    None,
                    Some(CT_RPr {
                        bold: Some(true),
                        color: Some("CC0000".to_owned()),
                        ..CT_RPr::default()
                    }),
                    None,
                    None,
                    None,
                ),
            )
            .expect("header run style is valid");
        let mut table = document.add_table(2, 1);
        table.set_style("HeaderRun");
        table.set_look(TableLook {
            first_row: true,
            last_row: false,
            first_column: false,
            last_column: false,
            horizontal_banding: false,
            vertical_banding: false,
        });
        table.row(0).unwrap().cell(0).unwrap().set_text("head");
        table.row(1).unwrap().cell(0).unwrap().set_text("body");

        let layout = document
            .layout_with_fonts_and_bundled_fallback(&[])
            .expect("deterministic run layout");
        let mut runs = f267_page_elements(&layout.layout.pages[0].elements)
            .into_iter()
            .filter_map(|element| match element {
                oxml_layout::PositionedElement::Text(run) if !run.text.is_empty() => Some((
                    (run.origin.y * 100.0).round() as i64,
                    run.text.clone(),
                    run.bold,
                    format!(
                        "{:02X}{:02X}{:02X}",
                        (run.color.r * 255.0).round() as u8,
                        (run.color.g * 255.0).round() as u8,
                        (run.color.b * 255.0).round() as u8
                    ),
                )),
                _ => None,
            })
            .collect::<Vec<_>>();
        runs.sort();
        assert_eq!(
            runs.iter()
                .map(|(_, text, bold, color)| (text.as_str(), *bold, color.as_str()))
                .collect::<Vec<_>>(),
            vec![("head", true, "CC0000"), ("body", false, "000000"),]
        );
    }

    #[test]
    fn paragraph_cnf_style_selects_conditional_regions() {
        let mut document = Document::new();
        document
            .add_style(
                StyleBuilder::table("ParagraphCnf", "Paragraph Cnf").conditional_table_style(
                    TableStyleRegion::FirstRow,
                    None,
                    None,
                    None,
                    None,
                    Some(CT_TcPr {
                        shading: Some(f267_shading("123456")),
                        ..CT_TcPr::default()
                    }),
                ),
            )
            .expect("paragraph selector style is valid");
        let mut table = document.add_table(2, 1);
        table.set_style("ParagraphCnf");
        // No region comes from the look, so only the paragraph selector can
        // reach the first-row region.
        table.set_look(TableLook {
            first_row: false,
            last_row: false,
            first_column: false,
            last_column: false,
            horizontal_banding: false,
            vertical_banding: false,
        });
        table.row(0).unwrap().cell(0).unwrap().set_text("plain");
        table.row(1).unwrap().cell(0).unwrap().set_text("selected");
        table
            .row(1)
            .unwrap()
            .cell(0)
            .unwrap()
            .paragraph_mut(0)
            .expect("the authored paragraph")
            .set_conditional_formatting(Some(TableConditionalFormatting {
                first_row: true,
                ..TableConditionalFormatting::default()
            }));

        assert_eq!(f267_cell_fills(&document), vec!["123456".to_owned()]);

        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let selected = reopened
            .table(0)
            .unwrap()
            .row(1)
            .unwrap()
            .cell(0)
            .unwrap()
            .paragraph(0)
            .unwrap()
            .conditional_formatting()
            .expect("the paragraph selector reopens");
        assert!(selected.first_row);
        assert!(!selected.last_row);
        assert_eq!(f267_cell_fills(&reopened), vec!["123456".to_owned()]);
    }

    #[test]
    fn conditional_region_run_and_row_layers_survive_reopen() {
        let mut document = Document::new();
        document
            .add_style(
                f267_authored_style()
                    .table_row_properties(CT_TrPr {
                        cant_split: Some(true),
                        ..CT_TrPr::default()
                    })
                    .table_cell_properties(CT_TcPr {
                        shading: Some(f267_shading("F0F0F0")),
                        ..CT_TcPr::default()
                    })
                    .table_properties(CT_TblPr {
                        row_band_size: Some(2),
                        column_band_size: Some(3),
                        ..CT_TblPr::default()
                    }),
            )
            .expect("authored regions are valid");
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let style = reopened.style("RegionGrid").unwrap();

        // The style's own base row and cell layers reopen beside its regions.
        assert_eq!(
            style
                .table_row_properties()
                .and_then(|properties| properties.cant_split),
            Some(true)
        );
        assert_eq!(
            style
                .table_cell_properties()
                .and_then(|properties| properties.shading.as_ref())
                .and_then(|shading| shading.fill.as_deref()),
            Some("F0F0F0")
        );
        let regions = style.conditional_table_styles();
        assert_eq!(regions.len(), 13);
        for (index, (region, _, size, _, _)) in F267_REGIONS.into_iter().enumerate() {
            assert_eq!(regions[index].region(), Some(region));
            assert_eq!(
                regions[index].run_properties().and_then(|rpr| rpr.sz),
                Some(HalfPoint(size)),
                "{}",
                region.to_str()
            );
            assert_eq!(
                regions[index]
                    .row_properties()
                    .and_then(|trpr| trpr.cant_split),
                Some(true),
                "{}",
                region.to_str()
            );
        }
        assert_eq!(
            regions[7].row_properties().and_then(|trpr| trpr.header),
            Some(true)
        );

        // Every region keeps the `pPr`, `rPr`, `tblPr`, `trPr`, `tcPr`
        // sequence the schema requires.
        let mut reopened = reopened;
        let package =
            OpcPackage::from_reader(std::io::Cursor::new(reopened.to_bytes().unwrap())).unwrap();
        let styles =
            String::from_utf8(package.get_part("/word/styles.xml").unwrap().to_vec()).unwrap();

        // The style's own base layers sit at their schema ranks, in `w:trPr`
        // then `w:tcPr` order and before the first region.
        let definition = styles
            .split(r#"<w:style w:type="table" w:styleId="RegionGrid">"#)
            .nth(1)
            .and_then(|tail| tail.split("</w:style>").next())
            .unwrap_or_else(|| panic!("{styles}"));
        let base = ["<w:trPr>", "<w:tcPr>", "<w:tblStylePr "].map(|tag| {
            definition
                .find(tag)
                .unwrap_or_else(|| panic!("{tag}: {definition}"))
        });
        assert!(
            base.windows(2).all(|pair| pair[0] < pair[1]),
            "{definition}"
        );

        // Removing one base layer during an update leaves the other.
        let mut updated = Document::from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        updated
            .set_style(
                StyleBuilder::table("RegionGrid", "Region Grid")
                    .clear_table_row_properties()
                    .table_properties(CT_TblPr {
                        shading: Some(f267_shading("FAFAFA")),
                        row_band_size: Some(5),
                        ..CT_TblPr::default()
                    }),
            )
            .expect("one base layer is removable");
        let style = updated.style("RegionGrid").unwrap();
        assert_eq!(style.table_row_properties(), None);
        assert!(style.table_cell_properties().is_some());

        // An updated band size reaches the merged style properties, and the
        // one the update did not mention keeps its existing value.
        assert_eq!(
            style
                .table_properties()
                .map(|properties| (properties.row_band_size, properties.column_band_size)),
            Some((Some(5), Some(3)))
        );
        for region in F267_REGIONS.map(|(region, _, _, _, _)| region) {
            let body = styles
                .split(&format!(r#"<w:tblStylePr w:type="{}">"#, region.to_str()))
                .nth(1)
                .and_then(|tail| tail.split("</w:tblStylePr>").next())
                .unwrap_or_else(|| panic!("{}: {styles}", region.to_str()));
            let order = ["<w:pPr>", "<w:rPr>", "<w:tblPr>", "<w:trPr>", "<w:tcPr>"]
                .map(|tag| body.find(tag).unwrap_or_else(|| panic!("{tag}: {body}")));
            assert!(order.windows(2).all(|pair| pair[0] < pair[1]), "{body}");
        }
    }

    #[test]
    fn paragraph_conditional_selector_survives_reopen() {
        let mut document = Document::new();
        let mut paragraph = document.add_paragraph("selector");
        paragraph.set_div_id_value(Some(7));
        paragraph.set_conditional_formatting(Some(TableConditionalFormatting {
            first_row: true,
            last_row_first_column: true,
            ..TableConditionalFormatting::default()
        }));
        paragraph.mark().set_bold(true);

        let package =
            OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
        let body =
            String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
        // Schema slot 32, between `w:divId` at 31 and `w:rPr` at 33.
        let div_id = body.find(r#"<w:divId w:val="7"/>"#).expect("divId");
        let cnf = body
            .find(r#"<w:cnfStyle w:val="100000000001"/>"#)
            .unwrap_or_else(|| panic!("{body}"));
        let rpr = body.find("<w:rPr>").expect("paragraph mark properties");
        assert!(div_id < cnf && cnf < rpr, "{body}");

        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let selector = reopened
            .paragraph(0)
            .unwrap()
            .conditional_formatting()
            .expect("the selector reopens");
        assert!(selector.first_row);
        assert!(selector.last_row_first_column);
        assert_eq!(reopened.paragraph(0).unwrap().div_id(), Some(7));

        // Unrelated producer XML at the same slot keeps its place.
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
        let injected = body.replace(
            r#"<w:cnfStyle w:val="100000000001"/>"#,
            r#"<x:marker xmlns:x="urn:producer" x:keep="1"/><w:cnfStyle w:val="100000000001" x:extra="kept" xmlns:x="urn:producer"/>"#,
        );
        package.set_part("/word/document.xml", injected.into_bytes());
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let mut carried = Document::from_bytes(bytes.get_ref()).expect("producer XML reopens");
        assert!(
            carried
                .paragraph(0)
                .unwrap()
                .conditional_formatting()
                .expect("the selector still projects")
                .first_row
        );
        let package =
            OpcPackage::from_reader(std::io::Cursor::new(carried.to_bytes().unwrap())).unwrap();
        let saved =
            String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
        assert!(
            saved.contains(r#"<x:marker xmlns:x="urn:producer" x:keep="1"/>"#),
            "{saved}"
        );
        assert!(saved.contains(r#"x:extra="kept""#), "{saved}");
    }

    #[test]
    fn untouched_conditional_regions_serialize_byte_for_byte() {
        let region = concat!(
            r#"<w:tblStylePr w:type="firstRow" x:note="kept" xmlns:x="urn:producer">"#,
            r#"<w:pPr><w:spacing w:after="40"/></w:pPr>"#,
            r#"<w:rPr><w:b/></w:rPr>"#,
            r#"<w:tblPr/>"#,
            r#"<w:trPr><w:cantSplit/></w:trPr>"#,
            r#"<w:tcPr><w:shd w:val="clear" w:fill="E8F1F8"/></w:tcPr>"#,
            r#"<x:unmodelled x:value="kept"/>"#,
            r#"</w:tblStylePr>"#
        );
        let mut seed = Document::new();
        seed.add_style(StyleBuilder::table("Preserved", "Preserved"))
            .expect("seed table style");
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
        let styles =
            String::from_utf8(package.get_part("/word/styles.xml").unwrap().to_vec()).unwrap();
        let injected = styles.replace(
            r#"<w:name w:val="Preserved"/>"#,
            &format!(r#"<w:name w:val="Preserved"/>{region}"#),
        );
        package.set_part("/word/styles.xml", injected.into_bytes());
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();

        let mut reopened = Document::from_bytes(bytes.get_ref()).expect("the region reopens");
        let style = reopened.style("Preserved").unwrap();
        let regions = style.conditional_table_styles();
        assert_eq!(regions[0].region(), Some(TableStyleRegion::FirstRow));
        assert_eq!(
            regions[0].run_properties().and_then(|rpr| rpr.bold),
            Some(true)
        );
        assert_eq!(
            regions[0].row_properties().and_then(|trpr| trpr.cant_split),
            Some(true)
        );

        let package =
            OpcPackage::from_reader(std::io::Cursor::new(reopened.to_bytes().unwrap())).unwrap();
        let saved =
            String::from_utf8(package.get_part("/word/styles.xml").unwrap().to_vec()).unwrap();
        assert!(saved.contains(region), "{saved}");
    }

    #[test]
    fn invalid_band_size_leaves_document_bytes_unchanged() {
        let mut document = Document::new();
        document
            .add_style(StyleBuilder::table("Sized", "Sized"))
            .expect("table style is valid");
        let mut table = document.add_table(1, 1);
        table.set_style("Sized");
        table.row(0).unwrap().cell(0).unwrap().set_text("x");
        let before = document.to_bytes().unwrap();

        let mut table = document.table_mut(0).unwrap();
        assert!(table.set_row_band_size(0).is_err());
        assert!(table.set_column_band_size(0).is_err());
        assert_eq!(document.to_bytes().unwrap(), before);
        assert_eq!(document.table(0).unwrap().row_band_size(), None);
        assert_eq!(document.table(0).unwrap().column_band_size(), None);

        let mut table = document.table_mut(0).unwrap();
        table.set_row_band_size(3).unwrap();
        table.set_column_band_size(2).unwrap();
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.table(0).unwrap().row_band_size(), Some(3));
        assert_eq!(reopened.table(0).unwrap().column_band_size(), Some(2));

        let mut cleared = reopened;
        cleared.table_mut(0).unwrap().clear_band_sizes();
        assert_eq!(cleared.table(0).unwrap().row_band_size(), None);
        assert_eq!(cleared.table(0).unwrap().column_band_size(), None);
    }
}

/// F-268a, advanced table authoring, content-driven autofit and row geometry.
mod advanced_table_authoring_and_geometry {
    use rdocx::table::{
        TableAnchor, TableFloatPosition, TableFloatX, TableFloatY, TableLayout, TableOverlap,
        TableTextDistance, TableWidth,
    };
    use rdocx::{Document, Length};
    use rdocx_oxml::table::{
        CT_Row, CT_Tbl, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc,
    };
    use rdocx_oxml::units::Twips;

    /// Parse one table out of a minimal document, which is the only public
    /// entry point that drives `CT_Tbl` from bytes.
    fn parse_single_table(table_xml: &str) -> CT_Tbl {
        let source = format!(
            r#"<w:document xmlns:w="{ns}"><w:body>{table_xml}</w:body></w:document>"#,
            ns = rdocx_oxml::namespace::W_NS
        );
        let document = rdocx_oxml::document::CT_Document::from_xml(source.as_bytes())
            .expect("document parses");
        document
            .body
            .content
            .into_iter()
            .find_map(|item| match item {
                rdocx_oxml::document::BodyContent::Table(table) => Some(table),
                _ => None,
            })
            .expect("document holds one table")
    }

    /// Collect every painted cell rectangle, rounded to two decimal places.
    ///
    /// Cell shading is what marks a painted cell, so a table whose cells are
    /// shaded reports one rectangle per cell in paint order.
    fn collect_cell_rectangles(
        elements: &[oxml_layout::PositionedElement],
        output: &mut Vec<(f64, f64, f64, f64)>,
    ) {
        let round = |value: f64| (value * 100.0).round() / 100.0;
        for element in elements {
            match element {
                oxml_layout::PositionedElement::FilledRect { rect, .. } => output.push((
                    round(rect.x),
                    round(rect.y),
                    round(rect.width),
                    round(rect.height),
                )),
                oxml_layout::PositionedElement::Group(group) => {
                    collect_cell_rectangles(&group.children, output)
                }
                oxml_layout::PositionedElement::MarkedContent { children, .. } => {
                    collect_cell_rectangles(children, output)
                }
                _ => {}
            }
        }
    }

    fn sample_float_position() -> TableFloatPosition {
        TableFloatPosition {
            horizontal_anchor: TableAnchor::Margin,
            vertical_anchor: TableAnchor::Page,
            horizontal: TableFloatX::Offset(Length::twips(720)),
            vertical: TableFloatY::Offset(Length::twips(-360)),
            distance_from_text: TableTextDistance {
                top: Length::twips(80),
                right: Length::twips(160),
                bottom: Length::twips(80),
                left: Length::twips(160),
            },
        }
    }

    fn layout_input(document: rdocx_oxml::document::CT_Document) -> rdocx_layout::LayoutInput {
        rdocx_layout::LayoutInput {
            automatic_hyphenation: false,
            mirror_margins: false,
            gutter_at_top: false,
            do_not_use_html_paragraph_auto_spacing: false,
            default_tab_stop: None,
            math_properties: None,
            document,
            styles: rdocx_oxml::styles::CT_Styles::new_default(),
            numbering: None,
            headers: std::collections::HashMap::new(),
            footers: std::collections::HashMap::new(),
            images: std::collections::HashMap::new(),
            charts: std::collections::HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: std::collections::HashMap::new(),
            footnotes: None,
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
            revision_view: rdocx_layout::RevisionView::Accepted,
        }
    }

    /// Lay one table out in deterministic font mode at `available_width`.
    fn lay_out(table: &CT_Tbl, available_width: f64) -> rdocx_layout::table::TableBlock {
        let input = layout_input(rdocx_oxml::document::CT_Document {
            body: rdocx_oxml::document::CT_Body {
                content: Vec::new(),
                sect_pr: None,
            },
            extra_namespaces: Vec::new(),
            background_xml: None,
            background_extra_xml: Vec::new(),
        });
        let media = rdocx_layout::MediaRegistry::new(&input.images);
        let mut fonts =
            oxml_layout::FontManager::new_deterministic().expect("deterministic fonts load");
        let mut numbering = rdocx_layout::style_resolver::NumberingState::new();
        let mut diagnostics = Vec::new();
        rdocx_layout::table::layout_table(
            table,
            available_width,
            &input.styles,
            &input,
            &media,
            &mut fonts,
            &mut numbering,
            &mut diagnostics,
            None,
        )
        .expect("table lays out")
    }

    /// A two-column table whose cells carry the given text.
    fn table_with_text(columns: &[i32], rows: &[&[&str]]) -> CT_Tbl {
        let mut table = CT_Tbl::new();
        table.grid = Some(CT_TblGrid {
            columns: columns
                .iter()
                .map(|width| CT_TblGridCol {
                    width: Twips(*width),
                })
                .collect(),
            ..CT_TblGrid::default()
        });
        for cells in rows {
            let mut row = CT_Row::new();
            for text in cells.iter() {
                let mut cell = CT_Tc::new();
                cell.paragraphs_mut()[0].add_run(text);
                row.cells.push(cell);
            }
            table.rows.push(row);
        }
        table
    }

    #[test]
    fn table_and_row_advanced_properties_survive_reopen() {
        let mut document = Document::new();
        {
            let mut table = document.add_table(2, 2);
            table
                .set_float_position(Some(sample_float_position()))
                .expect("float position is valid");
            table.set_overlap(Some(TableOverlap::Never));
            table.set_bidi_visual(Some(true));
            table
                .set_cell_spacing(Some(Length::twips(24)))
                .expect("cell spacing is valid");
            table
                // The caption carries a character the attribute writer must
                // escape rather than emit raw.
                .set_caption(Some("Totals & targets"))
                .expect("caption");
            table
                .set_description(Some("Region < quarter, by \"total\""))
                .expect("description");
            let mut row = table.row(0).expect("first row");
            row.set_width_before(Some(TableWidth::Fixed(Length::twips(360))))
                .expect("leading width is valid");
            row.set_width_after(Some(TableWidth::Percentage(10.0)))
                .expect("trailing width is valid");
            row.set_cell_spacing(Some(Length::twips(12)))
                .expect("row cell spacing is valid");
            row.set_hidden(Some(true));
        }

        let bytes = document.to_bytes().expect("document saves");
        let reopened = Document::from_bytes(&bytes).expect("document reopens");
        let table = reopened.table(0).expect("table reopens");
        assert_eq!(table.float_position(), Some(sample_float_position()));
        assert_eq!(table.overlap(), Some(TableOverlap::Never));
        assert_eq!(table.bidi_visual(), Some(true));
        assert_eq!(
            table.cell_spacing(),
            Some(TableWidth::Fixed(Length::twips(24)))
        );
        assert_eq!(table.caption(), Some("Totals & targets"));
        assert_eq!(table.description(), Some("Region < quarter, by \"total\""));
        let row = table.row(0).expect("row reopens");
        assert_eq!(
            row.width_before(),
            Some(TableWidth::Fixed(Length::twips(360)))
        );
        assert_eq!(row.width_after(), Some(TableWidth::Percentage(10.0)));
        assert_eq!(
            row.cell_spacing(),
            Some(TableWidth::Fixed(Length::twips(12)))
        );
        assert_eq!(row.hidden(), Some(true));
        assert!(!table.has_unmodeled_properties());
        assert!(!row.has_unmodeled_properties());

        // The ten new children land in their schema-sequence slots.
        let package =
            oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(bytes)).expect("package opens");
        let xml = String::from_utf8(
            package
                .get_part("/word/document.xml")
                .expect("document part")
                .to_vec(),
        )
        .expect("document xml is utf8");
        let at = |needle: &str| {
            xml.find(needle)
                .unwrap_or_else(|| panic!("{needle}: {xml}"))
        };
        assert!(at("<w:tblpPr") < at("<w:tblOverlap"));
        assert!(at("<w:tblOverlap") < at("<w:bidiVisual"));
        assert!(at("<w:bidiVisual") < at("<w:tblW"));
        assert!(at("<w:tblW") < at("<w:tblCellSpacing"));
        assert!(at("<w:tblCellSpacing") < at("<w:tblCaption"));
        assert!(at("<w:tblCaption") < at("<w:tblDescription"));
        // Free-text values are escaped on the way out and unescaped on the
        // way back in, unlike the token-valued `w:val` attributes beside them.
        assert!(
            xml.contains(r#"<w:tblCaption w:val="Totals &amp; targets"/>"#),
            "{xml}"
        );
        assert!(
            xml.contains(
                r#"<w:tblDescription w:val="Region &lt; quarter, by &quot;total&quot;"/>"#
            ),
            "{xml}"
        );
        assert!(at("<w:wBefore") < at("<w:wAfter"));
        assert!(at("<w:wAfter") < at("<w:hidden"));
    }

    #[test]
    fn tbl_ppr_attribute_matrix_is_prefix_tolerant_and_writes_a_fixed_prefix() {
        let source = format!(
            concat!(
                r#"<x:tbl xmlns:x="{ns}"><x:tblPr><x:tblpPr x:leftFromText="10""#,
                r#" x:rightFromText="20" x:topFromText="30" x:bottomFromText="40""#,
                r#" x:horzAnchor="page" x:vertAnchor="text" x:tblpX="120""#,
                r#" x:tblpXSpec="outside" x:tblpY="240" x:tblpYSpec="inside"/>"#,
                r#"<x:tblOverlap x:val="overlap"/></x:tblPr>"#,
                r#"<x:tblGrid><x:gridCol x:w="1000"/></x:tblGrid>"#,
                r#"<x:tr><x:tc><x:p/></x:tc></x:tr></x:tbl>"#
            ),
            ns = rdocx_oxml::namespace::W_NS
        );
        let table = parse_single_table(&source);
        let position = table
            .properties
            .as_ref()
            .expect("table properties")
            .float_position
            .as_deref()
            .expect("float position parses");
        assert_eq!(position.left_from_text, Some(Twips(10)));
        assert_eq!(position.right_from_text, Some(Twips(20)));
        assert_eq!(position.top_from_text, Some(Twips(30)));
        assert_eq!(position.bottom_from_text, Some(Twips(40)));
        assert_eq!(
            position.horz_anchor,
            Some(rdocx_oxml::table::ST_TblAnchor::Page)
        );
        assert_eq!(
            position.vert_anchor,
            Some(rdocx_oxml::table::ST_TblAnchor::Text)
        );
        assert_eq!(position.tbl_p_x, Some(Twips(120)));
        assert_eq!(
            position.tbl_p_x_spec,
            Some(rdocx_oxml::drawing::AnchorAlignH::Outside)
        );
        assert_eq!(position.tbl_p_y, Some(Twips(240)));
        assert_eq!(
            position.tbl_p_y_spec,
            Some(rdocx_oxml::table::ST_YAlign::Inside)
        );

        let mut writer = quick_xml::Writer::new(Vec::new());
        table.to_xml(&mut writer).expect("table serialises");
        let written = String::from_utf8(writer.into_inner()).expect("serialised table is utf8");
        for attribute in [
            r#"w:leftFromText="10""#,
            r#"w:rightFromText="20""#,
            r#"w:topFromText="30""#,
            r#"w:bottomFromText="40""#,
            r#"w:horzAnchor="page""#,
            r#"w:vertAnchor="text""#,
            r#"w:tblpXSpec="outside""#,
            r#"w:tblpX="120""#,
            r#"w:tblpYSpec="inside""#,
            r#"w:tblpY="240""#,
        ] {
            assert!(written.contains(attribute), "{attribute}: {written}");
        }
        assert!(
            written.contains(r#"<w:tblOverlap w:val="overlap"/>"#),
            "{written}"
        );

        // An unrecognised value falls back rather than inventing a position.
        let unknown = format!(
            concat!(
                r#"<w:tbl xmlns:w="{ns}"><w:tblPr>"#,
                r#"<w:tblpPr w:horzAnchor="elsewhere" w:tblpXSpec="sideways"/>"#,
                r#"<w:tblOverlap w:val="sometimes"/></w:tblPr>"#,
                r#"<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>"#,
                r#"<w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"#
            ),
            ns = rdocx_oxml::namespace::W_NS
        );
        let table = parse_single_table(&unknown);
        let properties = table.properties.as_ref().expect("table properties");
        let position = properties
            .float_position
            .as_deref()
            .expect("float position");
        assert_eq!(position.horz_anchor, None);
        assert_eq!(position.tbl_p_x_spec, None);
        assert_eq!(properties.overlap, None);
    }

    #[test]
    fn autofit_engages_only_for_an_auto_width_autofit_table() {
        let declared = [Twips(1440), Twips(4320)];
        let grid_points = |table: &CT_Tbl| -> Vec<f64> {
            lay_out(table, 360.0)
                .col_widths
                .iter()
                .map(|width| (width * 100.0).round() / 100.0)
                .collect()
        };
        let declared_points = declared
            .iter()
            .map(|width| width.to_pt())
            .collect::<Vec<_>>();

        let mut table = table_with_text(
            &[declared[0].0, declared[1].0],
            &[&["Region", "Quarterly revenue for the northern region"]],
        );

        // Absent width and absent layout mode keeps the declared grid. ECMA
        // makes autofit the default here, but engagement deliberately requires
        // the element, because an absent layout is the shape almost every
        // producer writes and treating it as autofit moves the pinned private
        // corpus reference page count.
        assert_eq!(grid_points(&table), declared_points);

        // An explicit fixed layout keeps the declared grid.
        table.properties = Some(CT_TblPr {
            layout: Some("fixed".to_owned()),
            ..CT_TblPr::default()
        });
        assert_eq!(grid_points(&table), declared_points);

        // An authored dxa width keeps the declared grid even with autofit.
        table.properties = Some(CT_TblPr {
            layout: Some("autofit".to_owned()),
            width: Some(CT_TblWidth::dxa(5760)),
            ..CT_TblPr::default()
        });
        assert_eq!(grid_points(&table), declared_points);

        // An authored percentage width keeps the declared grid.
        table.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::pct(5000)),
            ..CT_TblPr::default()
        });
        assert_eq!(grid_points(&table), declared_points);

        // An auto width with an autofit layout engages.
        table.properties = Some(CT_TblPr {
            layout: Some("autofit".to_owned()),
            width: Some(CT_TblWidth::auto()),
            ..CT_TblPr::default()
        });
        assert_ne!(grid_points(&table), declared_points);
    }

    #[test]
    fn autofit_distributes_available_width_between_measured_minima_and_maxima() {
        let mut table = table_with_text(
            &[2880, 2880],
            &[&[
                "ID",
                "A considerably longer heading that cannot fit on one line at this width",
            ]],
        );
        // Engagement requires the element, so this fixture opts in explicitly.
        table.properties = Some(CT_TblPr {
            layout: Some("autofit".to_owned()),
            ..CT_TblPr::default()
        });

        // Wide enough for every cell's natural width: columns stop at content.
        let roomy = lay_out(&table, 600.0);
        assert!(roomy.col_widths[0] < roomy.col_widths[1]);
        assert!(roomy.table_width < 600.0);

        // Narrow enough to force distribution: the columns fill the caller's
        // width and the narrow column keeps at least its measured minimum.
        let tight = lay_out(&table, 200.0);
        let total: f64 = tight.col_widths.iter().sum();
        assert!((total - 200.0).abs() < 0.01, "{:?}", tight.col_widths);
        assert!(tight.col_widths[0] > 0.0);
        assert!(tight.col_widths[0] < roomy.col_widths[0] + 0.01);
        assert!(tight.col_widths[1] > tight.col_widths[0]);
    }

    #[test]
    fn checked_table_and_row_setters_reject_invalid_values() {
        let mut document = Document::new();
        document.add_table(1, 1);
        let before = document.to_bytes().expect("document saves");

        {
            let mut table = document.table_mut(0).expect("table");
            assert!(table.set_cell_spacing(Some(Length::twips(-1))).is_err());
            assert!(table.set_caption(Some("   ")).is_err());
            assert!(table.set_description(Some("")).is_err());
            let mut invalid = sample_float_position();
            invalid.distance_from_text.top = Length::twips(-1);
            assert!(table.set_float_position(Some(invalid)).is_err());
        }
        {
            let mut table = document.table_mut(0).expect("table");
            let mut row = table.row(0).expect("row");
            assert!(
                row.set_width_before(Some(TableWidth::Percentage(140.0)))
                    .is_err()
            );
            assert!(
                row.set_width_after(Some(TableWidth::Fixed(Length::twips(-5))))
                    .is_err()
            );
            assert!(row.set_cell_spacing(Some(Length::twips(-3))).is_err());
        }

        assert_eq!(document.to_bytes().expect("document saves"), before);
        let table = document.table(0).expect("table");
        assert_eq!(table.cell_spacing(), None);
        assert_eq!(table.caption(), None);
        assert_eq!(table.description(), None);
        assert_eq!(table.float_position(), None);
        let row = table.row(0).expect("row");
        assert_eq!(row.width_before(), None);
        assert_eq!(row.width_after(), None);
        assert_eq!(row.cell_spacing(), None);
    }

    #[test]
    fn fixed_autofit_and_nested_table_geometry_matches_reviewed_word_pages() {
        let mut document = Document::new();
        {
            let mut fixed = document.add_table(2, 3);
            fixed.set_layout(TableLayout::Fixed);
            fixed
                .set_grid_widths(&[
                    Length::twips(2880),
                    Length::twips(2880),
                    Length::twips(3600),
                ])
                .expect("fixed grid widths");
            for row_index in 0..2 {
                let mut row = fixed.row(row_index).expect("fixed row");
                for cell_index in 0..3 {
                    let mut cell = row.cell(cell_index).expect("fixed cell");
                    cell.set_text("Fixed");
                    cell.set_shading("EEEEEE");
                }
            }
        }
        {
            let mut autofit = document.add_table(2, 2);
            autofit
                .set_width_mode(TableWidth::Auto)
                .expect("auto width");
            autofit.set_layout(TableLayout::AutoFit);
            let texts = [
                ["ID", "A much longer autofit heading than the first column"],
                ["7", "Short"],
            ];
            for (row_index, row_texts) in texts.iter().enumerate() {
                let mut row = autofit.row(row_index).expect("autofit row");
                for (cell_index, text) in row_texts.iter().enumerate() {
                    let mut cell = row.cell(cell_index).expect("autofit cell");
                    cell.set_text(text);
                    cell.set_shading("DDDDDD");
                }
            }
        }
        {
            let mut outer = document.add_table(1, 2);
            outer.set_layout(TableLayout::Fixed);
            let mut row = outer.row(0).expect("outer row");
            row.cell(0).expect("outer cell").set_text("Outer");
            let mut host = row.cell(1).expect("nested host");
            let mut nested = host.add_table(2, 2);
            nested.set_layout(TableLayout::Fixed);
            for nested_row in 0..2 {
                let mut nested_row = nested.row(nested_row).expect("nested row");
                for nested_cell in 0..2 {
                    let mut cell = nested_row.cell(nested_cell).expect("nested cell");
                    cell.set_text("N");
                    cell.set_shading("CCCCCC");
                }
            }
        }

        let layout = document.layout_deterministic().expect("document lays out");
        let mut origins = Vec::new();
        for page in &layout.layout.pages {
            collect_cell_rectangles(&page.elements, &mut origins);
        }
        assert_eq!(layout.layout.pages.len(), 1);
        assert_eq!(origins, GOLDEN_TABLE_GEOMETRY);
    }

    /// Reviewed page geometry for
    /// `fixed_autofit_and_nested_table_geometry_matches_reviewed_word_pages`,
    /// as `(x, y, width, height)` in points for every painted cell.
    ///
    /// Rows 1 to 6 are the fixed-grid table, which keeps its declared 144,
    /// 144 and 180 point columns. Rows 7 to 10 are the auto-width autofit
    /// table, whose narrow `ID` column measures 20.34 points against a 242.28
    /// point heading column and whose total stops short of the 468 point text
    /// column because the content fits. Rows 11 to 14 are the nested table,
    /// which resolves its own grid inside the owning cell content box.
    const GOLDEN_TABLE_GEOMETRY: &[(f64, f64, f64, f64)] = &[
        (72.0, 72.0, 144.0, 19.87),
        (216.0, 72.0, 144.0, 19.87),
        (360.0, 72.0, 180.0, 19.87),
        (72.0, 91.87, 144.0, 19.87),
        (216.0, 91.87, 144.0, 19.87),
        (360.0, 91.87, 180.0, 19.87),
        (72.0, 111.74, 20.34, 19.87),
        (92.34, 111.74, 242.28, 19.87),
        (72.0, 131.61, 20.34, 19.87),
        (92.34, 131.61, 242.28, 19.87),
        (311.4, 172.43, 111.6, 19.87),
        (423.0, 172.43, 111.6, 19.87),
        (311.4, 192.3, 111.6, 19.87),
        (423.0, 192.3, 111.6, 19.87),
    ];
}

/// F-266a, script identity and font slot resolution.
///
/// The gate is a recorded geometry digest over a deterministic mixed-script
/// page. It uses no rasteriser and no external oracle, because the properties
/// under test are glyph identity, glyph positioning, cluster mapping and
/// painted order, all of which the layout result already states exactly. A
/// pixel comparison would add an external dependency and prove less.
mod f266a_mixed_script_typography {
    use super::*;
    use oxml_layout::{PositionedElement, TextDirection, TextScript};
    use rdocx::RunFontSlot;
    use sha2::{Digest, Sha256};

    const LATIN: &str = "Mixed script page";
    const ARABIC: &str = "العربية";
    const HEBREW: &str = "שלום עולם";
    const KOREAN: &str = "안녕하세요 세계";
    const JAPANESE: &str = "こんにちは、カタカナ世界";
    const KANJI: &str = "世界";

    /// The recorded geometry of the mixed-script page.
    ///
    /// Re-record only with a stated reason. The digest covers every painted
    /// run on the page in paint order, with its font family, point size,
    /// origin, logical text, glyph ids and advances, and for a rich run also
    /// its direction, script, bidi embedding level, both offset axes and its
    /// cluster ranges.
    ///
    /// The serialisation is host-stable because the pipeline is f64
    /// throughout with no FMA contraction, the shaper is pure Rust, and every
    /// face is bundled, so each coordinate is an integer font unit scaled by
    /// one multiply and summed in a fixed order. Four decimal places is not
    /// what makes it stable. It is a guard band that keeps an ordinary
    /// representation difference away from the printed digits, and it is
    /// applied to a value whose sign of zero has been normalised, because
    /// `format!("{:.4}", -0.0)` renders `-0.0000`.
    pub(super) const MIXED_SCRIPT_GEOMETRY_DIGEST: &str =
        "516ebb6e45438731d3cb0983707ad00c9de55068401e073ef2a069a56f397402";

    /// One page holding all five scripts, authored through the public facade.
    ///
    /// Every script sets its font through the `w:rFonts` slot Word uses for
    /// it. The Kanji paragraph is what makes slot resolution load bearing
    /// here, because `Noto Sans SC` on `w:ascii` and `Noto Sans JP` on
    /// `w:eastAsia` both cover its text, so coverage fallback cannot choose
    /// between them and only the slot can. Every other paragraph has exactly
    /// one bundled face that covers it, so those prove script identity,
    /// reading order and geometry rather than slot resolution.
    pub(super) fn mixed_script_document() -> Document {
        let mut document = Document::new();

        let mut latin = document.add_paragraph("");
        latin.add_run(LATIN).font("Carlito").language("en-US");

        let mut arabic = document.add_paragraph("").right_to_left(true);
        {
            let mut run = arabic.add_run(ARABIC);
            run.set_slot_font(RunFontSlot::ComplexScript, Some("Noto Sans Arabic"));
            run.set_rtl_value(Some(true));
            run.set_complex_script_value(Some(true));
            run.set_language_bidi_value(Some("ar-SA"));
        }

        let mut hebrew = document.add_paragraph("").right_to_left(true);
        {
            let mut run = hebrew.add_run(HEBREW);
            run.set_slot_font(RunFontSlot::ComplexScript, Some("Noto Sans Hebrew"));
            run.set_rtl_value(Some(true));
            run.set_complex_script_value(Some(true));
            run.set_language_bidi_value(Some("he-IL"));
        }

        let mut korean = document.add_paragraph("");
        {
            let mut run = korean.add_run(KOREAN);
            run.set_slot_font(RunFontSlot::EastAsia, Some("Noto Sans KR"));
            run.set_language_east_asia_value(Some("ko-KR"));
        }

        let mut japanese = document.add_paragraph("");
        {
            let mut run = japanese.add_run(JAPANESE);
            run.set_slot_font(RunFontSlot::EastAsia, Some("Noto Sans JP"));
            run.set_language_east_asia_value(Some("ja-JP"));
        }

        let mut kanji = document.add_paragraph("");
        {
            let mut run = kanji.add_run(KANJI);
            run.set_slot_font(RunFontSlot::Ascii, Some("Noto Sans SC"));
            run.set_slot_font(RunFontSlot::EastAsia, Some("Noto Sans JP"));
            run.set_language_east_asia_value(Some("ja-JP"));
        }

        document
    }

    fn family_of(result: &rdocx_layout::WordLayoutResult, id: oxml_layout::FontId) -> String {
        result
            .layout
            .fonts
            .iter()
            .find(|font| font.id == id)
            .map(|font| font.family.clone())
            .expect("every painted run names a font in the result font table")
    }

    /// Every rich run on the page, which is every run that reached the shaper.
    fn rich_runs(
        result: &rdocx_layout::WordLayoutResult,
    ) -> Vec<oxml_layout::MultilingualGlyphRun> {
        let mut runs = Vec::new();
        for page in &result.layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| {
                if let PositionedElement::MultilingualText(run) = element {
                    runs.push(run.clone());
                }
            });
        }
        runs
    }

    /// Every legacy run on the page. Text with no complex script stays here,
    /// which is the path the Latin paragraph takes.
    fn legacy_runs(result: &rdocx_layout::WordLayoutResult) -> Vec<oxml_layout::GlyphRun> {
        let mut runs = Vec::new();
        for page in &result.layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| {
                if let PositionedElement::Text(run) = element {
                    runs.push(run.clone());
                }
            });
        }
        runs
    }

    /// One coordinate, with the sign of zero normalised first.
    ///
    /// A shaper that returns `-0.0` for an offset is arithmetically equal to
    /// one that returns `0.0`, but `format!("{:.4}", -0.0)` renders
    /// `-0.0000`, which would move the digest for no geometric reason.
    fn number(value: f64) -> String {
        format!("{:.4}", if value == 0.0 { 0.0 } else { value })
    }

    fn numbers(values: &[f64]) -> String {
        values
            .iter()
            .copied()
            .map(number)
            .collect::<Vec<_>>()
            .join(",")
    }

    fn glyphs(values: &[u16]) -> String {
        values
            .iter()
            .map(u16::to_string)
            .collect::<Vec<_>>()
            .join(",")
    }

    /// The canonical serialisation the digest is taken over.
    pub(super) fn canonical_geometry(result: &rdocx_layout::WordLayoutResult) -> String {
        let mut lines = Vec::new();
        for (page_index, page) in result.layout.pages.iter().enumerate() {
            let mut paint_index = 0usize;
            oxml_layout::walk(&page.elements, &mut |element, _| {
                let body = match element {
                    PositionedElement::Text(run) => format!(
                        "kind=legacy font={} size={} origin={},{} text={} \
                         glyphs={} adv={}",
                        family_of(result, run.font_id),
                        number(run.font_size),
                        number(run.origin.x),
                        number(run.origin.y),
                        run.text,
                        glyphs(&run.glyph_ids),
                        numbers(&run.advances),
                    ),
                    PositionedElement::MultilingualText(run) => format!(
                        "kind=rich font={} size={} origin={},{} dir={:?} \
                         script={:?} bidi={} logical={} text={} glyphs={} \
                         xadv={} yadv={} xoff={} yoff={} clusters={}",
                        family_of(result, run.font_id),
                        number(run.font_size),
                        number(run.origin.x),
                        number(run.origin.y),
                        run.direction,
                        run.script,
                        run.bidi_level,
                        run.logical_index,
                        run.logical_text,
                        glyphs(&run.glyph_ids),
                        numbers(&run.x_advances),
                        numbers(&run.y_advances),
                        numbers(&run.x_offsets),
                        numbers(&run.y_offsets),
                        run.clusters
                            .iter()
                            .map(|cluster| format!(
                                "{}:{}>{}:{}",
                                cluster.glyph_start,
                                cluster.glyph_end,
                                cluster.char_start,
                                cluster.char_end
                            ))
                            .collect::<Vec<_>>()
                            .join(","),
                    ),
                    _ => return,
                };
                lines.push(format!("page={page_index} paint={paint_index} {body}"));
                paint_index += 1;
            });
        }
        lines.join("\n")
    }

    pub(super) fn digest(text: &str) -> String {
        let mut hasher = Sha256::new();
        hasher.update(text.as_bytes());
        hasher
            .finalize()
            .iter()
            .map(|byte| format!("{byte:02x}"))
            .collect()
    }

    #[test]
    fn mixed_script_page_matches_the_pinned_geometry_and_reading_order() {
        let mut document = mixed_script_document();
        let result = document
            .layout_deterministic()
            .expect("deterministic mixed-script layout");
        assert_eq!(result.layout.pages.len(), 1, "the fixture is one page");

        let rich = rich_runs(&result);
        let legacy = legacy_runs(&result);
        assert!(!rich.is_empty(), "the complex scripts reach the shaper");

        // Reading order and identity are asserted one property at a time, so
        // a failure names the property that broke rather than only reporting
        // that a hash moved.
        //
        // The Latin paragraph has no complex script, so it stays on the
        // legacy path, which is correct and is asserted separately below.
        for (script, family) in [
            (TextScript::Arabic, "Noto Sans Arabic"),
            (TextScript::Hebrew, "Noto Sans Hebrew"),
            (TextScript::Hangul, "Noto Sans KR"),
            (TextScript::Kana, "Noto Sans JP"),
            (TextScript::Han, "Noto Sans JP"),
        ] {
            let script_runs = rich
                .iter()
                .filter(|run| run.script == script)
                .collect::<Vec<_>>();
            assert!(
                !script_runs.is_empty(),
                "{script:?} must reach the page with its own script identity"
            );
            for run in &script_runs {
                assert!(run.is_valid(), "{script:?} run is a complete rich run");
                assert_eq!(
                    family_of(&result, run.font_id),
                    family,
                    "{script:?} must resolve through its own w:rFonts slot"
                );
            }
        }

        // Logical order is exact, not merely contained. Each source paragraph
        // reassembles to its whole fixture string when its rich runs are read
        // back in logical index order, so a dropped or reordered span fails
        // here and names the paragraph.
        let mut by_paragraph = std::collections::BTreeMap::<u32, Vec<_>>::new();
        for run in &rich {
            let node = run
                .source
                .unwrap_or_else(|| panic!("rich run {:?} retains provenance", run.logical_text));
            by_paragraph.entry(node.node.get()).or_default().push(run);
        }
        let mut reassembled = by_paragraph
            .into_values()
            .map(|mut runs| {
                runs.sort_by_key(|run| run.logical_index);
                runs.iter()
                    .map(|run| run.logical_text.as_str())
                    .collect::<String>()
            })
            .collect::<Vec<_>>();
        reassembled.sort();
        let mut expected = vec![
            ARABIC.to_owned(),
            HEBREW.to_owned(),
            KOREAN.to_owned(),
            JAPANESE.to_owned(),
            KANJI.to_owned(),
        ];
        expected.sort();
        assert_eq!(
            reassembled, expected,
            "every complex-script paragraph reassembles to its whole fixture string"
        );

        // The Latin paragraph stays on the legacy path with its own family.
        let latin = legacy
            .iter()
            .filter(|run| !run.text.trim().is_empty())
            .collect::<Vec<_>>();
        assert_eq!(
            latin
                .iter()
                .map(|run| run.text.as_str())
                .collect::<String>(),
            LATIN,
            "the Latin paragraph keeps its text and its order"
        );
        for run in &latin {
            assert_eq!(family_of(&result, run.font_id), "Carlito");
        }

        // The two right-to-left paragraphs carry a right-to-left direction and
        // an odd bidi embedding level, which is what reordering acts on.
        for script in [TextScript::Arabic, TextScript::Hebrew] {
            for run in rich.iter().filter(|run| run.script == script) {
                assert_eq!(
                    run.direction,
                    TextDirection::RightToLeft,
                    "{script:?} paints right to left"
                );
                assert_eq!(
                    run.bidi_level % 2,
                    1,
                    "{script:?} carries an odd bidi embedding level"
                );
            }
        }

        // The East Asian paragraphs must not have been swept into the
        // right-to-left base direction.
        for script in [TextScript::Hangul, TextScript::Kana, TextScript::Han] {
            for run in rich.iter().filter(|run| run.script == script) {
                assert_eq!(
                    run.bidi_level % 2,
                    0,
                    "{script:?} keeps an even bidi embedding level"
                );
            }
        }

        // Reordering is a painting concern. The saved bytes stay logical.
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let saved = reopened
            .paragraphs()
            .iter()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>();
        assert_eq!(saved, vec![LATIN, ARABIC, HEBREW, KOREAN, JAPANESE, KANJI]);

        let geometry = canonical_geometry(&result);
        assert_eq!(
            digest(&geometry),
            MIXED_SCRIPT_GEOMETRY_DIGEST,
            "mixed-script page geometry moved:\n{geometry}"
        );
    }
}

/// F-266b, ruby phonetic guides and East Asian emphasis marks.
///
/// The gate is a recorded geometry digest over a deterministic page carrying
/// both, taken with the same canonical serialisation F-266a records. It uses
/// no rasteriser and no external oracle, because what is under test is glyph
/// identity, placement and painted order, all of which the layout result
/// already states exactly.
///
/// F-266a's digest is asserted unmoved in the same module, so this story
/// cannot quietly move its sibling's baseline while recording its own.
mod f266b_ruby_and_emphasis_typography {
    use super::f266a_mixed_script_typography::{
        MIXED_SCRIPT_GEOMETRY_DIGEST, canonical_geometry, digest, mixed_script_document,
    };
    use super::*;
    use rdocx::{RunFontSlot, ST_Em, ST_RubyAlign};
    use rdocx_oxml::units::HalfPoint;

    const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    const RUBY_BASE: &str = "漢字";
    const RUBY_TEXT: &str = "かんじ";
    const EMPHASIS_JAPANESE: &str = "強調";
    const EMPHASIS_KOREAN: &str = "강조";
    const EMPHASIS_LATIN: &str = "marked words";
    const LATIN_CONTROL: &str = "Ruby and emphasis page";

    /// The recorded geometry of the ruby and emphasis page.
    ///
    /// Re-record only with a stated reason. It covers every painted run on
    /// the page in paint order, including the runs inside the annotation
    /// groups, because the canonical serialisation walks the element tree
    /// rather than the top level. The serialisation is the one F-266a
    /// documents, and it is host-stable for the same reasons.
    pub(super) const RUBY_AND_EMPHASIS_GEOMETRY_DIGEST: &str =
        "b119714501d061f912bf9c05224f66dc8d4a30f3bdd195040038b89157e6fbf6";

    /// Wrap producer body XML in a package a `Document` can open.
    fn producer_document(body: &str) -> Vec<u8> {
        let mut seed = Document::new();
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
        package.set_part(
            "/word/document.xml",
            format!(
                r#"<w:document xmlns:w="{W_NS}"><w:body>{body}<w:sectPr/></w:body></w:document>"#
            )
            .into_bytes(),
        );
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        bytes.into_inner()
    }

    /// The saved `document.xml` with the writer's layout indentation removed,
    /// so an assertion states element order rather than pretty printing.
    fn saved_document_xml(bytes: &[u8]) -> String {
        let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        let xml =
            String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
        let mut compact = String::with_capacity(xml.len());
        let mut in_tag = false;
        let mut pending = String::new();
        for character in xml.chars() {
            match character {
                '<' => {
                    if !pending.trim().is_empty() {
                        compact.push_str(&pending);
                    }
                    pending.clear();
                    in_tag = true;
                    compact.push('<');
                }
                '>' => {
                    in_tag = false;
                    compact.push('>');
                }
                _ if in_tag => compact.push(character),
                _ => pending.push(character),
            }
        }
        compact.push_str(&pending);
        compact
    }

    /// Give one run the East Asian font slot a bundled subset face covers.
    fn east_asian_run(properties: &mut Option<CT_RPr>, family: &str, language: &str) {
        let rpr = properties.get_or_insert_with(CT_RPr::default);
        rpr.font_east_asia = Some(family.to_owned());
        rpr.language_east_asia = Some(language.to_owned());
    }

    /// One page carrying ruby-annotated Japanese, emphasis-marked Japanese
    /// and Korean, and a Latin control, all authored through the facade.
    pub(super) fn ruby_and_emphasis_document() -> Document {
        let mut document = Document::new();

        let mut latin = document.add_paragraph("");
        latin
            .add_run(LATIN_CONTROL)
            .font("Carlito")
            .language("en-US");

        let mut annotated = document.add_paragraph("");
        {
            let index = annotated.add_ruby(RUBY_BASE, RUBY_TEXT);
            {
                let mut base = annotated
                    .run_mut(0)
                    .expect("the base run is a paragraph run");
                base.set_slot_font(RunFontSlot::EastAsia, Some("Noto Sans JP"));
                base.set_language_east_asia_value(Some("ja-JP"));
            }
            let ruby = annotated.ruby_mut(index).expect("the ruby was just added");
            east_asian_run(&mut ruby.ruby_text[0].properties, "Noto Sans JP", "ja-JP");
            ruby.properties = Some(rdocx::CT_RubyPr {
                align: Some(ST_RubyAlign::DistributeSpace),
                hps: Some(HalfPoint(10)),
                hps_raise: Some(HalfPoint(24)),
                hps_base_text: Some(HalfPoint(22)),
                language: Some("ja-JP".to_owned()),
                dirty: None,
                raw_xml: Vec::new(),
            });
        }

        for (text, family, language, mark) in [
            (EMPHASIS_JAPANESE, "Noto Sans JP", "ja-JP", ST_Em::Dot),
            (EMPHASIS_KOREAN, "Noto Sans KR", "ko-KR", ST_Em::Circle),
        ] {
            let mut paragraph = document.add_paragraph("");
            let mut run = paragraph.add_run(text);
            run.set_slot_font(RunFontSlot::EastAsia, Some(family));
            run.set_language_east_asia_value(Some(language));
            run.set_emphasis_mark_value(Some(mark));
        }

        let mut latin_marked = document.add_paragraph("");
        {
            let mut run = latin_marked.add_run(EMPHASIS_LATIN);
            run.set_font("Carlito");
            run.set_emphasis_mark_value(Some(ST_Em::UnderDot));
        }

        document
    }

    /// **The test gate.** The ruby and emphasis page keeps its recorded
    /// geometry and its reading order, and F-266a's page is unmoved.
    #[test]
    fn ruby_and_emphasis_page_matches_the_pinned_geometry_and_reading_order() {
        let mut document = ruby_and_emphasis_document();
        let result = document
            .layout_deterministic()
            .expect("deterministic ruby and emphasis layout");
        assert_eq!(result.layout.pages.len(), 1, "the fixture is one page");

        // Every painted string on the page, in paint order.
        let mut painted = Vec::new();
        for page in &result.layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| match element {
                oxml_layout::PositionedElement::Text(run) if !run.text.trim().is_empty() => {
                    painted.push(run.text.clone());
                }
                oxml_layout::PositionedElement::MultilingualText(run)
                    if !run.logical_text.trim().is_empty() =>
                {
                    painted.push(run.logical_text.clone());
                }
                _ => {}
            });
        }

        // The base line and the phonetic line both reach the page, and the
        // base comes first because it is the content and the annotation is
        // painted over it.
        let base_at = painted
            .iter()
            .position(|text| text == RUBY_BASE)
            .expect("the ruby base line is painted");
        let phonetic_at = painted
            .iter()
            .position(|text| text == RUBY_TEXT)
            .expect("the ruby phonetic line is painted");
        assert!(
            base_at < phonetic_at,
            "the base line is painted before its annotation: {painted:?}"
        );

        // Both emphasis-marked strings reach the page with their marks.
        for text in [EMPHASIS_JAPANESE, EMPHASIS_KOREAN] {
            assert!(
                painted.iter().any(|painted| painted == text),
                "{text} reaches the page: {painted:?}"
            );
        }
        let marks = painted.iter().filter(|text| *text == "\u{2022}").count();
        assert_eq!(
            marks,
            EMPHASIS_JAPANESE.chars().count()
                + EMPHASIS_LATIN
                    .chars()
                    .filter(|c| !c.is_whitespace())
                    .count(),
            "one solid dot per non-space base character of the two dot-marked runs"
        );
        assert_eq!(
            painted.iter().filter(|text| *text == "\u{25CB}").count(),
            EMPHASIS_KOREAN.chars().count(),
            "one open circle per Korean base character"
        );

        // The phonetic line is an annotation, so the saved bytes and every
        // text projection carry the base text and nothing else.
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let saved = reopened
            .paragraphs()
            .iter()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>();
        assert_eq!(
            saved,
            vec![
                LATIN_CONTROL,
                RUBY_BASE,
                EMPHASIS_JAPANESE,
                EMPHASIS_KOREAN,
                EMPHASIS_LATIN,
            ]
        );

        let geometry = canonical_geometry(&result);
        assert_eq!(
            digest(&geometry),
            RUBY_AND_EMPHASIS_GEOMETRY_DIGEST,
            "ruby and emphasis page geometry moved:\n{geometry}"
        );

        // F-266a's page is measured again here so this story cannot move its
        // sibling's recorded baseline without failing.
        let sibling = mixed_script_document()
            .layout_deterministic()
            .expect("deterministic mixed-script layout");
        assert_eq!(
            digest(&canonical_geometry(&sibling)),
            MIXED_SCRIPT_GEOMETRY_DIGEST,
            "F-266a's recorded geometry moved"
        );
    }

    /// Every `ST_Em` value authors, saves and reopens typed, and `w:em` lands
    /// in its `EG_RPrBase` sequence slot between `w:cs` and
    /// `w:eastAsianLayout`.
    #[test]
    fn emphasis_marks_reopen_as_modeled_state() {
        let marks = [
            ST_Em::None,
            ST_Em::Dot,
            ST_Em::Comma,
            ST_Em::Circle,
            ST_Em::UnderDot,
        ];
        let mut document = Document::new();
        for mark in &marks {
            let mut paragraph = document.add_paragraph("");
            let mut run = paragraph.add_run("marked");
            run.set_complex_script_value(Some(true));
            run.set_emphasis_mark_value(Some(mark.clone()));
            run.set_east_asian_layout_value(Some(rdocx::CT_EastAsianLayout {
                vert: Some(true),
                ..Default::default()
            }));
        }

        let saved = document.to_bytes().unwrap();
        let xml = saved_document_xml(&saved);
        let cs = xml.find("<w:cs/>").expect("w:cs is written");
        let em = xml
            .find(r#"<w:em w:val="none"/>"#)
            .expect("w:em is written");
        let layout = xml
            .find("<w:eastAsianLayout")
            .expect("w:eastAsianLayout is written");
        assert!(
            cs < em && em < layout,
            "w:em sits between w:cs and w:eastAsianLayout: {xml}"
        );

        let reopened = Document::from_bytes(&saved).unwrap();
        let paragraphs = reopened.paragraphs();
        let reopened_marks = paragraphs
            .iter()
            .map(|paragraph| {
                paragraph
                    .runs()
                    .next()
                    .unwrap()
                    .emphasis_mark()
                    .unwrap()
                    .clone()
            })
            .collect::<Vec<_>>();
        assert_eq!(reopened_marks, marks);
    }

    /// A producer `w:em` naming a token outside the ECMA-376 inventory is
    /// kept rather than normalised away.
    #[test]
    fn a_producer_emphasis_mark_with_unknown_attributes_is_retained_verbatim() {
        let source = producer_document(concat!(
            r#"<w:p><w:r><w:rPr><w:em w:val="producerMark"/></w:rPr>"#,
            r#"<w:t>marked</w:t></w:r></w:p>"#,
        ));
        let mut document = Document::from_bytes(&source).unwrap();
        let saved = document.to_bytes().unwrap();
        assert!(
            saved_document_xml(&saved).contains(r#"<w:em w:val="producerMark"/>"#),
            "the producer token survives the round trip"
        );
        let reopened = Document::from_bytes(&saved).unwrap();
        let paragraphs = reopened.paragraphs();
        assert_eq!(
            paragraphs[0].runs().next().unwrap().emphasis_mark(),
            Some(&ST_Em::Other("producerMark".to_owned()))
        );
    }

    /// `w:ruby` round-trips typed, in `xsd:sequence` order, and a prefix
    /// aliased producer ruby writes back with the fixed `w:` prefix.
    #[test]
    fn ruby_authors_saves_and_reopens_with_base_and_phonetic_runs() {
        let mut document = Document::new();
        {
            let mut paragraph = document.add_paragraph("");
            paragraph.add_run("before ");
            let index = paragraph.add_ruby(RUBY_BASE, RUBY_TEXT);
            paragraph.ruby_mut(index).unwrap().properties = Some(rdocx::CT_RubyPr {
                align: Some(ST_RubyAlign::Center),
                hps: Some(HalfPoint(10)),
                hps_raise: Some(HalfPoint(22)),
                hps_base_text: Some(HalfPoint(21)),
                language: Some("ja-JP".to_owned()),
                dirty: Some(false),
                raw_xml: Vec::new(),
            });
            paragraph.add_run(" after");
        }
        let saved = document.to_bytes().unwrap();
        let xml = saved_document_xml(&saved);
        assert!(
            xml.contains(concat!(
                r#"<w:ruby><w:rubyPr><w:rubyAlign w:val="center"/><w:hps w:val="10"/>"#,
                r#"<w:hpsRaise w:val="22"/><w:hpsBaseText w:val="21"/><w:lid w:val="ja-JP"/>"#,
                r#"<w:dirty w:val="false"/></w:rubyPr><w:rt><w:r><w:t>かんじ</w:t></w:r></w:rt>"#,
                r#"<w:rubyBase><w:r><w:t>漢字</w:t></w:r></w:rubyBase></w:ruby>"#,
            )),
            "the ruby writes in schema order with the fixed prefix: {xml}"
        );

        let reopened = Document::from_bytes(&saved).unwrap();
        let paragraphs = reopened.paragraphs();
        let ruby = paragraphs[0].ruby(0).expect("the ruby reopens typed");
        assert_eq!(ruby.ruby_text.len(), 1);
        assert_eq!(ruby.ruby_text[0].text(), RUBY_TEXT);
        assert_eq!(ruby.base_range(), 1..2);
        assert_eq!(
            ruby.properties.as_ref().and_then(|p| p.hps_raise),
            Some(HalfPoint(22))
        );
        assert_eq!(paragraphs[0].text(), format!("before {RUBY_BASE} after"));

        // The same ruby behind a producer alias reads through that alias and
        // writes back with the fixed `w:` prefix.
        let aliased = producer_document(concat!(
            r#"<w:p><q:ruby xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
            r#"<q:rt><q:r><q:t>かんじ</q:t></q:r></q:rt>"#,
            r#"<q:rubyBase><q:r><q:t>漢字</q:t></q:r></q:rubyBase></q:ruby></w:p>"#,
        ));
        let mut document = Document::from_bytes(&aliased).unwrap();
        let xml = saved_document_xml(&document.to_bytes().unwrap());
        assert!(
            xml.contains(
                r#"<w:ruby><w:rt><w:r><w:t>かんじ</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>漢字</w:t></w:r></w:rubyBase></w:ruby>"#
            ),
            "an aliased producer ruby writes back with the fixed prefix: {xml}"
        );
    }

    /// An unmodelled `w:rubyPr` child is preserved rather than dropped.
    #[test]
    fn unmodelled_ruby_properties_survive_a_noop_save() {
        let source = producer_document(concat!(
            r#"<w:p><w:ruby><w:rubyPr><w:hps w:val="10"/>"#,
            r#"<w:producerOnly w:val="1"/></w:rubyPr>"#,
            r#"<w:rt><w:r><w:t>かんじ</w:t></w:r></w:rt>"#,
            r#"<w:rubyBase><w:r><w:t>漢字</w:t></w:r></w:rubyBase></w:ruby></w:p>"#,
        ));
        let mut document = Document::from_bytes(&source).unwrap();
        let xml = saved_document_xml(&document.to_bytes().unwrap());
        assert!(
            xml.contains(r#"<w:producerOnly w:val="1"/>"#),
            "the unmodelled property child is preserved: {xml}"
        );
    }

    /// A ruby line claims the raise and the phonetic size as real height, so
    /// the paginator sees the taller line and breaks the page on it.
    #[test]
    fn a_ruby_line_is_taller_and_paginates_on_its_real_height() {
        fn page_count(with_ruby: bool, paragraphs: usize) -> usize {
            let mut document = Document::new();
            for _ in 0..paragraphs {
                let mut paragraph = document.add_paragraph("");
                if with_ruby {
                    let index = paragraph.add_ruby(RUBY_BASE, RUBY_TEXT);
                    paragraph.ruby_mut(index).unwrap().properties = Some(rdocx::CT_RubyPr {
                        hps: Some(HalfPoint(22)),
                        hps_raise: Some(HalfPoint(120)),
                        ..Default::default()
                    });
                } else {
                    paragraph.add_run(RUBY_BASE);
                }
            }
            document
                .layout_deterministic()
                .expect("deterministic layout")
                .layout
                .pages
                .len()
        }

        // The plain paragraphs fit one page. The same count with a 60 point
        // raise does not, which is only true if the raise reached the line.
        assert_eq!(page_count(false, 30), 1);
        assert!(
            page_count(true, 30) > 1,
            "the raise and the phonetic size reach the line height"
        );
    }

    /// A mark codepoint the resolved font cannot draw records a diagnostic
    /// and paints nothing, leaving the base text untouched.
    #[test]
    fn an_undrawable_emphasis_mark_records_a_diagnostic_and_paints_nothing() {
        fn painted(mark: Option<ST_Em>) -> (Vec<String>, Vec<String>) {
            let mut document = Document::new();
            {
                let mut paragraph = document.add_paragraph("");
                let mut run = paragraph.add_run(EMPHASIS_LATIN);
                run.set_font("Carlito");
                run.set_emphasis_mark_value(mark);
            }
            let result = document
                .layout_deterministic()
                .expect("deterministic layout");
            let mut text = Vec::new();
            for page in &result.layout.pages {
                oxml_layout::walk(&page.elements, &mut |element, _| {
                    if let oxml_layout::PositionedElement::Text(run) = element
                        && !run.text.trim().is_empty()
                    {
                        text.push(run.text.clone());
                    }
                });
            }
            (
                text,
                result
                    .layout
                    .diagnostics
                    .iter()
                    .map(|diagnostic| diagnostic.message.clone())
                    .collect(),
            )
        }

        let (plain, plain_diagnostics) = painted(None);
        assert!(plain_diagnostics.is_empty(), "{plain_diagnostics:?}");
        let (marked, diagnostics) = painted(Some(ST_Em::Comma));
        assert_eq!(
            marked, plain,
            "an undrawable mark paints nothing and leaves the base text alone"
        );
        assert!(
            diagnostics
                .iter()
                .any(|message| message.contains("emphasis mark comma")),
            "the undrawable mark is diagnosed: {diagnostics:?}"
        );
    }
}

mod floating_table_placement_and_wrap {
    use rdocx::table::{
        TableAnchor, TableFloatPosition, TableFloatX, TableFloatY, TableTextDistance,
    };
    use rdocx::{Document, Length};

    /// Letter page, one inch margins, which is what `Document::new` produces.
    const MARGIN_LEFT: f64 = 72.0;
    const MARGIN_TOP: f64 = 72.0;

    /// A float position with the four from-text distances every case shares.
    ///
    /// 180 twips is 9 points and 80 twips is 4 points, so the reserved band is
    /// readable in the pinned geometry rather than an artefact of rounding.
    fn float_at(
        anchor_h: TableAnchor,
        anchor_v: TableAnchor,
        x: i32,
        y: i32,
    ) -> TableFloatPosition {
        TableFloatPosition {
            horizontal_anchor: anchor_h,
            vertical_anchor: anchor_v,
            horizontal: TableFloatX::Offset(Length::twips(x)),
            vertical: TableFloatY::Offset(Length::twips(y)),
            distance_from_text: TableTextDistance {
                top: Length::twips(80),
                right: Length::twips(180),
                bottom: Length::twips(80),
                left: Length::twips(180),
            },
        }
    }

    /// A two by two table 100 points wide, labelled so its cells are
    /// distinguishable from body text in the placed elements.
    fn add_float(document: &mut Document, label: &str, position: Option<TableFloatPosition>) {
        let mut table = document.add_table(2, 2);
        table.set_column_width(0, Length::twips(1000));
        table.set_column_width(1, Length::twips(1000));
        table.set_width(Length::twips(2000));
        table
            .set_float_position(position)
            .expect("float position is valid");
        for row in 0..2 {
            for column in 0..2 {
                table
                    .cell(row, column)
                    .expect("cell exists")
                    .set_text(&format!("{label}{row}{column}"));
            }
        }
    }

    fn prose(document: &mut Document, from: usize, to: usize) {
        for index in from..to {
            document.add_paragraph(&format!(
                "Body line {index:02} with enough words to wrap across the measure of the page and show the reservation."
            ));
        }
    }

    fn round(value: f64) -> f64 {
        (value * 100.0).round() / 100.0
    }

    /// Where one direct body item landed, rounded, as (page, x, y, w, h).
    fn placed(
        result: &rdocx_layout::WordLayoutResult,
        body_index: usize,
    ) -> Vec<(usize, f64, f64, f64, f64)> {
        result
            .body_layout_fragments(body_index)
            .expect("body index is in range")
            .iter()
            .map(|fragment| {
                (
                    fragment.physical_page,
                    round(fragment.x),
                    round(fragment.y),
                    round(fragment.width),
                    round(fragment.height),
                )
            })
            .collect()
    }

    /// Body line boxes of one page as (baseline, left edge, right edge).
    ///
    /// A float's own cell text is excluded by its label, so what remains is the
    /// prose that had to flow around it.
    fn body_line_boxes(page: &oxml_layout::PageFrame) -> Vec<(f64, f64, f64)> {
        let mut lines: Vec<(f64, f64, f64)> = Vec::new();
        oxml_layout::walk(&page.elements, &mut |element, _| {
            let oxml_layout::PositionedElement::Text(run) = element else {
                return;
            };
            let is_cell_label = run.text.len() == 3
                && run.text.starts_with(['M', 'P', 'T'])
                && run.text[1..].chars().all(|glyph| glyph.is_ascii_digit());
            if run.text.trim().is_empty() || is_cell_label {
                return;
            }
            let baseline = round(run.origin.y);
            let left = run.origin.x;
            let right = run.origin.x + run.advances.iter().sum::<f64>();
            match lines.iter_mut().find(|line| line.0 == baseline) {
                Some(line) => {
                    line.1 = line.1.min(left);
                    line.2 = line.2.max(right);
                }
                None => lines.push((baseline, left, right)),
            }
        });
        lines.sort_by(|left, right| left.0.total_cmp(&right.0));
        lines
            .into_iter()
            .map(|(baseline, left, right)| (baseline, round(left), round(right)))
            .collect()
    }

    /// The body line boxes whose baseline falls inside one float's keep-out
    /// band, which is exactly the text that float pushed aside.
    fn boxes_in_band(page: &oxml_layout::PageFrame, top: f64, bottom: f64) -> Vec<(f64, f64, f64)> {
        body_line_boxes(page)
            .into_iter()
            .filter(|(baseline, ..)| *baseline > top && *baseline < bottom)
            .collect()
    }

    /// The reviewed geometry of a page carrying a margin-anchored float, a
    /// page-anchored float and a text-anchored float.
    ///
    /// Recorded in deterministic font mode. The margin float and the text float
    /// sit on the left, so the lines beside them start at 181 points, which is
    /// the left margin plus the 100 point table plus its 9 point right
    /// clearance. The page float sits on the right, so the lines beside it keep
    /// their left edge and lose their right.
    #[test]
    fn floating_tables_match_reviewed_word_page_geometry_and_pagination() {
        let mut document = Document::new();
        prose(&mut document, 0, 4);
        add_float(
            &mut document,
            "M",
            Some(float_at(TableAnchor::Margin, TableAnchor::Margin, 0, 0)),
        );
        prose(&mut document, 4, 10);
        add_float(
            &mut document,
            "T",
            Some(float_at(TableAnchor::Text, TableAnchor::Text, 0, 0)),
        );
        prose(&mut document, 10, 20);
        add_float(
            &mut document,
            "P",
            Some(float_at(TableAnchor::Page, TableAnchor::Page, 7200, 9000)),
        );
        prose(&mut document, 20, 30);

        let result = document.layout_deterministic().expect("document lays out");
        assert_eq!(result.layout.pages.len(), 2, "page count moved");

        // Each float sits whole on one page, at the rect its anchor resolves
        // to, and none of them advanced the flow.
        assert_eq!(placed(&result, 4), [(1, 72.0, 72.0, 100.0, 39.74)]);
        assert_eq!(placed(&result, 11), [(1, 72.0, 294.45, 100.0, 39.74)]);
        assert_eq!(placed(&result, 22), [(1, 360.0, 450.0, 100.0, 39.74)]);

        let first = &result.layout.pages[0];
        assert_eq!(
            boxes_in_band(first, 68.0, 115.75),
            [
                (80.25, 181.0, 525.39),
                (92.12, 181.0, 277.95),
                (111.99, 181.0, 525.39),
            ],
            "margin-anchored float geometry moved"
        );
        assert_eq!(
            boxes_in_band(first, 290.45, 338.2),
            [
                (302.7, 181.0, 525.39),
                (314.57, 181.0, 277.95),
                (334.44, 181.0, 525.39),
            ],
            "text-anchored float geometry moved"
        );
        assert_eq!(
            boxes_in_band(first, 446.0, 493.75),
            [
                (457.54, 72.0, 241.43),
                (477.41, 72.0, 343.91),
                (489.28, 72.0, 241.43),
            ],
            "page-anchored float geometry moved"
        );
    }

    /// `w:horzAnchor` and `w:vertAnchor` name three frames, and each resolves
    /// to a different origin. Margin and text coincide horizontally, because
    /// the text column and the margin start at the same edge, so the vertical
    /// frame is what separates them.
    #[test]
    fn tblp_pr_anchors_map_onto_the_drawing_anchor_frames() {
        let origin_for = |anchor: TableAnchor| {
            let mut document = Document::new();
            document.add_paragraph("One line above the float.");
            add_float(&mut document, "M", Some(float_at(anchor, anchor, 0, 0)));
            let result = document.layout_deterministic().expect("document lays out");
            let fragments = placed(&result, 1);
            assert_eq!(fragments.len(), 1, "a float occupies one page");
            (fragments[0].1, fragments[0].2)
        };

        // The page frame starts at the physical page corner.
        assert_eq!(origin_for(TableAnchor::Page), (0.0, 0.0));
        // The margin frame starts at the top left of the text area.
        assert_eq!(origin_for(TableAnchor::Margin), (MARGIN_LEFT, MARGIN_TOP));
        // The text frame starts where the floating block itself landed, which
        // is below the paragraph above it.
        let (text_x, text_y) = origin_for(TableAnchor::Text);
        assert_eq!(text_x, MARGIN_LEFT);
        assert!(
            text_y > MARGIN_TOP,
            "a text anchor follows the flow, got {text_y}"
        );

        // `tblpYSpec="inline"` is how `w:tblpPr` spells "not floating", so the
        // table keeps the flow position an inline table would have had.
        let mut document = Document::new();
        document.add_paragraph("One line above the float.");
        let mut inline = float_at(TableAnchor::Margin, TableAnchor::Margin, 0, 0);
        inline.vertical = TableFloatY::Inline;
        add_float(&mut document, "M", Some(inline));
        let result = document.layout_deterministic().expect("document lays out");
        assert_eq!(placed(&result, 1), [(1, MARGIN_LEFT, text_y, 100.0, 39.74)]);
    }

    /// `w:tblpPr` is a position, and a position leaves no room for `w:tblInd`
    /// to contribute. An inline table with the same indent still takes it.
    #[test]
    fn a_floating_table_takes_its_origin_from_the_anchor_not_the_indent() {
        let indented = |position: Option<TableFloatPosition>| {
            let mut document = Document::new();
            document.add_paragraph("One line above the table.");
            {
                let mut table = document.add_table(2, 2);
                table.set_column_width(0, Length::twips(1000));
                table.set_column_width(1, Length::twips(1000));
                table.set_width(Length::twips(2000));
                table.set_indent(Length::twips(1440));
                table
                    .set_float_position(position)
                    .expect("float position is valid");
            }
            let result = document.layout_deterministic().expect("document lays out");
            placed(&result, 1)[0].1
        };

        let float = float_at(TableAnchor::Margin, TableAnchor::Margin, 0, 0);
        assert_eq!(indented(Some(float)), MARGIN_LEFT);
        assert_eq!(indented(None), MARGIN_LEFT + 72.0);
    }
}

/// F-266c, the East Asian character grid and vertical text.
///
/// The gate is a recorded geometry digest over a deterministic page carrying a
/// gridded Japanese section, a table with rotated and horizontal cells, and a
/// combined run. It uses no rasteriser and no external oracle, because what is
/// under test is glyph identity, placement, rotation and painted order, all of
/// which the layout result already states exactly.
///
/// The serialisation this module records is F-266a's with the accumulated
/// group transform added, because a rotation is invisible in a glyph run's own
/// group-local origin and this story's whole subject is rotation.
///
/// Both sibling digests are asserted unmoved in the same module, so this story
/// cannot quietly move F-266a's or F-266b's recorded baseline while recording
/// its own.
mod f266c_character_grid_and_vertical_text {
    use super::f266a_mixed_script_typography::{
        MIXED_SCRIPT_GEOMETRY_DIGEST, canonical_geometry, digest, mixed_script_document,
    };
    use super::f266b_ruby_and_emphasis_typography::{
        RUBY_AND_EMPHASIS_GEOMETRY_DIGEST, ruby_and_emphasis_document,
    };
    use super::*;
    use oxml_layout::{PositionedElement, Transform};
    use rdocx::table::CellTextDirection;
    use rdocx::{CT_DocGrid, CT_EastAsianLayout, RunFontSlot, ST_DocGrid, ST_Em};
    use rdocx_oxml::units::Twips;

    const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    const GRIDDED_JAPANESE: &str = "こんにちは世界";
    const COMBINED: &str = "世界";
    const VERTICAL_CELL: &str = "縦書き";
    const HORIZONTAL_CELL: &str = "Across";
    const LATIN_CONTROL: &str = "Grid and vertical page";
    const ROTATED_RUN: &str = "縦書き";
    /// One character of `ROTATED_RUN`. Shaping splits East Asian text at every
    /// character, so a whole-string match would find nothing.
    const ROTATED_GLYPH: &str = "縦";

    /// The recorded geometry of the gridded and vertical page.
    ///
    /// Re-record only with a stated reason. It covers every painted run on the
    /// page in paint order with the properties F-266a's serialisation records,
    /// and the six coefficients of the transform that maps the run into page
    /// space, which is what makes a lost or altered rotation fail here.
    const GRID_AND_VERTICAL_GEOMETRY_DIGEST: &str =
        "cb3043d53719f5dd9e16b61a001aff8c8827c19f96536972d4a17b9a626d2164";

    /// One coordinate, with the sign of zero normalised, as F-266a documents.
    fn number(value: f64) -> String {
        format!("{:.4}", if value == 0.0 { 0.0 } else { value })
    }

    /// F-266a's canonical serialisation, plus the transform each run carries.
    fn canonical_rotated_geometry(result: &rdocx_layout::WordLayoutResult) -> String {
        let mut lines = Vec::new();
        for (page_index, page) in result.layout.pages.iter().enumerate() {
            let mut paint_index = 0usize;
            oxml_layout::walk(&page.elements, &mut |element, transform| {
                let text = match element {
                    PositionedElement::Text(run) => run.text.clone(),
                    PositionedElement::MultilingualText(run) => run.logical_text.clone(),
                    _ => return,
                };
                let origin = match element {
                    PositionedElement::Text(run) => run.origin,
                    PositionedElement::MultilingualText(run) => run.origin,
                    _ => return,
                };
                lines.push(format!(
                    "page={page_index} paint={paint_index} text={text} \
                     origin={},{} transform={},{},{},{},{},{}",
                    number(origin.x),
                    number(origin.y),
                    number(transform.a),
                    number(transform.b),
                    number(transform.c),
                    number(transform.d),
                    number(transform.e),
                    number(transform.f),
                ));
                paint_index += 1;
            });
        }
        lines.join("\n")
    }

    /// Every painted string on the page, in paint order.
    fn painted(result: &rdocx_layout::WordLayoutResult) -> Vec<String> {
        let mut painted = Vec::new();
        for page in &result.layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| match element {
                PositionedElement::Text(run) if !run.text.trim().is_empty() => {
                    painted.push(run.text.clone());
                }
                PositionedElement::MultilingualText(run) if !run.logical_text.trim().is_empty() => {
                    painted.push(run.logical_text.clone());
                }
                _ => {}
            });
        }
        painted
    }

    /// The page-space x the painted run carrying `text` starts at.
    ///
    /// A run placed after another one starts where its predecessor's advance
    /// ended, so this is how far the run before it reached along the line.
    fn painted_origin_x(result: &rdocx_layout::WordLayoutResult, text: &str) -> f64 {
        let mut found = None;
        for page in &result.layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, transform| {
                let origin = match element {
                    PositionedElement::Text(run) if run.text.contains(text) => run.origin,
                    PositionedElement::MultilingualText(run) if run.logical_text.contains(text) => {
                        run.origin
                    }
                    _ => return,
                };
                if found.is_none() {
                    found = Some(transform.apply(origin).x);
                }
            });
        }
        found.unwrap_or_else(|| panic!("{text} reaches the page"))
    }

    /// The transform every painted run carrying `text` reached the page with.
    fn transforms_for(result: &rdocx_layout::WordLayoutResult, text: &str) -> Vec<Transform> {
        let mut found = Vec::new();
        for page in &result.layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, transform| {
                let matches = match element {
                    PositionedElement::Text(run) => run.text.contains(text),
                    PositionedElement::MultilingualText(run) => run.logical_text.contains(text),
                    _ => false,
                };
                if matches {
                    found.push(*transform);
                }
            });
        }
        found
    }

    /// Wrap producer body XML in a package a `Document` can open.
    fn producer_body(body: &str) -> Vec<u8> {
        let mut seed = Document::new();
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
        package.set_part(
            "/word/document.xml",
            format!(
                r#"<w:document xmlns:w="{W_NS}"><w:body>{body}<w:sectPr/></w:body></w:document>"#
            )
            .into_bytes(),
        );
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        bytes.into_inner()
    }

    /// A `linesAndChars` grid, the type that puts both axes on the grid.
    fn lines_and_chars_grid() -> CT_DocGrid {
        CT_DocGrid {
            grid_type: Some(ST_DocGrid::LinesAndChars),
            line_pitch: Some(Twips(360)),
            char_space: Some(120),
            extra_attributes: Vec::new(),
        }
    }

    /// One page holding a gridded Japanese section, a table with a rotated
    /// cell on each side of a horizontal control, a combined run and a Latin
    /// control, all authored through the public facade.
    fn grid_and_vertical_document() -> Document {
        let mut document = Document::new();

        let mut latin = document.add_paragraph("");
        latin
            .add_run(LATIN_CONTROL)
            .font("Carlito")
            .language("en-US");

        let mut japanese = document.add_paragraph("");
        {
            let mut run = japanese.add_run(GRIDDED_JAPANESE);
            run.set_slot_font(RunFontSlot::EastAsia, Some("Noto Sans JP"));
            run.set_language_east_asia_value(Some("ja-JP"));
        }

        let mut combined = document.add_paragraph("");
        {
            let mut run = combined.add_run(COMBINED);
            run.set_slot_font(RunFontSlot::EastAsia, Some("Noto Sans JP"));
            run.set_language_east_asia_value(Some("ja-JP"));
            run.set_east_asian_layout_value(Some(CT_EastAsianLayout {
                combine: Some(true),
                combine_brackets: Some("round".to_owned()),
                ..CT_EastAsianLayout::default()
            }));
        }

        {
            let mut table = document.add_table(1, 3);
            for (column, direction, text) in [
                (0, CellTextDirection::TopToBottomRightToLeft, VERTICAL_CELL),
                (
                    1,
                    CellTextDirection::LeftToRightTopToBottom,
                    HORIZONTAL_CELL,
                ),
                (2, CellTextDirection::BottomToTopLeftToRight, VERTICAL_CELL),
            ] {
                let mut cell = table.cell(0, column).expect("authored cell");
                cell.set_text_direction(Some(direction));
                cell.set_text(text);
                let mut paragraph = cell.paragraph_mut(0).expect("cell paragraph");
                let mut run = paragraph.run_mut(0).expect("cell run");
                run.set_slot_font(RunFontSlot::EastAsia, Some("Noto Sans JP"));
                run.set_slot_font(RunFontSlot::Ascii, Some("Carlito"));
            }
        }

        document
            .section_mut(0)
            .expect("final section")
            .set_doc_grid(Some(lines_and_chars_grid()));

        document
    }

    /// **The test gate.** The gridded and vertical page keeps its recorded
    /// geometry and its reading order, and both siblings' pages are unmoved.
    #[test]
    fn grid_and_vertical_page_matches_the_pinned_geometry_and_reading_order() {
        let mut document = grid_and_vertical_document();
        let result = document
            .layout_deterministic()
            .expect("deterministic grid and vertical layout");
        assert_eq!(result.layout.pages.len(), 1, "the fixture is one page");

        // Shaping splits a run at every break opportunity, so containment is
        // asserted over the whole painted page rather than one painted run.
        let painted = painted(&result);
        let whole = painted.concat();
        for text in [LATIN_CONTROL, GRIDDED_JAPANESE, HORIZONTAL_CELL] {
            assert!(whole.contains(text), "{text} reaches the page: {painted:?}");
        }
        assert_eq!(
            whole.matches(VERTICAL_CELL).count(),
            2,
            "both rotated cells paint their text: {painted:?}"
        );

        // The rotated cells reach the page through a rotation, and the
        // horizontal control does not. Shaping splits the vertical cell's
        // text at every East Asian break opportunity, so the transforms are
        // collected on one of its characters rather than the whole string.
        let rotated = transforms_for(&result, "縦");
        assert_eq!(
            rotated.len(),
            2,
            "both rotated cells reach the page: {rotated:?}"
        );
        for transform in &rotated {
            assert!(
                !transform.is_identity(),
                "a rotated cell carries a transform: {transform:?}"
            );
        }
        let horizontal = transforms_for(&result, HORIZONTAL_CELL);
        assert_eq!(horizontal.len(), 1, "the horizontal control paints once");
        for transform in &horizontal {
            assert!(
                transform.is_identity(),
                "the horizontal control keeps the untransformed path: {transform:?}"
            );
        }

        // Rotation and combining are painting concerns. The saved bytes stay
        // logical, and the cell text comes back in grid order.
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .paragraphs()
                .iter()
                .map(|paragraph| paragraph.text())
                .collect::<Vec<_>>(),
            vec![LATIN_CONTROL, GRIDDED_JAPANESE, COMBINED]
        );
        let tables = reopened.tables();
        let table = tables.first().expect("the authored table");
        assert_eq!(
            (0..3)
                .map(|column| table.cell(0, column).expect("cell").text())
                .collect::<Vec<_>>(),
            vec![VERTICAL_CELL, HORIZONTAL_CELL, VERTICAL_CELL]
        );

        let geometry = canonical_rotated_geometry(&result);
        assert_eq!(
            digest(&geometry),
            GRID_AND_VERTICAL_GEOMETRY_DIGEST,
            "grid and vertical page geometry moved:\n{geometry}"
        );

        // Both siblings' pages are measured again here, so this story cannot
        // move either recorded baseline without failing.
        let mixed = mixed_script_document()
            .layout_deterministic()
            .expect("deterministic mixed-script layout");
        assert_eq!(
            digest(&canonical_geometry(&mixed)),
            MIXED_SCRIPT_GEOMETRY_DIGEST,
            "F-266a's recorded geometry moved"
        );
        let ruby = ruby_and_emphasis_document()
            .layout_deterministic()
            .expect("deterministic ruby and emphasis layout");
        assert_eq!(
            digest(&canonical_geometry(&ruby)),
            RUBY_AND_EMPHASIS_GEOMETRY_DIGEST,
            "F-266b's recorded geometry moved"
        );
    }

    /// `w:docGrid` authors, saves and reopens typed at its `w:sectPr` sequence
    /// position, and the unrelated producer children stay byte for byte.
    #[test]
    fn doc_grid_reopens_on_its_section() {
        for (grid_type, expected) in [
            (ST_DocGrid::Default, "default"),
            (ST_DocGrid::Lines, "lines"),
            (ST_DocGrid::LinesAndChars, "linesAndChars"),
            (ST_DocGrid::SnapToChars, "snapToChars"),
        ] {
            let mut document = Document::new();
            document.add_paragraph("body");
            document
                .section_mut(0)
                .expect("final section")
                .set_doc_grid(Some(CT_DocGrid {
                    grid_type: Some(grid_type),
                    line_pitch: Some(Twips(312)),
                    char_space: Some(179),
                    extra_attributes: Vec::new(),
                }));

            let bytes = document.to_bytes().unwrap();
            let reopened = Document::from_bytes(&bytes).unwrap();
            let grid = reopened
                .sections()
                .next()
                .unwrap()
                .doc_grid()
                .expect("the grid reopens typed")
                .clone();
            assert_eq!(grid.grid_type, Some(grid_type));
            assert_eq!(grid.line_pitch, Some(Twips(312)));
            assert_eq!(grid.char_space, Some(179));

            let xml = String::from_utf8(
                OpcPackage::from_reader(std::io::Cursor::new(&bytes))
                    .unwrap()
                    .get_part("/word/document.xml")
                    .unwrap()
                    .to_vec(),
            )
            .unwrap();
            assert!(
                xml.contains(&format!(
                    "<w:docGrid w:type=\"{expected}\" w:linePitch=\"312\" w:charSpace=\"179\"/>"
                )),
                "the fixed `w:` prefix and the attribute order are written: {xml}"
            );
            let grid_at = xml.find("<w:docGrid").expect("w:docGrid is written");
            let sect_end = xml.find("</w:sectPr>").expect("the section closes");
            assert!(grid_at < sect_end, "w:docGrid sits inside w:sectPr");
        }
    }

    /// The whole `w:docGrid` sequence position, including the producer
    /// children that share its slot, is preserved byte for byte.
    #[test]
    fn an_unmodelled_doc_grid_attribute_survives_a_noop_save() {
        let sect_pr = concat!(
            "<w:sectPr>",
            "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>",
            "<w:textDirection w:val=\"tbRl\"/>",
            "<w:bidi w:val=\"0\"/>",
            "<w:rtlGutter w:val=\"0\"/>",
            "<w:docGrid xmlns:x=\"urn:producer\" x:kept=\"grid\" ",
            "w:type=\"linesAndChars\" w:linePitch=\"360\" w:charSpace=\"120\"/>",
            "<w:printerSettings r:id=\"rIdPrinter\" ",
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>",
            "</w:sectPr>",
        );
        let source = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p>{sect_pr}</w:body></w:document>"#
        );
        let mut seed = Document::new();
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
        package.set_part("/word/document.xml", source.into_bytes());
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();

        let mut document = Document::from_bytes(&bytes.into_inner()).unwrap();
        let grid = document
            .sections()
            .next()
            .unwrap()
            .doc_grid()
            .expect("the grid is typed")
            .clone();
        assert_eq!(grid.grid_type, Some(ST_DocGrid::LinesAndChars));
        assert_eq!(grid.line_pitch, Some(Twips(360)));
        assert_eq!(grid.char_space, Some(120));
        assert_eq!(
            grid.extra_attributes,
            vec![
                ("xmlns:x".to_owned(), "urn:producer".to_owned()),
                ("x:kept".to_owned(), "grid".to_owned()),
            ],
            "the producer attribute the model does not own is retained"
        );

        let saved = String::from_utf8(
            OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
                .unwrap()
                .get_part("/word/document.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        // The writer indents between children, so the comparison states child
        // order and attribute bytes rather than pretty printing.
        let saved = saved
            .split('\n')
            .map(str::trim_start)
            .collect::<Vec<_>>()
            .concat();
        let start = saved.find("<w:sectPr>").expect("the section is written");
        let end = saved.find("</w:sectPr>").expect("the section closes") + "</w:sectPr>".len();
        assert_eq!(&saved[start..end], sect_pr);
    }

    /// `w:eastAsianLayout` round-trips typed, prefix-aliased on read and with
    /// the fixed `w:` prefix on write.
    #[test]
    fn east_asian_layout_reopens_as_modeled_state() {
        let mut document = Document::new();
        {
            let mut paragraph = document.add_paragraph("");
            let mut run = paragraph.add_run("combined");
            run.set_east_asian_layout_value(Some(CT_EastAsianLayout {
                id: Some(7),
                combine: Some(true),
                combine_brackets: Some("square".to_owned()),
                vert: Some(true),
                vert_compress: Some(true),
                extra_attributes: Vec::new(),
            }));
        }
        let bytes = document.to_bytes().unwrap();
        let xml = String::from_utf8(
            OpcPackage::from_reader(std::io::Cursor::new(&bytes))
                .unwrap()
                .get_part("/word/document.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        assert!(
            xml.contains(concat!(
                "<w:eastAsianLayout w:id=\"7\" w:combine=\"1\" ",
                "w:combineBrackets=\"square\" w:vert=\"1\" w:vertCompress=\"1\"/>"
            )),
            "the fixed prefix and the attribute order are written: {xml}"
        );

        let reopened = Document::from_bytes(&bytes).unwrap();
        let paragraph = reopened.paragraphs().into_iter().next().unwrap();
        let run = paragraph.runs().next().expect("the authored run");
        let layout = run.east_asian_layout().expect("the layout reopens typed");
        assert_eq!(layout.id, Some(7));
        assert_eq!(layout.combine, Some(true));
        assert_eq!(layout.combine_brackets.as_deref(), Some("square"));
        assert_eq!(layout.vert, Some(true));
        assert_eq!(layout.vert_compress, Some(true));
    }

    /// `w:vert` rotates the run inside its line, `w:vertCompress` narrows the
    /// rotated run to one base advance, and `w:combine` wins when a run asks
    /// for both. A run carrying `w:em` as well takes the East Asian layout,
    /// because combining and rotating change the run's advance.
    #[test]
    fn a_vertical_run_rotates_inside_its_line_and_compresses_on_request() {
        let build = |layout: CT_EastAsianLayout, mark: Option<ST_Em>| {
            let mut document = Document::new();
            {
                let mut paragraph = document.add_paragraph("");
                {
                    // A bundled East Asian face, because `w:vertCompress`
                    // narrows a rotated run to one em and a face whose ascent
                    // and descent already sum to one em could not show it.
                    let mut run = paragraph.add_run(ROTATED_RUN);
                    run.set_slot_font(RunFontSlot::EastAsia, Some("Noto Sans JP"));
                    run.set_language_east_asia_value(Some("ja-JP"));
                    run.set_east_asian_layout_value(Some(layout));
                    if let Some(mark) = mark {
                        run.set_emphasis_mark_value(Some(mark));
                    }
                }
                // A trailing run starts where the run before it stopped, so
                // its origin states the advance the East Asian layout took.
                paragraph.add_run("END").set_font("Carlito");
            }
            document
                .layout_deterministic()
                .expect("deterministic rotated run layout")
        };

        // An ordinary run reaches the page untransformed.
        let plain = build(CT_EastAsianLayout::default(), None);
        let plain_transforms = transforms_for(&plain, ROTATED_GLYPH);
        assert!(!plain_transforms.is_empty(), "the plain run paints");
        assert!(
            plain_transforms
                .iter()
                .all(|transform| transform.is_identity()),
            "a run with neither combine nor vert keeps the ordinary path"
        );
        // A blank document is US Letter with one-inch margins, so the line
        // starts at 72 points and the advance is measured from there.
        const LINE_START: f64 = 72.0;
        let plain_advance = painted_origin_x(&plain, "END") - LINE_START;

        // `w:vert` rotates it 90 degrees within the line.
        let rotated = build(
            CT_EastAsianLayout {
                vert: Some(true),
                ..CT_EastAsianLayout::default()
            },
            None,
        );
        let transforms = transforms_for(&rotated, ROTATED_GLYPH);
        assert_eq!(transforms.len(), 1, "the rotated run paints once");
        assert!(
            (transforms[0].b - 1.0).abs() < 1e-9 && (transforms[0].a).abs() < 1e-9,
            "the run rotates 90 degrees inside its line: {:?}",
            transforms[0]
        );
        // Rotated, the run takes its own line height along the line rather
        // than its text length, so it advances far less than it did.
        let rotated_advance = painted_origin_x(&rotated, "END") - LINE_START;
        assert!(
            rotated_advance < plain_advance / 2.0,
            "the rotated run takes its line height along the line, \
             {rotated_advance} against {plain_advance}"
        );

        // `w:vertCompress` narrows the rotated run further.
        let compressed = build(
            CT_EastAsianLayout {
                vert: Some(true),
                vert_compress: Some(true),
                ..CT_EastAsianLayout::default()
            },
            None,
        );
        assert!(
            painted_origin_x(&compressed, "END") - LINE_START < rotated_advance,
            "vertCompress narrows the rotated run to one base advance"
        );

        // `w:combine` wins over `w:vert`, and both win over `w:em`.
        let combined = build(
            CT_EastAsianLayout {
                combine: Some(true),
                vert: Some(true),
                combine_brackets: Some("round".to_owned()),
                ..CT_EastAsianLayout::default()
            },
            Some(ST_Em::Dot),
        );
        let painted = painted(&combined).concat();
        assert!(
            painted.contains('('),
            "the combined run draws its brackets, so combine won: {painted}"
        );
        assert!(
            !painted.contains('\u{2022}'),
            "the East Asian layout wins over the emphasis mark: {painted}"
        );
    }

    /// The seven East Asian paragraph toggles F-264 left raw author, save,
    /// reopen and remove through the public paragraph surface.
    #[test]
    fn the_east_asian_paragraph_toggles_reopen_as_modeled_state() {
        let mut document = Document::new();
        {
            let mut paragraph = document.add_paragraph("gridded");
            paragraph.set_kinsoku_value(Some(false));
            paragraph.set_word_wrap_value(Some(false));
            paragraph.set_overflow_punct_value(Some(false));
            paragraph.set_top_line_punct_value(Some(true));
            paragraph.set_auto_space_de_value(Some(false));
            paragraph.set_auto_space_dn_value(Some(false));
            paragraph.set_snap_to_grid_value(Some(false));
        }
        let bytes = document.to_bytes().unwrap();
        let xml = String::from_utf8(
            OpcPackage::from_reader(std::io::Cursor::new(&bytes))
                .unwrap()
                .get_part("/word/document.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        let order = [
            "<w:kinsoku",
            "<w:wordWrap",
            "<w:overflowPunct",
            "<w:topLinePunct",
            "<w:autoSpaceDE",
            "<w:autoSpaceDN",
            "<w:snapToGrid",
        ];
        let mut previous = 0usize;
        for element in order {
            let at = xml
                .find(element)
                .unwrap_or_else(|| panic!("{element} is written: {xml}"));
            assert!(at > previous, "{element} follows its predecessor: {xml}");
            previous = at;
        }

        let reopened = Document::from_bytes(&bytes).unwrap();
        let paragraph = reopened.paragraphs().into_iter().next().unwrap();
        assert_eq!(paragraph.kinsoku_value(), Some(false));
        assert_eq!(paragraph.word_wrap_value(), Some(false));
        assert_eq!(paragraph.overflow_punct_value(), Some(false));
        assert_eq!(paragraph.top_line_punct_value(), Some(true));
        assert_eq!(paragraph.auto_space_de_value(), Some(false));
        assert_eq!(paragraph.auto_space_dn_value(), Some(false));
        assert_eq!(paragraph.snap_to_grid_value(), Some(false));

        let mut removable = Document::from_bytes(&bytes).unwrap();
        removable
            .paragraph_mut(0)
            .expect("the authored paragraph")
            .set_kinsoku_value(None);
        let removed = Document::from_bytes(&removable.to_bytes().unwrap()).unwrap();
        assert_eq!(
            removed
                .paragraphs()
                .into_iter()
                .next()
                .unwrap()
                .kinsoku_value(),
            None
        );
    }

    /// A producer spelling of the seven toggles reopens typed and is written
    /// in the canonical form, and an attribute the model does not own keeps
    /// the source element as its carrier.
    #[test]
    fn a_producer_spelling_of_an_east_asian_toggle_reopens_typed() {
        let body = concat!(
            "<w:p><w:pPr>",
            "<w:kinsoku w:val=\"true\"/>",
            "<w:wordWrap w:val=\"0\"/>",
            "<w:overflowPunct w:val=\"off\"/>",
            "<w:topLinePunct/>",
            "<w:autoSpaceDE w:val=\"false\"/>",
            "<w:autoSpaceDN w:val=\"1\"/>",
            "<w:snapToGrid w:val=\"on\"/>",
            "</w:pPr><w:r><w:t>body</w:t></w:r></w:p>",
        );
        let mut document = Document::from_bytes(&producer_body(body)).unwrap();
        {
            let paragraph = document.paragraphs().into_iter().next().unwrap();
            assert_eq!(paragraph.kinsoku_value(), Some(true));
            assert_eq!(paragraph.word_wrap_value(), Some(false));
            assert_eq!(paragraph.overflow_punct_value(), Some(false));
            assert_eq!(paragraph.top_line_punct_value(), Some(true));
            assert_eq!(paragraph.auto_space_de_value(), Some(false));
            assert_eq!(paragraph.auto_space_dn_value(), Some(true));
            assert_eq!(paragraph.snap_to_grid_value(), Some(true));
        }

        // A modeled toggle writes the canonical spelling, which is bare for
        // an on value and `w:val="false"` for an off one, exactly as the
        // toggles F-264 already modeled do.
        let saved = String::from_utf8(
            OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
                .unwrap()
                .get_part("/word/document.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        for expected in [
            "<w:kinsoku/>",
            "<w:wordWrap w:val=\"false\"/>",
            "<w:overflowPunct w:val=\"false\"/>",
            "<w:topLinePunct/>",
            "<w:autoSpaceDE w:val=\"false\"/>",
            "<w:autoSpaceDN/>",
            "<w:snapToGrid/>",
        ] {
            assert!(
                saved.contains(expected),
                "{expected} is written canonically: {saved}"
            );
        }
    }

    /// A `w:docGrid` attribute bound on an ancestor keeps its binding, because
    /// the root declaration that binds it is retained with the root.
    #[test]
    fn a_doc_grid_attribute_bound_on_the_root_keeps_its_binding() {
        let body = concat!(
            "<w:p><w:r><w:t>body</w:t></w:r></w:p>",
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>",
            "<w:docGrid w:type=\"lines\" w:linePitch=\"360\" x:kept=\"grid\"/>",
            "</w:sectPr>",
        );
        let source = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:x="urn:producer"><w:body>{body}</w:body></w:document>"#
        );
        let mut seed = Document::new();
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(seed.to_bytes().unwrap())).unwrap();
        package.set_part("/word/document.xml", source.into_bytes());
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();

        let mut document = Document::from_bytes(&bytes.into_inner()).unwrap();
        assert_eq!(
            document
                .sections()
                .next()
                .unwrap()
                .doc_grid()
                .expect("the grid is typed")
                .extra_attributes,
            vec![("x:kept".to_owned(), "grid".to_owned())]
        );
        let saved = String::from_utf8(
            OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap()))
                .unwrap()
                .get_part("/word/document.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        assert!(
            saved.contains("x:kept=\"grid\""),
            "the retained attribute is written: {saved}"
        );
        let root_start = saved.find("<w:document").expect("the root element opens");
        let root_end = root_start
            + saved[root_start..]
                .find('>')
                .expect("the root element closes");
        assert!(
            saved[root_start..root_end].contains("xmlns:x=\"urn:producer\""),
            "the ancestor binding that scopes it is still on the root: {saved}"
        );
    }

    /// Every `w:tcPr/w:textDirection` value projects to the rotation the
    /// rendering spec names, and upright stacking records its diagnostic.
    #[test]
    fn every_supported_cell_text_direction_renders_at_its_rotation() {
        let expected = [
            (CellTextDirection::LeftToRightTopToBottom, None, false),
            (
                CellTextDirection::TopToBottomRightToLeft,
                Some(90.0_f64),
                false,
            ),
            (
                CellTextDirection::BottomToTopLeftToRight,
                Some(-90.0),
                false,
            ),
            (
                CellTextDirection::LeftToRightTopToBottomVertical,
                Some(-90.0),
                true,
            ),
            (
                CellTextDirection::TopToBottomRightToLeftVertical,
                Some(90.0),
                true,
            ),
            (
                CellTextDirection::TopToBottomLeftToRightVertical,
                Some(-90.0),
                true,
            ),
        ];
        for (direction, rotation, stacks_upright) in expected {
            let mut document = Document::new();
            {
                let mut table = document.add_table(1, 1);
                let mut cell = table.cell(0, 0).expect("authored cell");
                cell.set_text_direction(Some(direction));
                cell.set_text("Vertical");
            }
            let result = document
                .layout_deterministic()
                .expect("deterministic vertical cell layout");
            let transforms = transforms_for(&result, "Vertical");
            assert_eq!(transforms.len(), 1, "the cell paints once: {direction:?}");
            match rotation {
                None => assert!(
                    transforms[0].is_identity(),
                    "{direction:?} keeps the untransformed path"
                ),
                Some(degrees) => {
                    let (sin, cos) = degrees.to_radians().sin_cos();
                    assert!(
                        (transforms[0].a - cos).abs() < 1e-9
                            && (transforms[0].b - sin).abs() < 1e-9,
                        "{direction:?} rotates {degrees} degrees: {:?}",
                        transforms[0]
                    );
                }
            }
            let diagnostics = result
                .layout
                .diagnostics
                .iter()
                .map(|diagnostic| diagnostic.message.clone())
                .collect::<Vec<_>>();
            assert_eq!(
                diagnostics.iter().any(|message| message
                    == "east Asian vertical text rendered as rotated vertical text"),
                stacks_upright,
                "{direction:?} records the upright-stacking fallback only when it asks for \
                 upright stacking: {diagnostics:?}"
            );
        }
    }

    /// A rotated cell's row height comes from the transposed box, so it is the
    /// text's own length rather than its stacked line height.
    #[test]
    fn a_vertical_cell_drives_the_row_height_from_its_transposed_box() {
        // The row height decides where the block after the table sits, which
        // is the only public statement of it that survives to the page.
        let row_height = |direction: CellTextDirection| {
            let mut document = Document::new();
            {
                let mut table = document.add_table(1, 2);
                let mut cell = table.cell(0, 0).expect("authored cell");
                cell.set_text_direction(Some(direction));
                cell.set_text("A much longer stretch of vertical cell text");
                table.cell(0, 1).expect("control cell").set_text("short");
            }
            document.add_paragraph("below");
            let result = document
                .layout_deterministic()
                .expect("deterministic vertical cell layout");
            let mut below = None;
            for page in &result.layout.pages {
                oxml_layout::walk(&page.elements, &mut |element, _| {
                    if let PositionedElement::Text(run) = element
                        && run.text.trim() == "below"
                    {
                        below = Some(run.origin.y);
                    }
                });
            }
            below.expect("the block after the table paints")
        };

        let horizontal = row_height(CellTextDirection::LeftToRightTopToBottom);
        let vertical = row_height(CellTextDirection::TopToBottomRightToLeft);
        assert!(
            vertical > horizontal + 1.0,
            "the rotated cell grows the row to its text length, {vertical} against {horizontal}"
        );

        // The height is the transposed box's measure, not merely larger.
        // Lengthening the rotated cell's single word moves the block after
        // the table down by exactly the width that word gained, because the
        // row height is the transposed measure and nothing else. Both numbers
        // are read from layout, so this states the rule rather than pinning a
        // recorded value.
        let vertical_row = |text: &str| {
            let mut document = Document::new();
            {
                let mut table = document.add_table(1, 1);
                let mut cell = table.cell(0, 0).expect("rotated cell");
                cell.set_text_direction(Some(CellTextDirection::TopToBottomRightToLeft));
                cell.set_text(text);
            }
            document.add_paragraph("below");
            let result = document
                .layout_deterministic()
                .expect("deterministic vertical cell layout");
            let mut below = None;
            let mut width = 0.0f64;
            for page in &result.layout.pages {
                oxml_layout::walk(&page.elements, &mut |element, _| {
                    if let PositionedElement::Text(run) = element {
                        if run.text.trim() == "below" {
                            below = Some(run.origin.y);
                        } else if text.contains(run.text.trim()) && !run.text.trim().is_empty() {
                            width += run.advances.iter().sum::<f64>();
                        }
                    }
                });
            }
            (below.expect("the block after the table paints"), width)
        };

        let (short_below, short_width) = vertical_row("Narrow");
        let (long_below, long_width) = vertical_row("NarrowNarrow");
        assert!(
            long_width > short_width,
            "the longer word is wider, {long_width} against {short_width}"
        );
        assert!(
            ((long_below - short_below) - (long_width - short_width)).abs() < 1e-6,
            "the row grew by exactly the transposed measure the word gained, \
             {} against {}",
            long_below - short_below,
            long_width - short_width
        );
    }

    /// Each grid type that snaps line advance does so, and `default` does not.
    #[test]
    fn a_gridded_section_snaps_line_advance_to_its_line_pitch() {
        let advance = |grid: Option<ST_DocGrid>| {
            let mut document = Document::new();
            {
                let mut paragraph = document.add_paragraph("first");
                paragraph.add_line_break();
                paragraph.add_run("second");
            }
            if let Some(grid_type) = grid {
                document
                    .section_mut(0)
                    .expect("final section")
                    .set_doc_grid(Some(CT_DocGrid {
                        grid_type: Some(grid_type),
                        line_pitch: Some(Twips(720)),
                        char_space: None,
                        extra_attributes: Vec::new(),
                    }));
            }
            let result = document
                .layout_deterministic()
                .expect("deterministic gridded layout");
            let mut origins = Vec::new();
            for page in &result.layout.pages {
                oxml_layout::walk(&page.elements, &mut |element, _| {
                    if let PositionedElement::Text(run) = element
                        && !run.text.trim().is_empty()
                    {
                        origins.push(run.origin.y);
                    }
                });
            }
            assert_eq!(origins.len(), 2, "both lines paint");
            origins[1] - origins[0]
        };

        let ungridded = advance(None);
        assert_eq!(
            advance(Some(ST_DocGrid::Default)),
            ungridded,
            "a default grid keeps the ungridded advance"
        );
        assert!(
            ungridded < 36.0,
            "the ungridded advance is inside one grid row, so the snap is visible"
        );
        // A 720 twip pitch is 36 points, which is more than one line of the
        // default face, so a snapping grid takes exactly one grid row.
        for grid_type in [
            ST_DocGrid::Lines,
            ST_DocGrid::LinesAndChars,
            ST_DocGrid::SnapToChars,
        ] {
            assert!(
                (advance(Some(grid_type)) - 36.0).abs() < 1e-9,
                "{grid_type:?} puts the next baseline on the grid pitch"
            );
        }
    }

    /// The section-level projection reads the property F-269 delivers and
    /// writes nothing back.
    #[test]
    fn section_text_direction_renders_over_the_property_f269_delivers() {
        let mut plain = Document::new();
        plain.add_paragraph("body text for the section");
        let plain_result = plain
            .layout_deterministic()
            .expect("deterministic horizontal layout");
        let plain_transforms = transforms_for(&plain_result, "section");
        assert_eq!(plain_transforms.len(), 1, "the horizontal body paints once");
        assert!(
            plain_transforms
                .iter()
                .all(|transform| transform.is_identity()),
            "a horizontal section keeps the untransformed path"
        );

        let mut document = Document::new();
        document.add_paragraph("body text for the section");
        document
            .section_mut(0)
            .expect("final section")
            .set_text_direction("tbRl");
        let result = document
            .layout_deterministic()
            .expect("deterministic vertical section layout");
        let transforms = transforms_for(&result, "section");
        assert_eq!(transforms.len(), 1, "the body paints once");
        assert!(
            (transforms[0].b - 1.0).abs() < 1e-9,
            "the body band rotates 90 degrees: {:?}",
            transforms[0]
        );

        // A vertical section fills one band, so a declared column layout is
        // dropped and the fact is recorded rather than painted wrong.
        let mut columns = Document::new();
        columns.add_paragraph("body text for the section");
        {
            let mut section = columns.section_mut(0).expect("final section");
            section.set_text_direction("tbRl");
            section.set_columns(2, Length::twips(360)).unwrap();
        }
        let columns_result = columns
            .layout_deterministic()
            .expect("deterministic vertical column layout");
        assert!(
            columns_result
                .layout
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message
                    == "vertical section text is laid out in one column track"),
            "the dropped column tracks are recorded: {:?}",
            columns_result.layout.diagnostics
        );

        // The projection is read-only. The saved property is exactly what was
        // authored, and reopening gives it back unchanged.
        let mut reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.sections().next().unwrap().text_direction(),
            Some("tbRl")
        );
        assert_eq!(
            reopened
                .section_mut(0)
                .expect("final section")
                .text_direction(),
            Some("tbRl")
        );
    }
}
