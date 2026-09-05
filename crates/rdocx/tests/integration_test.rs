//! Integration tests for rdocx — end-to-end document creation and round-trip.

use std::collections::HashMap;

use oxml_opc::OpcPackage;
use oxml_opc::relationship::rel_types;
use rdocx::paragraph::Alignment;
use rdocx::table::VerticalAlignment;
use rdocx::{
    BodyItemRef, BorderStyle, Length, ListLevel, MhtmlDiagnostic, ParagraphRef, RunPosition,
    RunRange, SectionBreak, StyleBuilder, TabAlignment, TabLeader, UnderlineStyle,
};
use rdocx::{Document, PackageReadLimits, RevisionKind, WordPackageClass};

const ODT_ORACLE_VERSION: &str = "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb";
const MHTML_ORACLE_VERSION: &str = "Microsoft Word 16.104 build 16.104.25121423";
const MHTML_ORACLE_HTML: &str = "<h1>Oracle title</h1><p><strong>bold</strong> <a href='https://example.test/'>link</a><img src='https://example.test/pixel.png' width='2' height='3'></p><ol><li>one</li><li>two</li></ol><table><tr><td>cell</td></tr></table>";

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

mod flat_opc_package_class_tests {
    use super::*;
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

    fn add_package_signature_graph(package: &mut OpcPackage) {
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

    fn has_package_signature_invalidation_marker(package: &OpcPackage) -> bool {
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
                package.content_types.overrides["/word/document.xml"],
                content_type(class)
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
fn settings_relationship_target_is_resolved_instead_of_assumed() {
    let settings = br#"<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:documentProtection w:edit="comments" w:enforcement="true" w:hash="custom-hash" w:salt="custom-salt"/></w:settings>"#;
    let mut seed = Document::new();
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
fn rtf_writer_diagnoses_unsupported_numbering_without_coercion() {
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
        !rtf.contains("\\ls1\\ilvl0{\\plain\\f0 unsupported format}"),
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
        [
            (
                "body[0]/ppr/numPr/ilvl".to_owned(),
                "numbering level above 8 was dropped during RTF export".to_owned(),
            ),
            (
                "numbering[numId=1]/level[0]/numFmt".to_owned(),
                "unsupported numbering format was dropped during RTF export".to_owned(),
            ),
        ]
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
    );

    doc.add_style(StyleBuilder::character("Emphasis", "Emphasis Style"));

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
    doc.add_style(StyleBuilder::paragraph("BlockQuote", "Block Quote").based_on("Normal"));

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
    package.content_types.overrides.remove("/docProps/core.xml");
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
fn comments_part_uses_its_existing_relationship_target() {
    let mut source = Document::new();
    source.add_paragraph("body");
    let bytes = source.to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let comments_xml = br#"<x:comments xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:comment x:id="9"><x:p><x:r><x:t>custom target</x:t></x:r></x:p></x:comment></x:comments>"#;
    package.set_part("/custom/comments-data.xml", comments_xml.to_vec());
    package.content_types.add_override(
        "/custom/comments-data.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    );
    package
        .get_or_create_part_rels("/word/document.xml")
        .add(rel_types::COMMENTS, "../custom/comments-data.xml");

    let mut input = std::io::Cursor::new(Vec::new());
    package.write_to(&mut input).unwrap();
    let mut document = Document::from_bytes(input.get_ref()).unwrap();
    let saved = document.to_bytes().unwrap();
    let saved_package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();

    assert!(saved_package.get_part("/word/comments.xml").is_none());
    let output = String::from_utf8(
        saved_package
            .get_part("/custom/comments-data.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(output.contains("<w:comments"));
    assert!(output.contains("custom target"));
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
        Some("application/vnd.ms-word.commentsExtended+xml")
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
    const COMMENTS_EXTENDED_CONTENT_TYPE: &str = "application/vnd.ms-word.commentsExtended+xml";

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
            package.content_types.overrides.get("/custom/producer.bin"),
            Some(&"application/x-producer-private".to_owned())
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
            assert!(document.legacy_form_fields().is_err(), "{case}");
        }
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
            "external",
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
            let id = relationships.add(rel_types::GLOSSARY_DOCUMENT, target);
            if case == "duplicate" {
                relationships.add(rel_types::GLOSSARY_DOCUMENT, "glossary/other.xml");
            }
            if case == "external" {
                relationships
                    .items
                    .iter_mut()
                    .find(|relationship| relationship.id == id)
                    .unwrap()
                    .target_mode = Some("External".into());
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
        let mut document = Document::from_bytes(&package_bytes(package)).unwrap();
        let before = document.to_bytes().unwrap();
        assert!(document.legacy_form_fields().is_err());
        assert!(
            document
                .set_legacy_form_field_value(
                    "/word/header-id-one.xml",
                    0,
                    LegacyFormFieldValue::Text("changed".to_owned()),
                )
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), before);
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
