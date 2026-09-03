//! High-level API for reading, writing, and converting DOCX documents.
//!
//! Provides a Python-docx-like interface for working with Word documents.
//!
//! # Examples
//!
//! ```no_run
//! use rdocx::Document;
//!
//! // Create a new document
//! let mut doc = Document::new();
//! doc.add_paragraph("Hello, World!");
//! doc.save("output.docx").unwrap();
//!
//! // Open an existing document
//! let doc = Document::open("existing.docx").unwrap();
//! for para in doc.paragraphs() {
//!     println!("{}", para.text());
//! }
//! ```

#![allow(clippy::too_many_arguments)]

mod comments;
mod comparison;
mod content_control;
mod document;
mod epub;
mod error;
mod field;
mod html;
mod math;
mod odt;
pub mod paragraph;
mod redaction;
mod revision;
mod rtf;
pub mod run;
pub mod style;
mod svg;
pub mod table;
mod template;

pub use comments::{BookmarkRef, CommentRef, RunPosition, RunRange};
pub use comparison::ComparisonDiagnostic;
pub use content_control::ContentControlRef;
pub use document::{
    AccessibilityIssue, BodyContentRef, BodyItemRef, Document, ImageInfo, IssueSeverity, LinkInfo,
    ListLevel, ListLevelSuffix, ListNumberFormat, NumberingFormat, NumberingLevel, OutlineNode,
    RenderOptions, UnsupportedXmlRef,
};
pub use epub::{EpubDiagnostic, EpubWriteResult};
pub use error::{Error, Result};
pub use field::{
    BarcodeCaseStyle, BarcodeField, BarcodeKind, BarcodePointOfSaleStyle, FieldDateTime,
    FieldEvaluation, FieldEvaluationContext, FieldOutcome, MailMergeControl, TcField,
    TocEntrySelection, TocField,
};
pub use html::{HtmlDiagnostic, HtmlReadResult};
pub use math::{
    MathConversionDiagnostic, MathConversionResult, equation_from_latex, equation_from_mathml,
    equation_to_latex, equation_to_mathml,
};
pub use odt::{OdtDiagnostic, OdtReadResult, OdtWriteResult};
pub use oxml_chart::{ChartData, ChartKind};
pub use oxml_core::Length;
pub use oxml_opc::PackageReadLimits;
#[cfg(feature = "digital-signatures")]
pub use oxml_opc::{
    CoveredRelationship, SignatureIssue, SignatureReport, SignerCertificateIdentity,
};
pub use oxml_pdf::{RasterFormat, RasterOptions, RasterOutput};
pub use paragraph::{
    Alignment, BorderStyle, HyperlinkItemRef, HyperlinkRef, Paragraph, ParagraphBorderRef,
    ParagraphItemRef, ParagraphRef, SectionBreak, TabAlignment, TabLeader,
};
pub use rdocx_layout::RevisionView;
pub use rdocx_oxml::math::{
    CT_OMath, CT_OMathPara, FractionType, LimitLocation, MathAccent, MathArgument, MathDelimiter,
    MathExpression, MathFraction, MathJustification, MathLimit, MathMatrix, MathMatrixProperties,
    MathMatrixRow, MathNary, MathParagraphProperties, MathPreSubSuperscript, MathProperties,
    MathRadical, MathRun, MathRunProperties, MathScript, MathScriptStyle, MathStyle,
    MathSubSuperscript, MatrixBaseJustification, OfficeMath,
};
pub use rdocx_oxml::settings::{
    CryptAlgorithmClass, CryptAlgorithmType, CryptProviderType, DocumentProtection, ProtectionMode,
};
pub use redaction::RedactionReport;
pub use revision::{RevisionKind, RevisionRef};
pub use rtf::{RtfDiagnostic, RtfReadResult, RtfWriteResult};
pub use run::{
    BreakKind, DrawingKind, DrawingRef, DrawingRelationshipKind, FieldDisplaySegmentRef, FieldKind,
    FieldRef, LegacyHorizontalRuleRef, Run, RunItemRef, RunProperties, RunRef, UnderlineStyle,
};
pub use style::{Style, StyleBuilder};
pub use svg::{SvgDiagnostic, SvgRenderResult};
pub use table::{
    Cell, CellItemRef, CellRef, Row, RowRef, Table, TableRef, VMerge, VerticalAlignment,
};

#[cfg(test)]
mod tests {
    use super::Length;

    #[test]
    fn html_import_collapses_whitespace_without_losing_inline_order() {
        let parsed = crate::Document::from_html(
            "direct <p> one  <strong>two</strong><br>three </p><pre> a  b\n c </pre>",
        )
        .expect("recoverable HTML");
        assert_eq!(
            parsed.document.text(),
            "direct\none two\nthree\n a  b\n c \n"
        );
    }

    #[test]
    fn html_import_restores_nested_inline_and_css_formatting() {
        let parsed = crate::Document::from_html(
            r#"<style>
                span { color: red }
                p.note > span#target { color: green; font-family: 'Arial'; font-size: 20px; font-weight: bold; font-style: italic; text-decoration: underline line-through; background-color: #abc }
            </style>
            <p class="note" style="text-align: center; margin-top: 6pt; margin-bottom: 8px; margin-left: 12pt; margin-right: 4pt; text-indent: -3pt">
                <span id="target" style="color: #123456">x</span><strong style="font-weight: normal">y</strong><sup>z</sup>
            </p>"#,
        )
        .expect("supported CSS");
        let paragraph = parsed.document.paragraph(0).unwrap();
        let run = paragraph.run(0).unwrap();
        assert_eq!(run.color(), Some("123456"));
        assert_eq!(run.bold_value(), Some(true));
        assert_eq!(run.italic_value(), Some(true));
        assert_eq!(run.strike_value(), Some(true));
        assert_eq!(run.underline_code_value(), Some(1));
        assert_eq!(run.font_name(), Some("Arial"));
        assert_eq!(run.size(), Some(15.0));
        assert_eq!(run.highlight(), Some("AABBCC".to_string()));
        assert_eq!(paragraph.run(1).unwrap().bold_value(), Some(false));
        assert_eq!(paragraph.alignment(), Some(crate::Alignment::Center));
        assert_eq!(paragraph.space_before(), Some(crate::Length::pt(6.0)));
        assert_eq!(paragraph.space_after(), Some(crate::Length::pt(6.0)));
        assert_eq!(paragraph.indent_left(), Some(crate::Length::pt(12.0)));
        assert_eq!(paragraph.indent_right(), Some(crate::Length::pt(4.0)));
        assert_eq!(paragraph.first_line_indent(), Some(crate::Length::pt(-3.0)));
    }

    #[test]
    fn public_length_is_the_shared_type() {
        let _: oxml_core::Length = Length::emu(1);
    }

    #[test]
    fn rtf_groups_restore_formatting_and_destination_state() {
        let parsed = crate::Document::from_rtf_bytes(
            br"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0\fcharset204 Cyrillic;}{\f1\fcharset0 Latin;}}\f1 plain {\b bold}{\*\unknown ignored} plain\par{\qc centered\par}tail {\f0\'c0} \'c0}",
        )
        .expect("valid RTF");
        assert_eq!(
            parsed.document.text(),
            "plain bold plain\ncentered\ntail А À\n"
        );
        let paragraph = parsed.document.paragraph(0).expect("paragraph");
        assert_eq!(paragraph.run(0).unwrap().bold_value(), None);
        assert_eq!(paragraph.run(1).unwrap().bold_value(), Some(true));
        assert_eq!(paragraph.run(2).unwrap().bold_value(), None);
        assert_eq!(
            parsed.document.paragraph(1).unwrap().alignment(),
            Some(crate::Alignment::Center)
        );
        assert_eq!(parsed.document.paragraph(2).unwrap().alignment(), None);
    }

    #[test]
    fn rtf_unicode_skips_the_declared_fallback_width() {
        let parsed = crate::Document::from_rtf_bytes(
            br"{\rtf1\ansi\ansicpg1252\uc1 A\u945? \uc2\u-10179?x\u-8704?x \'e9 \\ \{\}}",
        )
        .expect("valid Unicode RTF");
        assert_eq!(parsed.document.text(), "Aα 😀 é \\ {}\n");
    }

    #[test]
    fn rtf_malformed_input_fails_without_partial_document() {
        for input in [
            br"{\rtf1 unclosed".as_slice(),
            br"{\rtf1 \'zz}".as_slice(),
            br"{\rtf1\u}".as_slice(),
            br"{\rtf1\uc2\u65?}".as_slice(),
            br"{\rtf1{\pict\pngblip\bin67108865}}".as_slice(),
        ] {
            assert!(crate::Document::from_rtf_bytes(input).is_err());
        }

        let mut oversized_table = String::from("{\\rtf1\\trowd");
        for column in 0..257 {
            oversized_table.push_str(&format!("\\cellx{}", column + 1));
        }
        oversized_table.push_str("\\row}");
        assert!(crate::Document::from_rtf_bytes(oversized_table.as_bytes()).is_err());
    }

    #[test]
    fn rtf_writer_escapes_control_characters_and_signed_unicode() {
        let mut document = crate::Document::new();
        let mut paragraph = document.add_paragraph("");
        paragraph.add_run("slash \\ brace { } alpha α grin 😀");
        paragraph.add_tab();
        paragraph.add_line_break();
        paragraph.add_run("tail");

        let written = document.to_rtf_bytes().expect("RTF serializes");
        assert!(written.diagnostics.is_empty());
        let rtf = String::from_utf8(written.bytes).expect("RTF is ASCII");

        assert!(rtf.contains("slash \\\\ brace \\{ \\} alpha \\u945? grin \\u-10179?\\u-8704?"));
        assert!(rtf.contains("\\tab"));
        assert!(rtf.contains("\\line"));
    }

    #[test]
    fn save_rtf_writes_serialized_bytes_and_returns_diagnostics() {
        let path = std::env::temp_dir().join(format!(
            "rdocx-save-rtf-{}-{}.rtf",
            std::process::id(),
            "writer"
        ));
        let mut document = crate::Document::new();
        document.add_paragraph("saved");
        std::fs::write(&path, b"old RTF bytes").expect("old RTF file writes");

        let diagnostics = document.save_rtf(&path).expect("RTF file writes");
        let bytes = std::fs::read(&path).expect("RTF file exists");
        let _ = std::fs::remove_file(&path);

        assert!(diagnostics.is_empty());
        let rtf = String::from_utf8(bytes).unwrap();
        assert!(rtf.contains("saved"));
        assert!(!rtf.contains("old RTF bytes"));
    }

    #[test]
    fn save_rtf_preserves_existing_destination_when_staging_fails() {
        let root =
            std::env::temp_dir().join(format!("rdocx-save-rtf-atomic-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&root);
        std::fs::create_dir(&root).unwrap();
        let path = root.join("existing.rtf");
        std::fs::write(&path, b"existing destination").unwrap();
        for attempt in 0..128_u8 {
            let temporary = root.join(format!(
                ".existing.rtf.rdocx-{}-{attempt}.tmp",
                std::process::id()
            ));
            std::fs::write(temporary, b"occupied").unwrap();
        }

        let mut document = crate::Document::new();
        document.add_paragraph("replacement");
        let error = document.save_rtf(&path).unwrap_err();

        assert!(
            error
                .to_string()
                .contains("could not allocate RTF-save staging file")
        );
        assert_eq!(std::fs::read(&path).unwrap(), b"existing destination");
        std::fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn rtf_default_font_and_legacy_character_sets_are_preserved() {
        let parsed = crate::Document::from_rtf_bytes(
            br"{\rtf1\ansi\deff0{\fonttbl{\f0\fcharset0 Calibri;}}\plain text}",
        )
        .unwrap();
        assert_eq!(
            parsed
                .document
                .paragraph(0)
                .unwrap()
                .run(0)
                .unwrap()
                .font_name(),
            Some("Calibri")
        );

        for (input, expected) in [
            (br"{\rtf1\mac \'8e}".as_slice(), "é\n"),
            (br"{\rtf1\pc \'82}".as_slice(), "é\n"),
            (br"{\rtf1\pca \'9b}".as_slice(), "ø\n"),
            (
                br"{\rtf1\ansi{\fonttbl{\f0\fcharset254 Pc;}{\f1\fcharset255 Oem;}}\f0\'82 \f1\'9b}".as_slice(),
                "é ø\n",
            ),
        ] {
            assert_eq!(crate::Document::from_rtf_bytes(input).unwrap().document.text(), expected);
        }
    }

    #[test]
    fn rtf_unicode_alternatives_require_the_root_marker_and_reject_bad_numbers() {
        let parsed =
            crate::Document::from_rtf_bytes(br"{\rtf1\ansi{\upr{fallback}{\*\ud Unicode \u945?}}}")
                .unwrap();
        assert_eq!(parsed.document.text(), "Unicode α\n");
        assert!(crate::Document::from_rtf_bytes(br"{\ansi body{\rtf1}}").is_err());
        assert!(crate::Document::from_rtf_bytes(br"{\rtf1\ansi\b-foo}").is_err());

        for malformed in [
            br"{\rtf1{\upr{ansi}}}".as_slice(),
            br"{\rtf1{\upr{ansi}{Unicode}}}".as_slice(),
            br"{\rtf1{\upr{ansi}{\ud Unicode}}}".as_slice(),
            br"{\rtf1{\upr{ansi}{\* text\ud Unicode}}}".as_slice(),
            br"{\rtf1{\*\ud Unicode}}".as_slice(),
        ] {
            assert!(crate::Document::from_rtf_bytes(malformed).is_err());
        }
    }

    #[test]
    fn rtf_destinations_and_star_markers_must_start_their_group() {
        for malformed in [
            br"{\rtf1 body\fonttbl trailing}".as_slice(),
            br"{\rtf1 body\pict 00}".as_slice(),
            br"{\rtf1{\* text\unknown ignored}}".as_slice(),
        ] {
            assert!(crate::Document::from_rtf_bytes(malformed).is_err());
        }
    }

    #[test]
    fn rtf_font_charsets_reset_and_unsupported_johab_is_rejected() {
        let parsed = crate::Document::from_rtf_bytes(
            br"{\rtf1\ansi\ansicpg1251{\fonttbl\f0\fcharset1 Default;\f1\fcharset2 Symbol;}\f0\'c0 \f1a}",
        )
        .unwrap();
        assert_eq!(parsed.document.text(), "А a\n");

        let parsed = crate::Document::from_rtf_bytes(
            br"{\rtf1\ansi\ansicpg1252{\fonttbl\f0\fcharset204 Cyrillic;\f1 Plain;}\f0\'c0 \f1\'e9}",
        )
        .unwrap();
        assert_eq!(parsed.document.text(), "А é\n");

        assert!(
            crate::Document::from_rtf_bytes(
                br"{\rtf1\ansi{\fonttbl{\f0\fcharset130 Johab;}}\f0 text}"
            )
            .is_err()
        );
    }

    #[test]
    fn rtf_numeric_picture_list_and_reference_controls_are_strict() {
        const PNG: &[u8] = br"89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d49444154789c6360f8cff00000040101089d1de10000000049454e44ae426082";
        let mut missing_goal = br"{\rtf1{\pict\pngblip\picwgoal ".to_vec();
        missing_goal.extend_from_slice(PNG);
        missing_goal.extend_from_slice(b"}}");
        assert!(crate::Document::from_rtf_bytes(&missing_goal).is_err());
        assert!(crate::Document::from_rtf_bytes(br"{\rtf1{\pict\wmetafile data}}").is_err());
        assert!(crate::Document::from_rtf_bytes(br"{\rtf1\ilvl9 item}").is_err());
        assert!(crate::Document::from_rtf_bytes(br"{\rtf1\f999 missing}").is_err());
        assert!(crate::Document::from_rtf_bytes(br"{\rtf1\cf999 missing}").is_err());
        assert!(crate::Document::from_rtf_bytes(br"{\rtf1\highlight999 missing}").is_err());
    }

    #[test]
    fn rtf_control_symbols_cannot_interrupt_surrogate_pairs() {
        assert!(crate::Document::from_rtf_bytes(br"{\rtf1\uc1\u-10179?\~\u-8704?}").is_err());
        assert!(crate::Document::from_rtf_bytes(br"{\rtf1\uc1\u-10179?\b\u-8704?}").is_err());
    }

    #[test]
    fn rtf_word_special_characters_are_preserved_as_text() {
        let parsed = crate::Document::from_rtf_bytes(
            br"{\rtf1 one\emdash two\endash three\bullet four\lquote five\rquote six\ldblquote seven\rdblquote eight\emspace nine\enspace ten\qmspace eleven}",
        )
        .unwrap();
        assert_eq!(
            parsed.document.text(),
            "one—two–three•four‘five’six“seven”eight\u{2003}nine\u{2002}ten\u{2005}eleven\n"
        );
    }

    #[test]
    fn rtf_table_rows_require_matching_cells_and_boundaries() {
        for malformed in [
            br"{\rtf1\trowd\cellx1440 one\cell two\cell\row}".as_slice(),
            br"{\rtf1\trowd\cellx1440\cellx2880 one\cell\row}".as_slice(),
        ] {
            assert!(crate::Document::from_rtf_bytes(malformed).is_err());
        }
    }

    #[test]
    fn rtf_font_charset_controls_are_required_and_supported() {
        assert!(
            crate::Document::from_rtf_bytes(br"{\rtf1{\fonttbl{\f0\fcharset Missing;}}\f0 text}")
                .is_err()
        );
        assert!(
            crate::Document::from_rtf_bytes(
                br"{\rtf1{\fonttbl{\f0\fcharset999 Unknown;}}\f0 text}"
            )
            .is_err()
        );
    }

    #[test]
    fn rtf_diagnoses_each_dropped_list_property() {
        let parsed = crate::Document::from_rtf_bytes(
            br"{\rtf1{\*\listtable{\list{\listlevel\leveljc1\levelfollow2\levelspace120\levelindent240}\listid10}}}",
        )
        .unwrap();
        for property in ["leveljc", "levelfollow", "levelspace", "levelindent"] {
            assert!(parsed.diagnostics.iter().any(|diagnostic| {
                diagnostic.message == format!("RTF list property \\{property} was dropped")
            }));
        }
    }

    #[test]
    fn rtf_unicode_list_markers_remain_unicode() {
        let parsed =
            crate::Document::from_rtf_bytes(br"{\rtf1\ansi{\listtext\u8226?\tab}Item}").unwrap();
        let numbering = parsed.document.paragraph(0).unwrap().numbering().unwrap();
        assert_eq!(parsed.document.numbering_is_bullet(numbering.0), Some(true));
    }

    #[test]
    fn rtf_character_set_declarations_are_header_only() {
        for malformed in [
            br"{\rtf1 text\ansicpg1251 more}".as_slice(),
            br"{\rtf1\trowd\cellx1440\cell\row\mac text}".as_slice(),
            br"{\rtf1{\ansi}text}".as_slice(),
        ] {
            assert!(crate::Document::from_rtf_bytes(malformed).is_err());
        }
    }

    #[test]
    fn rtf_parser_output_collections_are_bounded() {
        let mut diagnostics = String::from("{\\rtf1 ");
        for _ in 0..10_001 {
            diagnostics.push_str("\\unknown ");
        }
        diagnostics.push('}');
        assert!(crate::Document::from_rtf_bytes(diagnostics.as_bytes()).is_err());

        let mut blocks = String::from("{\\rtf1 ");
        for _ in 0..100_001 {
            blocks.push_str("\\par ");
        }
        blocks.push('}');
        assert!(crate::Document::from_rtf_bytes(blocks.as_bytes()).is_err());

        let mut runs = String::from("{\\rtf1 ");
        for index in 0..100_001 {
            runs.push_str(if index % 2 == 0 { "\\b x" } else { "\\b0 x" });
        }
        runs.push('}');
        assert!(crate::Document::from_rtf_bytes(runs.as_bytes()).is_err());

        let mut cells = String::from("{\\rtf1 ");
        for _ in 0..196 {
            cells.push_str("\\trowd ");
            for column in 1..=256 {
                cells.push_str(&format!("\\cellx{column} "));
            }
            for _ in 0..256 {
                cells.push_str("\\cell ");
            }
            cells.push_str("\\row ");
        }
        cells.push('}');
        let error = match crate::Document::from_rtf_bytes(cells.as_bytes()) {
            Err(error) => error,
            Ok(_) => panic!("table cell amplification was accepted"),
        };
        assert!(error.to_string().contains("table cell limit"));

        let mut retained = br"{\rtf1\mac ".to_vec();
        retained.resize(retained.len() + (64 * 1024 * 1024 / 3) + 1, 0xa0);
        retained.push(b'}');
        let error = match crate::Document::from_rtf_bytes(&retained) {
            Err(error) => error,
            Ok(_) => panic!("retained output amplification was accepted"),
        };
        assert!(error.to_string().contains("retained output"));
    }

    #[test]
    fn open_rtf_rejects_an_oversized_file_before_reading_it() {
        let path =
            std::env::temp_dir().join(format!("rdocx-f176-oversized-{}.rtf", std::process::id()));
        let file = std::fs::File::create(&path).unwrap();
        file.set_len(64 * 1024 * 1024 + 1).unwrap();
        drop(file);
        let result = crate::Document::open_rtf(&path);
        std::fs::remove_file(path).unwrap();
        assert!(matches!(result, Err(crate::Error::Rtf { offset: 0, .. })));
    }

    #[cfg(feature = "digital-signatures")]
    #[test]
    fn native_document_forwards_signature_verification() {
        let document = crate::Document::new();
        assert!(document.verify_signatures().unwrap().is_empty());
    }

    #[cfg(all(feature = "digital-signatures", not(target_arch = "wasm32")))]
    #[test]
    fn native_document_exposes_atomic_signature_creation() {
        let sign: fn(&mut crate::Document, &[u8], &[u8]) -> crate::Result<crate::SignatureReport> =
            crate::Document::sign;
        let mut document = crate::Document::new();
        document.add_paragraph("atomic native signature creation");
        let before = document.to_bytes().unwrap();
        assert!(sign(&mut document, b"not-pkcs8", b"not-x509").is_err());
        assert_eq!(document.to_bytes().unwrap(), before);
        assert!(document.verify_signatures().unwrap().is_empty());
    }
}
