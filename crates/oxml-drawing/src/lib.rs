pub mod color;
pub mod effect;
pub mod fill;
pub mod geometry;
pub mod line;
pub mod namespace;
pub mod order;
mod preset_shape_data;
pub mod shape_props;
pub mod style_ref;
pub mod table;
pub mod text;
pub mod theme;
pub mod xfrm;

#[cfg(test)]
mod table_gate_tests {
    use crate::table::CT_Table;

    #[test]
    fn merged_table_round_trips_with_merge_origins_intact() {
        let xml = br#"<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblGrid><a:gridCol w="1000"/><a:gridCol w="2000"/><a:gridCol w="3000"/></a:tblGrid><a:tr h="4000"><a:tc rowSpan="2" gridSpan="2"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc><a:tc hMerge="1"><a:tcPr/></a:tc><a:tc><a:tcPr/></a:tc></a:tr><a:tr h="5000"><a:tc vMerge="1"><a:tcPr/></a:tc><a:tc hMerge="1" vMerge="1"><a:tcPr/></a:tc><a:tc><a:tcPr/></a:tc></a:tr></a:tbl>"#;
        let table = CT_Table::from_xml(xml).unwrap();
        assert_eq!(
            table
                .grid
                .columns
                .iter()
                .map(|width| width.0)
                .collect::<Vec<_>>(),
            vec![1000, 2000, 3000]
        );
        assert_eq!(table.rows[0].height.0, 4000);
        assert_eq!(table.rows[1].height.0, 5000);
        let origin = &table.rows[0].cells[0];
        assert_eq!((origin.row_span, origin.grid_span), (2, 2));
        assert!(!origin.horizontal_merge && !origin.vertical_merge);
        assert!(table.rows[0].cells[1].horizontal_merge);
        assert!(table.rows[1].cells[0].vertical_merge);
        assert!(table.rows[1].cells[1].horizontal_merge);
        assert!(table.rows[1].cells[1].vertical_merge);
        let written = table.to_xml().unwrap();
        assert_eq!(table, CT_Table::from_xml(&written).unwrap());
    }
}

#[cfg(test)]
mod tests {

    /// F-X024. The shared crates must not depend on a format crate.
    ///
    /// `oxml-drawing` used to be the one exception, hosting the Word theme
    /// adapter and therefore depending on `rdocx-oxml`. That single edge made
    /// the two publication trains mutually dependent, so neither could publish
    /// first once both carried breaking changes. The adapter now lives in
    /// `rdocx-oxml` and the rule has no exception, which this test keeps true.
    #[test]
    fn no_shared_crate_depends_on_a_format_crate() {
        let manifests: [(&str, &str); 9] = [
            ("oxml-core", include_str!("../../oxml-core/Cargo.toml")),
            ("oxml-opc", include_str!("../../oxml-opc/Cargo.toml")),
            ("oxml-media", include_str!("../../oxml-media/Cargo.toml")),
            ("oxml-layout", include_str!("../../oxml-layout/Cargo.toml")),
            ("oxml-drawing", include_str!("../Cargo.toml")),
            ("oxml-pdf", include_str!("../../oxml-pdf/Cargo.toml")),
            ("oxml-sml", include_str!("../../oxml-sml/Cargo.toml")),
            (
                "oxml-cli-support",
                include_str!("../../oxml-cli-support/Cargo.toml"),
            ),
            (
                "oxml-py-support",
                include_str!("../../oxml-py-support/Cargo.toml"),
            ),
        ];

        for (crate_name, manifest) in manifests {
            for line in manifest.lines() {
                let line = line.trim();
                if line.starts_with('#') || line.is_empty() {
                    continue;
                }
                // A dependency entry names the crate at the start of the line,
                // either as `rdocx-oxml.workspace = true` or as a table key.
                let names_format_crate = line.starts_with("rdocx")
                    || line.starts_with("rpptx")
                    || line.starts_with("\"rdocx")
                    || line.starts_with("\"rpptx");
                assert!(
                    !names_format_crate,
                    "{crate_name} depends on a format crate: {line}"
                );
            }
        }
    }
    use std::collections::HashSet;
    use std::path::PathBuf;
    use std::process::Command;

    use quick_xml::Reader;
    use quick_xml::events::Event;

    use super::geometry::CT_CustomGeometry2D;
    use super::namespace::{A_NS, A_PREFIX, PIC_NS, PIC_PREFIX};
    use super::preset_shape_data::preset_shape_definition;

    fn workspace_file(path: &str) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(path)
    }

    fn run_generator(arguments: &[&str]) -> std::process::Output {
        Command::new("python3")
            .arg(workspace_file("tools/gen-presets/generate.py"))
            .args(arguments)
            .current_dir(workspace_file("."))
            .output()
            .expect("run the preset shape generator")
    }

    #[test]
    fn generator_reproduces_checked_in_table() {
        let output = run_generator(&["--check"]);
        assert!(
            output.status.success(),
            "generator check failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }

    #[test]
    fn generated_table_covers_every_corpus_preset() {
        let corpus = std::env::var_os("RDOCX_PPTX_CORPUS_DIR")
            .map(PathBuf::from)
            .unwrap_or_else(|| workspace_file("corpus/pptx"));
        if !corpus.is_dir() {
            assert_ne!(
                std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").as_deref(),
                Some(std::ffi::OsStr::new("1")),
                "the pinned corpus is required but {} does not exist",
                corpus.display()
            );
            eprintln!("corpus gate skipped because {} is absent", corpus.display());
            return;
        }
        let corpus = corpus.to_str().expect("corpus path is UTF-8");
        let output = run_generator(&["--check", "--corpus", corpus]);
        assert!(
            output.status.success(),
            "corpus preset check failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }

    #[test]
    fn source_has_187_direct_definitions() {
        let source = std::fs::read(workspace_file(
            "tools/gen-presets/presetShapeDefinitions.xml",
        ))
        .expect("read the pinned official preset source");
        let mut reader = Reader::from_reader(source.as_slice());
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut direct_definitions = 0usize;
        let mut names = HashSet::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .expect("parse source XML")
            {
                Event::Start(element) => {
                    if depth == 1 {
                        direct_definitions += 1;
                        names.insert(element.name().as_ref().to_vec());
                    }
                    depth += 1;
                }
                Event::Empty(element) if depth == 1 => {
                    direct_definitions += 1;
                    names.insert(element.name().as_ref().to_vec());
                }
                Event::End(_) => depth -= 1,
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
        assert_eq!(direct_definitions, 187);
        assert_eq!(names.len(), 186);
        let source = String::from_utf8(source).expect("source XML is UTF-8");
        let duplicates: Vec<_> = source
            .split("<upDownArrow>")
            .skip(1)
            .map(|suffix| {
                suffix
                    .split("</upDownArrow>")
                    .next()
                    .expect("complete upDownArrow definition")
            })
            .collect();
        assert_eq!(duplicates.len(), 2);
        assert_eq!(duplicates[0].as_bytes(), duplicates[1].as_bytes());
    }

    #[test]
    fn generated_lookup_has_known_and_unknown_cases() {
        let rectangle = preset_shape_definition("rect").expect("known rectangle preset");
        assert!(CT_CustomGeometry2D::from_xml(rectangle).is_ok());
        assert!(preset_shape_definition("notAStandardPreset").is_none());
    }

    #[test]
    fn drawingml_namespace_uris_match_the_specification() {
        assert_eq!(
            A_NS,
            "http://schemas.openxmlformats.org/drawingml/2006/main"
        );
        assert_eq!(A_PREFIX, "a");
        assert_eq!(
            PIC_NS,
            "http://schemas.openxmlformats.org/drawingml/2006/picture"
        );
        assert_eq!(PIC_PREFIX, "pic");
    }

    #[test]
    fn oxml_drawing_is_an_explicit_publication_candidate() {
        let manifest = include_str!("../Cargo.toml");
        assert!(manifest.contains("version = \"0.11.0\""));
        assert!(manifest.contains("publish = true"));
    }
}
