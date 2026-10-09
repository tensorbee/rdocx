//! Bundled fallback fonts for standalone operation.
//!
//! The following fonts are embedded in the binary:
//!
//! - **Carlito** — metric-compatible with Calibri (Word's default font),
//!   licensed under the SIL Open Font License 1.1
//! - **Caladea** — metric-compatible with Cambria, licensed under the Apache
//!   License 2.0
//! - **Liberation Sans** — metric-compatible with Arial, licensed under the
//!   SIL Open Font License 1.1
//! - **Liberation Serif** — metric-compatible with Times New Roman, licensed
//!   under the SIL Open Font License 1.1
//! - **Liberation Mono** — metric-compatible with Courier New, licensed under
//!   the SIL Open Font License 1.1
//! - **Noto Sans Arabic**, **Noto Sans Devanagari**, **Noto Sans Thai**,
//!   **Noto Sans SC**, **Noto Sans Hebrew**, **Noto Sans KR**, and **Noto Sans
//!   JP** — deterministic complex-script fallbacks licensed under the SIL Open
//!   Font License 1.1
//! - **Gelasio** — metric-compatible with Georgia, licensed under the SIL Open
//!   Font License 1.1, with Georgia's line metrics as `fonts/NOTICE-Gelasio`
//!   records
//! - **Selawik** — metric-compatible with Segoe UI, by Microsoft, licensed
//!   under the SIL Open Font License 1.1, Regular and Bold only
//! - **Noto Sans Symbols 2** — the Unicode equivalents of the Symbol and
//!   Wingdings bullets that no other bundled face draws, licensed under the
//!   SIL Open Font License 1.1
//!
//! The Noto Sans SC, Hebrew, KR, JP and Symbols 2 faces are deterministic
//! subsets, not the full families. Each ships only its approved repertoire,
//! which its `fonts/SUBSET-*.md` record names and reproduces.

/// Returns bundled font data: `(family_name, font_bytes)` pairs.
///
/// Returns Carlito, Caladea, Liberation Sans, Liberation Serif, Liberation
/// Mono and Gelasio, each with Regular, Bold, Italic, and BoldItalic variants,
/// Selawik Regular and Bold, and one face of each Noto family.
pub fn bundled_font_data() -> Vec<(&'static str, &'static [u8])> {
    vec![
        // Carlito — metric-compatible replacement for Calibri
        (
            "Carlito",
            include_bytes!("../fonts/Carlito-Regular.ttf").as_slice(),
        ),
        (
            "Carlito",
            include_bytes!("../fonts/Carlito-Bold.ttf").as_slice(),
        ),
        (
            "Carlito",
            include_bytes!("../fonts/Carlito-Italic.ttf").as_slice(),
        ),
        (
            "Carlito",
            include_bytes!("../fonts/Carlito-BoldItalic.ttf").as_slice(),
        ),
        // Caladea — metric-compatible replacement for Cambria
        (
            "Caladea",
            include_bytes!("../fonts/Caladea-Regular.ttf").as_slice(),
        ),
        (
            "Caladea",
            include_bytes!("../fonts/Caladea-Bold.ttf").as_slice(),
        ),
        (
            "Caladea",
            include_bytes!("../fonts/Caladea-Italic.ttf").as_slice(),
        ),
        (
            "Caladea",
            include_bytes!("../fonts/Caladea-BoldItalic.ttf").as_slice(),
        ),
        // Liberation Sans — metric-compatible replacement for Arial
        (
            "Liberation Sans",
            include_bytes!("../fonts/LiberationSans-Regular.ttf").as_slice(),
        ),
        (
            "Liberation Sans",
            include_bytes!("../fonts/LiberationSans-Bold.ttf").as_slice(),
        ),
        (
            "Liberation Sans",
            include_bytes!("../fonts/LiberationSans-Italic.ttf").as_slice(),
        ),
        (
            "Liberation Sans",
            include_bytes!("../fonts/LiberationSans-BoldItalic.ttf").as_slice(),
        ),
        // Liberation Serif — metric-compatible replacement for Times New Roman
        (
            "Liberation Serif",
            include_bytes!("../fonts/LiberationSerif-Regular.ttf").as_slice(),
        ),
        (
            "Liberation Serif",
            include_bytes!("../fonts/LiberationSerif-Bold.ttf").as_slice(),
        ),
        (
            "Liberation Serif",
            include_bytes!("../fonts/LiberationSerif-Italic.ttf").as_slice(),
        ),
        (
            "Liberation Serif",
            include_bytes!("../fonts/LiberationSerif-BoldItalic.ttf").as_slice(),
        ),
        // Liberation Mono — metric-compatible replacement for Courier New
        (
            "Liberation Mono",
            include_bytes!("../fonts/LiberationMono-Regular.ttf").as_slice(),
        ),
        (
            "Liberation Mono",
            include_bytes!("../fonts/LiberationMono-Bold.ttf").as_slice(),
        ),
        (
            "Liberation Mono",
            include_bytes!("../fonts/LiberationMono-Italic.ttf").as_slice(),
        ),
        (
            "Liberation Mono",
            include_bytes!("../fonts/LiberationMono-BoldItalic.ttf").as_slice(),
        ),
        (
            "Noto Sans Arabic",
            include_bytes!("../fonts/NotoSansArabic.ttf").as_slice(),
        ),
        (
            "Noto Sans Devanagari",
            include_bytes!("../fonts/NotoSansDevanagari.ttf").as_slice(),
        ),
        (
            "Noto Sans Thai",
            include_bytes!("../fonts/NotoSansThai.ttf").as_slice(),
        ),
        (
            "Noto Sans SC",
            include_bytes!("../fonts/NotoSansSC-FX058-subset.ttf").as_slice(),
        ),
        (
            "Noto Sans Hebrew",
            include_bytes!("../fonts/NotoSansHebrew-F266a-subset.ttf").as_slice(),
        ),
        (
            "Noto Sans KR",
            include_bytes!("../fonts/NotoSansKR-F266a-subset.ttf").as_slice(),
        ),
        (
            "Noto Sans JP",
            include_bytes!("../fonts/NotoSansJP-F266a-subset.ttf").as_slice(),
        ),
        // Gelasio — metric-compatible replacement for Georgia
        (
            "Gelasio",
            include_bytes!("../fonts/Gelasio-Regular.ttf").as_slice(),
        ),
        (
            "Gelasio",
            include_bytes!("../fonts/Gelasio-Bold.ttf").as_slice(),
        ),
        (
            "Gelasio",
            include_bytes!("../fonts/Gelasio-Italic.ttf").as_slice(),
        ),
        (
            "Gelasio",
            include_bytes!("../fonts/Gelasio-BoldItalic.ttf").as_slice(),
        ),
        // Selawik — metric-compatible replacement for Segoe UI
        (
            "Selawik",
            include_bytes!("../fonts/Selawik-Regular.ttf").as_slice(),
        ),
        (
            "Selawik",
            include_bytes!("../fonts/Selawik-Bold.ttf").as_slice(),
        ),
        (
            "Noto Sans Symbols 2",
            include_bytes!("../fonts/NotoSansSymbols2-bullets-subset.ttf").as_slice(),
        ),
    ]
}

#[cfg(test)]
mod tests {
    use super::bundled_font_data;
    use std::collections::BTreeSet;
    use std::path::Path;

    #[test]
    fn every_bundled_font_family_has_a_licence_file() {
        let family_licences = [
            ("Caladea", "LICENSE-Caladea"),
            ("Carlito", "LICENSE-Carlito"),
            ("Gelasio", "LICENSE-Gelasio"),
            ("Liberation Mono", "LICENSE-Liberation"),
            ("Liberation Sans", "LICENSE-Liberation"),
            ("Liberation Serif", "LICENSE-Liberation"),
            ("Noto Sans Arabic", "LICENSE-Noto"),
            ("Noto Sans Devanagari", "LICENSE-Noto"),
            ("Noto Sans Hebrew", "LICENSE-Noto"),
            ("Noto Sans JP", "LICENSE-Noto"),
            ("Noto Sans KR", "LICENSE-Noto"),
            ("Noto Sans SC", "LICENSE-Noto"),
            ("Noto Sans Symbols 2", "LICENSE-Noto"),
            ("Noto Sans Thai", "LICENSE-Noto"),
            ("Selawik", "LICENSE-Selawik"),
        ];
        let bundled_families = bundled_font_data()
            .into_iter()
            .map(|(family, _)| family)
            .collect::<BTreeSet<_>>();
        let licensed_families = family_licences
            .iter()
            .map(|(family, _)| *family)
            .collect::<BTreeSet<_>>();

        assert_eq!(bundled_families, licensed_families);

        let fonts_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("fonts");
        for (family, licence_file) in family_licences {
            assert!(
                fonts_dir.join(licence_file).is_file(),
                "{family} is missing {licence_file}"
            );
        }

        for notice in [
            "NOTICE-Caladea",
            "NOTICE-Gelasio",
            "NOTICE-Noto",
            "NOTICE-Selawik",
        ] {
            assert!(fonts_dir.join(notice).is_file(), "{notice} is missing");
        }
        for subset_record in [
            "SUBSET-NotoSansSC.md",
            "SUBSET-NotoSansHebrew.md",
            "SUBSET-NotoSansKR.md",
            "SUBSET-NotoSansJP.md",
            "SUBSET-NotoSansSymbols2.md",
        ] {
            assert!(
                fonts_dir.join(subset_record).is_file(),
                "{subset_record} is missing"
            );
        }
    }

    /// Printable ASCII advances, U+0020 to U+007E, in 2048ths of an em, of
    /// Georgia 5.00x (macOS `/System/Library/Fonts/Supplemental`) and Segoe UI
    /// 5.67 (Office's cloud fonts), measured from the installed faces.
    const GEORGIA_REGULAR: [u16; 95] = [
        494, 678, 843, 1317, 1249, 1674, 1455, 441, 768, 768, 967, 1317, 552, 766, 552, 960, 1257,
        880, 1144, 1130, 1157, 1082, 1159, 1029, 1221, 1159, 640, 640, 1317, 1317, 1317, 980, 1902,
        1374, 1339, 1315, 1534, 1338, 1227, 1485, 1669, 798, 1060, 1422, 1236, 1899, 1571, 1524,
        1249, 1524, 1437, 1149, 1267, 1549, 1365, 1998, 1455, 1260, 1232, 768, 960, 768, 1317,
        1317, 1024, 1032, 1147, 930, 1176, 990, 666, 1043, 1192, 600, 598, 1097, 586, 1804, 1210,
        1104, 1170, 1146, 839, 885, 707, 1178, 1017, 1510, 1034, 1008, 909, 881, 768, 881, 1317,
    ];
    const GEORGIA_BOLD: [u16; 95] = [
        520, 771, 1044, 1440, 1312, 1801, 1637, 551, 915, 915, 987, 1440, 672, 776, 672, 966, 1436,
        1003, 1283, 1279, 1330, 1227, 1327, 1135, 1385, 1327, 752, 752, 1440, 1440, 1440, 1123,
        1980, 1553, 1551, 1465, 1708, 1477, 1375, 1653, 1870, 913, 1219, 1673, 1404, 2096, 1719,
        1679, 1436, 1679, 1633, 1329, 1401, 1707, 1561, 2307, 1656, 1499, 1412, 915, 966, 915,
        1440, 1440, 1024, 1220, 1322, 1088, 1358, 1171, 805, 1181, 1392, 724, 709, 1294, 705, 2080,
        1413, 1302, 1347, 1328, 1065, 1050, 814, 1386, 1161, 1768, 1204, 1151, 1076, 1024, 794,
        1024, 1440,
    ];
    const GEORGIA_ITALIC: [u16; 95] = [
        494, 678, 843, 1317, 1249, 1674, 1455, 441, 768, 768, 967, 1317, 552, 766, 552, 960, 1257,
        880, 1144, 1130, 1157, 1082, 1159, 1017, 1221, 1159, 786, 786, 1317, 1317, 1317, 980, 1902,
        1374, 1339, 1315, 1534, 1338, 1227, 1485, 1669, 798, 1060, 1422, 1236, 1899, 1571, 1496,
        1249, 1496, 1437, 1149, 1267, 1549, 1365, 1998, 1455, 1260, 1232, 768, 960, 768, 1317,
        1317, 1024, 1173, 1134, 929, 1178, 966, 673, 1173, 1152, 609, 596, 1081, 584, 1801, 1208,
        1100, 1184, 1137, 945, 883, 711, 1178, 1102, 1684, 1026, 1146, 909, 881, 768, 881, 1317,
    ];
    const GEORGIA_BOLD_ITALIC: [u16; 95] = [
        520, 771, 1044, 1440, 1312, 1801, 1637, 551, 915, 915, 987, 1440, 672, 776, 672, 966, 1436,
        1003, 1283, 1279, 1330, 1227, 1327, 1160, 1385, 1327, 752, 752, 1440, 1440, 1440, 1123,
        1980, 1553, 1555, 1465, 1708, 1477, 1375, 1653, 1870, 923, 1219, 1673, 1404, 2116, 1699,
        1679, 1446, 1679, 1633, 1337, 1401, 1707, 1561, 2307, 1643, 1499, 1412, 915, 966, 915,
        1440, 1440, 1024, 1352, 1329, 1097, 1357, 1141, 780, 1330, 1383, 749, 747, 1313, 726, 2052,
        1413, 1302, 1357, 1331, 1093, 1059, 854, 1403, 1254, 1912, 1195, 1371, 1059, 1024, 794,
        1024, 1440,
    ];
    const SEGOE_UI_REGULAR: [u16; 95] = [
        561, 582, 803, 1210, 1104, 1676, 1639, 471, 618, 618, 854, 1401, 444, 819, 444, 798, 1104,
        1104, 1104, 1104, 1104, 1104, 1104, 1104, 1104, 1104, 444, 444, 1401, 1401, 1401, 918,
        1956, 1321, 1174, 1268, 1436, 1036, 1000, 1405, 1454, 545, 731, 1188, 964, 1839, 1532,
        1544, 1147, 1544, 1225, 1088, 1073, 1407, 1272, 1913, 1208, 1132, 1168, 618, 776, 618,
        1401, 850, 549, 1042, 1204, 946, 1206, 1071, 641, 1206, 1159, 496, 496, 1018, 496, 1764,
        1159, 1200, 1204, 1206, 712, 869, 694, 1159, 981, 1480, 940, 991, 926, 618, 490, 618, 1401,
    ];
    const SEGOE_UI_BOLD: [u16; 95] = [
        565, 670, 1010, 1213, 1178, 1776, 1740, 600, 756, 756, 932, 1448, 555, 828, 555, 908, 1178,
        1178, 1178, 1178, 1178, 1178, 1178, 1178, 1178, 1178, 555, 555, 1448, 1448, 1448, 897,
        1954, 1440, 1313, 1278, 1510, 1090, 1065, 1456, 1569, 649, 912, 1329, 1047, 1960, 1618,
        1553, 1258, 1553, 1337, 1148, 1200, 1481, 1366, 2058, 1342, 1243, 1243, 756, 893, 756,
        1448, 850, 643, 1102, 1270, 983, 1268, 1108, 785, 1268, 1233, 582, 582, 1145, 582, 1876,
        1239, 1252, 1270, 1268, 815, 901, 797, 1239, 1110, 1633, 1131, 1102, 981, 756, 668, 756,
        1448,
    ];

    /// The printable ASCII advances of one face, in 2048ths of an em.
    fn ascii_advances(bytes: &[u8]) -> Vec<u16> {
        let face = ttf_parser::Face::parse(bytes, 0).expect("font parses");
        assert_eq!(face.units_per_em(), 2048);
        (0x20u8..0x7F)
            .map(|code| {
                let glyph = face.glyph_index(char::from(code)).expect("ASCII glyph");
                face.glyph_hor_advance(glyph).expect("advance")
            })
            .collect()
    }

    /// Gelasio and Selawik take every printable ASCII advance of the face
    /// they replace to within one unit in 2048, and Gelasio its Windows line
    /// metrics, so a Georgia or Segoe UI line breaks and paginates as in
    /// Office. Where the real face is installed, it is measured too.
    #[test]
    fn gelasio_and_selawik_are_metric_compatible_with_georgia_and_segoe_ui() {
        let fonts = Path::new(env!("CARGO_MANIFEST_DIR")).join("fonts");
        let georgia = "/System/Library/Fonts/Supplemental";
        for (clone, recorded, installed) in [
            (
                "Gelasio-Regular.ttf",
                &GEORGIA_REGULAR,
                format!("{georgia}/Georgia.ttf"),
            ),
            (
                "Gelasio-Bold.ttf",
                &GEORGIA_BOLD,
                format!("{georgia}/Georgia Bold.ttf"),
            ),
            (
                "Gelasio-Italic.ttf",
                &GEORGIA_ITALIC,
                format!("{georgia}/Georgia Italic.ttf"),
            ),
            (
                "Gelasio-BoldItalic.ttf",
                &GEORGIA_BOLD_ITALIC,
                format!("{georgia}/Georgia Bold Italic.ttf"),
            ),
            ("Selawik-Regular.ttf", &SEGOE_UI_REGULAR, String::new()),
            ("Selawik-Bold.ttf", &SEGOE_UI_BOLD, String::new()),
        ] {
            let clone_advances = ascii_advances(&std::fs::read(fonts.join(clone)).unwrap());
            let mut references = vec![recorded.to_vec()];
            if let Ok(bytes) = std::fs::read(&installed) {
                references.push(ascii_advances(&bytes));
            }
            for reference in references {
                for (index, (expected, actual)) in reference.iter().zip(&clone_advances).enumerate()
                {
                    assert!(
                        expected.abs_diff(*actual) <= 1,
                        "{clone} {:?}: {actual} against {expected}",
                        char::from(0x20 + index as u8)
                    );
                }
            }
        }

        for face in [
            "Gelasio-Regular.ttf",
            "Gelasio-Bold.ttf",
            "Gelasio-Italic.ttf",
            "Gelasio-BoldItalic.ttf",
        ] {
            let bytes = std::fs::read(fonts.join(face)).unwrap();
            let face = ttf_parser::Face::parse(&bytes, 0).unwrap();
            let os2 = face.tables().os2.unwrap();
            assert!(!os2.use_typographic_metrics());
            assert_eq!(
                (face.ascender(), face.descender(), face.line_gap()),
                (1878, -449, 0)
            );
        }
    }

    #[test]
    fn deterministic_complex_script_fonts_cover_the_approved_fixture_repertoire() {
        let fixtures = [
            ("Noto Sans Arabic", "العربية"),
            ("Noto Sans Devanagari", "कि"),
            ("Noto Sans Thai", "ภาษาไทยยินดีต้อนรับ"),
            ("Noto Sans SC", "〈中〉、你好世界"),
            ("Noto Sans Hebrew", "שלום עולם"),
            ("Noto Sans KR", "안녕하세요 세계"),
            ("Noto Sans JP", "こんにちは、カタカナ世界"),
        ];
        for (family, text) in fixtures {
            let bytes = bundled_font_data()
                .into_iter()
                .find_map(|(candidate, bytes)| (candidate == family).then_some(bytes))
                .expect("approved family is bundled");
            let face = ttf_parser::Face::parse(bytes, 0).expect("bundled font parses");
            assert!(
                text.chars()
                    .filter(|character| !character.is_whitespace())
                    .all(|character| face.glyph_index(character).is_some()),
                "{family} misses its approved fixture repertoire"
            );
        }
    }
}
