//! Bundled fallback fonts for standalone operation.
//!
//! When the `bundled-fonts` feature is enabled (default), the following fonts
//! are embedded in the binary:
//!
//! - **Carlito** — metric-compatible with Calibri (Word's default font)
//! - **Caladea** — metric-compatible with Cambria
//! - **Liberation Sans** — metric-compatible with Arial
//! - **Liberation Serif** — metric-compatible with Times New Roman
//! - **Liberation Mono** — metric-compatible with Courier New
//!
//! All fonts are licensed under the SIL Open Font License.

/// Returns bundled font data: `(family_name, font_bytes)` pairs.
///
/// When `bundled-fonts` feature is enabled, returns Carlito, Caladea,
/// Liberation Sans, Liberation Serif, and Liberation Mono (each with
/// Regular, Bold, Italic, and BoldItalic variants).
///
/// When the feature is disabled, returns an empty vec (system fonts are used).
pub fn bundled_font_data() -> Vec<(&'static str, &'static [u8])> {
    #[cfg(feature = "bundled-fonts")]
    {
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
        ]
    }
    #[cfg(not(feature = "bundled-fonts"))]
    {
        vec![]
    }
}
