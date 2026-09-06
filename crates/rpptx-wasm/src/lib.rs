//! WebAssembly bindings for rpptx.

use wasm_bindgen::prelude::*;

/// A PowerPoint-compatible presentation that can be opened, edited, and saved.
#[wasm_bindgen]
pub struct WasmPresentation {
    inner: rpptx::Presentation,
}

#[wasm_bindgen]
impl WasmPresentation {
    /// Creates a new presentation from the bundled default template.
    #[wasm_bindgen(constructor)]
    pub fn new() -> Result<WasmPresentation, JsValue> {
        let inner = rpptx::Presentation::new().map_err(to_js_error)?;
        Ok(Self { inner })
    }

    /// Opens a presentation from PPTX bytes.
    #[wasm_bindgen(js_name = "fromBytes")]
    pub fn from_bytes(data: &[u8]) -> Result<WasmPresentation, JsValue> {
        let inner = rpptx::Presentation::from_bytes(data).map_err(to_js_error)?;
        Ok(Self { inner })
    }

    /// Serializes the complete presentation package.
    #[wasm_bindgen(js_name = "toBytes")]
    pub fn to_bytes(&self) -> Result<Vec<u8>, JsValue> {
        self.inner.to_bytes().map_err(to_js_error)
    }

    /// Returns the number of slides.
    #[wasm_bindgen(js_name = "slideCount")]
    pub fn slide_count(&self) -> u32 {
        self.inner.len() as u32
    }

    /// Adds a slide from the zero-based layout index.
    #[wasm_bindgen(js_name = "addSlide")]
    pub fn add_slide(&mut self, layout_index: u32) -> Result<(), JsValue> {
        self.inner
            .add_slide(layout_index as usize)
            .map(|_| ())
            .map_err(to_js_error)
    }

    /// Renders the presentation to a deterministic PDF.
    #[cfg(feature = "render")]
    #[wasm_bindgen(js_name = "toPdf")]
    pub fn to_pdf(&self) -> Result<Vec<u8>, JsValue> {
        self.inner.to_pdf_deterministic().map_err(to_js_error)
    }
}

fn to_js_error(error: rpptx::Error) -> JsValue {
    JsValue::from_str(&error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use oxml_opc::OpcPackage;
    use std::collections::BTreeSet;
    use std::io::Cursor;
    use std::path::{Path, PathBuf};
    use std::process::Command;

    fn nontrivial_deck() -> Vec<u8> {
        let mut presentation = rpptx::Presentation::new().expect("create presentation");
        presentation.add_slide(0).expect("add slide");
        presentation.to_bytes().expect("serialize presentation")
    }

    fn assert_complete_package_round_trip(source: &[u8], output: &[u8]) {
        let source_package = OpcPackage::from_reader(Cursor::new(source)).expect("open source");
        let output_package = OpcPackage::from_reader(Cursor::new(output)).expect("open output");
        assert_eq!(
            source_package.parts.keys().collect::<BTreeSet<_>>(),
            output_package.parts.keys().collect::<BTreeSet<_>>(),
            "the complete part inventory must survive"
        );
        assert_eq!(
            source_package.content_types.defaults,
            output_package.content_types.defaults
        );
        assert_eq!(
            source_package.content_types.overrides,
            output_package.content_types.overrides
        );
        assert_eq!(
            normalized_relationships(&source_package.package_rels),
            normalized_relationships(&output_package.package_rels)
        );
        assert_eq!(
            source_package.part_rels.keys().collect::<BTreeSet<_>>(),
            output_package.part_rels.keys().collect::<BTreeSet<_>>()
        );
        for (part_name, relationships) in &source_package.part_rels {
            assert_eq!(
                normalized_relationships(relationships),
                normalized_relationships(&output_package.part_rels[part_name]),
                "relationships for {part_name} must survive"
            );
        }
        let reopened = rpptx::Presentation::from_bytes(output).expect("reopen output");
        assert_eq!(reopened.len(), 1);
    }

    fn normalized_relationships(
        relationships: &oxml_opc::Relationships,
    ) -> BTreeSet<(&str, &str, &str, Option<&str>)> {
        relationships
            .items
            .iter()
            .map(|relationship| {
                (
                    relationship.id.as_str(),
                    relationship.rel_type.as_str(),
                    relationship.target.as_str(),
                    relationship.target_mode.as_deref(),
                )
            })
            .collect()
    }

    const WASM_PACK_ARGS: &[&str] = &[
        "build",
        "<crate>",
        "--target",
        "nodejs",
        "--release",
        "--no-opt",
        "--out-dir",
        "<bindgen-output>",
    ];
    const WASM_OPT_ARGS: &[&str] = &[
        "-Oz",
        "--enable-bulk-memory",
        "<bindgen-wasm>",
        "-o",
        "<optimized-wasm>",
    ];
    const GZIP_ARGS: &[&str] = &["-n", "-9", "-c", "<optimized-wasm>"];

    struct SizePipelineEvidence {
        wasm_pack_version: String,
        wasm_opt_version: String,
        wasm_pack_args: Vec<String>,
        wasm_opt_args: Vec<String>,
        gzip_args: Vec<String>,
        optimized_wasm: Vec<u8>,
        gzip: Vec<u8>,
    }

    struct ScratchDir(PathBuf);

    impl ScratchDir {
        fn new() -> Self {
            let path = std::env::temp_dir().join(format!(
                "rpptx-wasm-size-gate-{}-{}",
                std::process::id(),
                std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .expect("system clock should follow Unix epoch")
                    .as_nanos()
            ));
            std::fs::create_dir(&path).expect("create size-gate scratch directory");
            Self(path)
        }
    }

    impl Drop for ScratchDir {
        fn drop(&mut self) {
            let _ = std::fs::remove_dir_all(&self.0);
        }
    }

    fn exact_size_pipeline() -> SizePipelineEvidence {
        let wasm_pack = std::env::var("RPPTX_WASM_PACK").expect("set RPPTX_WASM_PACK");
        let wasm_opt = std::env::var("RPPTX_WASM_OPT").expect("set RPPTX_WASM_OPT");
        let gzip = std::env::var("RPPTX_WASM_GZIP").unwrap_or_else(|_| "gzip".to_owned());
        let scratch = ScratchDir::new();
        let bindgen_output = scratch.0.join("bindgen");
        let optimized_wasm = scratch.0.join("rpptx_wasm_final.wasm");
        let crate_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
        let bindgen_wasm = bindgen_output.join("rpptx_wasm_bg.wasm");
        let wasm_pack_args = vec![
            "build".to_owned(),
            crate_dir.display().to_string(),
            "--target".to_owned(),
            "nodejs".to_owned(),
            "--release".to_owned(),
            "--no-opt".to_owned(),
            "--out-dir".to_owned(),
            bindgen_output.display().to_string(),
        ];
        let wasm_opt_args = vec![
            "-Oz".to_owned(),
            "--enable-bulk-memory".to_owned(),
            bindgen_wasm.display().to_string(),
            "-o".to_owned(),
            optimized_wasm.display().to_string(),
        ];
        let gzip_args = vec![
            "-n".to_owned(),
            "-9".to_owned(),
            "-c".to_owned(),
            optimized_wasm.display().to_string(),
        ];

        let wasm_pack_version = command_text(Command::new(&wasm_pack).arg("--version"));
        let wasm_opt_version = command_text(Command::new(&wasm_opt).arg("--version"));
        let status = Command::new(&wasm_pack)
            .args(&wasm_pack_args)
            .status()
            .expect("run wasm-pack default-profile build");
        assert!(status.success(), "wasm-pack default-profile build failed");
        let status = Command::new(&wasm_opt)
            .args(&wasm_opt_args)
            .status()
            .expect("run wasm-opt");
        assert!(status.success(), "wasm-opt failed");
        let gzip_output = Command::new(&gzip)
            .args(&gzip_args)
            .output()
            .expect("run gzip");
        assert!(gzip_output.status.success(), "gzip failed");

        SizePipelineEvidence {
            wasm_pack_version,
            wasm_opt_version,
            wasm_pack_args: normalize_args(
                wasm_pack_args,
                &[(1, "<crate>"), (7, "<bindgen-output>")],
            ),
            wasm_opt_args: normalize_args(
                wasm_opt_args,
                &[(2, "<bindgen-wasm>"), (4, "<optimized-wasm>")],
            ),
            gzip_args: normalize_args(gzip_args, &[(3, "<optimized-wasm>")]),
            optimized_wasm: std::fs::read(optimized_wasm).expect("read optimized WASM"),
            gzip: gzip_output.stdout,
        }
    }

    fn command_text(command: &mut Command) -> String {
        let output = command.output().expect("run version command");
        assert!(output.status.success(), "version command failed");
        String::from_utf8(output.stdout)
            .expect("version output is UTF-8")
            .trim()
            .to_owned()
    }

    fn normalize_args(mut args: Vec<String>, paths: &[(usize, &str)]) -> Vec<String> {
        for (index, placeholder) in paths {
            args[*index] = (*placeholder).to_owned();
        }
        args
    }

    fn validate_size_pipeline(
        evidence: &SizePipelineEvidence,
        gzip_program: &str,
    ) -> Result<usize, String> {
        if evidence.wasm_pack_version != "wasm-pack 0.15.0" {
            return Err("wasm-pack must be exactly 0.15.0".to_owned());
        }
        if evidence.wasm_opt_version != "wasm-opt version 125" {
            return Err("wasm-opt must be exactly version 125".to_owned());
        }
        let exact_args = |actual: &[String], expected: &[&str], tool: &str| {
            (actual == expected)
                .then_some(())
                .ok_or_else(|| format!("{tool} did not use the reviewed pipeline arguments"))
        };
        exact_args(&evidence.wasm_pack_args, WASM_PACK_ARGS, "wasm-pack")?;
        exact_args(&evidence.wasm_opt_args, WASM_OPT_ARGS, "wasm-opt")?;
        exact_args(&evidence.gzip_args, GZIP_ARGS, "gzip")?;
        if !evidence.optimized_wasm.starts_with(b"\0asm") {
            return Err("optimized output is not WebAssembly".to_owned());
        }
        if evidence.gzip.get(..10) != Some(&[0x1f, 0x8b, 0x08, 0, 0, 0, 0, 0, 2, 3]) {
            return Err("gzip header is not the exact gzip -n -9 form".to_owned());
        }
        let scratch = ScratchDir::new();
        let gzip_path = scratch.0.join("bound-artifact.wasm.gz");
        std::fs::write(&gzip_path, &evidence.gzip)
            .map_err(|error| format!("write gzip verification input: {error}"))?;
        let output = Command::new(gzip_program)
            .arg("-dc")
            .arg(gzip_path)
            .output()
            .map_err(|error| format!("run gzip verification: {error}"))?;
        if !output.status.success() || output.stdout != evidence.optimized_wasm {
            return Err("gzip is not bound to the freshly optimized WASM".to_owned());
        }
        if evidence.gzip.len() >= 1_000_000 {
            return Err(format!(
                "optimized default profile is {} bytes, expected less than 1,000,000",
                evidence.gzip.len()
            ));
        }
        Ok(evidence.gzip.len())
    }

    #[test]
    #[ignore = "builds and measures the final default WASM artifact"]
    fn default_profile_is_under_one_megabyte_and_round_trips_a_deck() {
        let evidence = exact_size_pipeline();
        let gzip = std::env::var("RPPTX_WASM_GZIP").unwrap_or_else(|_| "gzip".to_owned());
        let compressed_size =
            validate_size_pipeline(&evidence, &gzip).expect("exact default size pipeline");
        eprintln!("fresh optimized default gzip size: {compressed_size} bytes");
        assert!(compressed_size < 1_000_000);

        let source = nontrivial_deck();
        let presentation = WasmPresentation::from_bytes(&source).expect("open presentation");
        let output = presentation.to_bytes().expect("serialize presentation");
        assert_complete_package_round_trip(&source, &output);

        let mut created = WasmPresentation::new().expect("create presentation");
        created.add_slide(0).expect("add slide");
        assert_eq!(created.slide_count(), 1);
    }

    #[test]
    fn size_gate_rejects_unrelated_gzip_and_wrong_pipeline() {
        let optimized_wasm = b"\0asm\x01\0\0\0".to_vec();
        let unrelated_empty_gzip = vec![
            0x1f, 0x8b, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x03, 0x03, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];
        let mut evidence = SizePipelineEvidence {
            wasm_pack_version: "wasm-pack 0.15.0".to_owned(),
            wasm_opt_version: "wasm-opt version 125".to_owned(),
            wasm_pack_args: WASM_PACK_ARGS.iter().map(ToString::to_string).collect(),
            wasm_opt_args: WASM_OPT_ARGS.iter().map(ToString::to_string).collect(),
            gzip_args: GZIP_ARGS.iter().map(ToString::to_string).collect(),
            optimized_wasm,
            gzip: unrelated_empty_gzip,
        };
        assert_eq!(
            validate_size_pipeline(&evidence, "gzip").unwrap_err(),
            "gzip is not bound to the freshly optimized WASM"
        );

        evidence
            .wasm_pack_args
            .extend(["--".to_owned(), "--no-default-features".to_owned()]);
        assert_eq!(
            validate_size_pipeline(&evidence, "gzip").unwrap_err(),
            "wasm-pack did not use the reviewed pipeline arguments"
        );
        evidence.wasm_pack_args = WASM_PACK_ARGS.iter().map(ToString::to_string).collect();
        evidence.gzip_args = ["-n", "-1", "-c", "<optimized-wasm>"]
            .into_iter()
            .map(ToString::to_string)
            .collect();
        assert_eq!(
            validate_size_pipeline(&evidence, "gzip").unwrap_err(),
            "gzip did not use the reviewed pipeline arguments"
        );
    }

    #[test]
    fn default_profile_round_trips_a_nontrivial_deck() {
        let source = nontrivial_deck();
        let presentation = WasmPresentation::from_bytes(&source).expect("open presentation");
        let output = presentation.to_bytes().expect("serialize presentation");
        assert_complete_package_round_trip(&source, &output);
    }

    #[cfg(target_arch = "wasm32")]
    #[wasm_bindgen_test::wasm_bindgen_test]
    fn wasm_presentation_uses_the_real_facade_in_node() {
        use js_sys::{Function, Reflect, Uint8Array};
        use wasm_bindgen::JsCast;

        let source = nontrivial_deck();
        let probe = JsValue::from(WasmPresentation::new().expect("create presentation"));
        let add_slide: Function = Reflect::get(&probe, &JsValue::from_str("addSlide"))
            .expect("generated addSlide should be visible")
            .dyn_into()
            .expect("addSlide should be callable");
        add_slide
            .call1(&probe, &JsValue::from_f64(0.0))
            .expect("addSlide should accept a layout index");
        let probe_slide_count: Function = Reflect::get(&probe, &JsValue::from_str("slideCount"))
            .expect("generated slideCount should be visible")
            .dyn_into()
            .expect("slideCount should be callable");
        assert_eq!(probe_slide_count.call0(&probe).unwrap().as_f64(), Some(1.0));
        let constructor = Reflect::get(&probe, &JsValue::from_str("constructor"))
            .expect("generated class constructor should be visible");
        let from_bytes: Function = Reflect::get(&constructor, &JsValue::from_str("fromBytes"))
            .expect("generated fromBytes should be visible")
            .dyn_into()
            .expect("fromBytes should be callable");
        let input = Uint8Array::from(source.as_slice());
        let presentation = from_bytes
            .call1(&constructor, &input)
            .expect("fromBytes should accept Uint8Array");
        let slide_count: Function = Reflect::get(&presentation, &JsValue::from_str("slideCount"))
            .expect("generated slideCount should be visible")
            .dyn_into()
            .expect("slideCount should be callable");
        assert_eq!(
            slide_count.call0(&presentation).unwrap().as_f64(),
            Some(1.0)
        );
        let to_bytes: Function = Reflect::get(&presentation, &JsValue::from_str("toBytes"))
            .expect("generated toBytes should be visible")
            .dyn_into()
            .expect("toBytes should be callable");
        let output = to_bytes
            .call0(&presentation)
            .expect("toBytes should return bytes");
        assert!(output.is_instance_of::<Uint8Array>());
        assert_complete_package_round_trip(&source, &Uint8Array::new(&output).to_vec());
    }

    #[cfg(feature = "render")]
    #[test]
    fn render_profile_returns_a_complete_pdf() {
        let mut presentation = WasmPresentation::new().expect("create presentation");
        presentation.add_slide(0).expect("add slide");
        let pdf = presentation.to_pdf().expect("render presentation");
        assert!(pdf.starts_with(b"%PDF-"));
        assert!(pdf.ends_with(b"%%EOF"));
    }

    #[cfg(all(feature = "render", target_arch = "wasm32"))]
    #[wasm_bindgen_test::wasm_bindgen_test]
    fn render_profile_returns_a_complete_pdf_in_node() {
        use js_sys::{Function, Reflect, Uint8Array};
        use wasm_bindgen::JsCast;

        let presentation = JsValue::from(WasmPresentation::new().expect("create presentation"));
        let add_slide: Function = Reflect::get(&presentation, &JsValue::from_str("addSlide"))
            .expect("generated addSlide should be visible")
            .dyn_into()
            .expect("addSlide should be callable");
        add_slide
            .call1(&presentation, &JsValue::from_f64(0.0))
            .expect("addSlide should accept a layout index");
        let to_pdf: Function = Reflect::get(&presentation, &JsValue::from_str("toPdf"))
            .expect("generated toPdf should be visible")
            .dyn_into()
            .expect("toPdf should be callable");
        let pdf = to_pdf
            .call0(&presentation)
            .expect("toPdf should return bytes");
        assert!(pdf.is_instance_of::<Uint8Array>());
        let pdf = Uint8Array::new(&pdf).to_vec();
        assert!(pdf.starts_with(b"%PDF-"));
        assert!(pdf.ends_with(b"%%EOF"));
    }

    #[test]
    fn native_defaults_and_wasm_profiles_are_manifest_contracts() {
        let workspace_manifest = include_str!("../../../Cargo.toml");
        let facade_manifest = include_str!("../../rpptx/Cargo.toml");
        let wasm_manifest = include_str!("../Cargo.toml");

        assert!(workspace_manifest.contains(
            "rpptx = { path = \"crates/rpptx\", version = \"0.11.0\", default-features = false }"
        ));
        assert!(
            facade_manifest
                .contains("default = [\"default-template\", \"render\", \"system-fonts\"]")
        );
        for dependency in [
            "dep:gif",
            "dep:jpeg-encoder",
            "dep:miniz_oxide",
            "dep:oxml-pdf",
            "dep:rpptx-layout",
            "dep:rpptx-render",
            "dep:tiny-skia",
        ] {
            assert!(facade_manifest.contains(&format!("\"{dependency}\"")));
        }
        for dependency in [
            "gif = { workspace = true, optional = true }",
            "jpeg-encoder = { workspace = true, optional = true, features = [\"std\"] }",
            "tiny-skia = { workspace = true, optional = true }",
        ] {
            assert!(facade_manifest.contains(dependency));
            assert!(!wasm_manifest.contains(dependency));
        }
        assert!(wasm_manifest.contains("default = []"));
        assert!(wasm_manifest.contains("render = [\"rpptx/render\"]"));
        assert!(wasm_manifest.contains(
            "rpptx = { workspace = true, default-features = false, features = [\"default-template\"] }"
        ));
        assert!(!wasm_manifest.contains("default = [\"render\"]"));
        assert!(!wasm_manifest.contains("features = [\"default-template\", \"render\"]"));
        assert!(!wasm_manifest.contains("rpptx/system-fonts"));
    }
}
