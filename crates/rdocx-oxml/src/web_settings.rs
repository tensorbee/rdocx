//! Word web settings, the `w:webSettings` part.
//!
//! The part keeps the same source-preserving architecture as the settings
//! part. `w:frameset` and `w:divs` stay preservation-only, because the first
//! carries relationship targets and the second is a nested tree no story asked
//! for. A read-only `div_ids` projection over the retained `w:divs` subtree is
//! what makes a `w:divId` reference checkable without opening HTML division
//! authoring.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::{OxmlError, Result};
use crate::namespace::W_NS;
use crate::numbering::word_prefixes_at;
use crate::properties::is_word_element;
use crate::settings::{
    SettingsDiagnostic, Tally, assign, parse_decimal, parse_toggle, rewrite_ordered_child,
    supported_name, word_attribute, word_value, write_toggle_setting, write_valued_setting,
};

/// Every `CT_WebSettings` child in `xsd:sequence` order.
const WEB_SETTINGS_ORDER: &[&[u8]] = &[
    b"frameset",
    b"divs",
    b"encoding",
    b"optimizeForBrowser",
    b"relyOnVML",
    b"allowPNG",
    b"doNotRelyOnCSS",
    b"doNotSaveAsSingleFile",
    b"doNotOrganizeInFolder",
    b"doNotUseLongFileNames",
    b"pixelsPerInch",
    b"targetScreenSz",
    b"saveSmartTagsAsXml",
];

/// The closed set of `w:webSettings` children the typed model owns.
pub const SUPPORTED_WEB_SETTINGS: &[&str] = &[
    "encoding",
    "optimizeForBrowser",
    "relyOnVML",
    "allowPNG",
    "doNotRelyOnCSS",
    "doNotSaveAsSingleFile",
    "doNotOrganizeInFolder",
    "doNotUseLongFileNames",
    "pixelsPerInch",
    "targetScreenSz",
    "saveSmartTagsAsXml",
];

/// No supported web settings child may occur more than once.
const WEB_SETTINGS_REPEATABLE: &[&[&str]] = &[];

/// The typed contents of a Word web settings part.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CT_WebSettings {
    encoding: Option<String>,
    optimize_for_browser: Option<bool>,
    rely_on_vml: Option<bool>,
    allow_png: Option<bool>,
    do_not_rely_on_css: Option<bool>,
    do_not_save_as_single_file: Option<bool>,
    do_not_organize_in_folder: Option<bool>,
    do_not_use_long_file_names: Option<bool>,
    pixels_per_inch: Option<i32>,
    target_screen_size: Option<String>,
    save_smart_tags_as_xml: Option<bool>,
    div_ids: Vec<u32>,
    tally: Tally,
    diagnostics: Vec<SettingsDiagnostic>,
    /// Parsed parts keep their complete producer bytes as the serialization
    /// source, which is what retains `w:frameset` and `w:divs` verbatim.
    source_xml: Option<Vec<u8>>,
}

impl CT_WebSettings {
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse a complete Word web settings part.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut model = Self::default();
        let mut root_prefixes = Vec::new();
        let mut divs_prefixes = Vec::new();
        let mut divs_depth: Option<usize> = None;
        let mut saw_root = false;
        let mut depth = 0usize;
        let mut buffer = Vec::new();

        loop {
            let event = reader.read_event_into(&mut buffer)?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    let started = matches!(event, Event::Start(_));
                    let inherited = if divs_depth.is_some() {
                        &divs_prefixes
                    } else {
                        &root_prefixes
                    };
                    let prefixes = word_prefixes_at(element, inherited)?;
                    if !saw_root {
                        if !is_word_element(element.name().as_ref(), b"webSettings", &prefixes) {
                            return Err(OxmlError::MissingElement("webSettings root".to_owned()));
                        }
                        root_prefixes = prefixes.into_owned();
                        saw_root = true;
                        depth = usize::from(started);
                        buffer.clear();
                        continue;
                    }
                    if depth == 1 {
                        if is_word_element(element.name().as_ref(), b"divs", &prefixes) {
                            // A self-closing `w:divs` has no children, so its
                            // namespace scope must not leak onto its siblings.
                            if started {
                                divs_depth = Some(depth + 1);
                                divs_prefixes = prefixes.to_vec();
                            }
                        } else {
                            model.absorb(element, &prefixes)?;
                        }
                    } else if divs_depth.is_some_and(|child_depth| depth >= child_depth)
                        && is_word_element(element.name().as_ref(), b"div", &prefixes)
                        && let Some(id) = word_attribute(element, b"id", &prefixes)?
                        && let Ok(id) = id.parse::<u32>()
                    {
                        model.div_ids.push(id);
                    }
                    if started {
                        depth += 1;
                    }
                }
                Event::End(_) if depth > 0 => {
                    if divs_depth == Some(depth) {
                        divs_depth = None;
                        divs_prefixes.clear();
                    }
                    depth -= 1;
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }

        if !saw_root {
            return Err(OxmlError::MissingElement("webSettings root".to_owned()));
        }
        model.diagnostics = model.tally.diagnostics(WEB_SETTINGS_REPEATABLE);
        model.source_xml = Some(xml.to_vec());
        Ok(model)
    }

    /// Project one supported `w:webSettings` child into its typed member.
    fn absorb(&mut self, element: &BytesStart<'_>, prefixes: &[String]) -> Result<()> {
        let Some(name) = supported_name(SUPPORTED_WEB_SETTINGS, element, prefixes) else {
            return Ok(());
        };
        match name {
            "encoding" => {
                let value = word_value(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.encoding, value, total);
            }
            "optimizeForBrowser" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.optimize_for_browser, value, total);
            }
            "relyOnVML" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.rely_on_vml, value, total);
            }
            "allowPNG" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.allow_png, value, total);
            }
            "doNotRelyOnCSS" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.do_not_rely_on_css, value, total);
            }
            "doNotSaveAsSingleFile" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.do_not_save_as_single_file, value, total);
            }
            "doNotOrganizeInFolder" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.do_not_organize_in_folder, value, total);
            }
            "doNotUseLongFileNames" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.do_not_use_long_file_names, value, total);
            }
            "pixelsPerInch" => {
                let value = parse_decimal(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.pixels_per_inch, value, total);
            }
            "targetScreenSz" => {
                let value = word_value(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.target_screen_size, value, total);
            }
            "saveSmartTagsAsXml" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.save_smart_tags_as_xml, value, total);
            }
            _ => {}
        }
        Ok(())
    }

    /// Report every supported child the typed model could not own.
    pub fn diagnostics(&self) -> &[SettingsDiagnostic] {
        &self.diagnostics
    }

    /// Report every `w:div` the retained `w:divs` subtree declares.
    ///
    /// A `w:divId` in a paragraph resolves only when its value appears here,
    /// which is what makes the reference checkable without authoring `w:divs`.
    pub fn div_ids(&self) -> Vec<u32> {
        self.div_ids.clone()
    }

    fn finish_set(&mut self, name: &'static str, replacement: Vec<u8>) -> Result<()> {
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_ordered_child(
                source,
                WEB_SETTINGS_ORDER,
                name.as_bytes(),
                &replacement,
                is_word_element,
            )?);
        }
        self.tally.set(&[name], true);
        self.diagnostics = self.tally.diagnostics(WEB_SETTINGS_REPEATABLE);
        Ok(())
    }

    fn begin_removal(&mut self, name: &'static str) -> Result<bool> {
        let total = self.tally.total(&[name]);
        if total == 0 {
            return Ok(false);
        }
        if total > 1 || self.tally.typed(&[name]) != total {
            return Err(OxmlError::InvalidValue(format!(
                "cannot rewrite ambiguous or malformed w:{name}"
            )));
        }
        Ok(true)
    }

    fn finish_removal(&mut self, name: &'static str) -> Result<()> {
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_ordered_child(
                source,
                WEB_SETTINGS_ORDER,
                name.as_bytes(),
                &[],
                is_word_element,
            )?);
        }
        self.tally.set(&[name], false);
        self.diagnostics = self.tally.diagnostics(WEB_SETTINGS_REPEATABLE);
        Ok(())
    }

    /// Return the character encoding Word writes web pages with.
    pub fn encoding(&self) -> Option<&str> {
        self.encoding.as_deref()
    }

    pub fn set_encoding(&mut self, value: String) -> Result<()> {
        let replacement = write_valued_setting("w:encoding", &value)?;
        self.encoding = Some(value);
        self.finish_set("encoding", replacement)
    }

    pub fn remove_encoding(&mut self) -> Result<Option<String>> {
        if !self.begin_removal("encoding")? {
            return Ok(None);
        }
        let removed = self.encoding.take();
        self.finish_removal("encoding")?;
        Ok(removed)
    }

    /// Return the `w:optimizeForBrowser` toggle.
    pub fn optimize_for_browser(&self) -> Option<bool> {
        self.optimize_for_browser
    }

    pub fn set_optimize_for_browser(&mut self, enabled: bool) -> Result<()> {
        self.optimize_for_browser = Some(enabled);
        self.finish_set(
            "optimizeForBrowser",
            write_toggle_setting("w:optimizeForBrowser", enabled)?,
        )
    }

    pub fn remove_optimize_for_browser(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("optimizeForBrowser")? {
            return Ok(None);
        }
        let removed = self.optimize_for_browser.take();
        self.finish_removal("optimizeForBrowser")?;
        Ok(removed)
    }

    /// Return the `w:relyOnVML` toggle.
    pub fn rely_on_vml(&self) -> Option<bool> {
        self.rely_on_vml
    }

    pub fn set_rely_on_vml(&mut self, enabled: bool) -> Result<()> {
        self.rely_on_vml = Some(enabled);
        self.finish_set("relyOnVML", write_toggle_setting("w:relyOnVML", enabled)?)
    }

    pub fn remove_rely_on_vml(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("relyOnVML")? {
            return Ok(None);
        }
        let removed = self.rely_on_vml.take();
        self.finish_removal("relyOnVML")?;
        Ok(removed)
    }

    /// Return the `w:allowPNG` toggle.
    pub fn allow_png(&self) -> Option<bool> {
        self.allow_png
    }

    pub fn set_allow_png(&mut self, enabled: bool) -> Result<()> {
        self.allow_png = Some(enabled);
        self.finish_set("allowPNG", write_toggle_setting("w:allowPNG", enabled)?)
    }

    pub fn remove_allow_png(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("allowPNG")? {
            return Ok(None);
        }
        let removed = self.allow_png.take();
        self.finish_removal("allowPNG")?;
        Ok(removed)
    }

    /// Return the `w:doNotRelyOnCSS` toggle.
    pub fn do_not_rely_on_css(&self) -> Option<bool> {
        self.do_not_rely_on_css
    }

    pub fn set_do_not_rely_on_css(&mut self, enabled: bool) -> Result<()> {
        self.do_not_rely_on_css = Some(enabled);
        self.finish_set(
            "doNotRelyOnCSS",
            write_toggle_setting("w:doNotRelyOnCSS", enabled)?,
        )
    }

    pub fn remove_do_not_rely_on_css(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("doNotRelyOnCSS")? {
            return Ok(None);
        }
        let removed = self.do_not_rely_on_css.take();
        self.finish_removal("doNotRelyOnCSS")?;
        Ok(removed)
    }

    /// Return the `w:doNotSaveAsSingleFile` toggle.
    pub fn do_not_save_as_single_file(&self) -> Option<bool> {
        self.do_not_save_as_single_file
    }

    pub fn set_do_not_save_as_single_file(&mut self, enabled: bool) -> Result<()> {
        self.do_not_save_as_single_file = Some(enabled);
        self.finish_set(
            "doNotSaveAsSingleFile",
            write_toggle_setting("w:doNotSaveAsSingleFile", enabled)?,
        )
    }

    pub fn remove_do_not_save_as_single_file(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("doNotSaveAsSingleFile")? {
            return Ok(None);
        }
        let removed = self.do_not_save_as_single_file.take();
        self.finish_removal("doNotSaveAsSingleFile")?;
        Ok(removed)
    }

    /// Return the `w:doNotOrganizeInFolder` toggle.
    pub fn do_not_organize_in_folder(&self) -> Option<bool> {
        self.do_not_organize_in_folder
    }

    pub fn set_do_not_organize_in_folder(&mut self, enabled: bool) -> Result<()> {
        self.do_not_organize_in_folder = Some(enabled);
        self.finish_set(
            "doNotOrganizeInFolder",
            write_toggle_setting("w:doNotOrganizeInFolder", enabled)?,
        )
    }

    pub fn remove_do_not_organize_in_folder(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("doNotOrganizeInFolder")? {
            return Ok(None);
        }
        let removed = self.do_not_organize_in_folder.take();
        self.finish_removal("doNotOrganizeInFolder")?;
        Ok(removed)
    }

    /// Return the `w:doNotUseLongFileNames` toggle.
    pub fn do_not_use_long_file_names(&self) -> Option<bool> {
        self.do_not_use_long_file_names
    }

    pub fn set_do_not_use_long_file_names(&mut self, enabled: bool) -> Result<()> {
        self.do_not_use_long_file_names = Some(enabled);
        self.finish_set(
            "doNotUseLongFileNames",
            write_toggle_setting("w:doNotUseLongFileNames", enabled)?,
        )
    }

    pub fn remove_do_not_use_long_file_names(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("doNotUseLongFileNames")? {
            return Ok(None);
        }
        let removed = self.do_not_use_long_file_names.take();
        self.finish_removal("doNotUseLongFileNames")?;
        Ok(removed)
    }

    /// Return the target pixel density for exported web pages.
    pub fn pixels_per_inch(&self) -> Option<i32> {
        self.pixels_per_inch
    }

    pub fn set_pixels_per_inch(&mut self, value: i32) -> Result<()> {
        let replacement = write_valued_setting("w:pixelsPerInch", &value.to_string())?;
        self.pixels_per_inch = Some(value);
        self.finish_set("pixelsPerInch", replacement)
    }

    pub fn remove_pixels_per_inch(&mut self) -> Result<Option<i32>> {
        if !self.begin_removal("pixelsPerInch")? {
            return Ok(None);
        }
        let removed = self.pixels_per_inch.take();
        self.finish_removal("pixelsPerInch")?;
        Ok(removed)
    }

    /// Return the target screen size, such as `800x600`.
    pub fn target_screen_size(&self) -> Option<&str> {
        self.target_screen_size.as_deref()
    }

    pub fn set_target_screen_size(&mut self, value: String) -> Result<()> {
        let replacement = write_valued_setting("w:targetScreenSz", &value)?;
        self.target_screen_size = Some(value);
        self.finish_set("targetScreenSz", replacement)
    }

    pub fn remove_target_screen_size(&mut self) -> Result<Option<String>> {
        if !self.begin_removal("targetScreenSz")? {
            return Ok(None);
        }
        let removed = self.target_screen_size.take();
        self.finish_removal("targetScreenSz")?;
        Ok(removed)
    }

    /// Return the `w:saveSmartTagsAsXml` toggle.
    pub fn save_smart_tags_as_xml(&self) -> Option<bool> {
        self.save_smart_tags_as_xml
    }

    pub fn set_save_smart_tags_as_xml(&mut self, enabled: bool) -> Result<()> {
        self.save_smart_tags_as_xml = Some(enabled);
        self.finish_set(
            "saveSmartTagsAsXml",
            write_toggle_setting("w:saveSmartTagsAsXml", enabled)?,
        )
    }

    pub fn remove_save_smart_tags_as_xml(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("saveSmartTagsAsXml")? {
            return Ok(None);
        }
        let removed = self.save_smart_tags_as_xml.take();
        self.finish_removal("saveSmartTagsAsXml")?;
        Ok(removed)
    }

    /// Return whether the model has no typed or retained web settings content.
    pub fn is_empty(&self) -> bool {
        self.encoding.is_none()
            && self.optimize_for_browser.is_none()
            && self.rely_on_vml.is_none()
            && self.allow_png.is_none()
            && self.do_not_rely_on_css.is_none()
            && self.do_not_save_as_single_file.is_none()
            && self.do_not_organize_in_folder.is_none()
            && self.do_not_use_long_file_names.is_none()
            && self.pixels_per_inch.is_none()
            && self.target_screen_size.is_none()
            && self.save_smart_tags_as_xml.is_none()
            && self.div_ids.is_empty()
            && self.source_xml.is_none()
    }

    /// Serialize web settings with fixed Word prefixes and schema child order.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        if let Some(source) = &self.source_xml {
            return Ok(source.clone());
        }

        let mut writer = Writer::new(Vec::new());
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;
        let mut root = BytesStart::new("w:webSettings");
        root.push_attribute(("xmlns:w", W_NS));
        writer.write_event(Event::Start(root))?;
        let mut emit = |bytes: Vec<u8>| writer.get_mut().extend_from_slice(&bytes);
        if let Some(value) = &self.encoding {
            emit(write_valued_setting("w:encoding", value)?);
        }
        for (name, value) in [
            ("w:optimizeForBrowser", self.optimize_for_browser),
            ("w:relyOnVML", self.rely_on_vml),
            ("w:allowPNG", self.allow_png),
            ("w:doNotRelyOnCSS", self.do_not_rely_on_css),
            ("w:doNotSaveAsSingleFile", self.do_not_save_as_single_file),
            ("w:doNotOrganizeInFolder", self.do_not_organize_in_folder),
            ("w:doNotUseLongFileNames", self.do_not_use_long_file_names),
        ] {
            if let Some(enabled) = value {
                emit(write_toggle_setting(name, enabled)?);
            }
        }
        if let Some(value) = self.pixels_per_inch {
            emit(write_valued_setting("w:pixelsPerInch", &value.to_string())?);
        }
        if let Some(value) = &self.target_screen_size {
            emit(write_valued_setting("w:targetScreenSz", value)?);
        }
        if let Some(enabled) = self.save_smart_tags_as_xml {
            emit(write_toggle_setting("w:saveSmartTagsAsXml", enabled)?);
        }
        writer.write_event(Event::End(BytesEnd::new("w:webSettings")))?;
        Ok(writer.into_inner())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn every_supported_child() -> String {
        format!(
            concat!(
                r#"<w:webSettings xmlns:w="{word}" xmlns:r="urn:rel" xmlns:x="urn:foreign">"#,
                r#"<w:frameset><w:frame r:id="rId1"/></w:frameset>"#,
                r#"<w:divs><w:div w:id="1"><w:divsChild><w:div w:id="2"/></w:divsChild></w:div>"#,
                r#"<w:div w:id="3"/></w:divs>"#,
                r#"<w:encoding w:val="utf-8"/>"#,
                r#"<w:optimizeForBrowser/>"#,
                r#"<w:relyOnVML w:val="false"/>"#,
                r#"<w:allowPNG/>"#,
                r#"<w:doNotRelyOnCSS/>"#,
                r#"<w:doNotSaveAsSingleFile/>"#,
                r#"<w:doNotOrganizeInFolder/>"#,
                r#"<w:doNotUseLongFileNames/>"#,
                r#"<w:pixelsPerInch w:val="96"/>"#,
                r#"<w:targetScreenSz w:val="800x600"/>"#,
                r#"<w:saveSmartTagsAsXml/>"#,
                r#"<x:kept x:value="raw"/>"#,
                r#"</w:webSettings>"#,
            ),
            word = W_NS,
        )
    }

    #[test]
    fn every_supported_web_setting_projects_and_keeps_producer_bytes() {
        let xml = every_supported_child();
        let settings = CT_WebSettings::from_xml(xml.as_bytes()).unwrap();

        assert_eq!(settings.encoding(), Some("utf-8"));
        assert_eq!(settings.optimize_for_browser(), Some(true));
        assert_eq!(settings.rely_on_vml(), Some(false));
        assert_eq!(settings.allow_png(), Some(true));
        assert_eq!(settings.do_not_rely_on_css(), Some(true));
        assert_eq!(settings.do_not_save_as_single_file(), Some(true));
        assert_eq!(settings.do_not_organize_in_folder(), Some(true));
        assert_eq!(settings.do_not_use_long_file_names(), Some(true));
        assert_eq!(settings.pixels_per_inch(), Some(96));
        assert_eq!(settings.target_screen_size(), Some("800x600"));
        assert_eq!(settings.save_smart_tags_as_xml(), Some(true));
        assert_eq!(settings.diagnostics(), &[]);
        assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());
    }

    #[test]
    fn declared_divisions_are_reported_for_reference_checking() {
        let settings = CT_WebSettings::from_xml(every_supported_child().as_bytes()).unwrap();
        assert_eq!(settings.div_ids(), vec![1, 2, 3]);

        for xml in [
            format!(r#"<w:webSettings xmlns:w="{W_NS}"><w:divs/></w:webSettings>"#),
            format!(r#"<w:webSettings xmlns:w="{W_NS}"/>"#),
        ] {
            let settings = CT_WebSettings::from_xml(xml.as_bytes()).unwrap();
            assert!(settings.div_ids().is_empty(), "{xml}");
            assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());
        }
    }

    #[test]
    fn authored_web_settings_land_in_schema_order_and_keep_neighbours() {
        let xml = format!(
            r#"<q:webSettings xmlns:q="{W_NS}" xmlns:x="urn:foreign"><q:divs><q:div q:id="4"/></q:divs><x:kept/><q:pixelsPerInch q:val="96"/></q:webSettings>"#,
        );
        let mut settings = CT_WebSettings::from_xml(xml.as_bytes()).unwrap();
        settings.set_allow_png(true).unwrap();
        let output = String::from_utf8(settings.to_xml().unwrap()).unwrap();
        assert!(output.contains("<x:kept/>"), "{output}");
        assert!(output.contains(r#"<q:div q:id="4"/>"#), "{output}");
        assert!(
            output.find("<q:divs").unwrap() < output.find("<w:allowPNG/>").unwrap(),
            "{output}"
        );
        assert!(
            output.find("<w:allowPNG/>").unwrap() < output.find("pixelsPerInch").unwrap(),
            "{output}"
        );

        let mut reopened = CT_WebSettings::from_xml(output.as_bytes()).unwrap();
        assert_eq!(reopened.div_ids(), vec![4]);
        assert_eq!(reopened.remove_allow_png().unwrap(), Some(true));
        let cleared = String::from_utf8(reopened.to_xml().unwrap()).unwrap();
        assert!(!cleared.contains("allowPNG"), "{cleared}");
        assert!(cleared.contains("<x:kept/>"), "{cleared}");
        assert!(cleared.contains(r#"<q:div q:id="4"/>"#), "{cleared}");
    }

    #[test]
    fn a_self_closing_divs_element_does_not_swallow_later_children() {
        let xml = format!(
            r#"<w:webSettings xmlns:w="{W_NS}"><w:divs/><w:allowPNG/><w:pixelsPerInch w:val="96"/></w:webSettings>"#,
        );
        let settings = CT_WebSettings::from_xml(xml.as_bytes()).unwrap();
        assert!(settings.div_ids().is_empty());
        assert_eq!(settings.allow_png(), Some(true));
        assert_eq!(settings.pixels_per_inch(), Some(96));
        assert_eq!(settings.diagnostics(), &[]);
        assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());
    }

    #[test]
    fn duplicate_and_malformed_web_settings_report_diagnostics() {
        let xml = format!(
            r#"<w:webSettings xmlns:w="{W_NS}"><w:allowPNG/><w:allowPNG w:val="false"/><w:pixelsPerInch w:val="lots"/></w:webSettings>"#,
        );
        let mut settings = CT_WebSettings::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(settings.diagnostics().len(), 2);
        assert_eq!(settings.diagnostics()[0].path, vec!["allowPNG"]);
        assert_eq!(settings.diagnostics()[0].occurrences, 2);
        assert_eq!(settings.diagnostics()[1].path, vec!["pixelsPerInch"]);
        assert!(settings.remove_allow_png().is_err());
        assert!(settings.remove_pixels_per_inch().is_err());
        assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());
    }

    #[test]
    fn a_constructed_part_writes_the_fixed_prefix_in_schema_order() {
        let mut settings = CT_WebSettings::new();
        assert!(settings.is_empty());
        settings.set_save_smart_tags_as_xml(true).unwrap();
        settings.set_encoding("utf-8".to_owned()).unwrap();
        assert!(!settings.is_empty());
        let output = String::from_utf8(settings.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"<w:webSettings xmlns:w=""#), "{output}");
        assert!(
            output.find("encoding").unwrap() < output.find("saveSmartTagsAsXml").unwrap(),
            "{output}"
        );
        assert_eq!(
            CT_WebSettings::from_xml(output.as_bytes())
                .unwrap()
                .encoding(),
            Some("utf-8")
        );
    }
}
