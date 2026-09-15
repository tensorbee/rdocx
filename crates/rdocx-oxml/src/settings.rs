//! Word document settings and document-protection metadata.

use oxml_core::Twips;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::error::{OxmlError, Result};
use crate::math::{MathProperties, fixed_math_prefix_is_safe, is_math_element};
use crate::namespace::W_NS;
use crate::numbering::{namespace_bindings, word_prefixes_at};
use crate::properties::{is_word_attribute, is_word_element};
use crate::raw_xml::capture_element;

/// The editing operation permitted by `w:documentProtection`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtectionMode {
    ReadOnly,
    Comments,
    TrackedChanges,
    Forms,
}

impl ProtectionMode {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "readOnly" => Some(Self::ReadOnly),
            "comments" => Some(Self::Comments),
            "trackedChanges" => Some(Self::TrackedChanges),
            "forms" => Some(Self::Forms),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::ReadOnly => "readOnly",
            Self::Comments => "comments",
            Self::TrackedChanges => "trackedChanges",
            Self::Forms => "forms",
        }
    }
}

/// The cryptographic provider category recorded by Word.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CryptProviderType {
    RsaAes,
    RsaFull,
    Custom,
}

impl CryptProviderType {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "rsaAES" => Some(Self::RsaAes),
            "rsaFull" => Some(Self::RsaFull),
            "custom" => Some(Self::Custom),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::RsaAes => "rsaAES",
            Self::RsaFull => "rsaFull",
            Self::Custom => "custom",
        }
    }
}

/// The algorithm class recorded by Word.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CryptAlgorithmClass {
    Hash,
    Custom,
}

impl CryptAlgorithmClass {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "hash" => Some(Self::Hash),
            "custom" => Some(Self::Custom),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Hash => "hash",
            Self::Custom => "custom",
        }
    }
}

/// The algorithm type recorded by Word.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CryptAlgorithmType {
    Any,
    Custom,
}

impl CryptAlgorithmType {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "typeAny" => Some(Self::Any),
            "custom" => Some(Self::Custom),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Any => "typeAny",
            Self::Custom => "custom",
        }
    }
}

/// A read-only projection of one valid `w:documentProtection` element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentProtection {
    pub mode: ProtectionMode,
    pub enforcement: Option<bool>,
    pub formatting: Option<bool>,
    pub provider_type: Option<CryptProviderType>,
    pub algorithm_class: Option<CryptAlgorithmClass>,
    pub algorithm_type: Option<CryptAlgorithmType>,
    pub algorithm_sid: Option<u32>,
    pub spin_count: Option<u32>,
    pub hash: Option<String>,
    pub salt: Option<String>,
}

/// One valid document variable from `w:settings/w:docVars`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentVariable {
    pub name: String,
    pub value: String,
}

/// One `w:compatSetting` entry from the document compatibility settings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CompatibilitySetting {
    pub name: String,
    pub uri: String,
    pub value: String,
}

/// Document-wide character spacing compression behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterSpacingControl {
    DoNotCompress,
    CompressPunctuation,
    CompressPunctuationAndJapaneseKana,
}

impl CharacterSpacingControl {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "doNotCompress" => Some(Self::DoNotCompress),
            "compressPunctuation" => Some(Self::CompressPunctuation),
            "compressPunctuationAndJapaneseKana" => Some(Self::CompressPunctuationAndJapaneseKana),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::DoNotCompress => "doNotCompress",
            Self::CompressPunctuation => "compressPunctuation",
            Self::CompressPunctuationAndJapaneseKana => "compressPunctuationAndJapaneseKana",
        }
    }
}

/// Default document theme languages from `w:themeFontLang`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ThemeFontLanguage {
    pub latin: Option<String>,
    pub east_asia: Option<String>,
    pub bidi: Option<String>,
}

/// The typed contents of a Word settings part.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CT_Settings {
    document_protection: Option<DocumentProtection>,
    document_variables: Vec<DocumentVariable>,
    compatibility_settings: Vec<CompatibilitySetting>,
    default_tab_stop: Option<Twips>,
    character_spacing_control: Option<CharacterSpacingControl>,
    theme_font_language: Option<ThemeFontLanguage>,
    automatic_hyphenation: Option<bool>,
    even_and_odd_headers: Option<bool>,
    update_fields: Option<bool>,
    math_properties: Option<MathProperties>,
    /// Parsed parts keep their complete producer bytes as the serialization
    /// source. This retains root attributes, child order, whitespace, and all
    /// unmodelled content without interpreting it.
    source_xml: Option<Vec<u8>>,
}

impl CT_Settings {
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse a complete Word settings part.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut root_prefixes = Vec::new();
        let mut protection = None;
        let mut protection_count = 0usize;
        let mut document_variables = Vec::new();
        let mut compatibility_settings = Vec::new();
        let mut default_tab_stop = None;
        let mut default_tab_stop_count = 0usize;
        let mut character_spacing_control = None;
        let mut character_spacing_control_count = 0usize;
        let mut theme_font_language = None;
        let mut theme_font_language_count = 0usize;
        let mut automatic_hyphenation = None;
        let mut automatic_hyphenation_count = 0usize;
        let mut even_and_odd_headers = None;
        let mut even_and_odd_headers_count = 0usize;
        let mut update_fields = None;
        let mut update_fields_count = 0usize;
        let mut math_properties = None;
        let mut math_properties_count = 0usize;
        let mut doc_vars_depth = None;
        let mut doc_vars_prefixes = Vec::new();
        let mut compat_depth = None;
        let mut compat_prefixes = Vec::new();
        let mut saw_root = false;
        let mut depth = 0usize;
        let mut buffer = Vec::new();

        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(element) => {
                    let inherited = if doc_vars_depth.is_some() {
                        &doc_vars_prefixes
                    } else if compat_depth.is_some() {
                        &compat_prefixes
                    } else {
                        &root_prefixes
                    };
                    let prefixes = word_prefixes_at(&element, inherited)?;
                    if !saw_root {
                        if !is_word_element(element.name().as_ref(), b"settings", &prefixes) {
                            return Err(OxmlError::MissingElement("settings root".to_owned()));
                        }
                        root_prefixes = prefixes;
                        saw_root = true;
                        depth = 1;
                    } else {
                        if depth == 1
                            && is_math_element(element.name().as_ref(), b"mathPr", &prefixes)
                        {
                            math_properties_count += 1;
                            let raw = capture_element(&mut reader, &element)?;
                            let bindings = namespace_bindings(&prefixes);
                            if fixed_math_prefix_is_safe(&raw, &bindings)? {
                                math_properties = Some(MathProperties::from_raw(&raw, &bindings)?);
                            }
                            buffer.clear();
                            continue;
                        }
                        if depth == 1
                            && is_word_element(
                                element.name().as_ref(),
                                b"documentProtection",
                                &prefixes,
                            )
                        {
                            protection_count += 1;
                            protection = parse_document_protection(&element, &prefixes);
                        } else if depth == 1
                            && is_word_element(
                                element.name().as_ref(),
                                b"autoHyphenation",
                                &prefixes,
                            )
                        {
                            automatic_hyphenation_count += 1;
                            automatic_hyphenation = parse_toggle(&element, &prefixes)?;
                        } else if depth == 1
                            && is_word_element(
                                element.name().as_ref(),
                                b"evenAndOddHeaders",
                                &prefixes,
                            )
                        {
                            even_and_odd_headers_count += 1;
                            even_and_odd_headers = parse_toggle(&element, &prefixes)?;
                        } else if depth == 1
                            && is_word_element(element.name().as_ref(), b"updateFields", &prefixes)
                        {
                            update_fields_count += 1;
                            update_fields = parse_toggle(&element, &prefixes)?;
                        } else if depth == 1
                            && is_word_element(
                                element.name().as_ref(),
                                b"defaultTabStop",
                                &prefixes,
                            )
                        {
                            default_tab_stop_count += 1;
                            default_tab_stop = parse_twips(&element, &prefixes);
                        } else if depth == 1
                            && is_word_element(
                                element.name().as_ref(),
                                b"characterSpacingControl",
                                &prefixes,
                            )
                        {
                            character_spacing_control_count += 1;
                            character_spacing_control =
                                parse_character_spacing(&element, &prefixes);
                        } else if depth == 1
                            && is_word_element(element.name().as_ref(), b"themeFontLang", &prefixes)
                        {
                            theme_font_language_count += 1;
                            theme_font_language =
                                Some(parse_theme_font_language(&element, &prefixes)?);
                        } else if depth == 1
                            && is_word_element(element.name().as_ref(), b"compat", &prefixes)
                        {
                            compat_depth = Some(depth + 1);
                            compat_prefixes = prefixes.clone();
                        } else if depth == 1
                            && is_word_element(element.name().as_ref(), b"docVars", &prefixes)
                        {
                            doc_vars_depth = Some(depth + 1);
                            doc_vars_prefixes = prefixes.clone();
                        } else if doc_vars_depth == Some(depth)
                            && is_word_element(element.name().as_ref(), b"docVar", &prefixes)
                            && let Some(variable) = parse_document_variable(&element, &prefixes)
                        {
                            document_variables.push(variable);
                        } else if compat_depth == Some(depth)
                            && is_word_element(element.name().as_ref(), b"compatSetting", &prefixes)
                            && let Some(setting) = parse_compatibility_setting(&element, &prefixes)
                        {
                            compatibility_settings.push(setting);
                        }
                        depth += 1;
                    }
                }
                Event::Empty(element) => {
                    let inherited = if doc_vars_depth.is_some() {
                        &doc_vars_prefixes
                    } else if compat_depth.is_some() {
                        &compat_prefixes
                    } else {
                        &root_prefixes
                    };
                    let prefixes = word_prefixes_at(&element, inherited)?;
                    if !saw_root {
                        if !is_word_element(element.name().as_ref(), b"settings", &prefixes) {
                            return Err(OxmlError::MissingElement("settings root".to_owned()));
                        }
                        saw_root = true;
                    } else if depth == 1
                        && is_math_element(element.name().as_ref(), b"mathPr", &prefixes)
                    {
                        math_properties_count += 1;
                        let raw = capture_empty_element(&element)?;
                        let bindings = namespace_bindings(&prefixes);
                        if fixed_math_prefix_is_safe(&raw, &bindings)? {
                            math_properties = Some(MathProperties::from_raw(&raw, &bindings)?);
                        }
                    } else if depth == 1
                        && is_word_element(
                            element.name().as_ref(),
                            b"documentProtection",
                            &prefixes,
                        )
                    {
                        protection_count += 1;
                        protection = parse_document_protection(&element, &prefixes);
                    } else if depth == 1
                        && is_word_element(element.name().as_ref(), b"autoHyphenation", &prefixes)
                    {
                        automatic_hyphenation_count += 1;
                        automatic_hyphenation = parse_toggle(&element, &prefixes)?;
                    } else if depth == 1
                        && is_word_element(element.name().as_ref(), b"evenAndOddHeaders", &prefixes)
                    {
                        even_and_odd_headers_count += 1;
                        even_and_odd_headers = parse_toggle(&element, &prefixes)?;
                    } else if depth == 1
                        && is_word_element(element.name().as_ref(), b"updateFields", &prefixes)
                    {
                        update_fields_count += 1;
                        update_fields = parse_toggle(&element, &prefixes)?;
                    } else if depth == 1
                        && is_word_element(element.name().as_ref(), b"defaultTabStop", &prefixes)
                    {
                        default_tab_stop_count += 1;
                        default_tab_stop = parse_twips(&element, &prefixes);
                    } else if depth == 1
                        && is_word_element(
                            element.name().as_ref(),
                            b"characterSpacingControl",
                            &prefixes,
                        )
                    {
                        character_spacing_control_count += 1;
                        character_spacing_control = parse_character_spacing(&element, &prefixes);
                    } else if depth == 1
                        && is_word_element(element.name().as_ref(), b"themeFontLang", &prefixes)
                    {
                        theme_font_language_count += 1;
                        theme_font_language = Some(parse_theme_font_language(&element, &prefixes)?);
                    } else if doc_vars_depth == Some(depth)
                        && is_word_element(element.name().as_ref(), b"docVar", &prefixes)
                        && let Some(variable) = parse_document_variable(&element, &prefixes)
                    {
                        document_variables.push(variable);
                    } else if compat_depth == Some(depth)
                        && is_word_element(element.name().as_ref(), b"compatSetting", &prefixes)
                        && let Some(setting) = parse_compatibility_setting(&element, &prefixes)
                    {
                        compatibility_settings.push(setting);
                    }
                }
                Event::End(_) if depth > 0 => {
                    if doc_vars_depth == Some(depth) {
                        doc_vars_depth = None;
                        doc_vars_prefixes.clear();
                    }
                    if compat_depth == Some(depth) {
                        compat_depth = None;
                        compat_prefixes.clear();
                    }
                    depth -= 1;
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }

        if !saw_root {
            return Err(OxmlError::MissingElement("settings root".to_owned()));
        }
        if protection_count != 1 {
            protection = None;
        }
        if automatic_hyphenation_count != 1 {
            automatic_hyphenation = None;
        }
        if even_and_odd_headers_count != 1 {
            even_and_odd_headers = None;
        }
        if update_fields_count != 1 {
            update_fields = None;
        }
        if math_properties_count != 1 {
            math_properties = None;
        }
        if default_tab_stop_count != 1 {
            default_tab_stop = None;
        }
        if character_spacing_control_count != 1 {
            character_spacing_control = None;
        }
        if theme_font_language_count != 1 {
            theme_font_language = None;
        }
        Ok(Self {
            document_protection: protection,
            document_variables,
            compatibility_settings,
            default_tab_stop,
            character_spacing_control,
            theme_font_language,
            automatic_hyphenation,
            even_and_odd_headers,
            update_fields,
            math_properties,
            source_xml: Some(xml.to_vec()),
        })
    }

    /// Return valid document-protection metadata, when the part records it.
    pub fn document_protection(&self) -> Option<&DocumentProtection> {
        self.document_protection.as_ref()
    }

    /// Return every valid document variable in package order.
    pub fn document_variables(&self) -> &[DocumentVariable] {
        &self.document_variables
    }

    pub fn document_variable(&self, name: &str) -> Option<&str> {
        self.document_variables
            .iter()
            .find(|variable| variable.name == name)
            .map(|variable| variable.value.as_str())
    }

    pub fn set_document_variable(&mut self, name: String, value: String) -> Result<()> {
        if let Some(index) = self
            .document_variables
            .iter()
            .position(|variable| variable.name == name)
        {
            self.document_variables[index].value = value;
            let mut occurrence = 0usize;
            self.document_variables.retain(|variable| {
                if variable.name != name {
                    return true;
                }
                occurrence += 1;
                occurrence == 1
            });
        } else {
            self.document_variables
                .push(DocumentVariable { name, value });
        }
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_group_setting(
                source,
                b"docVars",
                b"docVar",
                write_document_variables(&self.document_variables)?,
            )?);
        }
        Ok(())
    }

    pub fn remove_document_variable(&mut self, name: &str) -> Result<Option<DocumentVariable>> {
        let Some(removed) = self
            .document_variables
            .iter()
            .find(|variable| variable.name == name)
            .cloned()
        else {
            return Ok(None);
        };
        self.document_variables
            .retain(|variable| variable.name != name);
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_group_setting(
                source,
                b"docVars",
                b"docVar",
                write_document_variables(&self.document_variables)?,
            )?);
        }
        Ok(Some(removed))
    }

    pub fn compatibility_settings(&self) -> &[CompatibilitySetting] {
        &self.compatibility_settings
    }

    pub fn set_compatibility_setting(&mut self, setting: CompatibilitySetting) -> Result<()> {
        if let Some(index) = self
            .compatibility_settings
            .iter()
            .position(|existing| existing.name == setting.name && existing.uri == setting.uri)
        {
            self.compatibility_settings[index] = setting.clone();
            let mut occurrence = 0usize;
            self.compatibility_settings.retain(|existing| {
                if existing.name != setting.name || existing.uri != setting.uri {
                    return true;
                }
                occurrence += 1;
                occurrence == 1
            });
        } else {
            self.compatibility_settings.push(setting);
        }
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_group_setting(
                source,
                b"compat",
                b"compatSetting",
                write_compatibility_settings(&self.compatibility_settings)?,
            )?);
        }
        Ok(())
    }

    pub fn remove_compatibility_setting(
        &mut self,
        name: &str,
        uri: &str,
    ) -> Result<Option<CompatibilitySetting>> {
        let Some(removed) = self
            .compatibility_settings
            .iter()
            .find(|setting| setting.name == name && setting.uri == uri)
            .cloned()
        else {
            return Ok(None);
        };
        self.compatibility_settings
            .retain(|setting| setting.name != name || setting.uri != uri);
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_group_setting(
                source,
                b"compat",
                b"compatSetting",
                write_compatibility_settings(&self.compatibility_settings)?,
            )?);
        }
        Ok(Some(removed))
    }

    pub fn default_tab_stop(&self) -> Option<Twips> {
        self.default_tab_stop
    }

    pub fn set_default_tab_stop(&mut self, value: Twips) -> Result<()> {
        if value.0 < 0 {
            return Err(OxmlError::InvalidValue(
                "default tab stop must be non-negative twips".to_owned(),
            ));
        }
        let replacement = write_valued_setting("w:defaultTabStop", &value.0.to_string())?;
        self.rewrite_scalar(b"defaultTabStop", replacement)?;
        self.default_tab_stop = Some(value);
        Ok(())
    }

    pub fn remove_default_tab_stop(&mut self) -> Result<Option<Twips>> {
        let removed = self.default_tab_stop.take();
        self.rewrite_scalar(b"defaultTabStop", Vec::new())?;
        Ok(removed)
    }

    pub fn character_spacing_control(&self) -> Option<CharacterSpacingControl> {
        self.character_spacing_control
    }

    pub fn set_character_spacing_control(&mut self, value: CharacterSpacingControl) -> Result<()> {
        let replacement = write_valued_setting("w:characterSpacingControl", value.as_str())?;
        self.rewrite_scalar(b"characterSpacingControl", replacement)?;
        self.character_spacing_control = Some(value);
        Ok(())
    }

    pub fn remove_character_spacing_control(&mut self) -> Result<Option<CharacterSpacingControl>> {
        let removed = self.character_spacing_control.take();
        self.rewrite_scalar(b"characterSpacingControl", Vec::new())?;
        Ok(removed)
    }

    pub fn theme_font_language(&self) -> Option<&ThemeFontLanguage> {
        self.theme_font_language.as_ref()
    }

    /// Return whether the model has no typed or retained settings content.
    pub fn is_empty(&self) -> bool {
        self.document_protection.is_none()
            && self.document_variables.is_empty()
            && self.compatibility_settings.is_empty()
            && self.default_tab_stop.is_none()
            && self.character_spacing_control.is_none()
            && self.theme_font_language.is_none()
            && self.automatic_hyphenation.is_none()
            && self.even_and_odd_headers.is_none()
            && self.update_fields.is_none()
            && self.math_properties.is_none()
            && self.source_xml.is_none()
    }

    pub fn set_theme_font_language(&mut self, value: ThemeFontLanguage) -> Result<()> {
        let replacement = write_theme_font_language(&value)?;
        self.rewrite_scalar(b"themeFontLang", replacement)?;
        self.theme_font_language = Some(value);
        Ok(())
    }

    pub fn remove_theme_font_language(&mut self) -> Result<Option<ThemeFontLanguage>> {
        let removed = self.theme_font_language.take();
        self.rewrite_scalar(b"themeFontLang", Vec::new())?;
        Ok(removed)
    }

    fn rewrite_scalar(&mut self, local: &[u8], replacement: Vec<u8>) -> Result<()> {
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_top_level_setting(source, local, &replacement)?);
        }
        Ok(())
    }

    /// Return whether Word automatic hyphenation is enabled.
    ///
    /// OOXML defines omission as disabled.
    pub fn automatic_hyphenation(&self) -> bool {
        self.automatic_hyphenation.unwrap_or(false)
    }

    /// Return document-wide OfficeMath defaults, when present and valid.
    pub fn math_properties(&self) -> Option<&MathProperties> {
        self.math_properties.as_ref()
    }

    /// Replace the single schema-positioned OfficeMath defaults subtree.
    pub fn set_math_properties(&mut self, properties: MathProperties) -> Result<()> {
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_math_properties(source, &properties)?);
        }
        self.math_properties = Some(properties);
        Ok(())
    }

    /// Set the document automatic-hyphenation toggle.
    ///
    /// Parsed settings retain every unrelated producer byte. The one modeled
    /// toggle is rewritten with the fixed `w:` prefix at its schema position.
    pub fn set_automatic_hyphenation(&mut self, enabled: bool) -> Result<()> {
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_automatic_hyphenation(source, enabled)?);
        }
        self.automatic_hyphenation = Some(enabled);
        Ok(())
    }

    /// Return whether Word selects distinct even-page header and footer stories.
    ///
    /// OOXML defines omission as disabled.
    pub fn even_and_odd_headers(&self) -> bool {
        self.even_and_odd_headers.unwrap_or(false)
    }

    /// Set distinct even-page header and footer selection.
    pub fn set_even_and_odd_headers(&mut self, enabled: bool) -> Result<()> {
        let replacement = write_toggle_setting("w:evenAndOddHeaders", enabled)?;
        self.rewrite_scalar(b"evenAndOddHeaders", replacement)?;
        self.even_and_odd_headers = Some(enabled);
        Ok(())
    }

    /// Return the `w:updateFields` toggle, which asks Word to update fields
    /// when it opens the document.
    ///
    /// `None` means the part omits the toggle, or repeats it ambiguously.
    pub fn update_fields(&self) -> Option<bool> {
        self.update_fields
    }

    /// Set the `w:updateFields` toggle, or remove it with `None`.
    pub fn set_update_fields(&mut self, value: Option<bool>) -> Result<()> {
        let replacement = match value {
            Some(enabled) => write_toggle_setting("w:updateFields", enabled)?,
            None => Vec::new(),
        };
        self.rewrite_scalar(b"updateFields", replacement)?;
        self.update_fields = value;
        Ok(())
    }

    /// Serialize settings with fixed Word prefixes and schema child order.
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
        let mut root = BytesStart::new("w:settings");
        root.push_attribute(("xmlns:w", W_NS));
        writer.write_event(Event::Start(root))?;
        if let Some(protection) = &self.document_protection {
            write_document_protection(&mut writer, protection)?;
        }
        if let Some(value) = self.default_tab_stop {
            writer.get_mut().extend_from_slice(&write_valued_setting(
                "w:defaultTabStop",
                &value.0.to_string(),
            )?);
        }
        if let Some(enabled) = self.automatic_hyphenation {
            write_toggle(&mut writer, "w:autoHyphenation", enabled)?;
        }
        if let Some(enabled) = self.even_and_odd_headers {
            write_toggle(&mut writer, "w:evenAndOddHeaders", enabled)?;
        }
        if let Some(value) = self.character_spacing_control {
            writer.get_mut().extend_from_slice(&write_valued_setting(
                "w:characterSpacingControl",
                value.as_str(),
            )?);
        }
        if let Some(enabled) = self.update_fields {
            write_toggle(&mut writer, "w:updateFields", enabled)?;
        }
        if !self.compatibility_settings.is_empty() {
            writer.write_event(Event::Start(BytesStart::new("w:compat")))?;
            writer
                .get_mut()
                .extend_from_slice(&write_compatibility_settings(&self.compatibility_settings)?);
            writer.write_event(Event::End(BytesEnd::new("w:compat")))?;
        }
        if !self.document_variables.is_empty() {
            writer.write_event(Event::Start(BytesStart::new("w:docVars")))?;
            writer
                .get_mut()
                .extend_from_slice(&write_document_variables(&self.document_variables)?);
            writer.write_event(Event::End(BytesEnd::new("w:docVars")))?;
        }
        if let Some(properties) = &self.math_properties {
            properties.write_xml(&mut writer)?;
        }
        if let Some(language) = &self.theme_font_language {
            writer
                .get_mut()
                .extend_from_slice(&write_theme_font_language(language)?);
        }
        writer.write_event(Event::End(BytesEnd::new("w:settings")))?;
        Ok(writer.into_inner())
    }
}

fn capture_empty_element(element: &BytesStart<'_>) -> Result<Vec<u8>> {
    let mut raw = Vec::new();
    Writer::new(&mut raw).write_event(Event::Empty(element.to_owned().into_owned()))?;
    Ok(raw)
}

fn parse_twips(element: &BytesStart<'_>, prefixes: &[String]) -> Option<Twips> {
    word_attribute(element, b"val", prefixes)
        .ok()
        .flatten()?
        .parse::<i32>()
        .ok()
        .filter(|value| *value >= 0)
        .map(Twips)
}

fn parse_character_spacing(
    element: &BytesStart<'_>,
    prefixes: &[String],
) -> Option<CharacterSpacingControl> {
    CharacterSpacingControl::parse(&word_attribute(element, b"val", prefixes).ok().flatten()?)
}

fn parse_theme_font_language(
    element: &BytesStart<'_>,
    prefixes: &[String],
) -> Result<ThemeFontLanguage> {
    Ok(ThemeFontLanguage {
        latin: word_attribute(element, b"val", prefixes)?,
        east_asia: word_attribute(element, b"eastAsia", prefixes)?,
        bidi: word_attribute(element, b"bidi", prefixes)?,
    })
}

fn parse_compatibility_setting(
    element: &BytesStart<'_>,
    prefixes: &[String],
) -> Option<CompatibilitySetting> {
    Some(CompatibilitySetting {
        name: word_attribute(element, b"name", prefixes).ok().flatten()?,
        uri: word_attribute(element, b"uri", prefixes).ok().flatten()?,
        value: word_attribute(element, b"val", prefixes).ok().flatten()?,
    })
}

fn write_valued_setting(name: &str, value: &str) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    let mut element = BytesStart::new(name);
    element.push_attribute(("w:val", value));
    writer.write_event(Event::Empty(element))?;
    Ok(writer.into_inner())
}

fn write_theme_font_language(language: &ThemeFontLanguage) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    let mut element = BytesStart::new("w:themeFontLang");
    if let Some(value) = &language.latin {
        element.push_attribute(("w:val", value.as_str()));
    }
    if let Some(value) = &language.east_asia {
        element.push_attribute(("w:eastAsia", value.as_str()));
    }
    if let Some(value) = &language.bidi {
        element.push_attribute(("w:bidi", value.as_str()));
    }
    writer.write_event(Event::Empty(element))?;
    Ok(writer.into_inner())
}

fn write_document_variables(variables: &[DocumentVariable]) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    for variable in variables {
        let mut element = BytesStart::new("w:docVar");
        element.push_attribute(("w:name", variable.name.as_str()));
        element.push_attribute(("w:val", variable.value.as_str()));
        writer.write_event(Event::Empty(element))?;
    }
    Ok(writer.into_inner())
}

fn write_compatibility_settings(settings: &[CompatibilitySetting]) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    for setting in settings {
        let mut element = BytesStart::new("w:compatSetting");
        element.push_attribute(("w:name", setting.name.as_str()));
        element.push_attribute(("w:uri", setting.uri.as_str()));
        element.push_attribute(("w:val", setting.value.as_str()));
        writer.write_event(Event::Empty(element))?;
    }
    Ok(writer.into_inner())
}

fn rewrite_group_setting(
    source: &[u8],
    group_local: &[u8],
    child_local: &[u8],
    modeled_children: Vec<u8>,
) -> Result<Vec<u8>> {
    let existing = find_top_level_settings(source, group_local)?;
    let replacement = if existing.is_empty() {
        if modeled_children.is_empty() {
            Vec::new()
        } else {
            let group_name = format!("w:{}", String::from_utf8_lossy(group_local));
            let mut writer = Writer::new(Vec::new());
            writer.write_event(Event::Start(BytesStart::new(&group_name)))?;
            writer.get_mut().extend_from_slice(&modeled_children);
            writer.write_event(Event::End(BytesEnd::new(&group_name)))?;
            writer.into_inner()
        }
    } else {
        let mut replacement = Vec::new();
        for (index, (raw, inherited_prefixes)) in existing.into_iter().enumerate() {
            replacement.extend_from_slice(&rewrite_group_children(
                &raw,
                child_local,
                if index == 0 { &modeled_children } else { &[] },
                &inherited_prefixes,
            )?);
        }
        replacement
    };
    rewrite_top_level_setting(source, group_local, &replacement)
}

fn find_top_level_settings(source: &[u8], local: &[u8]) -> Result<Vec<(Vec<u8>, Vec<String>)>> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut root_prefixes = Vec::new();
    let mut matches = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if depth == 0 => {
                root_prefixes = word_prefixes_at(&element, &[])?;
                depth = 1;
            }
            Event::Start(element) if depth == 1 => {
                let prefixes = word_prefixes_at(&element, &root_prefixes)?;
                if is_word_element(element.name().as_ref(), local, &prefixes) {
                    matches.push((
                        capture_element(&mut reader, &element)?,
                        root_prefixes.clone(),
                    ));
                } else {
                    depth += 1;
                }
            }
            Event::Empty(element) if depth == 1 => {
                let prefixes = word_prefixes_at(&element, &root_prefixes)?;
                if is_word_element(element.name().as_ref(), local, &prefixes) {
                    matches.push((capture_empty_element(&element)?, root_prefixes.clone()));
                }
            }
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => return Ok(matches),
            _ => {}
        }
        buffer.clear();
    }
}

fn rewrite_group_children(
    raw: &[u8],
    child_local: &[u8],
    modeled: &[u8],
    inherited_prefixes: &[String],
) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(raw.len() + modeled.len()));
    let mut prefixes = Vec::new();
    let mut depth = 0usize;
    let mut inserted = false;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if depth == 0 => {
                prefixes = word_prefixes_at(&element, inherited_prefixes)?;
                writer.write_event(Event::Start(element.into_owned()))?;
                depth = 1;
            }
            Event::Empty(element) if depth == 0 => {
                let name = std::str::from_utf8(element.name().as_ref())?.to_owned();
                writer.write_event(Event::Start(element.into_owned()))?;
                writer.get_mut().extend_from_slice(modeled);
                writer.write_event(Event::End(BytesEnd::new(name)))?;
                inserted = true;
            }
            Event::Start(element) if depth == 1 => {
                let child_prefixes = word_prefixes_at(&element, &prefixes)?;
                if is_word_element(element.name().as_ref(), child_local, &child_prefixes) {
                    if !inserted {
                        writer.get_mut().extend_from_slice(modeled);
                        inserted = true;
                    }
                    capture_element(&mut reader, &element)?;
                } else {
                    writer.write_event(Event::Start(element.into_owned()))?;
                    depth += 1;
                }
            }
            Event::Empty(element) if depth == 1 => {
                let child_prefixes = word_prefixes_at(&element, &prefixes)?;
                if is_word_element(element.name().as_ref(), child_local, &child_prefixes) {
                    if !inserted {
                        writer.get_mut().extend_from_slice(modeled);
                        inserted = true;
                    }
                } else {
                    writer.write_event(Event::Empty(element.into_owned()))?;
                }
            }
            Event::End(element) if depth == 1 => {
                if !inserted {
                    writer.get_mut().extend_from_slice(modeled);
                }
                writer.write_event(Event::End(element.into_owned()))?;
                depth = 0;
            }
            Event::Start(element) => {
                writer.write_event(Event::Start(element.into_owned()))?;
                depth += 1;
            }
            Event::End(element) => {
                writer.write_event(Event::End(element.into_owned()))?;
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            event => writer.write_event(event.into_owned())?,
        }
        buffer.clear();
    }
    Ok(writer.into_inner())
}

fn rewrite_top_level_setting(source: &[u8], local: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(source.len() + replacement.len()));
    let mut root_prefixes = Vec::new();
    let mut depth = 0usize;
    let mut inserted = false;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if depth == 0 => {
                root_prefixes = word_prefixes_at(&element, &[])?;
                let mut root = element.into_owned();
                ensure_fixed_word_prefix(&mut root)?;
                writer.write_event(Event::Start(root))?;
                depth = 1;
            }
            Event::Empty(element) if depth == 0 => {
                let root_name = std::str::from_utf8(element.name().as_ref())?.to_owned();
                let mut root = element.into_owned();
                ensure_fixed_word_prefix(&mut root)?;
                writer.write_event(Event::Start(root))?;
                writer.get_mut().extend_from_slice(replacement);
                writer.write_event(Event::End(BytesEnd::new(root_name)))?;
                inserted = true;
            }
            Event::Start(element) if depth == 1 => {
                let prefixes = word_prefixes_at(&element, &root_prefixes)?;
                if is_word_element(element.name().as_ref(), local, &prefixes) {
                    if !inserted {
                        writer.get_mut().extend_from_slice(replacement);
                        inserted = true;
                    }
                    capture_element(&mut reader, &element)?;
                } else {
                    if !inserted && setting_follows(local, &element, &prefixes) {
                        writer.get_mut().extend_from_slice(replacement);
                        inserted = true;
                    }
                    writer.write_event(Event::Start(element.into_owned()))?;
                    depth += 1;
                }
            }
            Event::Empty(element) if depth == 1 => {
                let prefixes = word_prefixes_at(&element, &root_prefixes)?;
                if is_word_element(element.name().as_ref(), local, &prefixes) {
                    if !inserted {
                        writer.get_mut().extend_from_slice(replacement);
                        inserted = true;
                    }
                } else {
                    if !inserted && setting_follows(local, &element, &prefixes) {
                        writer.get_mut().extend_from_slice(replacement);
                        inserted = true;
                    }
                    writer.write_event(Event::Empty(element.into_owned()))?;
                }
            }
            Event::End(element) if depth == 1 => {
                if !inserted {
                    writer.get_mut().extend_from_slice(replacement);
                }
                writer.write_event(Event::End(element.into_owned()))?;
                depth = 0;
            }
            Event::Start(element) => {
                writer.write_event(Event::Start(element.into_owned()))?;
                depth += 1;
            }
            Event::End(element) => {
                writer.write_event(Event::End(element.into_owned()))?;
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            event => writer.write_event(event.into_owned())?,
        }
        buffer.clear();
    }
    Ok(writer.into_inner())
}

fn setting_follows(local: &[u8], element: &BytesStart<'_>, prefixes: &[String]) -> bool {
    const ORDER: &[&[u8]] = &[
        b"writeProtection",
        b"view",
        b"zoom",
        b"removePersonalInformation",
        b"removeDateAndTime",
        b"doNotDisplayPageBoundaries",
        b"displayBackgroundShape",
        b"printPostScriptOverText",
        b"printFractionalCharacterWidth",
        b"printFormsData",
        b"embedTrueTypeFonts",
        b"embedSystemFonts",
        b"saveSubsetFonts",
        b"saveFormsData",
        b"mirrorMargins",
        b"alignBordersAndEdges",
        b"bordersDoNotSurroundHeader",
        b"bordersDoNotSurroundFooter",
        b"gutterAtTop",
        b"hideSpellingErrors",
        b"hideGrammaticalErrors",
        b"activeWritingStyle",
        b"proofState",
        b"formsDesign",
        b"attachedTemplate",
        b"linkStyles",
        b"stylePaneFormatFilter",
        b"stylePaneSortMethod",
        b"documentType",
        b"mailMerge",
        b"revisionView",
        b"trackRevisions",
        b"doNotTrackMoves",
        b"doNotTrackFormatting",
        b"documentProtection",
        b"autoFormatOverride",
        b"styleLockTheme",
        b"styleLockQFSet",
        b"defaultTabStop",
        b"autoHyphenation",
        b"consecutiveHyphenLimit",
        b"hyphenationZone",
        b"doNotHyphenateCaps",
        b"showEnvelope",
        b"summaryLength",
        b"clickAndTypeStyle",
        b"defaultTableStyle",
        b"evenAndOddHeaders",
        b"bookFoldRevPrinting",
        b"bookFoldPrinting",
        b"bookFoldPrintingSheets",
        b"drawingGridHorizontalSpacing",
        b"drawingGridVerticalSpacing",
        b"displayHorizontalDrawingGridEvery",
        b"displayVerticalDrawingGridEvery",
        b"doNotUseMarginsForDrawingGridOrigin",
        b"drawingGridHorizontalOrigin",
        b"drawingGridVerticalOrigin",
        b"doNotShadeFormData",
        b"noPunctuationKerning",
        b"characterSpacingControl",
        b"printTwoOnOne",
        b"strictFirstAndLastChars",
        b"noLineBreaksAfter",
        b"noLineBreaksBefore",
        b"savePreviewPicture",
        b"doNotValidateAgainstSchema",
        b"saveInvalidXml",
        b"ignoreMixedContent",
        b"alwaysShowPlaceholderText",
        b"doNotDemarcateInvalidXml",
        b"saveXmlDataOnly",
        b"useXSLTWhenSaving",
        b"saveThroughXslt",
        b"showXMLTags",
        b"alwaysMergeEmptyNamespace",
        b"updateFields",
        b"hdrShapeDefaults",
        b"footnotePr",
        b"endnotePr",
        b"compat",
        b"docVars",
        b"rsids",
        b"mathPr",
        b"attachedSchema",
        b"themeFontLang",
        b"clrSchemeMapping",
        b"doNotIncludeSubdocsInStats",
        b"doNotAutoCompressPictures",
        b"forceUpgrade",
        b"captions",
        b"readModeInkLockDown",
        b"smartTagType",
        b"schemaLibrary",
        b"shapeDefaults",
        b"doNotEmbedSmartTags",
        b"decimalSymbol",
        b"listSeparator",
    ];
    let Some(target_index) = ORDER.iter().position(|candidate| *candidate == local) else {
        return false;
    };
    ORDER
        .iter()
        .enumerate()
        .skip(target_index + 1)
        .any(|(_, candidate)| is_word_element(element.name().as_ref(), candidate, prefixes))
}

fn rewrite_math_properties(source: &[u8], properties: &MathProperties) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(source.len() + 128));
    let mut root_prefixes = Vec::new();
    let mut depth = 0usize;
    let mut inserted = false;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if depth == 0 => {
                root_prefixes = word_prefixes_at(&element, &[])?;
                let mut root = element.into_owned();
                ensure_fixed_word_prefix(&mut root)?;
                writer.write_event(Event::Start(root))?;
                depth = 1;
            }
            Event::Empty(element) if depth == 0 => {
                let root_name = std::str::from_utf8(element.name().as_ref())?.to_owned();
                let mut root = element.into_owned();
                ensure_fixed_word_prefix(&mut root)?;
                writer.write_event(Event::Start(root))?;
                properties.write_xml(&mut writer)?;
                writer.write_event(Event::End(BytesEnd::new(root_name)))?;
                inserted = true;
            }
            Event::Start(element) if depth == 1 => {
                let prefixes = word_prefixes_at(&element, &root_prefixes)?;
                if is_math_element(element.name().as_ref(), b"mathPr", &prefixes) {
                    if !inserted {
                        properties.write_xml(&mut writer)?;
                        inserted = true;
                    }
                    capture_element(&mut reader, &element)?;
                } else {
                    if !inserted && setting_follows_math_properties(&element, &prefixes) {
                        properties.write_xml(&mut writer)?;
                        inserted = true;
                    }
                    writer.write_event(Event::Start(element.into_owned()))?;
                    depth += 1;
                }
            }
            Event::Empty(element) if depth == 1 => {
                let prefixes = word_prefixes_at(&element, &root_prefixes)?;
                if is_math_element(element.name().as_ref(), b"mathPr", &prefixes) {
                    if !inserted {
                        properties.write_xml(&mut writer)?;
                        inserted = true;
                    }
                } else {
                    if !inserted && setting_follows_math_properties(&element, &prefixes) {
                        properties.write_xml(&mut writer)?;
                        inserted = true;
                    }
                    writer.write_event(Event::Empty(element.into_owned()))?;
                }
            }
            Event::End(element) if depth == 1 => {
                if !inserted {
                    properties.write_xml(&mut writer)?;
                }
                writer.write_event(Event::End(element.into_owned()))?;
                depth = 0;
            }
            Event::Start(element) => {
                writer.write_event(Event::Start(element.into_owned()))?;
                depth += 1;
            }
            Event::End(element) => {
                writer.write_event(Event::End(element.into_owned()))?;
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            event => writer.write_event(event.into_owned())?,
        }
        buffer.clear();
    }
    Ok(writer.into_inner())
}

fn setting_follows_math_properties(element: &BytesStart<'_>, prefixes: &[String]) -> bool {
    [
        b"attachedSchema".as_slice(),
        b"themeFontLang".as_slice(),
        b"clrSchemeMapping".as_slice(),
        b"doNotIncludeSubdocsInStats".as_slice(),
        b"doNotAutoCompressPictures".as_slice(),
        b"forceUpgrade".as_slice(),
        b"captions".as_slice(),
        b"readModeInkLockDown".as_slice(),
        b"smartTagType".as_slice(),
        b"schemaLibrary".as_slice(),
        b"shapeDefaults".as_slice(),
        b"doNotEmbedSmartTags".as_slice(),
        b"decimalSymbol".as_slice(),
        b"listSeparator".as_slice(),
    ]
    .iter()
    .any(|local| is_word_element(element.name().as_ref(), local, prefixes))
}

fn parse_toggle(element: &BytesStart<'_>, prefixes: &[String]) -> Result<Option<bool>> {
    Ok(match word_attribute(element, b"val", prefixes)? {
        Some(value) => parse_on_off(&value),
        None => Some(true),
    })
}

fn write_toggle(writer: &mut Writer<Vec<u8>>, name: &str, value: bool) -> Result<()> {
    let mut element = BytesStart::new(name);
    if !value {
        element.push_attribute(("w:val", "false"));
    }
    writer.write_event(Event::Empty(element))?;
    Ok(())
}

fn write_toggle_setting(name: &str, value: bool) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    write_toggle(&mut writer, name, value)?;
    Ok(writer.into_inner())
}

fn rewrite_automatic_hyphenation(source: &[u8], enabled: bool) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(source.len() + 64));
    let mut root_prefixes = Vec::new();
    let mut depth = 0usize;
    let mut inserted = false;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if depth == 0 => {
                root_prefixes = word_prefixes_at(&element, &[])?;
                let mut root = element.into_owned();
                ensure_fixed_word_prefix(&mut root)?;
                writer.write_event(Event::Start(root))?;
                depth = 1;
            }
            Event::Empty(element) if depth == 0 => {
                let root_name = std::str::from_utf8(element.name().as_ref())?.to_owned();
                let mut root = element.into_owned();
                ensure_fixed_word_prefix(&mut root)?;
                writer.write_event(Event::Start(root))?;
                write_toggle(&mut writer, "w:autoHyphenation", enabled)?;
                writer.write_event(Event::End(BytesEnd::new(root_name)))?;
                inserted = true;
            }
            Event::Start(element) if depth == 1 => {
                let prefixes = word_prefixes_at(&element, &root_prefixes)?;
                if is_word_element(element.name().as_ref(), b"autoHyphenation", &prefixes) {
                    if !inserted {
                        write_toggle(&mut writer, "w:autoHyphenation", enabled)?;
                        inserted = true;
                    }
                    capture_element(&mut reader, &element)?;
                } else {
                    if !inserted && setting_follows_auto_hyphenation(&element, &prefixes) {
                        write_toggle(&mut writer, "w:autoHyphenation", enabled)?;
                        inserted = true;
                    }
                    writer.write_event(Event::Start(element.into_owned()))?;
                    depth += 1;
                }
            }
            Event::Empty(element) if depth == 1 => {
                let prefixes = word_prefixes_at(&element, &root_prefixes)?;
                if is_word_element(element.name().as_ref(), b"autoHyphenation", &prefixes) {
                    if !inserted {
                        write_toggle(&mut writer, "w:autoHyphenation", enabled)?;
                        inserted = true;
                    }
                } else {
                    if !inserted && setting_follows_auto_hyphenation(&element, &prefixes) {
                        write_toggle(&mut writer, "w:autoHyphenation", enabled)?;
                        inserted = true;
                    }
                    writer.write_event(Event::Empty(element.into_owned()))?;
                }
            }
            Event::End(element) if depth == 1 => {
                if !inserted {
                    write_toggle(&mut writer, "w:autoHyphenation", enabled)?;
                }
                writer.write_event(Event::End(element.into_owned()))?;
                depth = 0;
            }
            Event::Start(element) => {
                writer.write_event(Event::Start(element.into_owned()))?;
                depth += 1;
            }
            Event::End(element) => {
                writer.write_event(Event::End(element.into_owned()))?;
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            event => writer.write_event(event.into_owned())?,
        }
        buffer.clear();
    }

    Ok(writer.into_inner())
}

fn ensure_fixed_word_prefix(root: &mut BytesStart<'_>) -> Result<()> {
    let mut fixed_binding = false;
    let mut conflicting_binding = false;
    for attribute in root.attributes() {
        let attribute = attribute?;
        if attribute.key.as_ref() == b"xmlns:w" {
            let value =
                attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, root.decoder())?;
            fixed_binding = value.as_bytes() == W_NS.as_bytes();
            conflicting_binding = !fixed_binding;
        }
    }
    if conflicting_binding {
        return Err(OxmlError::InvalidValue(
            "settings binds the reserved w prefix to a foreign namespace".to_owned(),
        ));
    }
    if !fixed_binding {
        root.push_attribute(("xmlns:w", W_NS));
    }
    Ok(())
}

fn setting_follows_auto_hyphenation(element: &BytesStart<'_>, prefixes: &[String]) -> bool {
    const FOLLOWING: &[&[u8]] = &[
        b"consecutiveHyphenLimit",
        b"hyphenationZone",
        b"doNotHyphenateCaps",
        b"showEnvelope",
        b"summaryLength",
        b"clickAndTypeStyle",
        b"defaultTableStyle",
        b"evenAndOddHeaders",
        b"bookFoldRevPrinting",
        b"bookFoldPrinting",
        b"bookFoldPrintingSheets",
        b"drawingGridHorizontalSpacing",
        b"drawingGridVerticalSpacing",
        b"displayHorizontalDrawingGridEvery",
        b"displayVerticalDrawingGridEvery",
        b"doNotUseMarginsForDrawingGridOrigin",
        b"drawingGridHorizontalOrigin",
        b"drawingGridVerticalOrigin",
        b"doNotShadeFormData",
        b"noPunctuationKerning",
        b"characterSpacingControl",
        b"printTwoOnOne",
        b"strictFirstAndLastChars",
        b"noLineBreaksAfter",
        b"noLineBreaksBefore",
        b"savePreviewPicture",
        b"doNotValidateAgainstSchema",
        b"saveInvalidXml",
        b"ignoreMixedContent",
        b"alwaysShowPlaceholderText",
        b"doNotDemarcateInvalidXml",
        b"saveXmlDataOnly",
        b"useXSLTWhenSaving",
        b"saveThroughXslt",
        b"showXMLTags",
        b"alwaysMergeEmptyNamespace",
        b"updateFields",
        b"hdrShapeDefaults",
        b"footnotePr",
        b"endnotePr",
        b"compat",
        b"docVars",
        b"rsids",
        b"mathPr",
        b"attachedSchema",
        b"themeFontLang",
        b"clrSchemeMapping",
        b"doNotIncludeSubdocsInStats",
        b"doNotAutoCompressPictures",
        b"forceUpgrade",
        b"captions",
        b"readModeInkLockDown",
        b"smartTagType",
        b"schemaLibrary",
        b"shapeDefaults",
        b"doNotEmbedSmartTags",
        b"decimalSymbol",
        b"listSeparator",
    ];
    FOLLOWING
        .iter()
        .any(|local| is_word_element(element.name().as_ref(), local, prefixes))
}

fn parse_document_variable(
    element: &BytesStart<'_>,
    prefixes: &[String],
) -> Option<DocumentVariable> {
    Some(DocumentVariable {
        name: word_attribute(element, b"name", prefixes).ok().flatten()?,
        value: word_attribute(element, b"val", prefixes).ok().flatten()?,
    })
}

fn parse_document_protection(
    element: &BytesStart<'_>,
    prefixes: &[String],
) -> Option<DocumentProtection> {
    let value = |name| word_attribute(element, name, prefixes).ok().flatten();
    let mode = ProtectionMode::parse(&value(b"edit")?)?;
    let enforcement = match value(b"enforcement") {
        Some(value) => Some(parse_on_off(&value)?),
        None => None,
    };
    let formatting = match value(b"formatting") {
        Some(value) => Some(parse_on_off(&value)?),
        None => None,
    };
    let provider_type = match value(b"cryptProviderType") {
        Some(value) => Some(CryptProviderType::parse(&value)?),
        None => None,
    };
    let algorithm_class = match value(b"cryptAlgorithmClass") {
        Some(value) => Some(CryptAlgorithmClass::parse(&value)?),
        None => None,
    };
    let algorithm_type = match value(b"cryptAlgorithmType") {
        Some(value) => Some(CryptAlgorithmType::parse(&value)?),
        None => None,
    };
    let algorithm_sid = match value(b"cryptAlgorithmSid") {
        Some(value) => Some(value.parse::<u32>().ok()?),
        None => None,
    };
    let spin_count = match value(b"cryptSpinCount") {
        Some(value) => Some(value.parse::<u32>().ok()?),
        None => None,
    };
    Some(DocumentProtection {
        mode,
        enforcement,
        formatting,
        provider_type,
        algorithm_class,
        algorithm_type,
        algorithm_sid,
        spin_count,
        hash: value(b"hash"),
        salt: value(b"salt"),
    })
}

fn word_attribute(
    element: &BytesStart<'_>,
    local: &[u8],
    prefixes: &[String],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        if is_word_attribute(attribute.key.as_ref(), local, prefixes) {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn parse_on_off(value: &str) -> Option<bool> {
    match value {
        "1" | "true" | "on" => Some(true),
        "0" | "false" | "off" => Some(false),
        _ => None,
    }
}

fn write_document_protection(
    writer: &mut Writer<Vec<u8>>,
    protection: &DocumentProtection,
) -> Result<()> {
    let mut element = BytesStart::new("w:documentProtection");
    element.push_attribute(("w:edit", protection.mode.as_str()));
    if let Some(formatting) = protection.formatting {
        element.push_attribute(("w:formatting", if formatting { "1" } else { "0" }));
    }
    if let Some(enforcement) = protection.enforcement {
        element.push_attribute(("w:enforcement", if enforcement { "1" } else { "0" }));
    }
    if let Some(provider_type) = protection.provider_type {
        element.push_attribute(("w:cryptProviderType", provider_type.as_str()));
    }
    if let Some(algorithm_class) = protection.algorithm_class {
        element.push_attribute(("w:cryptAlgorithmClass", algorithm_class.as_str()));
    }
    if let Some(algorithm_type) = protection.algorithm_type {
        element.push_attribute(("w:cryptAlgorithmType", algorithm_type.as_str()));
    }
    let algorithm_sid = protection.algorithm_sid.map(|value| value.to_string());
    if let Some(value) = &algorithm_sid {
        element.push_attribute(("w:cryptAlgorithmSid", value.as_str()));
    }
    let spin_count = protection.spin_count.map(|value| value.to_string());
    if let Some(value) = &spin_count {
        element.push_attribute(("w:cryptSpinCount", value.as_str()));
    }
    if let Some(value) = &protection.hash {
        element.push_attribute(("w:hash", value.as_str()));
    }
    if let Some(value) = &protection.salt {
        element.push_attribute(("w:salt", value.as_str()));
    }
    writer.write_event(Event::Empty(element))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn settings(mode: &str, enforcement: &str, formatting: &str) -> Vec<u8> {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:settings xmlns:x="{W_NS}" xmlns:p="urn:producer" p:root="kept"><p:before p:v="1"/><x:documentProtection x:edit="{mode}" x:enforcement="{enforcement}" x:formatting="{formatting}" x:cryptProviderType="rsaAES" x:cryptAlgorithmClass="hash" x:cryptAlgorithmType="typeAny" x:cryptAlgorithmSid="14" x:cryptSpinCount="100000" x:hash="HASH-{mode}" x:salt="SALT-{mode}"/><p:after>keep me</p:after></x:settings>"#
        )
        .into_bytes()
    }

    #[test]
    fn document_protection_modes_and_metadata_parse_through_aliases() {
        for (name, expected) in [
            ("readOnly", ProtectionMode::ReadOnly),
            ("comments", ProtectionMode::Comments),
            ("trackedChanges", ProtectionMode::TrackedChanges),
            ("forms", ProtectionMode::Forms),
        ] {
            let parsed = CT_Settings::from_xml(&settings(name, "true", "0")).unwrap();
            let protection = parsed.document_protection().unwrap();
            assert_eq!(protection.mode, expected);
            assert_eq!(protection.enforcement, Some(true));
            assert_eq!(protection.formatting, Some(false));
            assert_eq!(protection.provider_type, Some(CryptProviderType::RsaAes));
            assert_eq!(protection.algorithm_class, Some(CryptAlgorithmClass::Hash));
            assert_eq!(protection.algorithm_type, Some(CryptAlgorithmType::Any));
            assert_eq!(protection.algorithm_sid, Some(14));
            assert_eq!(protection.spin_count, Some(100_000));
            assert_eq!(
                protection.hash.as_deref(),
                Some(format!("HASH-{name}").as_str())
            );
            assert_eq!(
                protection.salt.as_deref(),
                Some(format!("SALT-{name}").as_str())
            );
        }

        let false_and_on = CT_Settings::from_xml(&settings("forms", "false", "on")).unwrap();
        let protection = false_and_on.document_protection().unwrap();
        assert_eq!(protection.enforcement, Some(false));
        assert_eq!(protection.formatting, Some(true));
    }

    #[test]
    fn settings_keep_document_protection_and_unmodelled_children_byte_identical() {
        for mode in ["readOnly", "comments", "trackedChanges", "forms"] {
            let xml = settings(mode, "1", "off");
            let parsed = CT_Settings::from_xml(&xml).unwrap();
            assert_eq!(parsed.to_xml().unwrap(), xml);
        }
    }

    #[test]
    fn document_variables_are_alias_safe_and_leave_settings_bytes_unchanged() {
        let xml = format!(
            r#"<?xml version="1.0"?><q:settings xmlns:q="{W_NS}" xmlns:p="urn:producer"><q:docVars xmlns:v="{W_NS}"><v:docVar v:name="Customer" v:val="Ada"/><v:docVar v:name="Region" v:val="West"></v:docVar><p:docVar p:name="Foreign" p:val="ignored"/><v:docVar v:name="Malformed"/></q:docVars><p:after/></q:settings>"#
        )
        .into_bytes();
        let settings = CT_Settings::from_xml(&xml).unwrap();
        assert_eq!(
            settings.document_variables(),
            [
                DocumentVariable {
                    name: "Customer".to_owned(),
                    value: "Ada".to_owned(),
                },
                DocumentVariable {
                    name: "Region".to_owned(),
                    value: "West".to_owned(),
                },
            ]
        );
        assert_eq!(settings.to_xml().unwrap(), xml);
    }

    #[test]
    fn constructed_settings_use_fixed_prefix_and_schema_order() {
        let settings = CT_Settings {
            document_protection: Some(DocumentProtection {
                mode: ProtectionMode::ReadOnly,
                enforcement: Some(true),
                formatting: Some(false),
                provider_type: Some(CryptProviderType::RsaAes),
                algorithm_class: Some(CryptAlgorithmClass::Hash),
                algorithm_type: Some(CryptAlgorithmType::Any),
                algorithm_sid: Some(14),
                spin_count: Some(100_000),
                hash: Some("HASH".to_owned()),
                salt: Some("SALT".to_owned()),
            }),
            document_variables: Vec::new(),
            compatibility_settings: Vec::new(),
            default_tab_stop: None,
            character_spacing_control: None,
            theme_font_language: None,
            automatic_hyphenation: None,
            even_and_odd_headers: None,
            update_fields: None,
            math_properties: None,
            source_xml: None,
        };
        let xml = String::from_utf8(settings.to_xml().unwrap()).unwrap();
        assert!(xml.contains("<w:settings xmlns:w="));
        assert!(xml.contains("<w:documentProtection w:edit=\"readOnly\""));
        assert!(xml.find("w:formatting").unwrap() < xml.find("w:enforcement").unwrap());
        assert!(xml.find("w:cryptAlgorithmSid").unwrap() < xml.find("w:cryptSpinCount").unwrap());
        assert!(xml.find("w:cryptSpinCount").unwrap() < xml.find("w:hash").unwrap());
    }

    #[test]
    fn automatic_hyphenation_defaults_off_and_parses_only_word_settings() {
        let omitted =
            CT_Settings::from_xml(format!(r#"<w:settings xmlns:w="{W_NS}"/>"#).as_bytes()).unwrap();
        assert!(!omitted.automatic_hyphenation());

        let xml = format!(
            r#"<q:settings xmlns:q="{W_NS}" xmlns:x="urn:foreign"><x:autoHyphenation/><q:autoHyphenation q:val="on"/></q:settings>"#,
        );
        let parsed = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        assert!(parsed.automatic_hyphenation());
        assert_eq!(parsed.to_xml().unwrap(), xml.as_bytes());
    }

    #[test]
    fn authored_automatic_hyphenation_uses_schema_order_and_preserves_raw_children() {
        let xml = format!(
            r#"<q:settings xmlns:q="{W_NS}" xmlns:x="urn:foreign"><q:defaultTabStop q:val="720"/><x:kept x:value="raw"/><q:consecutiveHyphenLimit q:val="2"/></q:settings>"#,
        );
        let mut parsed = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        parsed.set_automatic_hyphenation(true).unwrap();
        let output = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"<x:kept x:value="raw"/>"#));
        assert!(output.contains("<w:autoHyphenation/>"));
        assert!(output.find("defaultTabStop").unwrap() < output.find("autoHyphenation").unwrap());
        assert!(
            output.find("autoHyphenation").unwrap()
                < output.find("consecutiveHyphenLimit").unwrap()
        );
    }

    #[test]
    fn authored_automatic_hyphenation_expands_a_self_closing_settings_root() {
        let xml = format!(r#"<q:settings xmlns:q="{W_NS}"/>"#);
        let mut parsed = CT_Settings::from_xml(xml.as_bytes()).unwrap();

        parsed.set_automatic_hyphenation(true).unwrap();

        let output = parsed.to_xml().unwrap();
        assert!(
            std::str::from_utf8(&output)
                .unwrap()
                .contains("<w:autoHyphenation/>")
        );
        assert!(
            CT_Settings::from_xml(&output)
                .unwrap()
                .automatic_hyphenation()
        );
    }

    #[test]
    fn even_and_odd_headers_are_alias_safe_and_rewrite_in_schema_order() {
        let xml = format!(
            r#"<q:settings xmlns:q="{W_NS}" xmlns:x="urn:foreign"><q:defaultTableStyle q:val="TableNormal"/><x:evenAndOddHeaders/><q:evenAndOddHeaders q:val="off"/><q:bookFoldPrinting/><x:kept/></q:settings>"#,
        );
        let mut settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        assert!(!settings.even_and_odd_headers());
        settings.set_even_and_odd_headers(true).unwrap();
        let output = String::from_utf8(settings.to_xml().unwrap()).unwrap();
        assert!(output.contains("<w:evenAndOddHeaders/>"));
        assert!(output.contains("<x:evenAndOddHeaders/>"));
        assert!(output.contains("<x:kept/>"));
        assert!(
            output.find("defaultTableStyle").unwrap()
                < output.find("<w:evenAndOddHeaders").unwrap()
        );
        assert!(
            output.find("<w:evenAndOddHeaders").unwrap() < output.find("bookFoldPrinting").unwrap()
        );
        let reopened = CT_Settings::from_xml(output.as_bytes()).unwrap();
        assert!(reopened.even_and_odd_headers());
    }

    #[test]
    fn update_fields_toggle_reads_sets_and_removes_in_schema_order() {
        let xml = format!(
            r#"<q:settings xmlns:q="{W_NS}" xmlns:x="urn:foreign"><q:characterSpacingControl q:val="doNotCompress"/><q:updateFields q:val="true"/><x:updateFields/><x:kept/><q:compat/></q:settings>"#,
        );
        let mut settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(settings.update_fields(), Some(true));

        settings.set_update_fields(Some(false)).unwrap();
        let output = String::from_utf8(settings.to_xml().unwrap()).unwrap();
        assert!(
            output.contains(r#"<w:updateFields w:val="false"/>"#),
            "{output}"
        );
        assert!(!output.contains("<q:updateFields"), "{output}");
        assert!(output.contains("<x:updateFields/>"), "{output}");
        assert!(output.contains("<x:kept/>"), "{output}");
        let reopened = CT_Settings::from_xml(output.as_bytes()).unwrap();
        assert_eq!(reopened.update_fields(), Some(false));

        settings.set_update_fields(None).unwrap();
        let output = String::from_utf8(settings.to_xml().unwrap()).unwrap();
        assert!(!output.contains("<w:updateFields"), "{output}");
        assert!(output.contains("<x:updateFields/>"), "{output}");
        let reopened = CT_Settings::from_xml(output.as_bytes()).unwrap();
        assert_eq!(reopened.update_fields(), None);

        // A part without the toggle receives it at its schema position.
        let xml = format!(
            r#"<w:settings xmlns:w="{W_NS}"><w:characterSpacingControl w:val="doNotCompress"/><w:compat/></w:settings>"#,
        );
        let mut settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        settings.set_update_fields(Some(true)).unwrap();
        let output = String::from_utf8(settings.to_xml().unwrap()).unwrap();
        let position = output.find("<w:updateFields/>").unwrap();
        assert!(
            output.find("characterSpacingControl").unwrap() < position,
            "{output}"
        );
        assert!(position < output.find("<w:compat").unwrap(), "{output}");

        let mut authored = CT_Settings::new();
        authored.set_update_fields(Some(true)).unwrap();
        let output = String::from_utf8(authored.to_xml().unwrap()).unwrap();
        assert!(output.contains("<w:updateFields/>"), "{output}");
    }

    #[test]
    fn math_properties_accept_aliases_and_replace_in_schema_order() {
        let xml = format!(
            r#"<q:settings xmlns:q="{W_NS}" xmlns:z="{}" xmlns:x="urn:producer"><q:rsids/><z:mathPr x:keep="yes"><z:mathFont z:val="Cambria Math"/><x:inside/></z:mathPr><x:outside/><q:attachedSchema q:val="urn:test"/></q:settings>"#,
            crate::namespace::M_NS,
        );
        let mut settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(
            settings.math_properties().unwrap().math_font.as_deref(),
            Some("Cambria Math")
        );
        let mut properties = settings.math_properties().unwrap().clone();
        properties.math_font = Some("STIX Two Math".to_owned());
        properties.justification = Some(crate::math::MathJustification::CenterGroup);
        settings.set_math_properties(properties).unwrap();

        let output = String::from_utf8(settings.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"x:keep="yes""#));
        assert!(output.contains("<x:inside/>"));
        assert!(output.contains("<x:outside/>"));
        assert!(output.contains(r#"<m:mathFont m:val="STIX Two Math"/>"#));
        assert!(output.find("<q:rsids").unwrap() < output.find("<m:mathPr").unwrap());
        assert!(output.find("<m:mathPr").unwrap() < output.find("<q:attachedSchema").unwrap());
        assert_eq!(
            CT_Settings::from_xml(output.as_bytes())
                .unwrap()
                .math_properties()
                .unwrap()
                .justification,
            Some(crate::math::MathJustification::CenterGroup)
        );
    }

    #[test]
    fn math_properties_with_a_conflicting_m_binding_remain_untyped() {
        let xml = format!(
            r#"<w:settings xmlns:w="{W_NS}" xmlns:q="{}" xmlns:m="urn:producer"><q:mathPr><m:opaque/></q:mathPr></w:settings>"#,
            crate::namespace::M_NS,
        );
        let settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        assert!(settings.math_properties().is_none());
        assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());
    }
}
