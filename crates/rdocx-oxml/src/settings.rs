//! Word document settings and document-protection metadata.

use oxml_core::Twips;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::document::{CT_NoteProperties, CT_SectPr};
use crate::error::{OxmlError, Result};
use crate::math::{MathProperties, fixed_math_prefix_is_safe, is_math_element};
use crate::namespace::W_NS;
use crate::numbering::{namespace_bindings, word_prefixes_at};
use crate::properties::{is_word_attribute, is_word_element, parse_integer_measurement};
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

/// One valid `w:documentProtection` element.
///
/// Authoring records caller-supplied protection metadata verbatim. Deriving a
/// hash from a password, choosing a salt and picking a spin count are explicit
/// non-goals, so nothing here computes a credential.
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

/// The document view Word opens with, from `w:view`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentView {
    None,
    Print,
    Outline,
    MasterPages,
    Normal,
    Web,
}

impl DocumentView {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "none" => Some(Self::None),
            "print" => Some(Self::Print),
            "outline" => Some(Self::Outline),
            "masterPages" => Some(Self::MasterPages),
            "normal" => Some(Self::Normal),
            "web" => Some(Self::Web),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Print => "print",
            Self::Outline => "outline",
            Self::MasterPages => "masterPages",
            Self::Normal => "normal",
            Self::Web => "web",
        }
    }
}

/// The magnification preset recorded by `w:zoom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ZoomKind {
    None,
    FullPage,
    BestFit,
    TextFit,
}

impl ZoomKind {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "none" => Some(Self::None),
            "fullPage" => Some(Self::FullPage),
            "bestFit" => Some(Self::BestFit),
            "textFit" => Some(Self::TextFit),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::FullPage => "fullPage",
            Self::BestFit => "bestFit",
            Self::TextFit => "textFit",
        }
    }
}

/// The `w:zoom` magnification preset and its percentage.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DocumentZoom {
    pub kind: Option<ZoomKind>,
    pub percent: Option<u32>,
}

/// Whether Word considers a proofing pass current, from `w:proofState`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProofState {
    Clean,
    Dirty,
}

impl ProofState {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "clean" => Some(Self::Clean),
            "dirty" => Some(Self::Dirty),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Clean => "clean",
            Self::Dirty => "dirty",
        }
    }
}

/// The spelling and grammar proofing state recorded by `w:proofState`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DocumentProofState {
    pub spelling: Option<ProofState>,
    pub grammar: Option<ProofState>,
}

/// The mail-merge output category from `w:mainDocumentType`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeDocumentType {
    Catalog,
    Envelopes,
    MailingLabels,
    FormLetters,
    Email,
    Fax,
}

impl MailMergeDocumentType {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "catalog" => Some(Self::Catalog),
            "envelopes" => Some(Self::Envelopes),
            "mailingLabels" => Some(Self::MailingLabels),
            "formLetters" => Some(Self::FormLetters),
            "email" => Some(Self::Email),
            "fax" => Some(Self::Fax),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Catalog => "catalog",
            Self::Envelopes => "envelopes",
            Self::MailingLabels => "mailingLabels",
            Self::FormLetters => "formLetters",
            Self::Email => "email",
            Self::Fax => "fax",
        }
    }
}

/// The mail-merge destination from `w:destination`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeDestination {
    NewDocument,
    Printer,
    Email,
    Fax,
}

impl MailMergeDestination {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "newDocument" => Some(Self::NewDocument),
            "printer" => Some(Self::Printer),
            "email" => Some(Self::Email),
            "fax" => Some(Self::Fax),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::NewDocument => "newDocument",
            Self::Printer => "printer",
            Self::Email => "email",
            Self::Fax => "fax",
        }
    }
}

/// The bounded `w:mailMerge` subset this model authors.
///
/// `w:dataSource`, `w:headerSource` and `w:odso` stay preservation-only. The
/// first two carry `r:id` values whose rebasing belongs to the relationship
/// layer, and `w:odso` is a nested tree no story asked for.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MailMerge {
    pub main_document_type: Option<MailMergeDocumentType>,
    pub link_to_query: Option<bool>,
    pub data_type: Option<String>,
    pub connect_string: Option<String>,
    pub query: Option<String>,
    pub do_not_suppress_blank_lines: Option<bool>,
    pub destination: Option<MailMergeDestination>,
    pub address_field_name: Option<String>,
    pub mail_subject: Option<String>,
    pub mail_as_attachment: Option<bool>,
    pub view_merged_data: Option<bool>,
    pub active_record: Option<i32>,
    pub check_errors: Option<i32>,
}

/// Why a supported settings child produced no typed value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SettingsDiagnosticReason {
    /// More than one occurrence, so no single occurrence owns the value.
    Duplicated,
    /// One occurrence whose attributes did not parse to a typed value.
    Malformed,
}

/// One supported settings child the typed model could not own.
///
/// A name outside the closed supported set is never a diagnostic, which is what
/// makes "no unmodeled supported children" decidable rather than open-ended.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SettingsDiagnostic {
    /// The path from the part root, such as `["compat", "compatSetting"]`.
    pub path: Vec<&'static str>,
    pub occurrences: usize,
    pub reason: SettingsDiagnosticReason,
}

/// Occurrence counts for the closed supported set.
///
/// Parsing fills it, every setter and remover keeps it current, and both the
/// diagnostics and the removal ambiguity guard read it. That is one place to
/// look instead of one counter field per modeled child.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub(crate) struct Tally {
    entries: Vec<TallyEntry>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct TallyEntry {
    path: Vec<&'static str>,
    total: usize,
    typed: usize,
}

impl Tally {
    /// Count one occurrence and report the running total for that path.
    pub(crate) fn record(&mut self, path: &[&'static str], typed: bool) -> usize {
        let index = match self.entries.iter().position(|entry| entry.path == path) {
            Some(index) => index,
            None => {
                self.entries.push(TallyEntry {
                    path: path.to_vec(),
                    total: 0,
                    typed: 0,
                });
                self.entries.len() - 1
            }
        };
        let entry = &mut self.entries[index];
        entry.total += 1;
        entry.typed += usize::from(typed);
        entry.total
    }

    /// Replace the tally for one path with a single well-formed occurrence, or
    /// with no occurrence at all.
    pub(crate) fn set(&mut self, path: &[&'static str], present: bool) {
        let count = usize::from(present);
        match self.entries.iter_mut().find(|entry| entry.path == path) {
            Some(entry) => {
                entry.total = count;
                entry.typed = count;
            }
            None => {
                if present {
                    self.entries.push(TallyEntry {
                        path: path.to_vec(),
                        total: 1,
                        typed: 1,
                    });
                }
            }
        }
    }

    pub(crate) fn total(&self, path: &[&str]) -> usize {
        self.entries
            .iter()
            .find(|entry| entry.path == path)
            .map_or(0, |entry| entry.total)
    }

    pub(crate) fn typed(&self, path: &[&str]) -> usize {
        self.entries
            .iter()
            .find(|entry| entry.path == path)
            .map_or(0, |entry| entry.typed)
    }

    /// Report every supported child the typed model could not own.
    ///
    /// `repeatable` names the paths whose schema allows more than one
    /// occurrence, so repetition there is not a defect.
    pub(crate) fn diagnostics(&self, repeatable: &[&[&str]]) -> Vec<SettingsDiagnostic> {
        let mut diagnostics = Vec::new();
        for entry in &self.entries {
            let repeats = repeatable.contains(&entry.path.as_slice());
            if !repeats && entry.total > 1 {
                diagnostics.push(SettingsDiagnostic {
                    path: entry.path.clone(),
                    occurrences: entry.total,
                    reason: SettingsDiagnosticReason::Duplicated,
                });
            } else if entry.typed < entry.total {
                diagnostics.push(SettingsDiagnostic {
                    path: entry.path.clone(),
                    occurrences: entry.total - entry.typed,
                    reason: SettingsDiagnosticReason::Malformed,
                });
            }
        }
        diagnostics
    }
}

/// Assign a parsed member only while exactly one occurrence owns it.
///
/// A second occurrence clears the slot and keeps it clear, because no single
/// occurrence can then be called the value.
pub(crate) fn assign<T>(slot: &mut Option<T>, value: Option<T>, total: usize) {
    *slot = if total == 1 { value } else { None };
}

/// The typed contents of a Word settings part.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CT_Settings {
    view: Option<DocumentView>,
    zoom: Option<DocumentZoom>,
    remove_personal_information: Option<bool>,
    remove_date_and_time: Option<bool>,
    display_background_shape: Option<bool>,
    mirror_margins: Option<bool>,
    gutter_at_top: Option<bool>,
    proof_state: Option<DocumentProofState>,
    link_styles: Option<bool>,
    mail_merge: Option<MailMerge>,
    track_revisions: Option<bool>,
    do_not_track_moves: Option<bool>,
    do_not_track_formatting: Option<bool>,
    document_protection: Option<DocumentProtection>,
    footnote_properties: Option<CT_NoteProperties>,
    endnote_properties: Option<CT_NoteProperties>,
    default_tab_stop: Option<Twips>,
    automatic_hyphenation: Option<bool>,
    consecutive_hyphen_limit: Option<i32>,
    hyphenation_zone: Option<Twips>,
    do_not_hyphenate_caps: Option<bool>,
    default_table_style: Option<String>,
    even_and_odd_headers: Option<bool>,
    book_fold_rev_printing: Option<bool>,
    book_fold_printing: Option<bool>,
    book_fold_printing_sheets: Option<i32>,
    character_spacing_control: Option<CharacterSpacingControl>,
    update_fields: Option<bool>,
    compatibility_options: Vec<(CompatibilityOption, bool)>,
    compatibility_settings: Vec<CompatibilitySetting>,
    document_variables: Vec<DocumentVariable>,
    math_properties: Option<MathProperties>,
    theme_font_language: Option<ThemeFontLanguage>,
    decimal_symbol: Option<String>,
    list_separator: Option<String>,
    tally: Tally,
    diagnostics: Vec<SettingsDiagnostic>,
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
        let mut model = Self::default();
        let mut root_prefixes = Vec::new();
        let mut group: Option<(&'static str, usize, Vec<String>)> = None;
        let mut mail_merge = MailMerge::default();
        let mut saw_root = false;
        let mut depth = 0usize;
        let mut buffer = Vec::new();

        loop {
            let event = reader.read_event_into(&mut buffer)?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    let started = matches!(event, Event::Start(_));
                    let inherited = match &group {
                        Some((_, _, prefixes)) => prefixes,
                        None => &root_prefixes,
                    };
                    let prefixes = word_prefixes_at(element, inherited)?;
                    if !saw_root {
                        if !is_word_element(element.name().as_ref(), b"settings", &prefixes) {
                            return Err(OxmlError::MissingElement("settings root".to_owned()));
                        }
                        root_prefixes = prefixes;
                        saw_root = true;
                        depth = usize::from(started);
                        buffer.clear();
                        continue;
                    }
                    if depth == 1 && is_math_element(element.name().as_ref(), b"mathPr", &prefixes)
                    {
                        let raw = if started {
                            capture_element(&mut reader, element)?
                        } else {
                            capture_empty_element(element)?
                        };
                        let bindings = namespace_bindings(&prefixes);
                        let parsed = if fixed_math_prefix_is_safe(&raw, &bindings)? {
                            Some(MathProperties::from_raw(&raw, &bindings)?)
                        } else {
                            None
                        };
                        let total = model.tally.record(&["mathPr"], parsed.is_some());
                        assign(&mut model.math_properties, parsed, total);
                        buffer.clear();
                        continue;
                    }
                    if depth == 1
                        && (is_word_element(element.name().as_ref(), b"footnotePr", &prefixes)
                            || is_word_element(element.name().as_ref(), b"endnotePr", &prefixes))
                    {
                        let footnote =
                            is_word_element(element.name().as_ref(), b"footnotePr", &prefixes);
                        let name = if footnote { "footnotePr" } else { "endnotePr" };
                        let properties = if started {
                            CT_SectPr::parse_note_properties(
                                &mut reader,
                                &prefixes,
                                &[],
                                name.as_bytes(),
                            )?
                        } else {
                            CT_NoteProperties::default()
                        };
                        let total = model.tally.record(&[name], true);
                        let slot = if footnote {
                            &mut model.footnote_properties
                        } else {
                            &mut model.endnote_properties
                        };
                        assign(slot, Some(properties), total);
                        buffer.clear();
                        continue;
                    }
                    if depth == 1 {
                        model.absorb_top_level(element, &prefixes, &mut group, depth)?;
                        if !started {
                            // A self-closing group has no children, so its
                            // namespace scope must not leak onto its siblings.
                            group = None;
                        }
                    } else if let Some((name, child_depth, _)) = &group
                        && *child_depth == depth
                    {
                        let name = *name;
                        model.absorb_group_child(name, element, &prefixes, &mut mail_merge)?;
                    }
                    if started {
                        depth += 1;
                    }
                }
                Event::End(_) if depth > 0 => {
                    if let Some((_, child_depth, _)) = &group
                        && *child_depth == depth
                    {
                        group = None;
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
        if model.tally.total(&["mailMerge"]) == 1 {
            model.mail_merge = Some(mail_merge);
        }
        model.compatibility_options.sort_by_key(|entry| entry.0);
        model.diagnostics = model.tally.diagnostics(SETTINGS_REPEATABLE);
        model.source_xml = Some(xml.to_vec());
        Ok(model)
    }

    /// Project one top-level settings child into its typed member.
    fn absorb_top_level(
        &mut self,
        element: &BytesStart<'_>,
        prefixes: &[String],
        group: &mut Option<(&'static str, usize, Vec<String>)>,
        depth: usize,
    ) -> Result<()> {
        let Some(name) = supported_name(SUPPORTED_SETTINGS, element, prefixes) else {
            return Ok(());
        };
        match name {
            "view" => {
                let value = word_value(element, prefixes).and_then(|v| DocumentView::parse(&v));
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.view, value, total);
            }
            "zoom" => {
                let value = parse_zoom(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.zoom, value, total);
            }
            "removePersonalInformation" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.remove_personal_information, value, total);
            }
            "removeDateAndTime" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.remove_date_and_time, value, total);
            }
            "displayBackgroundShape" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.display_background_shape, value, total);
            }
            "mirrorMargins" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.mirror_margins, value, total);
            }
            "gutterAtTop" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.gutter_at_top, value, total);
            }
            "proofState" => {
                let value = parse_proof_state(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.proof_state, value, total);
            }
            "linkStyles" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.link_styles, value, total);
            }
            "mailMerge" => {
                self.tally.record(&[name], true);
                *group = Some(("mailMerge", depth + 1, prefixes.to_vec()));
            }
            "trackRevisions" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.track_revisions, value, total);
            }
            "doNotTrackMoves" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.do_not_track_moves, value, total);
            }
            "doNotTrackFormatting" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.do_not_track_formatting, value, total);
            }
            "documentProtection" => {
                let value = parse_document_protection(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.document_protection, value, total);
            }
            "defaultTabStop" => {
                let value = parse_twips(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.default_tab_stop, value, total);
            }
            "autoHyphenation" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.automatic_hyphenation, value, total);
            }
            "consecutiveHyphenLimit" => {
                let value = parse_decimal(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.consecutive_hyphen_limit, value, total);
            }
            "hyphenationZone" => {
                let value = parse_twips(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.hyphenation_zone, value, total);
            }
            "doNotHyphenateCaps" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.do_not_hyphenate_caps, value, total);
            }
            "defaultTableStyle" => {
                let value = word_value(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.default_table_style, value, total);
            }
            "evenAndOddHeaders" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.even_and_odd_headers, value, total);
            }
            "bookFoldRevPrinting" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.book_fold_rev_printing, value, total);
            }
            "bookFoldPrinting" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.book_fold_printing, value, total);
            }
            "bookFoldPrintingSheets" => {
                let value = parse_decimal(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.book_fold_printing_sheets, value, total);
            }
            "characterSpacingControl" => {
                let value =
                    word_value(element, prefixes).and_then(|v| CharacterSpacingControl::parse(&v));
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.character_spacing_control, value, total);
            }
            "updateFields" => {
                let value = parse_toggle(element, prefixes)?;
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.update_fields, value, total);
            }
            "compat" => {
                self.tally.record(&[name], true);
                *group = Some(("compat", depth + 1, prefixes.to_vec()));
            }
            "docVars" => {
                self.tally.record(&[name], true);
                *group = Some(("docVars", depth + 1, prefixes.to_vec()));
            }
            "themeFontLang" => {
                let value = Some(parse_theme_font_language(element, prefixes)?);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.theme_font_language, value, total);
            }
            "decimalSymbol" => {
                let value = word_value(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.decimal_symbol, value, total);
            }
            "listSeparator" => {
                let value = word_value(element, prefixes);
                let total = self.tally.record(&[name], value.is_some());
                assign(&mut self.list_separator, value, total);
            }
            // `mathPr` lives in the OfficeMath namespace and is matched before
            // the supported-name lookup runs, so it never reaches this match.
            _ => {}
        }
        Ok(())
    }

    /// Project one child of `w:compat`, `w:docVars` or `w:mailMerge`.
    fn absorb_group_child(
        &mut self,
        group: &str,
        element: &BytesStart<'_>,
        prefixes: &[String],
        mail_merge: &mut MailMerge,
    ) -> Result<()> {
        match group {
            "docVars" => {
                if is_word_element(element.name().as_ref(), b"docVar", prefixes) {
                    let variable = parse_document_variable(element, prefixes);
                    self.tally
                        .record(&["docVars", "docVar"], variable.is_some());
                    if let Some(variable) = variable {
                        self.document_variables.push(variable);
                    }
                }
            }
            "compat" => {
                if is_word_element(element.name().as_ref(), b"compatSetting", prefixes) {
                    let setting = parse_compatibility_setting(element, prefixes);
                    self.tally
                        .record(&["compat", "compatSetting"], setting.is_some());
                    if let Some(setting) = setting {
                        self.compatibility_settings.push(setting);
                    }
                } else if let Some(option) = CompatibilityOption::find(element, prefixes) {
                    let value = parse_toggle(element, prefixes)?;
                    let total = self
                        .tally
                        .record(&["compat", option.local_name()], value.is_some());
                    self.compatibility_options
                        .retain(|(existing, _)| *existing != option);
                    if total == 1
                        && let Some(value) = value
                    {
                        self.compatibility_options.push((option, value));
                    }
                }
            }
            "mailMerge" => {
                let Some(name) = supported_name(MAIL_MERGE_MEMBERS, element, prefixes) else {
                    return Ok(());
                };
                match name {
                    "mainDocumentType" => {
                        let value = word_value(element, prefixes)
                            .and_then(|v| MailMergeDocumentType::parse(&v));
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.main_document_type, value, total);
                    }
                    "linkToQuery" => {
                        let value = parse_toggle(element, prefixes)?;
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.link_to_query, value, total);
                    }
                    "dataType" => {
                        let value = word_value(element, prefixes);
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.data_type, value, total);
                    }
                    "connectString" => {
                        let value = word_value(element, prefixes);
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.connect_string, value, total);
                    }
                    "query" => {
                        let value = word_value(element, prefixes);
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.query, value, total);
                    }
                    "doNotSuppressBlankLines" => {
                        let value = parse_toggle(element, prefixes)?;
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.do_not_suppress_blank_lines, value, total);
                    }
                    "destination" => {
                        let value = word_value(element, prefixes)
                            .and_then(|v| MailMergeDestination::parse(&v));
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.destination, value, total);
                    }
                    "addressFieldName" => {
                        let value = word_value(element, prefixes);
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.address_field_name, value, total);
                    }
                    "mailSubject" => {
                        let value = word_value(element, prefixes);
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.mail_subject, value, total);
                    }
                    "mailAsAttachment" => {
                        let value = parse_toggle(element, prefixes)?;
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.mail_as_attachment, value, total);
                    }
                    "viewMergedData" => {
                        let value = parse_toggle(element, prefixes)?;
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.view_merged_data, value, total);
                    }
                    "activeRecord" => {
                        let value = parse_decimal(element, prefixes);
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.active_record, value, total);
                    }
                    "checkErrors" => {
                        let value = parse_decimal(element, prefixes);
                        let total = self.tally.record(&["mailMerge", name], value.is_some());
                        assign(&mut mail_merge.check_errors, value, total);
                    }
                    _ => {}
                }
            }
            _ => {}
        }
        Ok(())
    }
}

impl CT_Settings {
    /// Document-wide footnote policy, before a section override.
    pub fn footnote_properties(&self) -> Option<&CT_NoteProperties> {
        self.footnote_properties.as_ref()
    }

    /// Document-wide endnote policy, before a section override.
    pub fn endnote_properties(&self) -> Option<&CT_NoteProperties> {
        self.endnote_properties.as_ref()
    }

    /// Replace document-wide footnote policy in schema order.
    pub fn set_footnote_properties(&mut self, properties: CT_NoteProperties) -> Result<()> {
        let mut writer = Writer::new(Vec::new());
        properties.to_xml(&mut writer, "w:footnotePr")?;
        self.finish_set("footnotePr", writer.into_inner())?;
        self.footnote_properties = Some(properties);
        Ok(())
    }

    /// Replace document-wide endnote policy in schema order.
    pub fn set_endnote_properties(&mut self, properties: CT_NoteProperties) -> Result<()> {
        let mut writer = Writer::new(Vec::new());
        properties.to_xml(&mut writer, "w:endnotePr")?;
        self.finish_set("endnotePr", writer.into_inner())?;
        self.endnote_properties = Some(properties);
        Ok(())
    }

    /// Remove document-wide footnote policy.
    pub fn remove_footnote_properties(&mut self) -> Result<Option<CT_NoteProperties>> {
        if !self.begin_removal("footnotePr")? {
            return Ok(None);
        }
        self.finish_removal("footnotePr")?;
        Ok(self.footnote_properties.take())
    }

    /// Remove document-wide endnote policy.
    pub fn remove_endnote_properties(&mut self) -> Result<Option<CT_NoteProperties>> {
        if !self.begin_removal("endnotePr")? {
            return Ok(None);
        }
        self.finish_removal("endnotePr")?;
        Ok(self.endnote_properties.take())
    }

    /// Report every supported child the typed model could not own.
    ///
    /// A package authored entirely through the public API can only produce
    /// well-formed single occurrences, so its list is empty by construction.
    pub fn diagnostics(&self) -> &[SettingsDiagnostic] {
        &self.diagnostics
    }

    /// Replace one modeled top-level child at its schema position.
    fn rewrite_scalar(&mut self, local: &[u8], replacement: Vec<u8>) -> Result<()> {
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_ordered_child(
                source,
                SETTINGS_ORDER,
                local,
                &replacement,
                is_word_element,
            )?);
        }
        Ok(())
    }

    /// Record one well-formed occurrence and write it at its schema position.
    fn finish_set(&mut self, name: &'static str, replacement: Vec<u8>) -> Result<()> {
        self.rewrite_scalar(name.as_bytes(), replacement)?;
        self.tally.set(&[name], true);
        self.diagnostics = self.tally.diagnostics(SETTINGS_REPEATABLE);
        Ok(())
    }

    /// Report whether one modeled top-level child can be removed.
    ///
    /// Removal is never allowed to reinterpret an unowned occurrence, so a
    /// duplicated or malformed member is an error rather than a silent rewrite.
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

    /// Delete the single modeled occurrence and leave every neighbour alone.
    fn finish_removal(&mut self, name: &'static str) -> Result<()> {
        self.rewrite_scalar(name.as_bytes(), Vec::new())?;
        self.tally.set(&[name], false);
        self.diagnostics = self.tally.diagnostics(SETTINGS_REPEATABLE);
        Ok(())
    }

    /// Return the document view Word opens with.
    pub fn view(&self) -> Option<DocumentView> {
        self.view
    }

    pub fn set_view(&mut self, value: DocumentView) -> Result<()> {
        self.view = Some(value);
        self.finish_set("view", write_valued_setting("w:view", value.as_str())?)
    }

    pub fn remove_view(&mut self) -> Result<Option<DocumentView>> {
        if !self.begin_removal("view")? {
            return Ok(None);
        }
        let removed = self.view.take();
        self.finish_removal("view")?;
        Ok(removed)
    }

    /// Return the magnification preset and percentage.
    pub fn zoom(&self) -> Option<DocumentZoom> {
        self.zoom
    }

    pub fn set_zoom(&mut self, value: DocumentZoom) -> Result<()> {
        let replacement = write_zoom(&value)?;
        self.zoom = Some(value);
        self.finish_set("zoom", replacement)
    }

    pub fn remove_zoom(&mut self) -> Result<Option<DocumentZoom>> {
        if !self.begin_removal("zoom")? {
            return Ok(None);
        }
        let removed = self.zoom.take();
        self.finish_removal("zoom")?;
        Ok(removed)
    }

    /// Return the `w:removePersonalInformation` toggle.
    pub fn remove_personal_information(&self) -> Option<bool> {
        self.remove_personal_information
    }

    pub fn set_remove_personal_information(&mut self, enabled: bool) -> Result<()> {
        self.remove_personal_information = Some(enabled);
        self.finish_set(
            "removePersonalInformation",
            write_toggle_setting("w:removePersonalInformation", enabled)?,
        )
    }

    /// Remove the `w:removePersonalInformation` toggle itself.
    pub fn remove_remove_personal_information(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("removePersonalInformation")? {
            return Ok(None);
        }
        let removed = self.remove_personal_information.take();
        self.finish_removal("removePersonalInformation")?;
        Ok(removed)
    }

    /// Return the `w:removeDateAndTime` toggle.
    pub fn remove_date_and_time(&self) -> Option<bool> {
        self.remove_date_and_time
    }

    pub fn set_remove_date_and_time(&mut self, enabled: bool) -> Result<()> {
        self.remove_date_and_time = Some(enabled);
        self.finish_set(
            "removeDateAndTime",
            write_toggle_setting("w:removeDateAndTime", enabled)?,
        )
    }

    /// Remove the `w:removeDateAndTime` toggle itself.
    pub fn remove_remove_date_and_time(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("removeDateAndTime")? {
            return Ok(None);
        }
        let removed = self.remove_date_and_time.take();
        self.finish_removal("removeDateAndTime")?;
        Ok(removed)
    }

    /// Return the `w:displayBackgroundShape` toggle, which Word sets with a
    /// page colour so that the document background shows.
    pub fn display_background_shape(&self) -> Option<bool> {
        self.display_background_shape
    }

    pub fn set_display_background_shape(&mut self, enabled: bool) -> Result<()> {
        self.display_background_shape = Some(enabled);
        self.finish_set(
            "displayBackgroundShape",
            write_toggle_setting("w:displayBackgroundShape", enabled)?,
        )
    }

    pub fn remove_display_background_shape(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("displayBackgroundShape")? {
            return Ok(None);
        }
        let removed = self.display_background_shape.take();
        self.finish_removal("displayBackgroundShape")?;
        Ok(removed)
    }

    /// Return the `w:mirrorMargins` toggle.
    pub fn mirror_margins(&self) -> Option<bool> {
        self.mirror_margins
    }

    pub fn set_mirror_margins(&mut self, enabled: bool) -> Result<()> {
        self.mirror_margins = Some(enabled);
        self.finish_set(
            "mirrorMargins",
            write_toggle_setting("w:mirrorMargins", enabled)?,
        )
    }

    pub fn remove_mirror_margins(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("mirrorMargins")? {
            return Ok(None);
        }
        let removed = self.mirror_margins.take();
        self.finish_removal("mirrorMargins")?;
        Ok(removed)
    }

    /// Return the `w:gutterAtTop` toggle.
    pub fn gutter_at_top(&self) -> Option<bool> {
        self.gutter_at_top
    }

    pub fn set_gutter_at_top(&mut self, enabled: bool) -> Result<()> {
        self.gutter_at_top = Some(enabled);
        self.finish_set(
            "gutterAtTop",
            write_toggle_setting("w:gutterAtTop", enabled)?,
        )
    }

    pub fn remove_gutter_at_top(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("gutterAtTop")? {
            return Ok(None);
        }
        let removed = self.gutter_at_top.take();
        self.finish_removal("gutterAtTop")?;
        Ok(removed)
    }

    /// Return the spelling and grammar proofing state.
    pub fn proof_state(&self) -> Option<DocumentProofState> {
        self.proof_state
    }

    pub fn set_proof_state(&mut self, value: DocumentProofState) -> Result<()> {
        let replacement = write_proof_state(&value)?;
        self.proof_state = Some(value);
        self.finish_set("proofState", replacement)
    }

    pub fn remove_proof_state(&mut self) -> Result<Option<DocumentProofState>> {
        if !self.begin_removal("proofState")? {
            return Ok(None);
        }
        let removed = self.proof_state.take();
        self.finish_removal("proofState")?;
        Ok(removed)
    }

    /// Return the `w:linkStyles` toggle.
    pub fn link_styles(&self) -> Option<bool> {
        self.link_styles
    }

    pub fn set_link_styles(&mut self, enabled: bool) -> Result<()> {
        self.link_styles = Some(enabled);
        self.finish_set("linkStyles", write_toggle_setting("w:linkStyles", enabled)?)
    }

    pub fn remove_link_styles(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("linkStyles")? {
            return Ok(None);
        }
        let removed = self.link_styles.take();
        self.finish_removal("linkStyles")?;
        Ok(removed)
    }

    /// Return the bounded mail-merge configuration.
    pub fn mail_merge(&self) -> Option<&MailMerge> {
        self.mail_merge.as_ref()
    }

    /// Replace the modeled mail-merge members at their schema positions.
    ///
    /// `w:dataSource`, `w:headerSource` and `w:odso` keep their producer bytes,
    /// which is why each modeled member is written in place rather than the
    /// group being replaced whole.
    pub fn set_mail_merge(&mut self, value: MailMerge) -> Result<()> {
        let children = mail_merge_children(&value)?;
        if let Some(source) = &self.source_xml {
            let mut rewritten = ensure_top_level_group(source, b"mailMerge")?;
            for (local, replacement) in &children {
                rewritten = rewrite_group_child(
                    &rewritten,
                    b"mailMerge",
                    local.as_bytes(),
                    replacement,
                    MAIL_MERGE_ORDER,
                )?;
            }
            self.source_xml = Some(rewritten);
        }
        self.tally.set(&["mailMerge"], true);
        for (local, replacement) in &children {
            self.tally
                .set(&["mailMerge", local], !replacement.is_empty());
        }
        self.mail_merge = Some(value);
        self.diagnostics = self.tally.diagnostics(SETTINGS_REPEATABLE);
        Ok(())
    }

    /// Remove every modeled mail-merge member.
    ///
    /// The group itself goes only when nothing preserved is left inside it,
    /// because `CT_MailMerge` requires members this model does not author.
    pub fn remove_mail_merge(&mut self) -> Result<Option<MailMerge>> {
        let total = self.tally.total(&["mailMerge"]);
        if total == 0 {
            return Ok(None);
        }
        if total > 1 {
            return Err(OxmlError::InvalidValue(
                "cannot rewrite ambiguous or malformed w:mailMerge".to_owned(),
            ));
        }
        let removed = self.mail_merge.take();
        let mut group_remains = false;
        if let Some(source) = &self.source_xml {
            let mut rewritten = source.clone();
            for local in MAIL_MERGE_MEMBERS {
                rewritten = rewrite_group_child(
                    &rewritten,
                    b"mailMerge",
                    local.as_bytes(),
                    &[],
                    MAIL_MERGE_ORDER,
                )?;
            }
            if top_level_group_is_childless(&rewritten, b"mailMerge")? {
                rewritten = rewrite_ordered_child(
                    &rewritten,
                    SETTINGS_ORDER,
                    b"mailMerge",
                    &[],
                    is_word_element,
                )?;
            } else {
                group_remains = true;
            }
            self.source_xml = Some(rewritten);
        }
        for local in MAIL_MERGE_MEMBERS {
            self.tally.set(&["mailMerge", local], false);
        }
        self.tally.set(&["mailMerge"], group_remains);
        if group_remains {
            self.mail_merge = Some(MailMerge::default());
        }
        self.diagnostics = self.tally.diagnostics(SETTINGS_REPEATABLE);
        Ok(removed)
    }

    /// Return the `w:trackRevisions` toggle.
    pub fn track_revisions(&self) -> Option<bool> {
        self.track_revisions
    }

    pub fn set_track_revisions(&mut self, enabled: bool) -> Result<()> {
        self.track_revisions = Some(enabled);
        self.finish_set(
            "trackRevisions",
            write_toggle_setting("w:trackRevisions", enabled)?,
        )
    }

    pub fn remove_track_revisions(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("trackRevisions")? {
            return Ok(None);
        }
        let removed = self.track_revisions.take();
        self.finish_removal("trackRevisions")?;
        Ok(removed)
    }

    /// Return the `w:doNotTrackMoves` toggle.
    pub fn do_not_track_moves(&self) -> Option<bool> {
        self.do_not_track_moves
    }

    pub fn set_do_not_track_moves(&mut self, enabled: bool) -> Result<()> {
        self.do_not_track_moves = Some(enabled);
        self.finish_set(
            "doNotTrackMoves",
            write_toggle_setting("w:doNotTrackMoves", enabled)?,
        )
    }

    pub fn remove_do_not_track_moves(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("doNotTrackMoves")? {
            return Ok(None);
        }
        let removed = self.do_not_track_moves.take();
        self.finish_removal("doNotTrackMoves")?;
        Ok(removed)
    }

    /// Return the `w:doNotTrackFormatting` toggle.
    pub fn do_not_track_formatting(&self) -> Option<bool> {
        self.do_not_track_formatting
    }

    pub fn set_do_not_track_formatting(&mut self, enabled: bool) -> Result<()> {
        self.do_not_track_formatting = Some(enabled);
        self.finish_set(
            "doNotTrackFormatting",
            write_toggle_setting("w:doNotTrackFormatting", enabled)?,
        )
    }

    pub fn remove_do_not_track_formatting(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("doNotTrackFormatting")? {
            return Ok(None);
        }
        let removed = self.do_not_track_formatting.take();
        self.finish_removal("doNotTrackFormatting")?;
        Ok(removed)
    }

    /// Return valid document-protection metadata, when the part records it.
    pub fn document_protection(&self) -> Option<&DocumentProtection> {
        self.document_protection.as_ref()
    }

    /// Record caller-supplied document-protection metadata verbatim.
    ///
    /// Nothing here derives a hash, chooses a salt or picks a spin count.
    pub fn set_document_protection(&mut self, value: DocumentProtection) -> Result<()> {
        let replacement = write_document_protection(&value)?;
        self.document_protection = Some(value);
        self.finish_set("documentProtection", replacement)
    }

    pub fn remove_document_protection(&mut self) -> Result<Option<DocumentProtection>> {
        if !self.begin_removal("documentProtection")? {
            return Ok(None);
        }
        let removed = self.document_protection.take();
        self.finish_removal("documentProtection")?;
        Ok(removed)
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
        self.rewrite_group_list(b"docVars", b"docVar")?;
        self.tally.set(&["docVars"], true);
        self.diagnostics = self.tally.diagnostics(SETTINGS_REPEATABLE);
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
        self.rewrite_group_list(b"docVars", b"docVar")?;
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
        self.rewrite_group_list(b"compat", b"compatSetting")?;
        self.tally.set(&["compat"], true);
        self.diagnostics = self.tally.diagnostics(SETTINGS_REPEATABLE);
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
        self.rewrite_group_list(b"compat", b"compatSetting")?;
        Ok(Some(removed))
    }

    /// Replace every modeled occurrence of one repeated group child.
    fn rewrite_group_list(&mut self, group_local: &[u8], child_local: &[u8]) -> Result<()> {
        let Some(source) = &self.source_xml else {
            return Ok(());
        };
        let modeled = if child_local == b"docVar" {
            write_document_variables(&self.document_variables)?
        } else {
            write_compatibility_settings(&self.compatibility_settings)?
        };
        self.source_xml = Some(rewrite_group_setting(
            source,
            group_local,
            child_local,
            modeled,
        )?);
        Ok(())
    }

    /// Return every modeled `w:compat` toggle option in schema order.
    pub fn compatibility_options(&self) -> &[(CompatibilityOption, bool)] {
        &self.compatibility_options
    }

    pub fn compatibility_option(&self, option: CompatibilityOption) -> Option<bool> {
        self.compatibility_options
            .iter()
            .find(|(existing, _)| *existing == option)
            .map(|(_, value)| *value)
    }

    pub fn set_compatibility_option(
        &mut self,
        option: CompatibilityOption,
        enabled: bool,
    ) -> Result<()> {
        let local = option.local_name();
        let replacement = write_toggle_setting(&format!("w:{local}"), enabled)?;
        if let Some(source) = &self.source_xml {
            let ensured = ensure_top_level_group(source, b"compat")?;
            self.source_xml = Some(rewrite_group_child(
                &ensured,
                b"compat",
                local.as_bytes(),
                &replacement,
                COMPAT_ORDER,
            )?);
        }
        self.compatibility_options
            .retain(|(existing, _)| *existing != option);
        self.compatibility_options.push((option, enabled));
        self.compatibility_options.sort_by_key(|entry| entry.0);
        self.tally.set(&["compat"], true);
        self.tally.set(&["compat", local], true);
        self.diagnostics = self.tally.diagnostics(SETTINGS_REPEATABLE);
        Ok(())
    }

    pub fn remove_compatibility_option(
        &mut self,
        option: CompatibilityOption,
    ) -> Result<Option<bool>> {
        let local = option.local_name();
        let total = self.tally.total(&["compat", local]);
        if total == 0 {
            return Ok(None);
        }
        if total > 1 || self.tally.typed(&["compat", local]) != total {
            return Err(OxmlError::InvalidValue(format!(
                "cannot rewrite ambiguous or malformed w:{local}"
            )));
        }
        let removed = self.compatibility_option(option);
        self.compatibility_options
            .retain(|(existing, _)| *existing != option);
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_group_child(
                source,
                b"compat",
                local.as_bytes(),
                &[],
                COMPAT_ORDER,
            )?);
        }
        self.tally.set(&["compat", local], false);
        self.diagnostics = self.tally.diagnostics(SETTINGS_REPEATABLE);
        Ok(removed)
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
        self.default_tab_stop = Some(value);
        self.finish_set("defaultTabStop", replacement)
    }

    pub fn remove_default_tab_stop(&mut self) -> Result<Option<Twips>> {
        if !self.begin_removal("defaultTabStop")? {
            return Ok(None);
        }
        let removed = self.default_tab_stop.take();
        self.finish_removal("defaultTabStop")?;
        Ok(removed)
    }

    /// Return whether Word automatic hyphenation is enabled.
    ///
    /// OOXML defines omission as disabled.
    pub fn automatic_hyphenation(&self) -> bool {
        self.automatic_hyphenation.unwrap_or(false)
    }

    /// Set the document automatic-hyphenation toggle.
    ///
    /// Parsed settings retain every unrelated producer byte. The one modeled
    /// toggle is rewritten with the fixed `w:` prefix at its schema position.
    pub fn set_automatic_hyphenation(&mut self, enabled: bool) -> Result<()> {
        self.automatic_hyphenation = Some(enabled);
        self.finish_set(
            "autoHyphenation",
            write_toggle_setting("w:autoHyphenation", enabled)?,
        )
    }

    pub fn remove_automatic_hyphenation(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("autoHyphenation")? {
            return Ok(None);
        }
        let removed = self.automatic_hyphenation.take();
        self.finish_removal("autoHyphenation")?;
        Ok(removed)
    }

    /// Return the maximum count of consecutive hyphenated lines.
    pub fn consecutive_hyphen_limit(&self) -> Option<i32> {
        self.consecutive_hyphen_limit
    }

    pub fn set_consecutive_hyphen_limit(&mut self, value: i32) -> Result<()> {
        let replacement = write_valued_setting("w:consecutiveHyphenLimit", &value.to_string())?;
        self.consecutive_hyphen_limit = Some(value);
        self.finish_set("consecutiveHyphenLimit", replacement)
    }

    pub fn remove_consecutive_hyphen_limit(&mut self) -> Result<Option<i32>> {
        if !self.begin_removal("consecutiveHyphenLimit")? {
            return Ok(None);
        }
        let removed = self.consecutive_hyphen_limit.take();
        self.finish_removal("consecutiveHyphenLimit")?;
        Ok(removed)
    }

    /// Return the hyphenation zone width.
    pub fn hyphenation_zone(&self) -> Option<Twips> {
        self.hyphenation_zone
    }

    pub fn set_hyphenation_zone(&mut self, value: Twips) -> Result<()> {
        if value.0 < 0 {
            return Err(OxmlError::InvalidValue(
                "hyphenation zone must be non-negative twips".to_owned(),
            ));
        }
        let replacement = write_valued_setting("w:hyphenationZone", &value.0.to_string())?;
        self.hyphenation_zone = Some(value);
        self.finish_set("hyphenationZone", replacement)
    }

    pub fn remove_hyphenation_zone(&mut self) -> Result<Option<Twips>> {
        if !self.begin_removal("hyphenationZone")? {
            return Ok(None);
        }
        let removed = self.hyphenation_zone.take();
        self.finish_removal("hyphenationZone")?;
        Ok(removed)
    }

    /// Return the `w:doNotHyphenateCaps` toggle.
    pub fn do_not_hyphenate_caps(&self) -> Option<bool> {
        self.do_not_hyphenate_caps
    }

    pub fn set_do_not_hyphenate_caps(&mut self, enabled: bool) -> Result<()> {
        self.do_not_hyphenate_caps = Some(enabled);
        self.finish_set(
            "doNotHyphenateCaps",
            write_toggle_setting("w:doNotHyphenateCaps", enabled)?,
        )
    }

    pub fn remove_do_not_hyphenate_caps(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("doNotHyphenateCaps")? {
            return Ok(None);
        }
        let removed = self.do_not_hyphenate_caps.take();
        self.finish_removal("doNotHyphenateCaps")?;
        Ok(removed)
    }

    /// Return the default table style identifier.
    pub fn default_table_style(&self) -> Option<&str> {
        self.default_table_style.as_deref()
    }

    pub fn set_default_table_style(&mut self, value: String) -> Result<()> {
        let replacement = write_valued_setting("w:defaultTableStyle", &value)?;
        self.default_table_style = Some(value);
        self.finish_set("defaultTableStyle", replacement)
    }

    pub fn remove_default_table_style(&mut self) -> Result<Option<String>> {
        if !self.begin_removal("defaultTableStyle")? {
            return Ok(None);
        }
        let removed = self.default_table_style.take();
        self.finish_removal("defaultTableStyle")?;
        Ok(removed)
    }

    /// Return whether Word selects distinct even-page header and footer stories.
    ///
    /// OOXML defines omission as disabled.
    pub fn even_and_odd_headers(&self) -> bool {
        self.even_and_odd_headers.unwrap_or(false)
    }

    /// Set distinct even-page header and footer selection.
    pub fn set_even_and_odd_headers(&mut self, enabled: bool) -> Result<()> {
        self.even_and_odd_headers = Some(enabled);
        self.finish_set(
            "evenAndOddHeaders",
            write_toggle_setting("w:evenAndOddHeaders", enabled)?,
        )
    }

    pub fn remove_even_and_odd_headers(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("evenAndOddHeaders")? {
            return Ok(None);
        }
        let removed = self.even_and_odd_headers.take();
        self.finish_removal("evenAndOddHeaders")?;
        Ok(removed)
    }

    /// Return the `w:bookFoldRevPrinting` toggle.
    pub fn book_fold_rev_printing(&self) -> Option<bool> {
        self.book_fold_rev_printing
    }

    pub fn set_book_fold_rev_printing(&mut self, enabled: bool) -> Result<()> {
        self.book_fold_rev_printing = Some(enabled);
        self.finish_set(
            "bookFoldRevPrinting",
            write_toggle_setting("w:bookFoldRevPrinting", enabled)?,
        )
    }

    pub fn remove_book_fold_rev_printing(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("bookFoldRevPrinting")? {
            return Ok(None);
        }
        let removed = self.book_fold_rev_printing.take();
        self.finish_removal("bookFoldRevPrinting")?;
        Ok(removed)
    }

    /// Return the `w:bookFoldPrinting` toggle.
    pub fn book_fold_printing(&self) -> Option<bool> {
        self.book_fold_printing
    }

    pub fn set_book_fold_printing(&mut self, enabled: bool) -> Result<()> {
        self.book_fold_printing = Some(enabled);
        self.finish_set(
            "bookFoldPrinting",
            write_toggle_setting("w:bookFoldPrinting", enabled)?,
        )
    }

    pub fn remove_book_fold_printing(&mut self) -> Result<Option<bool>> {
        if !self.begin_removal("bookFoldPrinting")? {
            return Ok(None);
        }
        let removed = self.book_fold_printing.take();
        self.finish_removal("bookFoldPrinting")?;
        Ok(removed)
    }

    /// Return the signed sheet count per book-fold booklet.
    pub fn book_fold_printing_sheets(&self) -> Option<i32> {
        self.book_fold_printing_sheets
    }

    pub fn set_book_fold_printing_sheets(&mut self, value: i32) -> Result<()> {
        let replacement = write_valued_setting("w:bookFoldPrintingSheets", &value.to_string())?;
        self.book_fold_printing_sheets = Some(value);
        self.finish_set("bookFoldPrintingSheets", replacement)
    }

    pub fn remove_book_fold_printing_sheets(&mut self) -> Result<Option<i32>> {
        if !self.begin_removal("bookFoldPrintingSheets")? {
            return Ok(None);
        }
        let removed = self.book_fold_printing_sheets.take();
        self.finish_removal("bookFoldPrintingSheets")?;
        Ok(removed)
    }

    pub fn character_spacing_control(&self) -> Option<CharacterSpacingControl> {
        self.character_spacing_control
    }

    pub fn set_character_spacing_control(&mut self, value: CharacterSpacingControl) -> Result<()> {
        let replacement = write_valued_setting("w:characterSpacingControl", value.as_str())?;
        self.character_spacing_control = Some(value);
        self.finish_set("characterSpacingControl", replacement)
    }

    pub fn remove_character_spacing_control(&mut self) -> Result<Option<CharacterSpacingControl>> {
        if !self.begin_removal("characterSpacingControl")? {
            return Ok(None);
        }
        let removed = self.character_spacing_control.take();
        self.finish_removal("characterSpacingControl")?;
        Ok(removed)
    }

    /// Return the `w:updateFields` toggle, which asks Word to update fields
    /// when it opens the document.
    ///
    /// `None` means the part omits the toggle or carries an unmodelled form.
    pub fn update_fields(&self) -> Option<bool> {
        self.update_fields
    }

    /// Set the `w:updateFields` toggle, or remove it with `None`.
    pub fn set_update_fields(&mut self, value: Option<bool>) -> Result<()> {
        let total = self.tally.total(&["updateFields"]);
        if total > 1 || (total == 1 && self.update_fields.is_none()) {
            return Err(OxmlError::InvalidValue(
                "cannot rewrite ambiguous or malformed w:updateFields".to_owned(),
            ));
        }
        match value {
            Some(enabled) => {
                self.update_fields = Some(enabled);
                self.finish_set(
                    "updateFields",
                    write_toggle_setting("w:updateFields", enabled)?,
                )
            }
            None => {
                if total == 0 {
                    return Ok(());
                }
                self.update_fields = None;
                self.finish_removal("updateFields")
            }
        }
    }

    /// Return document-wide OfficeMath defaults, when present and valid.
    pub fn math_properties(&self) -> Option<&MathProperties> {
        self.math_properties.as_ref()
    }

    /// Replace the single schema-positioned OfficeMath defaults subtree.
    pub fn set_math_properties(&mut self, properties: MathProperties) -> Result<()> {
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_math_properties(source, Some(&properties))?);
        }
        self.math_properties = Some(properties);
        self.tally.set(&["mathPr"], true);
        self.diagnostics = self.tally.diagnostics(SETTINGS_REPEATABLE);
        Ok(())
    }

    pub fn remove_math_properties(&mut self) -> Result<Option<MathProperties>> {
        if !self.begin_removal("mathPr")? {
            return Ok(None);
        }
        let removed = self.math_properties.take();
        if let Some(source) = &self.source_xml {
            self.source_xml = Some(rewrite_math_properties(source, None)?);
        }
        self.tally.set(&["mathPr"], false);
        self.diagnostics = self.tally.diagnostics(SETTINGS_REPEATABLE);
        Ok(removed)
    }

    pub fn theme_font_language(&self) -> Option<&ThemeFontLanguage> {
        self.theme_font_language.as_ref()
    }

    pub fn set_theme_font_language(&mut self, value: ThemeFontLanguage) -> Result<()> {
        let replacement = write_theme_font_language(&value)?;
        self.theme_font_language = Some(value);
        self.finish_set("themeFontLang", replacement)
    }

    pub fn remove_theme_font_language(&mut self) -> Result<Option<ThemeFontLanguage>> {
        if !self.begin_removal("themeFontLang")? {
            return Ok(None);
        }
        let removed = self.theme_font_language.take();
        self.finish_removal("themeFontLang")?;
        Ok(removed)
    }

    /// Return the decimal separator Word uses in fields.
    pub fn decimal_symbol(&self) -> Option<&str> {
        self.decimal_symbol.as_deref()
    }

    pub fn set_decimal_symbol(&mut self, value: String) -> Result<()> {
        let replacement = write_valued_setting("w:decimalSymbol", &value)?;
        self.decimal_symbol = Some(value);
        self.finish_set("decimalSymbol", replacement)
    }

    pub fn remove_decimal_symbol(&mut self) -> Result<Option<String>> {
        if !self.begin_removal("decimalSymbol")? {
            return Ok(None);
        }
        let removed = self.decimal_symbol.take();
        self.finish_removal("decimalSymbol")?;
        Ok(removed)
    }

    /// Return the list separator Word uses in fields.
    pub fn list_separator(&self) -> Option<&str> {
        self.list_separator.as_deref()
    }

    pub fn set_list_separator(&mut self, value: String) -> Result<()> {
        let replacement = write_valued_setting("w:listSeparator", &value)?;
        self.list_separator = Some(value);
        self.finish_set("listSeparator", replacement)
    }

    pub fn remove_list_separator(&mut self) -> Result<Option<String>> {
        if !self.begin_removal("listSeparator")? {
            return Ok(None);
        }
        let removed = self.list_separator.take();
        self.finish_removal("listSeparator")?;
        Ok(removed)
    }

    /// Return whether the model has no typed or retained settings content.
    pub fn is_empty(&self) -> bool {
        self.view.is_none()
            && self.zoom.is_none()
            && self.remove_personal_information.is_none()
            && self.remove_date_and_time.is_none()
            && self.display_background_shape.is_none()
            && self.mirror_margins.is_none()
            && self.gutter_at_top.is_none()
            && self.proof_state.is_none()
            && self.link_styles.is_none()
            && self.mail_merge.is_none()
            && self.track_revisions.is_none()
            && self.do_not_track_moves.is_none()
            && self.do_not_track_formatting.is_none()
            && self.document_protection.is_none()
            && self.footnote_properties.is_none()
            && self.endnote_properties.is_none()
            && self.default_tab_stop.is_none()
            && self.automatic_hyphenation.is_none()
            && self.consecutive_hyphen_limit.is_none()
            && self.hyphenation_zone.is_none()
            && self.do_not_hyphenate_caps.is_none()
            && self.default_table_style.is_none()
            && self.even_and_odd_headers.is_none()
            && self.book_fold_rev_printing.is_none()
            && self.book_fold_printing.is_none()
            && self.book_fold_printing_sheets.is_none()
            && self.character_spacing_control.is_none()
            && self.update_fields.is_none()
            && self.compatibility_options.is_empty()
            && self.compatibility_settings.is_empty()
            && self.document_variables.is_empty()
            && self.math_properties.is_none()
            && self.theme_font_language.is_none()
            && self.decimal_symbol.is_none()
            && self.list_separator.is_none()
            && self.source_xml.is_none()
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
        let mut emit = |bytes: Vec<u8>| writer.get_mut().extend_from_slice(&bytes);
        if let Some(value) = self.view {
            emit(write_valued_setting("w:view", value.as_str())?);
        }
        if let Some(value) = &self.zoom {
            emit(write_zoom(value)?);
        }
        for (name, value) in [
            (
                "w:removePersonalInformation",
                self.remove_personal_information,
            ),
            ("w:removeDateAndTime", self.remove_date_and_time),
            ("w:displayBackgroundShape", self.display_background_shape),
            ("w:mirrorMargins", self.mirror_margins),
            ("w:gutterAtTop", self.gutter_at_top),
        ] {
            if let Some(enabled) = value {
                emit(write_toggle_setting(name, enabled)?);
            }
        }
        if let Some(value) = &self.proof_state {
            emit(write_proof_state(value)?);
        }
        if let Some(enabled) = self.link_styles {
            emit(write_toggle_setting("w:linkStyles", enabled)?);
        }
        if let Some(value) = &self.mail_merge {
            let children = mail_merge_children(value)?;
            if children.iter().any(|(_, bytes)| !bytes.is_empty()) {
                emit(wrap_group("w:mailMerge", &concatenate(&children))?);
            }
        }
        for (name, value) in [
            ("w:trackRevisions", self.track_revisions),
            ("w:doNotTrackMoves", self.do_not_track_moves),
            ("w:doNotTrackFormatting", self.do_not_track_formatting),
        ] {
            if let Some(enabled) = value {
                emit(write_toggle_setting(name, enabled)?);
            }
        }
        if let Some(protection) = &self.document_protection {
            emit(write_document_protection(protection)?);
        }
        if let Some(properties) = &self.footnote_properties {
            let mut child = Writer::new(Vec::new());
            properties.to_xml(&mut child, "w:footnotePr")?;
            emit(child.into_inner());
        }
        if let Some(properties) = &self.endnote_properties {
            let mut child = Writer::new(Vec::new());
            properties.to_xml(&mut child, "w:endnotePr")?;
            emit(child.into_inner());
        }
        if let Some(value) = self.default_tab_stop {
            emit(write_valued_setting(
                "w:defaultTabStop",
                &value.0.to_string(),
            )?);
        }
        if let Some(enabled) = self.automatic_hyphenation {
            emit(write_toggle_setting("w:autoHyphenation", enabled)?);
        }
        if let Some(value) = self.consecutive_hyphen_limit {
            emit(write_valued_setting(
                "w:consecutiveHyphenLimit",
                &value.to_string(),
            )?);
        }
        if let Some(value) = self.hyphenation_zone {
            emit(write_valued_setting(
                "w:hyphenationZone",
                &value.0.to_string(),
            )?);
        }
        if let Some(enabled) = self.do_not_hyphenate_caps {
            emit(write_toggle_setting("w:doNotHyphenateCaps", enabled)?);
        }
        if let Some(value) = &self.default_table_style {
            emit(write_valued_setting("w:defaultTableStyle", value)?);
        }
        for (name, value) in [
            ("w:evenAndOddHeaders", self.even_and_odd_headers),
            ("w:bookFoldRevPrinting", self.book_fold_rev_printing),
            ("w:bookFoldPrinting", self.book_fold_printing),
        ] {
            if let Some(enabled) = value {
                emit(write_toggle_setting(name, enabled)?);
            }
        }
        if let Some(value) = self.book_fold_printing_sheets {
            emit(write_valued_setting(
                "w:bookFoldPrintingSheets",
                &value.to_string(),
            )?);
        }
        if let Some(value) = self.character_spacing_control {
            emit(write_valued_setting(
                "w:characterSpacingControl",
                value.as_str(),
            )?);
        }
        if let Some(enabled) = self.update_fields {
            emit(write_toggle_setting("w:updateFields", enabled)?);
        }
        if !self.compatibility_options.is_empty() || !self.compatibility_settings.is_empty() {
            let mut children = Vec::new();
            for (option, enabled) in &self.compatibility_options {
                children.extend_from_slice(&write_toggle_setting(
                    &format!("w:{}", option.local_name()),
                    *enabled,
                )?);
            }
            children
                .extend_from_slice(&write_compatibility_settings(&self.compatibility_settings)?);
            emit(wrap_group("w:compat", &children)?);
        }
        if !self.document_variables.is_empty() {
            emit(wrap_group(
                "w:docVars",
                &write_document_variables(&self.document_variables)?,
            )?);
        }
        if let Some(properties) = &self.math_properties {
            properties.write_xml(&mut writer)?;
        }
        let mut emit = |bytes: Vec<u8>| writer.get_mut().extend_from_slice(&bytes);
        if let Some(language) = &self.theme_font_language {
            emit(write_theme_font_language(language)?);
        }
        if let Some(value) = &self.decimal_symbol {
            emit(write_valued_setting("w:decimalSymbol", value)?);
        }
        if let Some(value) = &self.list_separator {
            emit(write_valued_setting("w:listSeparator", value)?);
        }
        writer.write_event(Event::End(BytesEnd::new("w:settings")))?;
        Ok(writer.into_inner())
    }
}

/// The complete closed `CT_Compat` on-off child set.
///
/// Every `w:compat` child except `w:compatSetting` is a `CT_OnOff` of
/// identical shape, so one ordered enum carries the family in one place
/// instead of roughly eighty near-identical struct fields. The variants are
/// declared in `xsd:sequence` order, which is what `Ord` sorts by.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum CompatibilityOption {
    UseSingleBorderforContiguousCells,
    WpJustification,
    NoTabHangInd,
    NoLeading,
    SpaceForUL,
    NoColumnBalance,
    BalanceSingleByteDoubleByteWidth,
    NoExtraLineSpacing,
    DoNotLeaveBackslashAlone,
    UlTrailSpace,
    DoNotExpandShiftReturn,
    SpacingInWholePoints,
    LineWrapLikeWord6,
    PrintBodyTextBeforeHeader,
    PrintColBlack,
    WpSpaceWidth,
    ShowBreaksInFrames,
    SubFontBySize,
    SuppressBottomSpacing,
    SuppressTopSpacing,
    SuppressSpacingAtTopOfPage,
    SuppressTopSpacingWP,
    SuppressSpBfAfterPgBrk,
    SwapBordersFacingPages,
    ConvMailMergeEsc,
    TruncateFontHeightsLikeWP6,
    MwSmallCaps,
    UsePrinterMetrics,
    DoNotSuppressParagraphBorders,
    WrapTrailSpaces,
    FootnoteLayoutLikeWW8,
    ShapeLayoutLikeWW8,
    AlignTablesRowByRow,
    ForgetLastTabAlignment,
    AdjustLineHeightInTable,
    AutoSpaceLikeWord95,
    NoSpaceRaiseLower,
    DoNotUseHTMLParagraphAutoSpacing,
    LayoutRawTableWidth,
    LayoutTableRowsApart,
    UseWord97LineBreakRules,
    DoNotBreakWrappedTables,
    DoNotSnapToGridInCell,
    SelectFldWithFirstOrLastChar,
    ApplyBreakingRules,
    DoNotWrapTextWithPunct,
    DoNotUseEastAsianBreakRules,
    UseWord2002TableStyleRules,
    GrowAutofit,
    UseFELayout,
    UseNormalStyleForList,
    DoNotUseIndentAsNumberingTabStop,
    UseAltKinsokuLineBreakRules,
    AllowSpaceOfSameStyleInTable,
    DoNotSuppressIndentation,
    DoNotAutofitConstrainedTables,
    AutofitToFirstFixedWidthCell,
    UnderlineTabInNumList,
    DisplayHangulFixedWidth,
    SplitPgBreakAndParaMark,
    DoNotVertAlignCellWithSp,
    DoNotBreakConstrainedForcedTable,
    DoNotVertAlignInTxbx,
    UseAnsiKerningPairs,
    CachedColBalance,
}

impl CompatibilityOption {
    /// Every option in `xsd:sequence` order.
    pub const ALL: &'static [Self] = &[
        Self::UseSingleBorderforContiguousCells,
        Self::WpJustification,
        Self::NoTabHangInd,
        Self::NoLeading,
        Self::SpaceForUL,
        Self::NoColumnBalance,
        Self::BalanceSingleByteDoubleByteWidth,
        Self::NoExtraLineSpacing,
        Self::DoNotLeaveBackslashAlone,
        Self::UlTrailSpace,
        Self::DoNotExpandShiftReturn,
        Self::SpacingInWholePoints,
        Self::LineWrapLikeWord6,
        Self::PrintBodyTextBeforeHeader,
        Self::PrintColBlack,
        Self::WpSpaceWidth,
        Self::ShowBreaksInFrames,
        Self::SubFontBySize,
        Self::SuppressBottomSpacing,
        Self::SuppressTopSpacing,
        Self::SuppressSpacingAtTopOfPage,
        Self::SuppressTopSpacingWP,
        Self::SuppressSpBfAfterPgBrk,
        Self::SwapBordersFacingPages,
        Self::ConvMailMergeEsc,
        Self::TruncateFontHeightsLikeWP6,
        Self::MwSmallCaps,
        Self::UsePrinterMetrics,
        Self::DoNotSuppressParagraphBorders,
        Self::WrapTrailSpaces,
        Self::FootnoteLayoutLikeWW8,
        Self::ShapeLayoutLikeWW8,
        Self::AlignTablesRowByRow,
        Self::ForgetLastTabAlignment,
        Self::AdjustLineHeightInTable,
        Self::AutoSpaceLikeWord95,
        Self::NoSpaceRaiseLower,
        Self::DoNotUseHTMLParagraphAutoSpacing,
        Self::LayoutRawTableWidth,
        Self::LayoutTableRowsApart,
        Self::UseWord97LineBreakRules,
        Self::DoNotBreakWrappedTables,
        Self::DoNotSnapToGridInCell,
        Self::SelectFldWithFirstOrLastChar,
        Self::ApplyBreakingRules,
        Self::DoNotWrapTextWithPunct,
        Self::DoNotUseEastAsianBreakRules,
        Self::UseWord2002TableStyleRules,
        Self::GrowAutofit,
        Self::UseFELayout,
        Self::UseNormalStyleForList,
        Self::DoNotUseIndentAsNumberingTabStop,
        Self::UseAltKinsokuLineBreakRules,
        Self::AllowSpaceOfSameStyleInTable,
        Self::DoNotSuppressIndentation,
        Self::DoNotAutofitConstrainedTables,
        Self::AutofitToFirstFixedWidthCell,
        Self::UnderlineTabInNumList,
        Self::DisplayHangulFixedWidth,
        Self::SplitPgBreakAndParaMark,
        Self::DoNotVertAlignCellWithSp,
        Self::DoNotBreakConstrainedForcedTable,
        Self::DoNotVertAlignInTxbx,
        Self::UseAnsiKerningPairs,
        Self::CachedColBalance,
    ];

    pub fn local_name(self) -> &'static str {
        match self {
            Self::UseSingleBorderforContiguousCells => "useSingleBorderforContiguousCells",
            Self::WpJustification => "wpJustification",
            Self::NoTabHangInd => "noTabHangInd",
            Self::NoLeading => "noLeading",
            Self::SpaceForUL => "spaceForUL",
            Self::NoColumnBalance => "noColumnBalance",
            Self::BalanceSingleByteDoubleByteWidth => "balanceSingleByteDoubleByteWidth",
            Self::NoExtraLineSpacing => "noExtraLineSpacing",
            Self::DoNotLeaveBackslashAlone => "doNotLeaveBackslashAlone",
            Self::UlTrailSpace => "ulTrailSpace",
            Self::DoNotExpandShiftReturn => "doNotExpandShiftReturn",
            Self::SpacingInWholePoints => "spacingInWholePoints",
            Self::LineWrapLikeWord6 => "lineWrapLikeWord6",
            Self::PrintBodyTextBeforeHeader => "printBodyTextBeforeHeader",
            Self::PrintColBlack => "printColBlack",
            Self::WpSpaceWidth => "wpSpaceWidth",
            Self::ShowBreaksInFrames => "showBreaksInFrames",
            Self::SubFontBySize => "subFontBySize",
            Self::SuppressBottomSpacing => "suppressBottomSpacing",
            Self::SuppressTopSpacing => "suppressTopSpacing",
            Self::SuppressSpacingAtTopOfPage => "suppressSpacingAtTopOfPage",
            Self::SuppressTopSpacingWP => "suppressTopSpacingWP",
            Self::SuppressSpBfAfterPgBrk => "suppressSpBfAfterPgBrk",
            Self::SwapBordersFacingPages => "swapBordersFacingPages",
            Self::ConvMailMergeEsc => "convMailMergeEsc",
            Self::TruncateFontHeightsLikeWP6 => "truncateFontHeightsLikeWP6",
            Self::MwSmallCaps => "mwSmallCaps",
            Self::UsePrinterMetrics => "usePrinterMetrics",
            Self::DoNotSuppressParagraphBorders => "doNotSuppressParagraphBorders",
            Self::WrapTrailSpaces => "wrapTrailSpaces",
            Self::FootnoteLayoutLikeWW8 => "footnoteLayoutLikeWW8",
            Self::ShapeLayoutLikeWW8 => "shapeLayoutLikeWW8",
            Self::AlignTablesRowByRow => "alignTablesRowByRow",
            Self::ForgetLastTabAlignment => "forgetLastTabAlignment",
            Self::AdjustLineHeightInTable => "adjustLineHeightInTable",
            Self::AutoSpaceLikeWord95 => "autoSpaceLikeWord95",
            Self::NoSpaceRaiseLower => "noSpaceRaiseLower",
            Self::DoNotUseHTMLParagraphAutoSpacing => "doNotUseHTMLParagraphAutoSpacing",
            Self::LayoutRawTableWidth => "layoutRawTableWidth",
            Self::LayoutTableRowsApart => "layoutTableRowsApart",
            Self::UseWord97LineBreakRules => "useWord97LineBreakRules",
            Self::DoNotBreakWrappedTables => "doNotBreakWrappedTables",
            Self::DoNotSnapToGridInCell => "doNotSnapToGridInCell",
            Self::SelectFldWithFirstOrLastChar => "selectFldWithFirstOrLastChar",
            Self::ApplyBreakingRules => "applyBreakingRules",
            Self::DoNotWrapTextWithPunct => "doNotWrapTextWithPunct",
            Self::DoNotUseEastAsianBreakRules => "doNotUseEastAsianBreakRules",
            Self::UseWord2002TableStyleRules => "useWord2002TableStyleRules",
            Self::GrowAutofit => "growAutofit",
            Self::UseFELayout => "useFELayout",
            Self::UseNormalStyleForList => "useNormalStyleForList",
            Self::DoNotUseIndentAsNumberingTabStop => "doNotUseIndentAsNumberingTabStop",
            Self::UseAltKinsokuLineBreakRules => "useAltKinsokuLineBreakRules",
            Self::AllowSpaceOfSameStyleInTable => "allowSpaceOfSameStyleInTable",
            Self::DoNotSuppressIndentation => "doNotSuppressIndentation",
            Self::DoNotAutofitConstrainedTables => "doNotAutofitConstrainedTables",
            Self::AutofitToFirstFixedWidthCell => "autofitToFirstFixedWidthCell",
            Self::UnderlineTabInNumList => "underlineTabInNumList",
            Self::DisplayHangulFixedWidth => "displayHangulFixedWidth",
            Self::SplitPgBreakAndParaMark => "splitPgBreakAndParaMark",
            Self::DoNotVertAlignCellWithSp => "doNotVertAlignCellWithSp",
            Self::DoNotBreakConstrainedForcedTable => "doNotBreakConstrainedForcedTable",
            Self::DoNotVertAlignInTxbx => "doNotVertAlignInTxbx",
            Self::UseAnsiKerningPairs => "useAnsiKerningPairs",
            Self::CachedColBalance => "cachedColBalance",
        }
    }

    /// Return the option one `w:compat` child names, when it names one.
    fn find(element: &BytesStart<'_>, prefixes: &[String]) -> Option<Self> {
        Self::ALL.iter().copied().find(|option| {
            is_word_element(
                element.name().as_ref(),
                option.local_name().as_bytes(),
                prefixes,
            )
        })
    }
}

/// `CT_Compat` child order, the on-off family followed by `w:compatSetting`.
const COMPAT_ORDER: &[&[u8]] = &[
    b"useSingleBorderforContiguousCells",
    b"wpJustification",
    b"noTabHangInd",
    b"noLeading",
    b"spaceForUL",
    b"noColumnBalance",
    b"balanceSingleByteDoubleByteWidth",
    b"noExtraLineSpacing",
    b"doNotLeaveBackslashAlone",
    b"ulTrailSpace",
    b"doNotExpandShiftReturn",
    b"spacingInWholePoints",
    b"lineWrapLikeWord6",
    b"printBodyTextBeforeHeader",
    b"printColBlack",
    b"wpSpaceWidth",
    b"showBreaksInFrames",
    b"subFontBySize",
    b"suppressBottomSpacing",
    b"suppressTopSpacing",
    b"suppressSpacingAtTopOfPage",
    b"suppressTopSpacingWP",
    b"suppressSpBfAfterPgBrk",
    b"swapBordersFacingPages",
    b"convMailMergeEsc",
    b"truncateFontHeightsLikeWP6",
    b"mwSmallCaps",
    b"usePrinterMetrics",
    b"doNotSuppressParagraphBorders",
    b"wrapTrailSpaces",
    b"footnoteLayoutLikeWW8",
    b"shapeLayoutLikeWW8",
    b"alignTablesRowByRow",
    b"forgetLastTabAlignment",
    b"adjustLineHeightInTable",
    b"autoSpaceLikeWord95",
    b"noSpaceRaiseLower",
    b"doNotUseHTMLParagraphAutoSpacing",
    b"layoutRawTableWidth",
    b"layoutTableRowsApart",
    b"useWord97LineBreakRules",
    b"doNotBreakWrappedTables",
    b"doNotSnapToGridInCell",
    b"selectFldWithFirstOrLastChar",
    b"applyBreakingRules",
    b"doNotWrapTextWithPunct",
    b"doNotUseEastAsianBreakRules",
    b"useWord2002TableStyleRules",
    b"growAutofit",
    b"useFELayout",
    b"useNormalStyleForList",
    b"doNotUseIndentAsNumberingTabStop",
    b"useAltKinsokuLineBreakRules",
    b"allowSpaceOfSameStyleInTable",
    b"doNotSuppressIndentation",
    b"doNotAutofitConstrainedTables",
    b"autofitToFirstFixedWidthCell",
    b"underlineTabInNumList",
    b"displayHangulFixedWidth",
    b"splitPgBreakAndParaMark",
    b"doNotVertAlignCellWithSp",
    b"doNotBreakConstrainedForcedTable",
    b"doNotVertAlignInTxbx",
    b"useAnsiKerningPairs",
    b"cachedColBalance",
    b"compatSetting",
];

/// Every top-level `w:settings` child in `xsd:sequence` order.
///
/// One table drives both schema-position insertion and the closed supported
/// set, so a new member is added in one place rather than three.
pub(crate) const SETTINGS_ORDER: &[&[u8]] = &[
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

/// The closed set of top-level `w:settings` children the typed model owns.
///
/// "Supported" is a checkable constant rather than a judgement. A name outside
/// this list is preservation-only and is never reported as a diagnostic, which
/// is what makes "no unmodeled supported children" decidable. The nested
/// supported names are `compat/compatSetting`, every `CompatibilityOption`
/// local name under `compat`, `docVars/docVar` and `MAIL_MERGE_MEMBERS`.
pub const SUPPORTED_SETTINGS: &[&str] = &[
    "view",
    "zoom",
    "removePersonalInformation",
    "removeDateAndTime",
    "displayBackgroundShape",
    "mirrorMargins",
    "gutterAtTop",
    "proofState",
    "linkStyles",
    "mailMerge",
    "trackRevisions",
    "doNotTrackMoves",
    "doNotTrackFormatting",
    "documentProtection",
    "defaultTabStop",
    "autoHyphenation",
    "consecutiveHyphenLimit",
    "hyphenationZone",
    "doNotHyphenateCaps",
    "defaultTableStyle",
    "evenAndOddHeaders",
    "bookFoldRevPrinting",
    "bookFoldPrinting",
    "bookFoldPrintingSheets",
    "characterSpacingControl",
    "updateFields",
    "compat",
    "docVars",
    "mathPr",
    "themeFontLang",
    "decimalSymbol",
    "listSeparator",
];

/// Supported paths whose schema allows more than one occurrence.
const SETTINGS_REPEATABLE: &[&[&str]] = &[&["docVars", "docVar"], &["compat", "compatSetting"]];

/// `CT_MailMerge` child order, including the three preservation-only members.
const MAIL_MERGE_ORDER: &[&[u8]] = &[
    b"mainDocumentType",
    b"linkToQuery",
    b"dataType",
    b"connectString",
    b"query",
    b"dataSource",
    b"headerSource",
    b"doNotSuppressBlankLines",
    b"destination",
    b"addressFieldName",
    b"mailSubject",
    b"mailAsAttachment",
    b"viewMergedData",
    b"activeRecord",
    b"checkErrors",
    b"odso",
];

/// The bounded `w:mailMerge` children this model authors, in schema order.
pub const MAIL_MERGE_MEMBERS: &[&str] = &[
    "mainDocumentType",
    "linkToQuery",
    "dataType",
    "connectString",
    "query",
    "doNotSuppressBlankLines",
    "destination",
    "addressFieldName",
    "mailSubject",
    "mailAsAttachment",
    "viewMergedData",
    "activeRecord",
    "checkErrors",
];

/// Return the supported name one element carries, when it carries one.
pub(crate) fn supported_name(
    names: &[&'static str],
    element: &BytesStart<'_>,
    prefixes: &[String],
) -> Option<&'static str> {
    names
        .iter()
        .copied()
        .find(|name| is_word_element(element.name().as_ref(), name.as_bytes(), prefixes))
}

pub(crate) fn capture_empty_element(element: &BytesStart<'_>) -> Result<Vec<u8>> {
    let mut raw = Vec::new();
    Writer::new(&mut raw).write_event(Event::Empty(element.to_owned().into_owned()))?;
    Ok(raw)
}

pub(crate) fn word_value(element: &BytesStart<'_>, prefixes: &[String]) -> Option<String> {
    word_attribute(element, b"val", prefixes).ok().flatten()
}

fn parse_twips(element: &BytesStart<'_>, prefixes: &[String]) -> Option<Twips> {
    let value = word_value(element, prefixes)?;
    parse_integer_measurement::<i32>(element.name().as_ref(), b"w:val", &value)
        .ok()
        .filter(|value| *value >= 0)
        .map(Twips)
}

pub(crate) fn parse_decimal(element: &BytesStart<'_>, prefixes: &[String]) -> Option<i32> {
    word_value(element, prefixes)?.parse::<i32>().ok()
}

fn parse_zoom(element: &BytesStart<'_>, prefixes: &[String]) -> Option<DocumentZoom> {
    let kind = match word_value(element, prefixes) {
        Some(value) => Some(ZoomKind::parse(&value)?),
        None => None,
    };
    let percent = match word_attribute(element, b"percent", prefixes).ok().flatten() {
        Some(value) => Some(value.trim_end_matches('%').parse::<u32>().ok()?),
        None => None,
    };
    Some(DocumentZoom { kind, percent })
}

fn parse_proof_state(element: &BytesStart<'_>, prefixes: &[String]) -> Option<DocumentProofState> {
    let attribute = |local| word_attribute(element, local, prefixes).ok().flatten();
    let spelling = match attribute(b"spelling") {
        Some(value) => Some(ProofState::parse(&value)?),
        None => None,
    };
    let grammar = match attribute(b"grammar") {
        Some(value) => Some(ProofState::parse(&value)?),
        None => None,
    };
    Some(DocumentProofState { spelling, grammar })
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
        value: word_value(element, prefixes)?,
    })
}

pub(crate) fn write_valued_setting(name: &str, value: &str) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    let mut element = BytesStart::new(name);
    element.push_attribute(("w:val", value));
    writer.write_event(Event::Empty(element))?;
    Ok(writer.into_inner())
}

fn write_zoom(value: &DocumentZoom) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    let mut element = BytesStart::new("w:zoom");
    if let Some(kind) = value.kind {
        element.push_attribute(("w:val", kind.as_str()));
    }
    let percent = value.percent.map(|percent| percent.to_string());
    if let Some(percent) = &percent {
        element.push_attribute(("w:percent", percent.as_str()));
    }
    writer.write_event(Event::Empty(element))?;
    Ok(writer.into_inner())
}

fn write_proof_state(value: &DocumentProofState) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    let mut element = BytesStart::new("w:proofState");
    if let Some(spelling) = value.spelling {
        element.push_attribute(("w:spelling", spelling.as_str()));
    }
    if let Some(grammar) = value.grammar {
        element.push_attribute(("w:grammar", grammar.as_str()));
    }
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

/// Wrap modeled children in one fixed-prefix group element.
fn wrap_group(name: &str, children: &[u8]) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::with_capacity(children.len() + 2 * name.len() + 5));
    writer.write_event(Event::Start(BytesStart::new(name)))?;
    writer.get_mut().extend_from_slice(children);
    writer.write_event(Event::End(BytesEnd::new(name)))?;
    Ok(writer.into_inner())
}

/// Join every serialized group child in order.
fn concatenate(children: &[(&'static str, Vec<u8>)]) -> Vec<u8> {
    let mut joined = Vec::new();
    for (_, bytes) in children {
        joined.extend_from_slice(bytes);
    }
    joined
}

/// Serialize every modeled `w:mailMerge` member, empty where absent.
fn mail_merge_children(value: &MailMerge) -> Result<Vec<(&'static str, Vec<u8>)>> {
    let valued = |local: &str, text: Option<&str>| -> Result<Vec<u8>> {
        match text {
            Some(text) => write_valued_setting(&format!("w:{local}"), text),
            None => Ok(Vec::new()),
        }
    };
    let toggled = |local: &str, enabled: Option<bool>| -> Result<Vec<u8>> {
        match enabled {
            Some(enabled) => write_toggle_setting(&format!("w:{local}"), enabled),
            None => Ok(Vec::new()),
        }
    };
    let active_record = value.active_record.map(|value| value.to_string());
    let check_errors = value.check_errors.map(|value| value.to_string());
    Ok(vec![
        (
            "mainDocumentType",
            valued(
                "mainDocumentType",
                value.main_document_type.map(MailMergeDocumentType::as_str),
            )?,
        ),
        ("linkToQuery", toggled("linkToQuery", value.link_to_query)?),
        ("dataType", valued("dataType", value.data_type.as_deref())?),
        (
            "connectString",
            valued("connectString", value.connect_string.as_deref())?,
        ),
        ("query", valued("query", value.query.as_deref())?),
        (
            "doNotSuppressBlankLines",
            toggled("doNotSuppressBlankLines", value.do_not_suppress_blank_lines)?,
        ),
        (
            "destination",
            valued(
                "destination",
                value.destination.map(MailMergeDestination::as_str),
            )?,
        ),
        (
            "addressFieldName",
            valued("addressFieldName", value.address_field_name.as_deref())?,
        ),
        (
            "mailSubject",
            valued("mailSubject", value.mail_subject.as_deref())?,
        ),
        (
            "mailAsAttachment",
            toggled("mailAsAttachment", value.mail_as_attachment)?,
        ),
        (
            "viewMergedData",
            toggled("viewMergedData", value.view_merged_data)?,
        ),
        (
            "activeRecord",
            valued("activeRecord", active_record.as_deref())?,
        ),
        (
            "checkErrors",
            valued("checkErrors", check_errors.as_deref())?,
        ),
    ])
}

/// Replace every modeled occurrence of one repeated child inside one group.
fn rewrite_group_setting(
    source: &[u8],
    group_local: &[u8],
    child_local: &[u8],
    modeled_children: Vec<u8>,
) -> Result<Vec<u8>> {
    let existing = find_top_level_children(source, group_local)?;
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
    rewrite_ordered_child(
        source,
        SETTINGS_ORDER,
        group_local,
        &replacement,
        is_word_element,
    )
}

/// Capture every top-level child with one local name, with its inherited scope.
pub(crate) fn find_top_level_children(
    source: &[u8],
    local: &[u8],
) -> Result<Vec<(Vec<u8>, Vec<String>)>> {
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

/// Replace one modeled child at its `xsd:sequence` position inside a group.
fn rewrite_group_child(
    source: &[u8],
    group_local: &[u8],
    child_local: &[u8],
    replacement: &[u8],
    order: &[&[u8]],
) -> Result<Vec<u8>> {
    let existing = find_top_level_children(source, group_local)?;
    if existing.is_empty() {
        return Ok(source.to_vec());
    }
    let mut rebuilt = Vec::new();
    for (index, (raw, inherited_prefixes)) in existing.into_iter().enumerate() {
        rebuilt.extend_from_slice(&rewrite_child_in_raw(
            &raw,
            child_local,
            if index == 0 { replacement } else { &[] },
            &inherited_prefixes,
            order,
        )?);
    }
    rewrite_ordered_child(
        source,
        SETTINGS_ORDER,
        group_local,
        &rebuilt,
        is_word_element,
    )
}

fn rewrite_child_in_raw(
    raw: &[u8],
    child_local: &[u8],
    replacement: &[u8],
    inherited_prefixes: &[String],
    order: &[&[u8]],
) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(raw.len() + replacement.len()));
    let mut prefixes = Vec::new();
    let mut depth = 0usize;
    let mut inserted = replacement.is_empty();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if depth == 0 => {
                prefixes = word_prefixes_at(&element, inherited_prefixes)?;
                writer.write_event(Event::Start(element.into_owned()))?;
                depth = 1;
            }
            Event::Empty(element) if depth == 0 => {
                if inserted {
                    writer.write_event(Event::Empty(element.into_owned()))?;
                } else {
                    let name = std::str::from_utf8(element.name().as_ref())?.to_owned();
                    writer.write_event(Event::Start(element.into_owned()))?;
                    writer.get_mut().extend_from_slice(replacement);
                    writer.write_event(Event::End(BytesEnd::new(name)))?;
                    inserted = true;
                }
            }
            Event::Start(element) if depth == 1 => {
                let child_prefixes = word_prefixes_at(&element, &prefixes)?;
                if is_word_element(element.name().as_ref(), child_local, &child_prefixes) {
                    if !inserted {
                        writer.get_mut().extend_from_slice(replacement);
                        inserted = true;
                    }
                    capture_element(&mut reader, &element)?;
                } else {
                    if !inserted && element_follows(order, child_local, &element, &child_prefixes) {
                        writer.get_mut().extend_from_slice(replacement);
                        inserted = true;
                    }
                    writer.write_event(Event::Start(element.into_owned()))?;
                    depth += 1;
                }
            }
            Event::Empty(element) if depth == 1 => {
                let child_prefixes = word_prefixes_at(&element, &prefixes)?;
                if is_word_element(element.name().as_ref(), child_local, &child_prefixes) {
                    if !inserted {
                        writer.get_mut().extend_from_slice(replacement);
                        inserted = true;
                    }
                } else {
                    if !inserted && element_follows(order, child_local, &element, &child_prefixes) {
                        writer.get_mut().extend_from_slice(replacement);
                        inserted = true;
                    }
                    writer.write_event(Event::Empty(element.into_owned()))?;
                }
            }
            Event::End(element) if depth == 1 => {
                if !inserted {
                    writer.get_mut().extend_from_slice(replacement);
                    inserted = true;
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

/// Create one empty top-level group at its schema position when it is absent.
fn ensure_top_level_group(source: &[u8], local: &[u8]) -> Result<Vec<u8>> {
    if !find_top_level_children(source, local)?.is_empty() {
        return Ok(source.to_vec());
    }
    let name = format!("w:{}", std::str::from_utf8(local)?);
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Start(BytesStart::new(name.as_str())))?;
    writer.write_event(Event::End(BytesEnd::new(name.as_str())))?;
    let group = writer.into_inner();
    rewrite_ordered_child(source, SETTINGS_ORDER, local, &group, is_word_element)
}

/// Report whether every top-level group with one name holds no element child.
fn top_level_group_is_childless(source: &[u8], local: &[u8]) -> Result<bool> {
    for (raw, _) in find_top_level_children(source, local)? {
        let mut reader = Reader::from_reader(raw.as_slice());
        reader.config_mut().trim_text(false);
        let mut seen_root = false;
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(_) | Event::Empty(_) if !seen_root => seen_root = true,
                Event::Start(_) | Event::Empty(_) => return Ok(false),
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
    }
    Ok(true)
}

/// Replace one child at its `xsd:sequence` position under a part root.
///
/// `matches` selects the modeled element, which lets the OfficeMath defaults
/// reuse the same insertion walk from a different namespace.
pub(crate) fn rewrite_ordered_child(
    source: &[u8],
    order: &[&[u8]],
    local: &[u8],
    replacement: &[u8],
    matches: fn(&[u8], &[u8], &[String]) -> bool,
) -> Result<Vec<u8>> {
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
                if matches(element.name().as_ref(), local, &prefixes) {
                    if !inserted {
                        writer.get_mut().extend_from_slice(replacement);
                        inserted = true;
                    }
                    capture_element(&mut reader, &element)?;
                } else {
                    if !inserted && element_follows(order, local, &element, &prefixes) {
                        writer.get_mut().extend_from_slice(replacement);
                        inserted = true;
                    }
                    writer.write_event(Event::Start(element.into_owned()))?;
                    depth += 1;
                }
            }
            Event::Empty(element) if depth == 1 => {
                let prefixes = word_prefixes_at(&element, &root_prefixes)?;
                if matches(element.name().as_ref(), local, &prefixes) {
                    if !inserted {
                        writer.get_mut().extend_from_slice(replacement);
                        inserted = true;
                    }
                } else {
                    if !inserted && element_follows(order, local, &element, &prefixes) {
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

/// Report whether one element sits after `local` in an `xsd:sequence` table.
pub(crate) fn element_follows(
    order: &[&[u8]],
    local: &[u8],
    element: &BytesStart<'_>,
    prefixes: &[String],
) -> bool {
    let Some(target_index) = order.iter().position(|candidate| *candidate == local) else {
        return false;
    };
    order
        .iter()
        .skip(target_index + 1)
        .any(|candidate| is_word_element(element.name().as_ref(), candidate, prefixes))
}

fn rewrite_math_properties(source: &[u8], properties: Option<&MathProperties>) -> Result<Vec<u8>> {
    let replacement = match properties {
        Some(properties) => {
            let mut writer = Writer::new(Vec::new());
            properties.write_xml(&mut writer)?;
            writer.into_inner()
        }
        None => Vec::new(),
    };
    rewrite_ordered_child(
        source,
        SETTINGS_ORDER,
        b"mathPr",
        &replacement,
        is_math_element,
    )
}

pub(crate) fn parse_toggle(element: &BytesStart<'_>, prefixes: &[String]) -> Result<Option<bool>> {
    Ok(match word_attribute(element, b"val", prefixes)? {
        Some(value) => parse_on_off(&value),
        None => Some(true),
    })
}

pub(crate) fn write_toggle_setting(name: &str, value: bool) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    let mut element = BytesStart::new(name);
    if !value {
        element.push_attribute(("w:val", "false"));
    }
    writer.write_event(Event::Empty(element))?;
    Ok(writer.into_inner())
}

pub(crate) fn ensure_fixed_word_prefix(root: &mut BytesStart<'_>) -> Result<()> {
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

fn parse_document_variable(
    element: &BytesStart<'_>,
    prefixes: &[String],
) -> Option<DocumentVariable> {
    Some(DocumentVariable {
        name: word_attribute(element, b"name", prefixes).ok().flatten()?,
        value: word_value(element, prefixes)?,
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

pub(crate) fn word_attribute(
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

pub(crate) fn parse_on_off(value: &str) -> Option<bool> {
    match value {
        "1" | "true" | "on" => Some(true),
        "0" | "false" | "off" => Some(false),
        _ => None,
    }
}

fn write_document_protection(protection: &DocumentProtection) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
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
    Ok(writer.into_inner())
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

    /// One well-formed instance of every supported name, in schema order.
    fn every_supported_child() -> String {
        format!(
            concat!(
                r#"<w:settings xmlns:w="{word}" xmlns:m="{math}">"#,
                r#"<w:view w:val="print"/>"#,
                r#"<w:zoom w:val="fullPage" w:percent="120"/>"#,
                r#"<w:removePersonalInformation/>"#,
                r#"<w:removeDateAndTime w:val="false"/>"#,
                r#"<w:displayBackgroundShape/>"#,
                r#"<w:mirrorMargins/>"#,
                r#"<w:gutterAtTop/>"#,
                r#"<w:proofState w:spelling="clean" w:grammar="dirty"/>"#,
                r#"<w:linkStyles/>"#,
                r#"<w:mailMerge>"#,
                r#"<w:mainDocumentType w:val="formLetters"/>"#,
                r#"<w:linkToQuery/>"#,
                r#"<w:dataType w:val="native"/>"#,
                r#"<w:connectString w:val="DSN=Contacts"/>"#,
                r#"<w:query w:val="SELECT * FROM People"/>"#,
                r#"<w:doNotSuppressBlankLines/>"#,
                r#"<w:destination w:val="printer"/>"#,
                r#"<w:addressFieldName w:val="Address"/>"#,
                r#"<w:mailSubject w:val="Invitation"/>"#,
                r#"<w:mailAsAttachment/>"#,
                r#"<w:viewMergedData/>"#,
                r#"<w:activeRecord w:val="3"/>"#,
                r#"<w:checkErrors w:val="2"/>"#,
                r#"</w:mailMerge>"#,
                r#"<w:trackRevisions/>"#,
                r#"<w:doNotTrackMoves/>"#,
                r#"<w:doNotTrackFormatting/>"#,
                r#"<w:documentProtection w:edit="readOnly"/>"#,
                r#"<w:defaultTabStop w:val="720"/>"#,
                r#"<w:autoHyphenation/>"#,
                r#"<w:consecutiveHyphenLimit w:val="2"/>"#,
                r#"<w:hyphenationZone w:val="360"/>"#,
                r#"<w:doNotHyphenateCaps/>"#,
                r#"<w:defaultTableStyle w:val="TableNormal"/>"#,
                r#"<w:evenAndOddHeaders/>"#,
                r#"<w:bookFoldRevPrinting/>"#,
                r#"<w:bookFoldPrinting/>"#,
                r#"<w:bookFoldPrintingSheets w:val="4"/>"#,
                r#"<w:characterSpacingControl w:val="doNotCompress"/>"#,
                r#"<w:updateFields/>"#,
                r#"<w:compat><w:noTabHangInd/><w:cachedColBalance w:val="false"/>"#,
                r#"<w:compatSetting w:name="compatibilityMode" w:uri="urn:office" w:val="15"/>"#,
                r#"</w:compat>"#,
                r#"<w:docVars><w:docVar w:name="Customer" w:val="Ada"/></w:docVars>"#,
                r#"<m:mathPr><m:mathFont m:val="Cambria Math"/></m:mathPr>"#,
                r#"<w:themeFontLang w:val="en-US"/>"#,
                r#"<w:decimalSymbol w:val="."/>"#,
                r#"<w:listSeparator w:val=","/>"#,
                r#"</w:settings>"#,
            ),
            word = W_NS,
            math = crate::namespace::M_NS,
        )
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
        let mut settings = CT_Settings::new();
        settings
            .set_document_protection(DocumentProtection {
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
            })
            .unwrap();
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

        let xml = format!(
            r#"<q:settings xmlns:q="{W_NS}" xmlns:x="urn:foreign"><x:kept/></q:settings>"#,
        );
        let mut absent = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        absent.set_update_fields(None).unwrap();
        assert_eq!(absent.to_xml().unwrap(), xml.as_bytes());
    }

    #[test]
    fn update_fields_ambiguous_forms_are_not_rewritten() {
        for xml in [
            format!(
                r#"<w:settings xmlns:w="{W_NS}"><w:updateFields/><w:updateFields w:val="false"/></w:settings>"#,
            ),
            format!(
                r#"<w:settings xmlns:w="{W_NS}"><w:updateFields w:val="not-a-toggle"/></w:settings>"#,
            ),
        ] {
            let mut settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
            assert_eq!(settings.update_fields(), None);
            let before = settings.to_xml().unwrap();
            assert!(settings.set_update_fields(Some(true)).is_err());
            assert_eq!(settings.to_xml().unwrap(), before);
        }
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

    #[test]
    fn supported_settings_is_a_strict_subsequence_of_the_order_table() {
        let mut seen = Vec::new();
        for name in SUPPORTED_SETTINGS {
            assert!(!seen.contains(name), "duplicate supported name {name}");
            seen.push(name);
        }
        assert_eq!(SUPPORTED_SETTINGS.len(), 32);

        let mut cursor = 0usize;
        for name in SUPPORTED_SETTINGS {
            let position = SETTINGS_ORDER
                .iter()
                .position(|candidate| *candidate == name.as_bytes())
                .unwrap_or_else(|| panic!("{name} is missing from SETTINGS_ORDER"));
            assert!(
                position >= cursor,
                "{name} breaks the schema order of SETTINGS_ORDER"
            );
            cursor = position + 1;
        }
    }

    #[test]
    fn every_supported_name_is_projected_by_from_xml() {
        let xml = every_supported_child();
        let settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();

        assert_eq!(settings.view(), Some(DocumentView::Print));
        assert_eq!(
            settings.zoom(),
            Some(DocumentZoom {
                kind: Some(ZoomKind::FullPage),
                percent: Some(120),
            })
        );
        assert_eq!(settings.remove_personal_information(), Some(true));
        assert_eq!(settings.remove_date_and_time(), Some(false));
        assert_eq!(settings.display_background_shape(), Some(true));
        assert_eq!(settings.mirror_margins(), Some(true));
        assert_eq!(settings.gutter_at_top(), Some(true));
        assert_eq!(
            settings.proof_state(),
            Some(DocumentProofState {
                spelling: Some(ProofState::Clean),
                grammar: Some(ProofState::Dirty),
            })
        );
        assert_eq!(settings.link_styles(), Some(true));
        let mail_merge = settings.mail_merge().unwrap();
        assert_eq!(
            mail_merge.main_document_type,
            Some(MailMergeDocumentType::FormLetters)
        );
        assert_eq!(mail_merge.link_to_query, Some(true));
        assert_eq!(mail_merge.data_type.as_deref(), Some("native"));
        assert_eq!(mail_merge.connect_string.as_deref(), Some("DSN=Contacts"));
        assert_eq!(mail_merge.query.as_deref(), Some("SELECT * FROM People"));
        assert_eq!(mail_merge.do_not_suppress_blank_lines, Some(true));
        assert_eq!(mail_merge.destination, Some(MailMergeDestination::Printer));
        assert_eq!(mail_merge.address_field_name.as_deref(), Some("Address"));
        assert_eq!(mail_merge.mail_subject.as_deref(), Some("Invitation"));
        assert_eq!(mail_merge.mail_as_attachment, Some(true));
        assert_eq!(mail_merge.view_merged_data, Some(true));
        assert_eq!(mail_merge.active_record, Some(3));
        assert_eq!(mail_merge.check_errors, Some(2));
        assert_eq!(settings.track_revisions(), Some(true));
        assert_eq!(settings.do_not_track_moves(), Some(true));
        assert_eq!(settings.do_not_track_formatting(), Some(true));
        assert_eq!(
            settings.document_protection().map(|value| value.mode),
            Some(ProtectionMode::ReadOnly)
        );
        assert_eq!(settings.default_tab_stop(), Some(Twips(720)));
        assert!(settings.automatic_hyphenation());
        assert_eq!(settings.consecutive_hyphen_limit(), Some(2));
        assert_eq!(settings.hyphenation_zone(), Some(Twips(360)));
        assert_eq!(settings.do_not_hyphenate_caps(), Some(true));
        assert_eq!(settings.default_table_style(), Some("TableNormal"));
        assert!(settings.even_and_odd_headers());
        assert_eq!(settings.book_fold_rev_printing(), Some(true));
        assert_eq!(settings.book_fold_printing(), Some(true));
        assert_eq!(settings.book_fold_printing_sheets(), Some(4));
        assert_eq!(
            settings.character_spacing_control(),
            Some(CharacterSpacingControl::DoNotCompress)
        );
        assert_eq!(settings.update_fields(), Some(true));
        assert_eq!(
            settings.compatibility_options(),
            [
                (CompatibilityOption::NoTabHangInd, true),
                (CompatibilityOption::CachedColBalance, false),
            ]
        );
        assert_eq!(settings.compatibility_settings().len(), 1);
        assert_eq!(settings.document_variable("Customer"), Some("Ada"));
        assert_eq!(
            settings.math_properties().unwrap().math_font.as_deref(),
            Some("Cambria Math")
        );
        assert_eq!(
            settings.theme_font_language().unwrap().latin.as_deref(),
            Some("en-US")
        );
        assert_eq!(settings.decimal_symbol(), Some("."));
        assert_eq!(settings.list_separator(), Some(","));

        assert_eq!(settings.diagnostics(), &[]);
        assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());
    }

    #[test]
    fn decimal_twips_settings_round_and_keep_their_bytes() {
        let xml = format!(
            r#"<w:settings xmlns:w="{W_NS}"><w:defaultTabStop w:val="720.0"/><w:hyphenationZone w:val="356.5"/></w:settings>"#,
        );
        let settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();

        assert_eq!(settings.default_tab_stop(), Some(Twips(720)));
        assert_eq!(settings.hyphenation_zone(), Some(Twips(357)));
        assert!(settings.diagnostics().is_empty());
        assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());
    }

    #[test]
    fn duplicate_and_malformed_supported_children_report_diagnostics() {
        let xml = format!(
            r#"<w:settings xmlns:w="{W_NS}"><w:view w:val="print"/><w:view w:val="normal"/><w:zoom w:val="nonsense"/><w:defaultTabStop w:val="-12"/><w:compat><w:noTabHangInd w:val="maybe"/></w:compat><w:docVars><w:docVar w:name="A"/></w:docVars></w:settings>"#,
        );
        let settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();

        assert_eq!(
            settings.diagnostics(),
            &[
                SettingsDiagnostic {
                    path: vec!["view"],
                    occurrences: 2,
                    reason: SettingsDiagnosticReason::Duplicated,
                },
                SettingsDiagnostic {
                    path: vec!["zoom"],
                    occurrences: 1,
                    reason: SettingsDiagnosticReason::Malformed,
                },
                SettingsDiagnostic {
                    path: vec!["defaultTabStop"],
                    occurrences: 1,
                    reason: SettingsDiagnosticReason::Malformed,
                },
                SettingsDiagnostic {
                    path: vec!["compat", "noTabHangInd"],
                    occurrences: 1,
                    reason: SettingsDiagnosticReason::Malformed,
                },
                SettingsDiagnostic {
                    path: vec!["docVars", "docVar"],
                    occurrences: 1,
                    reason: SettingsDiagnosticReason::Malformed,
                },
            ]
        );
        assert!(settings.view().is_none());
        assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());
    }

    #[test]
    fn unsupported_settings_children_are_never_diagnostics() {
        let xml = format!(
            r#"<w:settings xmlns:w="{W_NS}" xmlns:w15="urn:w15" xmlns:x="urn:foreign"><w:rsids><w:rsid w:val="00AA00AA"/></w:rsids><w:clrSchemeMapping w:bg1="light1"/><w:shapeDefaults><x:anything/></w:shapeDefaults><w15:docId w15:val="{{GUID}}"/><x:view x:val="print"/><x:view x:val="normal"/></w:settings>"#,
        );
        let settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(settings.diagnostics(), &[]);
        assert!(settings.view().is_none());
        assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());
    }

    #[test]
    fn compatibility_option_covers_the_complete_compat_on_off_set() {
        assert_eq!(CompatibilityOption::ALL.len(), 65);
        let mut names = Vec::new();
        for option in CompatibilityOption::ALL {
            let name = option.local_name();
            assert!(!names.contains(&name), "duplicate compat option {name}");
            names.push(name);
        }
        assert_eq!(COMPAT_ORDER.len(), names.len() + 1);
        for (index, name) in names.iter().enumerate() {
            assert_eq!(COMPAT_ORDER[index], name.as_bytes());
        }
        assert_eq!(COMPAT_ORDER[names.len()], b"compatSetting");

        let children = names
            .iter()
            .map(|name| format!("<w:{name}/>"))
            .collect::<String>();
        let xml =
            format!(r#"<w:settings xmlns:w="{W_NS}"><w:compat>{children}</w:compat></w:settings>"#);
        let settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(settings.diagnostics(), &[]);
        assert_eq!(settings.compatibility_options().len(), names.len());
        for option in CompatibilityOption::ALL {
            assert_eq!(settings.compatibility_option(*option), Some(true));
        }
    }

    #[test]
    fn document_protection_authoring_records_caller_metadata_verbatim() {
        let protection = DocumentProtection {
            mode: ProtectionMode::Forms,
            enforcement: Some(true),
            formatting: None,
            provider_type: Some(CryptProviderType::RsaAes),
            algorithm_class: Some(CryptAlgorithmClass::Hash),
            algorithm_type: Some(CryptAlgorithmType::Any),
            algorithm_sid: Some(4),
            spin_count: Some(100_000),
            hash: Some("CALLER-HASH".to_owned()),
            salt: Some("CALLER-SALT".to_owned()),
        };
        let xml =
            format!(r#"<w:settings xmlns:w="{W_NS}"><w:defaultTabStop w:val="720"/></w:settings>"#);

        let mut first = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        first.set_document_protection(protection.clone()).unwrap();
        let mut second = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        second.set_document_protection(protection.clone()).unwrap();
        assert_eq!(first.to_xml().unwrap(), second.to_xml().unwrap());

        let output = String::from_utf8(first.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"w:hash="CALLER-HASH""#), "{output}");
        assert!(output.contains(r#"w:salt="CALLER-SALT""#), "{output}");
        assert!(output.contains(r#"w:cryptSpinCount="100000""#), "{output}");
        assert!(
            output.find("<w:documentProtection").unwrap() < output.find("defaultTabStop").unwrap(),
            "{output}"
        );

        let reopened = CT_Settings::from_xml(output.as_bytes()).unwrap();
        assert_eq!(reopened.document_protection(), Some(&protection));
        assert_eq!(reopened.diagnostics(), &[]);
    }

    #[test]
    fn removing_a_duplicated_member_fails_and_changes_nothing() {
        let xml = format!(
            r#"<w:settings xmlns:w="{W_NS}"><w:mirrorMargins/><w:mirrorMargins w:val="false"/></w:settings>"#,
        );
        let mut settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        assert!(settings.remove_mirror_margins().is_err());
        assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());

        let absent = format!(r#"<w:settings xmlns:w="{W_NS}"><w:kept/></w:settings>"#);
        let mut settings = CT_Settings::from_xml(absent.as_bytes()).unwrap();
        assert_eq!(settings.remove_mirror_margins().unwrap(), None);
        assert_eq!(settings.to_xml().unwrap(), absent.as_bytes());
    }

    #[test]
    fn a_self_closing_group_does_not_leak_its_namespace_scope() {
        // `xmlns:z` is declared on the self-closing group, so it is out of
        // scope for every following sibling.
        let xml = format!(
            r#"<w:settings xmlns:w="{W_NS}"><w:compat xmlns:z="{W_NS}"/><z:view z:val="print"/><w:mirrorMargins/></w:settings>"#,
        );
        let settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(settings.view(), None);
        assert_eq!(settings.mirror_margins(), Some(true));
        assert_eq!(settings.diagnostics(), &[]);
        assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());
    }

    #[test]
    fn mail_merge_members_land_between_their_preserved_neighbours() {
        let xml = format!(
            concat!(
                r#"<w:settings xmlns:w="{word}" xmlns:r="urn:rel">"#,
                r#"<w:mailMerge><w:dataSource r:id="rId7"/><w:odso><w:udl w:val="keep"/></w:odso></w:mailMerge>"#,
                r#"</w:settings>"#
            ),
            word = W_NS,
        );
        let mut settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        settings
            .set_mail_merge(MailMerge {
                main_document_type: Some(MailMergeDocumentType::FormLetters),
                check_errors: Some(2),
                ..MailMerge::default()
            })
            .unwrap();
        let output = String::from_utf8(settings.to_xml().unwrap()).unwrap();
        assert!(
            output.contains(r#"<w:dataSource r:id="rId7"/>"#),
            "{output}"
        );
        assert!(output.contains(r#"<w:udl w:val="keep"/>"#), "{output}");
        assert!(
            output.find("mainDocumentType").unwrap() < output.find("dataSource").unwrap(),
            "{output}"
        );
        assert!(
            output.find("dataSource").unwrap() < output.find("checkErrors").unwrap(),
            "{output}"
        );
        assert!(
            output.find("checkErrors").unwrap() < output.find("<w:odso").unwrap(),
            "{output}"
        );

        let reopened = CT_Settings::from_xml(output.as_bytes()).unwrap();
        assert_eq!(reopened.mail_merge().unwrap().check_errors, Some(2));
        assert_eq!(reopened.diagnostics(), &[]);
    }

    #[test]
    fn note_policy_preserves_unknown_note_and_section_xml() {
        let xml = format!(
            r#"<q:settings xmlns:q="{W_NS}" xmlns:x="urn:producer"><q:footnotePr><q:pos q:val="beneathText"/><x:keep x:flag="exact"/><q:numFmt q:val="decimal"/><q:endnote q:id="9"/><q:footnote q:id="-1"/></q:footnotePr><q:endnotePr><q:numStart q:val="4"/></q:endnotePr></q:settings>"#
        );
        let mut settings = CT_Settings::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(settings.to_xml().unwrap(), xml.as_bytes());
        let mut footnote = settings.footnote_properties().unwrap().clone();
        assert_eq!(footnote.pos.as_deref(), Some("beneathText"));
        assert_eq!(footnote.special_references, vec![-1]);
        footnote.num_start = Some(3);
        settings.set_footnote_properties(footnote).unwrap();
        let written = String::from_utf8(settings.to_xml().unwrap()).unwrap();
        assert!(written.contains(r#"<x:keep x:flag="exact"/>"#), "{written}");
        assert!(written.contains(r#"<q:endnote q:id="9"/>"#), "{written}");
        assert!(written.contains(r#"<w:numStart w:val="3"/>"#), "{written}");
        let reopened = CT_Settings::from_xml(written.as_bytes()).unwrap();
        assert_eq!(reopened.footnote_properties().unwrap().num_start, Some(3));
        assert_eq!(
            reopened.footnote_properties().unwrap().special_references,
            vec![-1]
        );
        assert_eq!(reopened.endnote_properties().unwrap().num_start, Some(4));

        let section_xml = format!(
            r#"<q:document xmlns:q="{W_NS}" xmlns:x="urn:producer"><q:body><q:p/><q:sectPr><q:footnotePr><q:pos q:val="pageBottom"/><x:keep x:flag="section-exact"/><q:numFmt q:val="decimal"/></q:footnotePr></q:sectPr></q:body></q:document>"#
        );
        let mut document = crate::document::CT_Document::from_xml(section_xml.as_bytes()).unwrap();
        document
            .body
            .sect_pr
            .as_mut()
            .unwrap()
            .footnote_pr
            .as_mut()
            .unwrap()
            .num_start = Some(5);
        let rewritten = String::from_utf8(document.to_xml().unwrap()).unwrap();
        assert!(
            rewritten.contains(r#"<x:keep x:flag="section-exact"/>"#),
            "{rewritten}"
        );
        assert!(
            rewritten.contains(r#"<w:numStart w:val="5"/>"#),
            "{rewritten}"
        );
        let reopened_section =
            crate::document::CT_Document::from_xml(rewritten.as_bytes()).unwrap();
        assert_eq!(
            reopened_section
                .body
                .sect_pr
                .unwrap()
                .footnote_pr
                .unwrap()
                .num_start,
            Some(5)
        );
    }
}
