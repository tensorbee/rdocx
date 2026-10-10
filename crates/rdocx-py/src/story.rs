//! Header and footer stories, document settings, notes and page decoration.
//!
//! A header or footer handle names its section and variant, as python-docx's
//! `section.header` does, and resolves the story it shows each time it is
//! used. Paragraph and run handles inside one carry that slot in their path,
//! so the body `Paragraph` and `Run` types serve every story.

use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList};
use smallvec::{SmallVec, smallvec};

use crate::document::PyDocument;
use crate::paragraph::{ParagraphLocation, PyParagraph, style_id_of_type};
use crate::run::PyRun;
use crate::{rdocx_to_pyerr, stale_to_pyerr};

/// One header or footer variant of one section.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct StorySlot {
    pub(crate) section: usize,
    pub(crate) kind: rdocx::HeaderFooterKind,
    pub(crate) variant: rdocx::HdrFtrType,
}

impl StorySlot {
    pub(crate) fn new(
        section: usize,
        kind: rdocx::HeaderFooterKind,
        variant: rdocx::HdrFtrType,
    ) -> Self {
        Self {
            section,
            kind,
            variant,
        }
    }

    /// The path segment that carries this slot.
    pub(crate) fn segment(self) -> PathSeg {
        let kind = usize::from(self.kind != rdocx::HeaderFooterKind::Header);
        let variant = match self.variant {
            rdocx::HdrFtrType::Default => 0,
            rdocx::HdrFtrType::First => 1,
            rdocx::HdrFtrType::Even => 2,
        };
        PathSeg::Story(self.section * 6 + kind * 3 + variant)
    }

    /// The slot a [`Self::segment`] carries.
    pub(crate) fn from_segment(code: usize) -> Self {
        let kind = if (code % 6) / 3 == 0 {
            rdocx::HeaderFooterKind::Header
        } else {
            rdocx::HeaderFooterKind::Footer
        };
        let variant = match code % 3 {
            0 => rdocx::HdrFtrType::Default,
            1 => rdocx::HdrFtrType::First,
            _ => rdocx::HdrFtrType::Even,
        };
        Self::new(code / 6, kind, variant)
    }

    /// The python-docx `Section` attribute that names this slot.
    fn attribute(self) -> &'static str {
        let header = self.kind == rdocx::HeaderFooterKind::Header;
        match (header, self.variant) {
            (true, rdocx::HdrFtrType::Default) => "header",
            (true, rdocx::HdrFtrType::First) => "first_page_header",
            (true, rdocx::HdrFtrType::Even) => "even_page_header",
            (false, rdocx::HdrFtrType::Default) => "footer",
            (false, rdocx::HdrFtrType::First) => "first_page_footer",
            (false, rdocx::HdrFtrType::Even) => "even_page_footer",
        }
    }

    fn kind_name(self) -> &'static str {
        if self.kind == rdocx::HeaderFooterKind::Header {
            "header"
        } else {
            "footer"
        }
    }

    /// How to fetch a fresh handle for a paragraph of this story.
    pub(crate) fn recovery_hint(self, at: rdocx::HeaderFooterParagraph) -> String {
        let story = format!("doc.sections[{}].{}", self.section, self.attribute());
        match at {
            rdocx::HeaderFooterParagraph::Direct(paragraph) => {
                format!("Re-fetch it with {story}.paragraphs[{paragraph}].")
            }
            rdocx::HeaderFooterParagraph::Cell {
                table,
                row,
                cell,
                paragraph,
            } => format!(
                "Re-fetch it with {story}.tables[{table}].cell({row}, {cell}).paragraphs[{paragraph}]."
            ),
        }
    }

    /// The story this slot shows now, its own or the one it inherits.
    pub(crate) fn story(
        self,
        py: Python<'_>,
        document: &rdocx::Document,
    ) -> PyResult<Option<rdocx::StoryId>> {
        if self.section >= document.section_count() {
            return Err(PyIndexError::new_err("section index out of range"));
        }
        document
            .section_story(self.section, self.kind, self.variant)
            .map(|story| story.map(|story| story.story().clone()))
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    /// Capture the path of one paragraph of this story.
    pub(crate) fn path(
        self,
        document: &PyDocument,
        at: rdocx::HeaderFooterParagraph,
    ) -> ContentPath {
        let mut segments: SmallVec<[PathSeg; 5]> = smallvec![self.segment()];
        match at {
            rdocx::HeaderFooterParagraph::Direct(paragraph) => {
                segments.push(PathSeg::Para(paragraph));
            }
            rdocx::HeaderFooterParagraph::Cell {
                table,
                row,
                cell,
                paragraph,
            } => segments.extend([
                PathSeg::Body(table),
                PathSeg::Row(row),
                PathSeg::Cell(cell),
                PathSeg::Para(paragraph),
            ]),
        }
        document.revisions.capture(segments)
    }
}

/// Read one header or footer paragraph, or `None` when the story has no
/// paragraph there.
pub(crate) fn read_paragraph<R>(
    py: Python<'_>,
    document: &PyDocument,
    slot: StorySlot,
    at: rdocx::HeaderFooterParagraph,
    read: impl FnOnce(rdocx::ParagraphRef<'_>) -> R,
) -> PyResult<Option<R>> {
    let Some(story) = slot.story(py, &document.inner)? else {
        return Ok(None);
    };
    document
        .inner
        .read_header_footer_paragraph(&story, at, read)
        .map_err(|error| rdocx_to_pyerr(py, error))
}

/// Edit one header or footer paragraph in place, or return `None` when the
/// story has no paragraph there.
pub(crate) fn edit_paragraph<R>(
    py: Python<'_>,
    document: &mut PyDocument,
    slot: StorySlot,
    at: rdocx::HeaderFooterParagraph,
    edit: impl FnOnce(&mut rdocx::Paragraph<'_>) -> R,
) -> PyResult<Option<R>> {
    let Some(story) = slot.story(py, &document.inner)? else {
        return Ok(None);
    };
    let result = document
        .inner
        .edit_header_footer_paragraph(&story, at, edit)
        .map_err(|error| rdocx_to_pyerr(py, error))?;
    if result.is_some() {
        show_written_variant(py, document, slot)?;
    }
    Ok(result)
}

/// Replace text within one header or footer paragraph, as
/// `Paragraph.replace_text` does in the body.
pub(crate) fn replace_text(
    py: Python<'_>,
    document: &mut PyDocument,
    slot: StorySlot,
    at: rdocx::HeaderFooterParagraph,
    old: &str,
    new: &str,
    expect: Option<usize>,
) -> PyResult<usize> {
    let story = slot
        .story(py, &document.inner)?
        .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
    let count = document
        .inner
        .try_replace_text_in_header_footer_paragraph(&story, at, old, new, expect)
        .map_err(|error| rdocx_to_pyerr(py, error))?
        .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?
        .map_err(|mismatch| crate::replacement_count_to_pyerr(py, &mismatch, false))?;
    if count > 0 {
        show_written_variant(py, document, slot)?;
        document.revisions.bump();
    }
    Ok(count)
}

/// The relationship of an external hyperlink from a header or footer, which
/// its own part owns.
pub(crate) fn hyperlink_relationship(
    py: Python<'_>,
    document: &mut PyDocument,
    slot: StorySlot,
    at: rdocx::HeaderFooterParagraph,
    url: &str,
) -> PyResult<String> {
    let story = slot.story(py, &document.inner)?;
    let story = story
        .filter(|story| {
            matches!(
                document
                    .inner
                    .read_header_footer_paragraph(story, at, |_| ()),
                Ok(Some(()))
            )
        })
        .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
    document
        .inner
        .add_hyperlink_relationship_to_story(&story, url)
        .map_err(|error| rdocx_to_pyerr(py, error))
}

/// Make Word show the variant just written: a first-page story needs its
/// section's different first page, and an even-page story needs the
/// document's odd and even pages setting.
fn show_written_variant(
    py: Python<'_>,
    document: &mut PyDocument,
    slot: StorySlot,
) -> PyResult<()> {
    match slot.variant {
        rdocx::HdrFtrType::First => {
            let shown = document
                .inner
                .section(slot.section)
                .and_then(|section| section.different_first_page())
                .unwrap_or(false);
            if !shown && let Some(mut section) = document.inner.section_mut(slot.section) {
                section.set_different_first_page(true);
            }
        }
        rdocx::HdrFtrType::Even if !document.inner.even_and_odd_headers() => {
            document
                .inner
                .set_even_and_odd_headers(true)
                .map_err(|error| rdocx_to_pyerr(py, error))?;
        }
        rdocx::HdrFtrType::Default | rdocx::HdrFtrType::Even => {}
    }
    Ok(())
}

/// The story a slot shows, giving it one when no section up to it has one.
///
/// As in python-docx, a missing default story is added to the first section,
/// which every later section then inherits. A first-page story goes to the
/// slot's own section instead, since a native first-page story also turns on
/// its section's title page, which would change the first page of every
/// earlier section. A new story holds one empty paragraph, since a header or
/// footer part needs at least one block.
fn story_or_create(
    py: Python<'_>,
    document: &mut PyDocument,
    slot: StorySlot,
) -> PyResult<rdocx::StoryId> {
    if let Some(story) = slot.story(py, &document.inner)? {
        return Ok(story);
    }
    let section = match slot.variant {
        rdocx::HdrFtrType::First => slot.section,
        rdocx::HdrFtrType::Default | rdocx::HdrFtrType::Even => 0,
    };
    install_story(py, document, StorySlot { section, ..slot })?;
    slot.story(py, &document.inner)?
        .ok_or_else(|| PyValueError::new_err("the new header or footer story was not found"))
}

/// Give one section variant a new story holding one empty paragraph.
fn install_story(py: Python<'_>, document: &mut PyDocument, slot: StorySlot) -> PyResult<()> {
    let inner = &mut document.inner;
    let story = py
        .detach(|| inner.create_section_story(slot.section, slot.kind, slot.variant))
        .map_err(|error| rdocx_to_pyerr(py, error))?;
    let index = document
        .inner
        .add_header_footer_paragraph(&story, "")
        .map_err(|error| rdocx_to_pyerr(py, error))?;
    if let Some(style) = story_paragraph_style(&document.inner, slot) {
        document
            .inner
            .edit_header_footer_paragraph(
                &story,
                rdocx::HeaderFooterParagraph::Direct(index),
                |paragraph| paragraph.set_style(&style),
            )
            .map_err(|error| rdocx_to_pyerr(py, error))?;
    }
    document.revisions.bump();
    Ok(())
}

/// Word's "Header" or "Footer" paragraph style, which a new header or footer
/// paragraph takes when the document defines it.
fn story_paragraph_style(document: &rdocx::Document, slot: StorySlot) -> Option<String> {
    let name = if slot.kind == rdocx::HeaderFooterKind::Header {
        "Header"
    } else {
        "Footer"
    };
    style_id_of_type(document, name, rdocx::StyleType::Paragraph).ok()
}

/// Split a page-number template into literal text and field names.
fn page_number_pieces(template: &str) -> PyResult<Vec<(bool, String)>> {
    const FIELDS: [&str; 3] = ["PAGE", "NUMPAGES", "SECTIONPAGES"];
    let mut pieces = Vec::new();
    let mut rest = template;
    while let Some(open) = rest.find('{') {
        let close = rest[open..]
            .find('}')
            .map(|offset| open + offset)
            .ok_or_else(|| {
                PyValueError::new_err("page number template has a '{' without its closing '}'")
            })?;
        if open > 0 {
            pieces.push((false, rest[..open].to_owned()));
        }
        let name = &rest[open + 1..close];
        if !FIELDS.contains(&name) {
            return Err(PyValueError::new_err(format!(
                "page number template field {{{name}}} must be {{PAGE}}, {{NUMPAGES}} or {{SECTIONPAGES}}"
            )));
        }
        pieces.push((true, name.to_owned()));
        rest = &rest[close + 1..];
    }
    if rest.contains('}') {
        return Err(PyValueError::new_err(
            "page number template has a '}' without its opening '{'",
        ));
    }
    if !rest.is_empty() {
        pieces.push((false, rest.to_owned()));
    }
    if !pieces.iter().any(|(field, _)| *field) {
        return Err(PyValueError::new_err(
            "page number template names no field, such as \"Page {PAGE} of {NUMPAGES}\"",
        ));
    }
    Ok(pieces)
}

/// The sections a handle saw: inserting or removing a section renumbers
/// them, so a handle that names a section by its index goes stale.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct SectionLayout {
    epoch: u64,
    count: usize,
}

impl SectionLayout {
    pub(crate) fn of(document: &PyDocument) -> Self {
        Self {
            epoch: document.section_epoch,
            count: document.inner.section_count(),
        }
    }

    /// Raise `StaleElementError` when the document's sections changed since
    /// the handle was made.
    pub(crate) fn check(
        self,
        py: Python<'_>,
        document: &PyDocument,
        kind: &str,
        section: usize,
        attribute: &str,
    ) -> PyResult<()> {
        if Self::of(document) == self {
            return Ok(());
        }
        Err(stale_to_pyerr(
            py,
            oxml_py_support::StaleElementError {
                element_kind: kind.to_owned(),
                captured_revision: self.epoch,
                current_revision: document.section_epoch,
                recovery_hint: format!(
                    "A section was inserted or removed. Re-fetch it with doc.sections[{section}]{attribute}."
                ),
            },
        ))
    }
}

/// A section's header or footer, python-docx's `_Header` and `_Footer`.
///
/// The handle names a section and a variant, not a story part, so it stays
/// valid across edits. Its story is the section's own, or the one it
/// inherits from an earlier section when `is_linked_to_previous` is true.
/// Writing into an inherited story edits the earlier section's story, as in
/// python-docx and Word.
#[pyclass(name = "HeaderFooter", frozen)]
pub struct PyHeaderFooter {
    document: Py<PyDocument>,
    slot: StorySlot,
    layout: SectionLayout,
}

impl PyHeaderFooter {
    pub(crate) fn new(document: Py<PyDocument>, slot: StorySlot, layout: SectionLayout) -> Self {
        Self {
            document,
            slot,
            layout,
        }
    }

    /// Refuse to act once the sections this handle numbers have changed.
    fn check(&self, py: Python<'_>) -> PyResult<()> {
        self.layout.check(
            py,
            &self.document.borrow(py),
            "header or footer",
            self.slot.section,
            &format!(".{}", self.slot.attribute()),
        )
    }

    fn paragraph_handle(
        &self,
        py: Python<'_>,
        document: &PyDocument,
        at: rdocx::HeaderFooterParagraph,
    ) -> PyResult<Py<PyParagraph>> {
        Py::new(
            py,
            PyParagraph::new(self.document.clone_ref(py), self.slot.path(document, at)),
        )
    }
}

#[pymethods]
impl PyHeaderFooter {
    /// `"header"` or `"footer"`.
    #[getter]
    fn kind(&self) -> &'static str {
        self.slot.kind_name()
    }

    /// `"default"`, `"first"` or `"even"`.
    #[getter]
    fn variant(&self) -> &'static str {
        self.slot.variant.to_str()
    }

    #[getter]
    fn section_index(&self) -> usize {
        self.slot.section
    }

    /// Whether this section shows the story of an earlier section rather
    /// than its own. Setting it to true drops the section's own story.
    /// Setting it to false gives the section a new story holding one empty
    /// paragraph, as python-docx does.
    #[getter]
    fn is_linked_to_previous(&self, py: Python<'_>) -> PyResult<bool> {
        self.check(py)?;
        let document = self.document.borrow(py);
        if self.slot.section >= document.inner.section_count() {
            return Err(PyIndexError::new_err("section index out of range"));
        }
        let story = document
            .inner
            .section_story(self.slot.section, self.slot.kind, self.slot.variant)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        Ok(story.is_none_or(|story| story.is_inherited()))
    }

    #[setter]
    fn set_is_linked_to_previous(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        if self.is_linked_to_previous(py)? == value {
            return Ok(());
        }
        let mut document = self.document.borrow_mut(py);
        let slot = self.slot;
        if value {
            let inner = &mut document.inner;
            py.detach(|| inner.inherit_section_story(slot.section, slot.kind, slot.variant))
                .map_err(|error| rdocx_to_pyerr(py, error))?;
            document.revisions.bump();
            Ok(())
        } else {
            install_story(py, &mut document, slot)
        }
    }

    /// The story's direct paragraphs, as python-docx lists them. A section
    /// without any story up to it gets one first, as in python-docx.
    #[getter]
    fn paragraphs<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        self.check(py)?;
        let count = {
            let mut document = self.document.borrow_mut(py);
            let story = story_or_create(py, &mut document, self.slot)?;
            document
                .inner
                .header_footer_paragraph_count(&story)
                .map_err(|error| rdocx_to_pyerr(py, error))?
        };
        let document = self.document.borrow(py);
        let items = (0..count)
            .map(|index| {
                self.paragraph_handle(py, &document, rdocx::HeaderFooterParagraph::Direct(index))
            })
            .collect::<PyResult<Vec<_>>>()?;
        PyList::new(py, items)
    }

    /// Append a paragraph after the story's last block, as python-docx does.
    #[pyo3(signature = (text = "", style = None))]
    fn add_paragraph(
        &self,
        py: Python<'_>,
        text: &str,
        style: Option<&str>,
    ) -> PyResult<Py<PyParagraph>> {
        let at = self.append_paragraph(py, text, style)?;
        let document = self.document.borrow(py);
        self.paragraph_handle(py, &document, at)
    }

    /// Append a `rows` by `cols` table after the story's last block. The
    /// columns share `width`, or 6.5 inches when it is not given.
    #[pyo3(signature = (rows, cols, width = None))]
    fn add_table(
        &self,
        py: Python<'_>,
        rows: usize,
        cols: usize,
        width: Option<i64>,
    ) -> PyResult<Py<PyHeaderFooterTable>> {
        self.check(py)?;
        let mut document = self.document.borrow_mut(py);
        let story = story_or_create(py, &mut document, self.slot)?;
        let table = document
            .inner
            .add_header_footer_table(&story, rows, cols, width.map(rdocx::Length::emu))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        show_written_variant(py, &mut document, self.slot)?;
        // An appended table moves no paragraph or table index, so live
        // handles stay valid.
        let path = document
            .revisions
            .capture(smallvec![self.slot.segment(), PathSeg::Body(table)]);
        Py::new(
            py,
            PyHeaderFooterTable {
                document: self.document.clone_ref(py),
                slot: self.slot,
                table,
                path,
            },
        )
    }

    /// The story's direct tables.
    #[getter]
    fn tables<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        self.check(py)?;
        let document = self.document.borrow(py);
        let count = match self.slot.story(py, &document.inner)? {
            Some(story) => document
                .inner
                .header_footer_table_count(&story)
                .map_err(|error| rdocx_to_pyerr(py, error))?,
            None => 0,
        };
        let items = (0..count)
            .map(|table| {
                Py::new(
                    py,
                    PyHeaderFooterTable {
                        document: self.document.clone_ref(py),
                        slot: self.slot,
                        table,
                        path: document
                            .revisions
                            .capture(smallvec![self.slot.segment(), PathSeg::Body(table)]),
                    },
                )
            })
            .collect::<PyResult<Vec<_>>>()?;
        PyList::new(py, items)
    }

    /// Append a paragraph that shows the page number, such as "Page 3 of 7".
    ///
    /// `template` is literal text with `{PAGE}`, `{NUMPAGES}` and
    /// `{SECTIONPAGES}` fields, which Word, Google Docs and the rdocx renderer
    /// fill in on every page. Use it on `section.footer`, and also on
    /// `section.first_page_footer` when the section has a different first
    /// page, or on `section.even_page_footer` when odd and even pages differ.
    #[pyo3(signature = (template = "Page {PAGE} of {NUMPAGES}", *, alignment = Some(1)))]
    fn add_page_number(
        &self,
        py: Python<'_>,
        template: &str,
        alignment: Option<i32>,
    ) -> PyResult<Py<PyParagraph>> {
        let pieces = page_number_pieces(template)?;
        let alignment = alignment
            .map(crate::formatting::alignment_from_int)
            .transpose()?;
        let at = self.page_number_paragraph(py)?;
        {
            let mut document = self.document.borrow_mut(py);
            let unstyled = read_paragraph(py, &document, self.slot, at, |paragraph| {
                paragraph.style_id().is_none()
            })?
            .unwrap_or(false);
            let style = story_paragraph_style(&document.inner, self.slot).filter(|_| unstyled);
            let written = edit_paragraph(py, &mut document, self.slot, at, |paragraph| {
                if let Some(style) = style.as_deref() {
                    paragraph.set_style(style);
                }
                if let Some(alignment) = alignment {
                    paragraph.set_alignment(alignment);
                }
                for (field, value) in &pieces {
                    if *field {
                        paragraph.add_run("").add_field(value, "1")?;
                    } else {
                        paragraph.add_run(value);
                    }
                }
                Ok::<_, rdocx::Error>(())
            })?
            .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
            written.map_err(|error| rdocx_to_pyerr(py, error))?;
        }
        let document = self.document.borrow(py);
        self.paragraph_handle(py, &document, at)
    }
}

impl PyHeaderFooter {
    /// The paragraph a page number goes in: the story's lone empty
    /// paragraph, such as the one a new story holds, or a new last one.
    fn page_number_paragraph(&self, py: Python<'_>) -> PyResult<rdocx::HeaderFooterParagraph> {
        self.check(py)?;
        {
            let mut document = self.document.borrow_mut(py);
            let story = story_or_create(py, &mut document, self.slot)?;
            let lone = rdocx::HeaderFooterParagraph::Direct(0);
            let only_paragraph = document
                .inner
                .header_footer_paragraph_count(&story)
                .map_err(|error| rdocx_to_pyerr(py, error))?
                == 1
                && document
                    .inner
                    .header_footer_table_count(&story)
                    .map_err(|error| rdocx_to_pyerr(py, error))?
                    == 0;
            let empty = document
                .inner
                .read_header_footer_paragraph(&story, lone, |paragraph| {
                    paragraph.run_count() == 0 && paragraph.text().is_empty()
                })
                .map_err(|error| rdocx_to_pyerr(py, error))?
                .unwrap_or(false);
            if only_paragraph && empty {
                return Ok(lone);
            }
        }
        self.append_paragraph(py, "", None)
    }

    /// Append one paragraph and return where it is.
    fn append_paragraph(
        &self,
        py: Python<'_>,
        text: &str,
        style: Option<&str>,
    ) -> PyResult<rdocx::HeaderFooterParagraph> {
        self.check(py)?;
        let mut document = self.document.borrow_mut(py);
        // An explicit style wins; otherwise the paragraph takes the story's
        // "Header" or "Footer" style when the document defines it.
        let style = match style {
            Some(style) => Some(style_id_of_type(
                &document.inner,
                style,
                rdocx::StyleType::Paragraph,
            )?),
            None => story_paragraph_style(&document.inner, self.slot),
        };
        let story = story_or_create(py, &mut document, self.slot)?;
        let index = document
            .inner
            .add_header_footer_paragraph(&story, text)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let at = rdocx::HeaderFooterParagraph::Direct(index);
        if let Some(style) = style {
            document
                .inner
                .edit_header_footer_paragraph(&story, at, |paragraph| paragraph.set_style(&style))
                .map_err(|error| rdocx_to_pyerr(py, error))?;
        }
        show_written_variant(py, &mut document, self.slot)?;
        // An appended paragraph moves no paragraph or table index, so live
        // handles stay valid.
        Ok(at)
    }
}

/// A table of a header or footer story.
#[pyclass(name = "HeaderFooterTable", frozen)]
pub struct PyHeaderFooterTable {
    document: Py<PyDocument>,
    slot: StorySlot,
    table: usize,
    path: ContentPath,
}

impl PyHeaderFooterTable {
    fn validate(&self, py: Python<'_>) -> PyResult<()> {
        let document = self.document.borrow(py);
        self.path
            .validate_revision(
                document.revisions.current(),
                "header or footer table",
                &format!(
                    "Re-fetch it with doc.sections[{}].{}.tables[{}].",
                    self.slot.section,
                    self.slot.attribute(),
                    self.table
                ),
            )
            .map_err(|error| stale_to_pyerr(py, error))
    }

    fn read<R>(&self, py: Python<'_>, read: impl FnOnce(rdocx::TableRef<'_>) -> R) -> PyResult<R> {
        self.validate(py)?;
        let document = self.document.borrow(py);
        let story = self
            .slot
            .story(py, &document.inner)?
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
        document
            .inner
            .read_header_footer_table(&story, self.table, read)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))
    }

    fn cell_handle(
        &self,
        py: Python<'_>,
        row: usize,
        col: usize,
    ) -> PyResult<Py<PyHeaderFooterCell>> {
        Py::new(
            py,
            PyHeaderFooterCell {
                document: self.document.clone_ref(py),
                slot: self.slot,
                table: self.table,
                row,
                cell: col,
                path: self.path.clone(),
            },
        )
    }
}

#[pymethods]
impl PyHeaderFooterTable {
    /// The rows, each with its `cells`, as python-docx `Table.rows` lists them.
    #[getter]
    fn rows<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let counts = self.read(py, |table| {
            (0..table.row_count())
                .map(|row| table.row(row).map_or(0, |row| row.cell_count()))
                .collect::<Vec<_>>()
        })?;
        let rows = counts
            .into_iter()
            .enumerate()
            .map(|(row, cells)| {
                let cells = (0..cells)
                    .map(|col| self.cell_handle(py, row, col))
                    .collect::<PyResult<Vec<_>>>()?;
                Py::new(py, PyHeaderFooterRow { cells })
            })
            .collect::<PyResult<Vec<_>>>()?;
        PyList::new(py, rows)
    }

    /// The table style ID, or `None`.
    #[getter]
    fn style(&self, py: Python<'_>) -> PyResult<Option<String>> {
        self.read(py, |table| table.style_id().map(str::to_owned))
    }

    /// Set the table style by ID or name, such as "Table Grid".
    #[setter]
    fn set_style(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let style = style_id_of_type(&document.inner, value, rdocx::StyleType::Table)?;
        let story = self
            .slot
            .story(py, &document.inner)?
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
        document
            .inner
            .edit_header_footer_table(&story, self.table, |table| table.set_style(&style))
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
        show_written_variant(py, &mut document, self.slot)
    }

    /// The number of rows.
    #[getter]
    fn row_count(&self, py: Python<'_>) -> PyResult<usize> {
        self.read(py, |table| table.row_count())
    }

    /// The number of grid columns.
    #[getter]
    fn column_count(&self, py: Python<'_>) -> PyResult<usize> {
        self.read(py, |table| table.column_count())
    }

    /// The cell at `row` and `col`, as python-docx `Table.cell` returns it.
    fn cell(&self, py: Python<'_>, row: usize, col: usize) -> PyResult<Py<PyHeaderFooterCell>> {
        let exists = self.read(py, |table| table.cell(row, col).is_some())?;
        if !exists {
            return Err(PyIndexError::new_err("cell index out of range"));
        }
        self.cell_handle(py, row, col)
    }
}

/// A row of a header or footer table, python-docx's `_Row` with its `cells`.
#[pyclass(name = "HeaderFooterRow", frozen)]
pub struct PyHeaderFooterRow {
    cells: Vec<Py<PyHeaderFooterCell>>,
}

#[pymethods]
impl PyHeaderFooterRow {
    #[getter]
    fn cells<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        PyList::new(py, self.cells.iter().map(|cell| cell.clone_ref(py)))
    }
}

/// A cell of a header or footer table.
#[pyclass(name = "HeaderFooterCell", frozen)]
pub struct PyHeaderFooterCell {
    document: Py<PyDocument>,
    slot: StorySlot,
    table: usize,
    row: usize,
    cell: usize,
    path: ContentPath,
}

impl PyHeaderFooterCell {
    fn story(&self, py: Python<'_>, document: &PyDocument) -> PyResult<rdocx::StoryId> {
        self.path
            .validate_revision(
                document.revisions.current(),
                "header or footer cell",
                &format!(
                    "Re-fetch it with doc.sections[{}].{}.tables[{}].cell({}, {}).",
                    self.slot.section,
                    self.slot.attribute(),
                    self.table,
                    self.row,
                    self.cell
                ),
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        self.slot
            .story(py, &document.inner)?
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))
    }

    fn edit<R>(&self, py: Python<'_>, edit: impl FnOnce(&mut rdocx::Cell<'_>) -> R) -> PyResult<R> {
        let mut document = self.document.borrow_mut(py);
        let story = self.story(py, &document)?;
        let (row, cell) = (self.row, self.cell);
        let result = document
            .inner
            .edit_header_footer_table(&story, self.table, |table| {
                table.cell(row, cell).map(|mut cell| edit(&mut cell))
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .flatten()
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?;
        show_written_variant(py, &mut document, self.slot)?;
        Ok(result)
    }

    fn paragraph_count(&self, py: Python<'_>) -> PyResult<usize> {
        let document = self.document.borrow(py);
        let story = self.story(py, &document)?;
        let (row, cell) = (self.row, self.cell);
        document
            .inner
            .read_header_footer_table(&story, self.table, |table| {
                table.cell(row, cell).map(|cell| cell.paragraph_count())
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .flatten()
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))
    }

    fn paragraph_handle(&self, py: Python<'_>, paragraph: usize) -> PyResult<Py<PyParagraph>> {
        let document = self.document.borrow(py);
        let at = rdocx::HeaderFooterParagraph::Cell {
            table: self.table,
            row: self.row,
            cell: self.cell,
            paragraph,
        };
        Py::new(
            py,
            PyParagraph::new(self.document.clone_ref(py), self.slot.path(&document, at)),
        )
    }
}

#[pymethods]
impl PyHeaderFooterCell {
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        let document = self.document.borrow(py);
        let story = self.story(py, &document)?;
        let (row, cell) = (self.row, self.cell);
        document
            .inner
            .read_header_footer_table(&story, self.table, |table| {
                table.cell(row, cell).map(|cell| cell.text())
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .flatten()
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))
    }

    /// Replace the cell content with one paragraph holding `text`.
    #[setter]
    fn set_text(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        check_xml_text("cell text", text)?;
        // Only this cell's paragraphs change, and a handle to one of them
        // reaches the paragraph now at its index, or raises IndexError.
        self.edit(py, |cell| cell.set_text(text))
    }

    #[getter]
    fn paragraphs<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let count = self.paragraph_count(py)?;
        let items = (0..count)
            .map(|paragraph| self.paragraph_handle(py, paragraph))
            .collect::<PyResult<Vec<_>>>()?;
        PyList::new(py, items)
    }

    #[pyo3(signature = (text = ""))]
    fn add_paragraph(&self, py: Python<'_>, text: &str) -> PyResult<Py<PyParagraph>> {
        check_xml_text("paragraph text", text)?;
        let paragraph = self.edit(py, |cell| {
            let index = cell.paragraph_count();
            cell.add_paragraph(text);
            index
        })?;
        self.paragraph_handle(py, paragraph)
    }
}

/// Document settings, python-docx's `Settings`.
#[pyclass(name = "Settings", frozen)]
pub struct PySettings {
    document: Py<PyDocument>,
}

impl PySettings {
    pub(crate) fn new(document: Py<PyDocument>) -> Self {
        Self { document }
    }
}

#[pymethods]
impl PySettings {
    /// Whether even pages show the even-page headers and footers, Word's
    /// `w:evenAndOddHeaders`. Writing into an even-page header or footer
    /// turns it on.
    #[getter]
    fn odd_and_even_pages_header_footer(&self, py: Python<'_>) -> bool {
        self.document.borrow(py).inner.even_and_odd_headers()
    }

    #[setter]
    fn set_odd_and_even_pages_header_footer(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.document
            .borrow_mut(py)
            .inner
            .set_even_and_odd_headers(value)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }
}

/// The run break a python-docx `WD_BREAK` value names.
pub(crate) fn break_kind(value: i32) -> PyResult<rdocx::run::BreakKind> {
    use rdocx::run::BreakKind;
    match value {
        6 => Ok(BreakKind::Line),
        7 => Ok(BreakKind::Page),
        8 => Ok(BreakKind::Column),
        9 => Ok(BreakKind::TextWrapping(rdocx::BreakClear::Left)),
        10 => Ok(BreakKind::TextWrapping(rdocx::BreakClear::Right)),
        11 => Ok(BreakKind::TextWrapping(rdocx::BreakClear::All)),
        _ => Err(PyValueError::new_err(
            "break type must be WD_BREAK.LINE, PAGE, COLUMN, LINE_CLEAR_LEFT, \
             LINE_CLEAR_RIGHT or LINE_CLEAR_ALL, start a section with Document.insert_section",
        )),
    }
}

/// Reject a page or column break in a header or footer run, which Word
/// does not honour there.
pub(crate) fn check_break_location(
    location: ParagraphLocation,
    kind: rdocx::run::BreakKind,
) -> PyResult<()> {
    if matches!(location, ParagraphLocation::Story { .. })
        && matches!(
            kind,
            rdocx::run::BreakKind::Page | rdocx::run::BreakKind::Column
        )
    {
        return Err(PyValueError::new_err(
            "a page or column break has no effect in a header or footer, use WD_BREAK.LINE",
        ));
    }
    Ok(())
}

/// The body paragraph a note reference goes at the end of: a paragraph, or
/// its last run.
pub(crate) fn note_reference_paragraph(
    py: Python<'_>,
    document: &Py<PyDocument>,
    target: &Bound<'_, PyAny>,
) -> PyResult<usize> {
    let (location, run) = if let Ok(paragraph) = target.cast::<PyParagraph>() {
        let paragraph = paragraph.borrow();
        if !paragraph.belongs_to(py, document) {
            return Err(PyValueError::new_err(
                "the paragraph belongs to a different document",
            ));
        }
        (paragraph.validate(py)?, None)
    } else if let Ok(run) = target.cast::<PyRun>() {
        let run = run.borrow();
        if !run.belongs_to(py, document) {
            return Err(PyValueError::new_err(
                "the run belongs to a different document",
            ));
        }
        let (location, index) = run.validate(py)?;
        (location, Some(index))
    } else {
        return Err(PyTypeError::new_err(
            "a note reference goes after a Paragraph or a Run",
        ));
    };
    let paragraph = match location {
        ParagraphLocation::Body(paragraph) => paragraph,
        ParagraphLocation::Cell { .. } => {
            return Err(PyValueError::new_err(
                "a note reference must be in a body paragraph, not in a table cell",
            ));
        }
        ParagraphLocation::Story { .. } => {
            return Err(PyValueError::new_err(
                "a note reference must be in a body paragraph, not in a header or footer",
            ));
        }
    };
    if let Some(run) = run {
        let count = document
            .borrow(py)
            .inner
            .paragraph(paragraph)
            .map_or(0, |paragraph| paragraph.run_count());
        if run + 1 != count {
            return Err(PyValueError::new_err(
                "a note reference goes at the end of its paragraph: pass the paragraph, or call \
                 add_footnote right after adding the run it follows",
            ));
        }
    }
    Ok(paragraph)
}

pub(crate) fn note_family(name: &str) -> PyResult<rdocx::NoteFamily> {
    match name {
        "footnote" => Ok(rdocx::NoteFamily::Footnote),
        "endnote" => Ok(rdocx::NoteFamily::Endnote),
        _ => Err(PyValueError::new_err(format!(
            "note kind must be footnote or endnote, not {name:?}"
        ))),
    }
}

pub(crate) fn note_policy(
    family: rdocx::NoteFamily,
    number_format: &str,
    start: u32,
    restart: &str,
    placement: Option<&str>,
) -> PyResult<rdocx::NotePolicy> {
    let format = match number_format {
        "decimal" => rdocx::NoteNumberFormat::Decimal,
        "upperRoman" => rdocx::NoteNumberFormat::UpperRoman,
        "lowerRoman" => rdocx::NoteNumberFormat::LowerRoman,
        "upperLetter" => rdocx::NoteNumberFormat::UpperLetter,
        "lowerLetter" => rdocx::NoteNumberFormat::LowerLetter,
        _ => {
            return Err(PyValueError::new_err(
                "number format must be decimal, upperRoman, lowerRoman, upperLetter or lowerLetter",
            ));
        }
    };
    let restart = match restart {
        "continuous" => rdocx::NoteRestart::Continuous,
        "eachSect" => rdocx::NoteRestart::EachSection,
        "eachPage" => rdocx::NoteRestart::EachPage,
        _ => {
            return Err(PyValueError::new_err(
                "restart must be continuous, eachSect or eachPage",
            ));
        }
    };
    let placement = match (family, placement) {
        (rdocx::NoteFamily::Footnote, None | Some("pageBottom")) => {
            rdocx::NotePlacement::PageBottom
        }
        (rdocx::NoteFamily::Footnote, Some("beneathText")) => rdocx::NotePlacement::BeneathText,
        (rdocx::NoteFamily::Endnote, None | Some("docEnd")) => rdocx::NotePlacement::DocumentEnd,
        (rdocx::NoteFamily::Endnote, Some("sectEnd")) => rdocx::NotePlacement::SectionEnd,
        (rdocx::NoteFamily::Footnote, Some(_)) => {
            return Err(PyValueError::new_err(
                "footnote placement must be pageBottom or beneathText",
            ));
        }
        (rdocx::NoteFamily::Endnote, Some(_)) => {
            return Err(PyValueError::new_err(
                "endnote placement must be docEnd or sectEnd",
            ));
        }
    };
    if start == 0 {
        return Err(PyValueError::new_err("note numbering starts at 1 or more"));
    }
    Ok(rdocx::NotePolicy {
        format,
        start,
        restart,
        placement,
    })
}

/// Six upper-case hexadecimal digits from an `RGBColor` or a string such as
/// `"FFF2CC"` or `"#fff2cc"`, or `auto` when `allow_auto` is set.
pub(crate) fn hex_color(value: &Bound<'_, PyAny>, allow_auto: bool) -> PyResult<String> {
    if let Ok((red, green, blue)) = value.extract::<(u8, u8, u8)>() {
        return Ok(format!("{red:02X}{green:02X}{blue:02X}"));
    }
    let Ok(text) = value.extract::<String>() else {
        return Err(PyTypeError::new_err(
            "a colour is an RGBColor or six hexadecimal digits, such as \"FFF2CC\"",
        ));
    };
    if allow_auto && text.eq_ignore_ascii_case("auto") {
        return Ok("auto".to_owned());
    }
    let digits = text.strip_prefix('#').unwrap_or(&text);
    if digits.len() == 6 && digits.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Ok(digits.to_ascii_uppercase());
    }
    Err(PyValueError::new_err(format!(
        "colour {text:?} must be an RGBColor or six hexadecimal digits, such as \"FFF2CC\"{}",
        if allow_auto { " or \"auto\"" } else { "" }
    )))
}

/// A page border side.
pub(crate) fn page_border_edge(
    style: &str,
    width: i64,
    color: Option<String>,
    space: i64,
) -> PyResult<rdocx_oxml::CT_BorderEdge> {
    let value = rdocx_oxml::shared::ST_Border::from_str(style)
        .map_err(|_| PyValueError::new_err(format!("unknown border style {style:?}")))?;
    // Word draws a line border from a quarter point to six points, in
    // eighths of a point, and its space from 0 to 31 points.
    let eighths = width * 8 / 12_700;
    if !(2..=96).contains(&eighths) {
        return Err(PyValueError::new_err(
            "page border width must be from Pt(0.25) to Pt(12)",
        ));
    }
    let points = space / 12_700;
    if !(0..=31).contains(&points) {
        return Err(PyValueError::new_err(
            "page border space must be from Pt(0) to Pt(31)",
        ));
    }
    Ok(rdocx_oxml::CT_BorderEdge {
        val: value,
        sz: Some(eighths as u32),
        space: Some(points as u32),
        color: Some(color.unwrap_or_else(|| "auto".to_owned())),
        extra_attributes: Vec::new(),
    })
}

/// Reject text that XML 1.0 cannot hold, before it reaches a part.
fn check_xml_text(name: &str, text: &str) -> PyResult<()> {
    match text.chars().find(|&c| {
        c < ' ' && !matches!(c, '\t' | '\n' | '\r') || matches!(c, '\u{FFFE}' | '\u{FFFF}')
    }) {
        Some(c) => Err(PyValueError::new_err(format!(
            "{name} holds U+{:04X}, which XML cannot carry",
            c as u32
        ))),
        None => Ok(()),
    }
}
