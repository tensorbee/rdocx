//! Shared path, revision, and unit support for Python OOXML bindings.

use smallvec::SmallVec;
use thiserror::Error;

use oxml_core::Length;

/// One index step in a Word or PowerPoint content handle.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PathSeg {
    /// A slide in a presentation.
    Slide(usize),
    /// A shape in a slide or group shape.
    Shape(usize),
    /// An item in the document body's block content.
    Body(usize),
    /// A row in a table.
    Row(usize),
    /// A cell in a table row.
    Cell(usize),
    /// A paragraph in its parent container.
    Para(usize),
    /// A run in a paragraph.
    Run(usize),
    /// A story outside the main flow, such as a Word header or footer. The
    /// binding that captures the path defines what the number names.
    Story(usize),
}

/// An index path paired with the document revision at which it was captured.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ContentPath {
    /// Ordered index steps from the document root to the content.
    pub segs: SmallVec<[PathSeg; 5]>,
    /// Document revision at handle creation.
    pub revision: u64,
}

impl ContentPath {
    /// Capture a path at `revision`.
    pub const fn new(segs: SmallVec<[PathSeg; 5]>, revision: u64) -> Self {
        Self { segs, revision }
    }

    /// Reject a handle captured before or after `current_revision`.
    pub fn validate_revision(
        &self,
        current_revision: u64,
        element_kind: &str,
        recovery_hint: &str,
    ) -> Result<(), StaleElementError> {
        if self.revision == current_revision {
            return Ok(());
        }

        Err(StaleElementError {
            element_kind: element_kind.to_owned(),
            captured_revision: self.revision,
            current_revision,
            recovery_hint: recovery_hint.to_owned(),
        })
    }
}

/// A content handle whose captured revision no longer matches its document.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[error(
    "{element_kind} handle was created at document revision {captured_revision}, but the document is now at revision {current_revision} (a structural change invalidated it). {recovery_hint}"
)]
pub struct StaleElementError {
    /// Human-readable handle kind, such as `paragraph` or `run`.
    pub element_kind: String,
    /// Revision captured when the handle was created.
    pub captured_revision: u64,
    /// Revision owned by the document when the handle was used.
    pub current_revision: u64,
    /// Caller-supplied guidance for obtaining a fresh handle.
    pub recovery_hint: String,
}

/// Monotonic revision owned by a Python document wrapper.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub struct RevisionCounter {
    current: u64,
}

impl RevisionCounter {
    /// Start at revision zero.
    pub const fn new() -> Self {
        Self { current: 0 }
    }

    /// Return the current document revision.
    pub const fn current(self) -> u64 {
        self.current
    }

    /// Capture `segs` at the current document revision.
    pub const fn capture(&self, segs: SmallVec<[PathSeg; 5]>) -> ContentPath {
        ContentPath::new(segs, self.current)
    }

    /// Bump after a successful structural mutation and return the new revision.
    pub fn bump(&mut self) -> u64 {
        self.current = self
            .current
            .checked_add(1)
            .expect("document revision counter overflowed");
        self.current
    }
}

/// Convert inches to EMU with the canonical truncation behavior.
pub fn emu_from_inches(value: f64) -> i64 {
    Length::inches(value).to_emu()
}

/// Convert centimetres to EMU with the canonical truncation behavior.
pub fn emu_from_centimetres(value: f64) -> i64 {
    Length::cm(value).to_emu()
}

/// Convert millimetres to EMU with the canonical truncation behavior.
pub fn emu_from_millimetres(value: f64) -> i64 {
    Length::mm(value).to_emu()
}

/// Convert points to EMU with the canonical truncation behavior.
pub fn emu_from_points(value: f64) -> i64 {
    Length::pt(value).to_emu()
}

/// Convert twips to EMU.
pub fn emu_from_twips(value: i32) -> i64 {
    Length::twips(value).to_emu()
}

/// Convert EMU to inches.
pub fn inches_from_emu(value: i64) -> f64 {
    Length::emu(value).to_inches()
}

/// Convert EMU to centimetres.
pub fn centimetres_from_emu(value: i64) -> f64 {
    Length::emu(value).to_cm()
}

/// Convert EMU to millimetres.
pub fn millimetres_from_emu(value: i64) -> f64 {
    Length::emu(value).to_cm() * 10.0
}

/// Convert EMU to points.
pub fn points_from_emu(value: i64) -> f64 {
    Length::emu(value).to_pt()
}

/// Convert EMU to twips.
pub fn twips_from_emu(value: i64) -> i32 {
    Length::emu(value).to_twips()
}

#[cfg(test)]
mod tests {
    use super::*;
    use smallvec::smallvec;

    #[test]
    fn stale_path_reports_both_revisions() {
        let path = ContentPath::new(smallvec![PathSeg::Body(0), PathSeg::Para(3)], 4);

        let error: StaleElementError = path
            .validate_revision(5, "paragraph", "Re-fetch it with doc.paragraphs[i].")
            .unwrap_err();

        assert_eq!(error.captured_revision, 4);
        assert_eq!(error.current_revision, 5);
        assert_eq!(
            error.to_string(),
            "paragraph handle was created at document revision 4, but the document is now at revision 5 (a structural change invalidated it). Re-fetch it with doc.paragraphs[i]."
        );
    }

    #[test]
    fn matching_revision_accepts_content_path() {
        let path = ContentPath::new(smallvec![PathSeg::Body(0), PathSeg::Para(3)], 4);

        assert!(
            path.validate_revision(4, "paragraph", "Re-fetch it with doc.paragraphs[i].")
                .is_ok()
        );
    }

    #[test]
    fn revision_counter_bumps_after_successful_structure_change() {
        let mut revisions = RevisionCounter::default();
        let before = revisions.capture(smallvec![PathSeg::Body(0)]);

        revisions.bump();
        let after = revisions.capture(smallvec![PathSeg::Body(0)]);

        assert_eq!(before.revision, 0);
        assert_eq!(after.revision, 1);
    }

    #[test]
    fn word_path_segments_preserve_nested_order() {
        let segments = smallvec![
            PathSeg::Body(0),
            PathSeg::Row(1),
            PathSeg::Cell(2),
            PathSeg::Para(3),
            PathSeg::Run(4),
        ];

        let path = ContentPath::new(segments.clone(), 7);

        assert_eq!(path.segs, segments);
    }

    #[test]
    fn presentation_path_segments_preserve_repeatable_shape_order() {
        let segments = smallvec![
            PathSeg::Slide(1),
            PathSeg::Shape(2),
            PathSeg::Shape(3),
            PathSeg::Para(4),
            PathSeg::Run(5),
        ];

        let path = ContentPath::new(segments.clone(), 7);

        assert_eq!(path.segs, segments);
    }

    #[test]
    fn python_length_helpers_preserve_rust_truncation() {
        assert_eq!(emu_from_inches(1.75 / 914_400.0), 1);
        assert_eq!(emu_from_inches(-1.75 / 914_400.0), -1);
        assert_eq!(emu_from_centimetres(1.75 / 360_000.0), 1);
        assert_eq!(emu_from_centimetres(-1.75 / 360_000.0), -1);
        assert_eq!(emu_from_millimetres(1.75 / 36_000.0), 1);
        assert_eq!(emu_from_millimetres(-1.75 / 36_000.0), -1);
        assert_eq!(emu_from_points(1.75 / 12_700.0), 1);
        assert_eq!(emu_from_points(-1.75 / 12_700.0), -1);
        assert_eq!(emu_from_twips(1_440), 914_400);
        assert_eq!(emu_from_twips(-1_440), -914_400);
        assert_eq!(inches_from_emu(914_400), 1.0);
        assert_eq!(centimetres_from_emu(914_400), 2.54);
        assert_eq!(millimetres_from_emu(914_400), 25.4);
        assert_eq!(points_from_emu(914_400), 72.0);
        assert_eq!(twips_from_emu(914_400), 1_440);
    }
}
