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

mod document;
mod error;
mod length;
pub mod paragraph;
pub mod run;
pub mod style;
pub mod table;

pub use document::{
    AccessibilityIssue, Document, ImageInfo, IssueSeverity, LinkInfo, OutlineNode,
};
pub use error::{Error, Result};
pub use length::Length;
pub use paragraph::{Alignment, BorderStyle, Paragraph, ParagraphRef, SectionBreak, TabAlignment, TabLeader};
pub use run::{Run, RunRef, UnderlineStyle};
pub use style::{Style, StyleBuilder};
pub use table::{Cell, CellRef, Row, RowRef, Table, TableRef, VerticalAlignment};
