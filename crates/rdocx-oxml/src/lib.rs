//! WordprocessingML XML element types for OOXML.
//!
//! This crate provides 1:1 Rust struct mappings for OOXML elements,
//! with manual `quick-xml` parsing and serialization.
//!
//! We use OOXML naming conventions (CT_P, CT_R, ST_Jc, etc.) for
//! direct correspondence with the specification.

#![allow(non_camel_case_types)]
#![allow(clippy::large_enum_variant)]
#![allow(clippy::should_implement_trait)]

pub mod borders;
pub mod core_properties;
pub mod document;
pub mod footnotes;
pub mod drawing;
pub mod error;
pub mod header_footer;
pub mod namespace;
pub mod numbering;
pub mod placeholder;
pub mod properties;
pub mod raw_xml;
pub mod shared;
pub mod styles;
pub mod table;
pub mod text;
pub mod theme;
pub mod units;

pub use borders::{CT_BorderEdge, CT_PBdr, CT_TabStop, CT_Tabs};
pub use document::{BodyContent, CT_Body, CT_Document, CT_SectPr};
pub use error::{OxmlError, Result};
pub use properties::{CT_PPr, CT_RPr};
pub use shared::{ST_Border, ST_Jc, ST_OnOff, ST_PageOrientation, ST_SectionType, ST_TabJc, ST_TabLeader};
pub use styles::{CT_DocDefaults, CT_Style, CT_Styles};
pub use table::{CT_Row, CT_Tbl, CT_TblPr, CT_Tc, CT_TcPr, CT_TblGrid};
pub use numbering::{CT_AbstractNum, CT_Lvl, CT_Num, CT_Numbering, ST_NumberFormat};
pub use text::{CT_P, CT_R, CT_Text};
pub use units::{Emu, HalfPoint, Twips};
