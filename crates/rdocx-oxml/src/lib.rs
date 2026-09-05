#![doc = include_str!("../README.md")]
#![allow(non_camel_case_types)]
#![allow(clippy::large_enum_variant)]
#![allow(clippy::should_implement_trait)]

pub mod borders;
pub mod comments;
pub mod comments_extended;
pub mod content_control;
pub mod document;
pub mod drawing;
pub mod footnotes;
pub mod glossary;
pub mod header_footer;
pub mod math;
pub mod namespace;
pub mod numbering;
pub mod placeholder;
pub mod properties;
pub mod revision;
pub mod settings;
pub mod shared;
pub mod styles;
pub mod table;
pub mod text;
pub mod theme;

pub use error::{OxmlError, Result};
pub(crate) use oxml_core::xml_text;
pub use oxml_core::{core_properties, error, raw_xml, units};

pub use borders::{CT_BorderEdge, CT_PBdr, CT_TabStop, CT_Tabs};
pub use content_control::{CT_DataBinding, CT_Sdt, CT_SdtPr, SdtContent, SdtType};
pub use document::{BodyContent, CT_Body, CT_Document, CT_SectPr};
pub use math::{CT_OMath, CT_OMathPara, MathArgument, MathExpression, MathProperties, OfficeMath};
pub use numbering::{CT_AbstractNum, CT_Lvl, CT_Num, CT_Numbering, ST_NumberFormat};
pub use properties::{CT_PPr, CT_RPr};
pub use revision::{CT_Revision, RevisionContent, RevisionKind};
pub use shared::{
    ST_Border, ST_Jc, ST_OnOff, ST_PageOrientation, ST_SectionType, ST_TabJc, ST_TabLeader,
};
pub use styles::{CT_DocDefaults, CT_Style, CT_Styles};
pub use table::{CT_Row, CT_Tbl, CT_TblGrid, CT_TblPr, CT_Tc, CT_TcPr};
pub use text::{CT_P, CT_R, CT_Text};
pub use units::{Emu, HalfPoint, Twips};
