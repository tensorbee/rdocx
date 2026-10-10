//! Raw XML of one slide, layout, shape or text body: read it and replace it
//! with checked XML, the way out when the model has no API for a feature.
//!
//! A replacement must have the root element of what it replaces, parse with
//! the model, keep every element it names, place the elements PowerPoint
//! checks where the schema allows them, and reference only relationships of
//! the right type that the owning part already has. Anything else is
//! refused before the presentation changes.

use std::collections::{BTreeMap, BTreeSet};

use quick_xml::events::{BytesText, Event};
use quick_xml::name::{QName, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, Writer};
use rpptx_oxml::connector::CT_ConnectionShape;
use rpptx_oxml::graphic_frame::CT_GraphicFrame;
use rpptx_oxml::namespace::{P_NS, R_NS};
use rpptx_oxml::picture::CT_Picture;
use rpptx_oxml::shape_tree::{CT_GroupShape, CT_Shape, ShapeTreeChild};
use rpptx_oxml::slide_parts::{CT_Slide, CT_SlideLayout};

use crate::{A_NS, Error, Presentation, Result};

const OPERATION: &str = "replace_xml";
const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn refused(message: impl Into<String>) -> Error {
    Error::InvalidShapeMutation {
        operation: OPERATION,
        message: message.into(),
    }
}

impl Presentation {
    /// Return slide `index`'s `p:sld` part XML.
    pub fn slide_xml(&self, index: usize) -> Result<Vec<u8>> {
        let record = self
            .slides
            .get(index)
            .ok_or_else(|| refused("slide index out of range"))?;
        record
            .slide
            .to_xml()
            .map_err(|error| refused(error.to_string()))
    }

    /// Replace slide `index` with one `p:sld` element given as XML.
    ///
    /// The slide keeps its notes, comments and relationships, so the XML may
    /// reference only relationship ids the slide part already has.
    pub fn replace_slide_xml(&mut self, index: usize, xml: &[u8]) -> Result<()> {
        let part = self
            .slides
            .get(index)
            .ok_or_else(|| refused("slide index out of range"))?
            .part_name
            .clone();
        let xml = with_root_declarations(xml, "sld")?;
        let slide = CT_Slide::from_xml(&xml).map_err(|error| refused(parse_message(&error)))?;
        let written = slide.to_xml().map_err(|error| refused(error.to_string()))?;
        check_kept(&xml, &written)?;
        self.check_relationships(&part, &xml)?;
        self.slides[index].slide = slide;
        Ok(())
    }

    /// Return slide layout `index`'s `p:sldLayout` part XML.
    pub fn layout_xml(&self, index: usize) -> Result<Vec<u8>> {
        let record = self
            .layouts
            .get(index)
            .ok_or_else(|| refused("slide layout index out of range"))?;
        record
            .layout
            .to_xml()
            .map_err(|error| refused(error.to_string()))
    }

    /// Replace slide layout `index` with one `p:sldLayout` element given as
    /// XML, checked as [`Self::replace_slide_xml`] checks a slide.
    pub fn replace_layout_xml(&mut self, index: usize, xml: &[u8]) -> Result<()> {
        let part = self
            .layouts
            .get(index)
            .ok_or_else(|| refused("slide layout index out of range"))?
            .part_name
            .clone();
        let xml = with_root_declarations(xml, "sldLayout")?;
        let layout =
            CT_SlideLayout::from_xml(&xml).map_err(|error| refused(parse_message(&error)))?;
        let written = layout
            .to_xml()
            .map_err(|error| refused(error.to_string()))?;
        check_kept(&xml, &written)?;
        self.check_relationships(&part, &xml)?;
        // Layouts are written from the package, which only this call edits,
        // so the part takes the new XML along with the model.
        self.package.set_part(&part, written);
        self.layouts[index].layout = layout;
        Ok(())
    }

    /// Replace the shape at `shape_path` (one z-order index per group
    /// level) on slide `slide` with one element of the same kind given as
    /// XML: `p:sp`, `p:pic`, `p:graphicFrame`, `p:grpSp` or `p:cxnSp`.
    pub fn replace_shape_xml(
        &mut self,
        slide: usize,
        shape_path: &[usize],
        xml: &[u8],
    ) -> Result<()> {
        let part = self
            .slides
            .get(slide)
            .ok_or_else(|| refused("slide index out of range"))?
            .part_name
            .clone();
        let current = self.shape_child_mut(slide, shape_path)?;
        let expected = match current {
            ShapeTreeChild::Shape(_) => "sp",
            ShapeTreeChild::Picture(_) => "pic",
            ShapeTreeChild::GraphicFrame(_) => "graphicFrame",
            ShapeTreeChild::GroupShape(_) => "grpSp",
            ShapeTreeChild::Connector(_) => "cxnSp",
            ShapeTreeChild::AlternateContent(_) => {
                return Err(refused(
                    "the shape is alternate content, replace it through the slide XML instead",
                ));
            }
        };
        let xml = with_root_declarations(xml, expected)?;
        let parse = |error: oxml_core::OxmlError| refused(parse_message(&error));
        let child = match expected {
            "sp" => ShapeTreeChild::Shape(CT_Shape::from_xml(&xml).map_err(parse)?),
            "pic" => ShapeTreeChild::Picture(CT_Picture::from_xml(&xml).map_err(parse)?),
            "graphicFrame" => ShapeTreeChild::GraphicFrame(Box::new(
                CT_GraphicFrame::from_xml(&xml).map_err(parse)?,
            )),
            "grpSp" => {
                ShapeTreeChild::GroupShape(Box::new(CT_GroupShape::from_xml(&xml).map_err(parse)?))
            }
            _ => ShapeTreeChild::Connector(CT_ConnectionShape::from_xml(&xml).map_err(parse)?),
        };
        let written = crate::shape_ref(&child).xml()?;
        check_kept(&xml, &written)?;
        self.check_relationships(&part, &xml)?;
        *self.shape_child_mut(slide, shape_path)? = child;
        Ok(())
    }

    /// Return the `p:txBody` of the shape at `shape_path` as standalone XML,
    /// or `None` when the shape has no text body.
    pub fn text_body_xml(&self, slide: usize, shape_path: &[usize]) -> Result<Option<Vec<u8>>> {
        let shape = self.shape_xml_at(slide, shape_path)?;
        let Some((start, end)) = child_span(&shape, "txBody")? else {
            return Ok(None);
        };
        let mut body = shape[start..end].to_vec();
        for (prefix, uri) in root_declarations(&shape)? {
            body = declare(&body, &prefix, &uri)?;
        }
        Ok(Some(body))
    }

    /// Replace the `p:txBody` of the shape at `shape_path` with one
    /// `p:txBody` element given as XML, checked as the whole shape is by
    /// [`Self::replace_shape_xml`].
    pub fn replace_text_body_xml(
        &mut self,
        slide: usize,
        shape_path: &[usize],
        xml: &[u8],
    ) -> Result<()> {
        let shape = self.shape_xml_at(slide, shape_path)?;
        let (start, end) = child_span(&shape, "txBody")?
            .ok_or_else(|| refused("the shape has no p:txBody to replace"))?;
        let body = with_root_declarations(xml, "txBody")?;
        let mut spliced = shape[..start].to_vec();
        spliced.extend_from_slice(&body);
        spliced.extend_from_slice(&shape[end..]);
        self.replace_shape_xml(slide, shape_path, &spliced)
    }

    fn shape_xml_at(&self, slide: usize, shape_path: &[usize]) -> Result<Vec<u8>> {
        let missing = || refused("shape index out of range");
        let slide = self.slide(slide).ok_or_else(missing)?;
        let (first, rest) = shape_path.split_first().ok_or_else(missing)?;
        let mut shape = slide.shape(*first).ok_or_else(missing)?;
        for index in rest {
            shape = shape.child(*index).ok_or_else(missing)?;
        }
        shape.xml()
    }

    fn shape_child_mut(
        &mut self,
        slide: usize,
        shape_path: &[usize],
    ) -> Result<&mut ShapeTreeChild> {
        let missing = || refused("shape index out of range");
        let record = self.slides.get_mut(slide).ok_or_else(missing)?;
        let (first, rest) = shape_path.split_first().ok_or_else(missing)?;
        let mut child = record
            .slide
            .common_slide_data
            .shape_tree
            .children
            .get_mut(*first)
            .ok_or_else(missing)?;
        for index in rest {
            child = match child {
                ShapeTreeChild::GroupShape(group) => {
                    group.children.get_mut(*index).ok_or_else(missing)?
                }
                _ => return Err(missing()),
            };
        }
        Ok(child)
    }

    /// Fail when `xml` references a relationship id that `part` does not
    /// have.
    fn check_relationships(&self, part: &str, xml: &[u8]) -> Result<()> {
        let known = self
            .package
            .get_part_rels(part)
            .map(|relationships| {
                relationships
                    .items
                    .iter()
                    .map(|relationship| (relationship.id.clone(), relationship.rel_type.clone()))
                    .collect::<BTreeMap<_, _>>()
            })
            .unwrap_or_default();
        let mut reader = NsReader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(|error| refused(format!("the XML is malformed: {error}")))?;
            let element = match &event {
                Event::Start(element) | Event::Empty(element) => element,
                Event::Eof => return Ok(()),
                _ => {
                    buffer.clear();
                    continue;
                }
            };
            for attribute in element.attributes() {
                let attribute =
                    attribute.map_err(|error| refused(format!("the XML is malformed: {error}")))?;
                let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                if !matches!(namespace, ResolveResult::Bound(uri) if uri.as_ref() == R_NS.as_bytes())
                {
                    continue;
                }
                let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                // PowerPoint writes an empty id for an action without a target,
                // such as a jump to the next slide.
                if value.is_empty() {
                    continue;
                }
                if let Some(rel_type) = known.get(&value) {
                    let kind = rel_type.rsplit('/').next().unwrap_or_default();
                    if let Some(expected) =
                        expected_relationship(element.local_name().as_ref(), local.as_ref())
                        && !expected.contains(&kind)
                    {
                        return Err(refused(format!(
                            "r:{}=\"{value}\" on {} names a {kind} relationship of {part} where \
                             PowerPoint expects {}, reference a relationship of that type",
                            String::from_utf8_lossy(local.as_ref()),
                            String::from_utf8_lossy(element.name().as_ref()),
                            expected.join(" or ")
                        )));
                    }
                } else {
                    return Err(refused(format!(
                        "r:{}=\"{value}\" names no relationship of {part}, add the target first \
                         (for example with add_picture or a hyperlink address) and reference \
                         the id it creates",
                        String::from_utf8_lossy(local.as_ref())
                    )));
                }
            }
            buffer.clear();
        }
    }
}

fn parse_message(error: &oxml_core::OxmlError) -> String {
    format!("the XML does not parse as PresentationML: {error}")
}

/// The fragment with an XML declaration removed, checked to hold one root
/// element named `local`, with the `p`, `a` and `r` prefixes declared on it
/// when it uses them without declaring them.
fn with_root_declarations(xml: &[u8], local: &str) -> Result<Vec<u8>> {
    let xml = strip_xml_declaration(xml);
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root = None;
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| refused(format!("the XML is malformed: {error}")))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) if depth == 0 => {
                if root.is_some() {
                    return Err(refused(format!(
                        "expected one {local} element, the XML holds more than one root element"
                    )));
                }
                root = Some((
                    String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                ));
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                }
            }
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::DocType(_) => {
                return Err(refused(
                    "a DOCTYPE is not allowed, write the text its entities stand for",
                ));
            }
            Event::Text(text) if depth == 0 && !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(refused(format!(
                    "expected one {local} element, the XML holds text outside it"
                )));
            }
            Event::Eof if depth > 0 => {
                return Err(refused("the XML is malformed: an element is not closed"));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let (name, root_local) =
        root.ok_or_else(|| refused(format!("expected one {local} element, the XML is empty")))?;
    if root_local != local {
        return Err(refused(format!("expected one {local} element, got {name}")));
    }
    let mut declared = cdata_as_text(xml)?;
    for (prefix, uri) in [("p", P_NS), ("a", A_NS), ("r", R_NS)] {
        declared = declare(&declared, prefix, uri)?;
    }
    check_vocabulary(&declared).map_err(refused)?;
    Ok(declared)
}

/// The relationship types an `r:` attribute must name, by the local names
/// of its element and attribute, or `None` when only its existence is
/// checked.
fn expected_relationship(element: &[u8], attribute: &[u8]) -> Option<&'static [&'static str]> {
    Some(match (element, attribute) {
        (b"hlinkClick" | b"hlinkHover" | b"hlinkMouseOver", b"id") => &["hyperlink", "slide"],
        (b"blip", b"embed" | b"link") => &["image"],
        (b"chart", b"id") => &["chart"],
        (b"relIds", b"dm") => &["diagramData"],
        (b"relIds", b"lo") => &["diagramLayout"],
        (b"relIds", b"qs") => &["diagramQuickStyle"],
        (b"relIds", b"cs") => &["diagramColors"],
        (b"videoFile" | b"quickTimeFile", b"link") => &["video"],
        (b"audioFile", b"link") => &["audio"],
        (b"media", b"embed" | b"link") => &["media"],
        (b"oleObj", b"id") => &["oleObject", "package"],
        _ => return None,
    })
}

/// `xml` with every CDATA section written as escaped text, which is what it
/// stands for: the model would otherwise keep the markers as literal text.
fn cdata_as_text(xml: &[u8]) -> Result<Vec<u8>> {
    if !xml.windows(9).any(|window| window == b"<![CDATA[") {
        return Ok(xml.to_vec());
    }
    let malformed = |error: quick_xml::Error| refused(format!("the XML is malformed: {error}"));
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::new());
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(malformed)?;
        let written = match event {
            Event::Eof => return Ok(writer.into_inner()),
            Event::CData(text) => {
                let text = String::from_utf8_lossy(text.as_ref()).into_owned();
                writer.write_event(Event::Text(BytesText::new(&text)))
            }
            event => writer.write_event(event),
        };
        written.map_err(|error| malformed(error.into()))?;
        buffer.clear();
    }
}

/// The namespace of a vocabulary entry: PresentationML or DrawingML.
#[derive(Clone, Copy, PartialEq)]
enum Ns {
    P,
    A,
}

/// One group of children an element may hold.
type Children = &'static [(Ns, &'static str)];

const SHAPE_TREE: Children = &[
    (Ns::P, "nvGrpSpPr"),
    (Ns::P, "grpSpPr"),
    (Ns::P, "sp"),
    (Ns::P, "grpSp"),
    (Ns::P, "graphicFrame"),
    (Ns::P, "cxnSp"),
    (Ns::P, "pic"),
    (Ns::P, "contentPart"),
    (Ns::P, "extLst"),
];
const TEXT_BODY: Children = &[(Ns::A, "bodyPr"), (Ns::A, "lstStyle"), (Ns::A, "p")];
const RUN_PROPERTIES: Children = &[
    (Ns::A, "ln"),
    (Ns::A, "noFill"),
    (Ns::A, "solidFill"),
    (Ns::A, "gradFill"),
    (Ns::A, "blipFill"),
    (Ns::A, "pattFill"),
    (Ns::A, "grpFill"),
    (Ns::A, "effectLst"),
    (Ns::A, "effectDag"),
    (Ns::A, "highlight"),
    (Ns::A, "uLnTx"),
    (Ns::A, "uLn"),
    (Ns::A, "uFillTx"),
    (Ns::A, "uFill"),
    (Ns::A, "latin"),
    (Ns::A, "ea"),
    (Ns::A, "cs"),
    (Ns::A, "sym"),
    (Ns::A, "hlinkClick"),
    (Ns::A, "hlinkMouseOver"),
    (Ns::A, "rtl"),
    (Ns::A, "extLst"),
];
const PARAGRAPH_PROPERTIES: Children = &[
    (Ns::A, "lnSpc"),
    (Ns::A, "spcBef"),
    (Ns::A, "spcAft"),
    (Ns::A, "buClrTx"),
    (Ns::A, "buClr"),
    (Ns::A, "buSzTx"),
    (Ns::A, "buSzPct"),
    (Ns::A, "buSzPts"),
    (Ns::A, "buFontTx"),
    (Ns::A, "buFont"),
    (Ns::A, "buNone"),
    (Ns::A, "buAutoNum"),
    (Ns::A, "buChar"),
    (Ns::A, "buBlip"),
    (Ns::A, "tabLst"),
    (Ns::A, "defRPr"),
    (Ns::A, "extLst"),
];
const BODY_PROPERTIES: Children = &[
    (Ns::A, "prstTxWarp"),
    (Ns::A, "noAutofit"),
    (Ns::A, "normAutofit"),
    (Ns::A, "spAutoFit"),
    (Ns::A, "scene3d"),
    (Ns::A, "sp3d"),
    (Ns::A, "flatTx"),
    (Ns::A, "extLst"),
];
const LIST_STYLE: Children = &[
    (Ns::A, "defPPr"),
    (Ns::A, "lvl1pPr"),
    (Ns::A, "lvl2pPr"),
    (Ns::A, "lvl3pPr"),
    (Ns::A, "lvl4pPr"),
    (Ns::A, "lvl5pPr"),
    (Ns::A, "lvl6pPr"),
    (Ns::A, "lvl7pPr"),
    (Ns::A, "lvl8pPr"),
    (Ns::A, "lvl9pPr"),
    (Ns::A, "extLst"),
];
/// `EG_FillProperties`.
const FILLS: Children = &[
    (Ns::A, "noFill"),
    (Ns::A, "solidFill"),
    (Ns::A, "gradFill"),
    (Ns::A, "blipFill"),
    (Ns::A, "pattFill"),
    (Ns::A, "grpFill"),
];
/// `EG_EffectProperties`.
const EFFECTS: Children = &[(Ns::A, "effectLst"), (Ns::A, "effectDag")];
/// `EG_ColorChoice`.
const COLOR: Children = &[
    (Ns::A, "scrgbClr"),
    (Ns::A, "srgbClr"),
    (Ns::A, "hslClr"),
    (Ns::A, "sysClr"),
    (Ns::A, "schemeClr"),
    (Ns::A, "prstClr"),
];
/// `EG_ColorTransform`.
const COLOR_TRANSFORMS: Children = &[
    (Ns::A, "tint"),
    (Ns::A, "shade"),
    (Ns::A, "comp"),
    (Ns::A, "inv"),
    (Ns::A, "gray"),
    (Ns::A, "alpha"),
    (Ns::A, "alphaOff"),
    (Ns::A, "alphaMod"),
    (Ns::A, "hue"),
    (Ns::A, "hueOff"),
    (Ns::A, "hueMod"),
    (Ns::A, "sat"),
    (Ns::A, "satOff"),
    (Ns::A, "satMod"),
    (Ns::A, "lum"),
    (Ns::A, "lumOff"),
    (Ns::A, "lumMod"),
    (Ns::A, "red"),
    (Ns::A, "redOff"),
    (Ns::A, "redMod"),
    (Ns::A, "green"),
    (Ns::A, "greenOff"),
    (Ns::A, "greenMod"),
    (Ns::A, "blue"),
    (Ns::A, "blueOff"),
    (Ns::A, "blueMod"),
    (Ns::A, "gamma"),
    (Ns::A, "invGamma"),
];
const SHAPE_PROPERTIES: Children = &[
    (Ns::A, "xfrm"),
    (Ns::A, "custGeom"),
    (Ns::A, "prstGeom"),
    (Ns::A, "ln"),
    (Ns::A, "scene3d"),
    (Ns::A, "sp3d"),
    (Ns::A, "extLst"),
];
const LINE: Children = &[
    (Ns::A, "noFill"),
    (Ns::A, "solidFill"),
    (Ns::A, "gradFill"),
    (Ns::A, "pattFill"),
    (Ns::A, "prstDash"),
    (Ns::A, "custDash"),
    (Ns::A, "round"),
    (Ns::A, "bevel"),
    (Ns::A, "miter"),
    (Ns::A, "headEnd"),
    (Ns::A, "tailEnd"),
    (Ns::A, "extLst"),
];
const BLIP_FILL: Children = &[
    (Ns::A, "blip"),
    (Ns::A, "srcRect"),
    (Ns::A, "tile"),
    (Ns::A, "stretch"),
];
const BLIP: Children = &[
    (Ns::A, "alphaBiLevel"),
    (Ns::A, "alphaCeiling"),
    (Ns::A, "alphaFloor"),
    (Ns::A, "alphaInv"),
    (Ns::A, "alphaMod"),
    (Ns::A, "alphaModFix"),
    (Ns::A, "alphaRepl"),
    (Ns::A, "biLevel"),
    (Ns::A, "blur"),
    (Ns::A, "clrChange"),
    (Ns::A, "clrRepl"),
    (Ns::A, "duotone"),
    (Ns::A, "fillOverlay"),
    (Ns::A, "grayscl"),
    (Ns::A, "hsl"),
    (Ns::A, "lum"),
    (Ns::A, "tint"),
    (Ns::A, "extLst"),
];
const PATH: Children = &[
    (Ns::A, "fillToRect"),
    (Ns::A, "close"),
    (Ns::A, "moveTo"),
    (Ns::A, "lnTo"),
    (Ns::A, "arcTo"),
    (Ns::A, "quadBezTo"),
    (Ns::A, "cubicBezTo"),
];
const EFFECT_LIST: Children = &[
    (Ns::A, "blur"),
    (Ns::A, "fillOverlay"),
    (Ns::A, "glow"),
    (Ns::A, "innerShdw"),
    (Ns::A, "outerShdw"),
    (Ns::A, "prstShdw"),
    (Ns::A, "reflection"),
    (Ns::A, "softEdge"),
];
const CELL_PROPERTIES: Children = &[
    (Ns::A, "lnL"),
    (Ns::A, "lnR"),
    (Ns::A, "lnT"),
    (Ns::A, "lnB"),
    (Ns::A, "lnTlToBr"),
    (Ns::A, "lnBlToTr"),
    (Ns::A, "cell3D"),
    (Ns::A, "headers"),
    (Ns::A, "extLst"),
];
const EXTENSIONS: Children = &[(Ns::A, "extLst")];
/// DrawingML elements that hold no child element.
const DRAWING_LEAVES: &[&[u8]] = &[
    b"off",
    b"chOff",
    b"chExt",
    b"spcPct",
    b"spcPts",
    b"buClrTx",
    b"buSzTx",
    b"buSzPct",
    b"buSzPts",
    b"buFontTx",
    b"buFont",
    b"buNone",
    b"buAutoNum",
    b"buChar",
    b"latin",
    b"ea",
    b"cs",
    b"sym",
    b"tab",
    b"snd",
    b"noFill",
    b"grpFill",
    b"gd",
    b"pt",
    b"pos",
    b"close",
    b"arcTo",
    b"prstDash",
    b"ds",
    b"round",
    b"bevel",
    b"miter",
    b"headEnd",
    b"tailEnd",
    b"srcRect",
    b"fillRect",
    b"tileRect",
    b"fillToRect",
    b"lin",
    b"tile",
    b"rot",
    b"anchor",
    b"norm",
    b"up",
    b"bevelT",
    b"bevelB",
    b"noAutofit",
    b"normAutofit",
    b"spAutoFit",
    b"flatTx",
    b"stCxn",
    b"endCxn",
    b"tableStyleId",
    b"masterClrMapping",
    b"tint",
    b"shade",
    b"comp",
    b"inv",
    b"gray",
    b"alpha",
    b"alphaOff",
    b"hue",
    b"hueOff",
    b"hueMod",
    b"sat",
    b"satOff",
    b"satMod",
    b"lum",
    b"lumOff",
    b"lumMod",
    b"red",
    b"redOff",
    b"redMod",
    b"green",
    b"greenOff",
    b"greenMod",
    b"blue",
    b"blueOff",
    b"blueMod",
    b"gamma",
    b"invGamma",
];

/// The children a PresentationML or DrawingML element PowerPoint checks may
/// hold, at any depth of a slide, layout, shape or text body, as groups.
/// An element missing here, such as `a:graphicData` or `p:timing`, is not
/// checked, but its descendants that are listed still are.
fn presentation_children(namespace: Ns, parent: &[u8]) -> Option<&'static [Children]> {
    Some(match (namespace, parent) {
        (Ns::P, b"sld") => &[&[
            (Ns::P, "cSld"),
            (Ns::P, "clrMapOvr"),
            (Ns::P, "transition"),
            (Ns::P, "timing"),
            (Ns::P, "extLst"),
        ]],
        (Ns::P, b"sldLayout") => &[&[
            (Ns::P, "cSld"),
            (Ns::P, "clrMapOvr"),
            (Ns::P, "transition"),
            (Ns::P, "timing"),
            (Ns::P, "hf"),
            (Ns::P, "extLst"),
        ]],
        (Ns::P, b"cSld") => &[&[
            (Ns::P, "bg"),
            (Ns::P, "spTree"),
            (Ns::P, "custDataLst"),
            (Ns::P, "controls"),
            (Ns::P, "extLst"),
        ]],
        (Ns::P, b"bg") => &[&[(Ns::P, "bgPr"), (Ns::P, "bgRef")]],
        (Ns::P, b"bgPr") => &[FILLS, EFFECTS, &[(Ns::P, "extLst")]],
        (Ns::P, b"bgRef") => &[COLOR],
        (Ns::P, b"clrMapOvr") => &[&[(Ns::A, "masterClrMapping"), (Ns::A, "overrideClrMapping")]],
        (Ns::P, b"custDataLst") => &[&[(Ns::P, "custData"), (Ns::P, "tags")]],
        (Ns::P, b"hf" | b"ph") => &[&[(Ns::P, "extLst")]],
        (Ns::P, b"extLst") => &[&[(Ns::P, "ext")]],
        (Ns::P, b"spTree" | b"grpSp") => &[SHAPE_TREE],
        (Ns::P, b"sp") => &[&[
            (Ns::P, "nvSpPr"),
            (Ns::P, "spPr"),
            (Ns::P, "style"),
            (Ns::P, "txBody"),
            (Ns::P, "extLst"),
        ]],
        (Ns::P, b"nvSpPr") => &[&[(Ns::P, "cNvPr"), (Ns::P, "cNvSpPr"), (Ns::P, "nvPr")]],
        (Ns::P, b"pic") => &[&[
            (Ns::P, "nvPicPr"),
            (Ns::P, "blipFill"),
            (Ns::P, "spPr"),
            (Ns::P, "style"),
            (Ns::P, "extLst"),
        ]],
        (Ns::P, b"nvPicPr") => &[&[(Ns::P, "cNvPr"), (Ns::P, "cNvPicPr"), (Ns::P, "nvPr")]],
        (Ns::P, b"cxnSp") => &[&[
            (Ns::P, "nvCxnSpPr"),
            (Ns::P, "spPr"),
            (Ns::P, "style"),
            (Ns::P, "extLst"),
        ]],
        (Ns::P, b"nvCxnSpPr") => &[&[(Ns::P, "cNvPr"), (Ns::P, "cNvCxnSpPr"), (Ns::P, "nvPr")]],
        (Ns::P, b"graphicFrame") => &[&[
            (Ns::P, "nvGraphicFramePr"),
            (Ns::P, "xfrm"),
            (Ns::A, "graphic"),
            (Ns::P, "extLst"),
        ]],
        (Ns::P, b"nvGraphicFramePr") => &[&[
            (Ns::P, "cNvPr"),
            (Ns::P, "cNvGraphicFramePr"),
            (Ns::P, "nvPr"),
        ]],
        (Ns::P, b"nvGrpSpPr") => &[&[(Ns::P, "cNvPr"), (Ns::P, "cNvGrpSpPr"), (Ns::P, "nvPr")]],
        (Ns::P, b"cNvPr") => &[&[
            (Ns::A, "hlinkClick"),
            (Ns::A, "hlinkHover"),
            (Ns::A, "extLst"),
        ]],
        (Ns::P, b"cNvSpPr") => &[&[(Ns::A, "spLocks"), (Ns::A, "extLst")]],
        (Ns::P, b"cNvPicPr") => &[&[(Ns::A, "picLocks"), (Ns::A, "extLst")]],
        (Ns::P, b"cNvGrpSpPr") => &[&[(Ns::A, "grpSpLocks"), (Ns::A, "extLst")]],
        (Ns::P, b"cNvCxnSpPr") => &[&[
            (Ns::A, "cxnSpLocks"),
            (Ns::A, "stCxn"),
            (Ns::A, "endCxn"),
            (Ns::A, "extLst"),
        ]],
        (Ns::P, b"cNvGraphicFramePr") => &[&[(Ns::A, "graphicFrameLocks"), (Ns::A, "extLst")]],
        (Ns::P, b"nvPr") => &[&[
            (Ns::P, "ph"),
            (Ns::A, "audioCd"),
            (Ns::A, "wavAudioFile"),
            (Ns::A, "audioFile"),
            (Ns::A, "videoFile"),
            (Ns::A, "quickTimeFile"),
            (Ns::P, "custDataLst"),
            (Ns::P, "extLst"),
        ]],
        (Ns::P, b"spPr") => &[SHAPE_PROPERTIES, FILLS, EFFECTS],
        (Ns::P, b"grpSpPr") => &[
            &[(Ns::A, "xfrm"), (Ns::A, "scene3d"), (Ns::A, "extLst")],
            FILLS,
            EFFECTS,
        ],
        (Ns::P, b"style") => &[&[
            (Ns::A, "lnRef"),
            (Ns::A, "fillRef"),
            (Ns::A, "effectRef"),
            (Ns::A, "fontRef"),
        ]],
        (Ns::P | Ns::A, b"blipFill") => &[BLIP_FILL],
        (Ns::P, b"xfrm") => &[&[(Ns::A, "off"), (Ns::A, "ext")]],
        (Ns::P | Ns::A, b"txBody") => &[TEXT_BODY],
        (Ns::A, b"bodyPr") => &[BODY_PROPERTIES],
        (Ns::A, b"prstTxWarp" | b"prstGeom") => &[&[(Ns::A, "avLst")]],
        (Ns::A, b"lstStyle") => &[LIST_STYLE],
        (
            Ns::A,
            b"pPr" | b"defPPr" | b"lvl1pPr" | b"lvl2pPr" | b"lvl3pPr" | b"lvl4pPr" | b"lvl5pPr"
            | b"lvl6pPr" | b"lvl7pPr" | b"lvl8pPr" | b"lvl9pPr",
        ) => &[PARAGRAPH_PROPERTIES],
        (Ns::A, b"lnSpc" | b"spcBef" | b"spcAft") => &[&[(Ns::A, "spcPct"), (Ns::A, "spcPts")]],
        (Ns::A, b"tabLst") => &[&[(Ns::A, "tab")]],
        (Ns::A, b"buBlip") => &[&[(Ns::A, "blip")]],
        (Ns::A, b"p") => &[&[
            (Ns::A, "pPr"),
            (Ns::A, "r"),
            (Ns::A, "br"),
            (Ns::A, "fld"),
            (Ns::A, "endParaRPr"),
        ]],
        (Ns::A, b"r") => &[&[(Ns::A, "rPr"), (Ns::A, "t")]],
        (Ns::A, b"br") => &[&[(Ns::A, "rPr")]],
        (Ns::A, b"fld") => &[&[(Ns::A, "rPr"), (Ns::A, "pPr"), (Ns::A, "t")]],
        (Ns::A, b"rPr" | b"endParaRPr" | b"defRPr") => &[RUN_PROPERTIES],
        (Ns::A, b"hlinkClick" | b"hlinkHover" | b"hlinkMouseOver") => {
            &[&[(Ns::A, "snd"), (Ns::A, "extLst")]]
        }
        (Ns::A, b"ln" | b"uLn" | b"lnL" | b"lnR" | b"lnT" | b"lnB" | b"lnTlToBr" | b"lnBlToTr") => {
            &[LINE]
        }
        (Ns::A, b"custDash") => &[&[(Ns::A, "ds")]],
        (Ns::A, b"uFill" | b"fillOverlay") => &[FILLS],
        (
            Ns::A,
            b"solidFill" | b"gs" | b"fgClr" | b"bgClr" | b"buClr" | b"highlight" | b"outerShdw"
            | b"innerShdw" | b"prstShdw" | b"glow" | b"extrusionClr" | b"contourClr" | b"lnRef"
            | b"fillRef" | b"effectRef" | b"fontRef",
        ) => &[COLOR],
        (Ns::A, b"scrgbClr" | b"srgbClr" | b"hslClr" | b"sysClr" | b"schemeClr" | b"prstClr") => {
            &[COLOR_TRANSFORMS]
        }
        (Ns::A, b"gradFill") => &[&[
            (Ns::A, "gsLst"),
            (Ns::A, "lin"),
            (Ns::A, "path"),
            (Ns::A, "tileRect"),
        ]],
        (Ns::A, b"gsLst") => &[&[(Ns::A, "gs")]],
        (Ns::A, b"pattFill") => &[&[(Ns::A, "fgClr"), (Ns::A, "bgClr")]],
        (Ns::A, b"stretch") => &[&[(Ns::A, "fillRect")]],
        (Ns::A, b"blip") => &[BLIP],
        (Ns::A, b"xfrm") => &[&[
            (Ns::A, "off"),
            (Ns::A, "ext"),
            (Ns::A, "chOff"),
            (Ns::A, "chExt"),
        ]],
        (Ns::A, b"avLst" | b"gdLst") => &[&[(Ns::A, "gd")]],
        (Ns::A, b"custGeom") => &[&[
            (Ns::A, "avLst"),
            (Ns::A, "gdLst"),
            (Ns::A, "ahLst"),
            (Ns::A, "cxnLst"),
            (Ns::A, "rect"),
            (Ns::A, "pathLst"),
        ]],
        (Ns::A, b"ahLst") => &[&[(Ns::A, "ahXY"), (Ns::A, "ahPolar")]],
        (Ns::A, b"ahXY" | b"ahPolar" | b"cxn") => &[&[(Ns::A, "pos")]],
        (Ns::A, b"cxnLst") => &[&[(Ns::A, "cxn")]],
        (Ns::A, b"pathLst") => &[&[(Ns::A, "path")]],
        (Ns::A, b"path") => &[PATH],
        (Ns::A, b"moveTo" | b"lnTo" | b"quadBezTo" | b"cubicBezTo") => &[&[(Ns::A, "pt")]],
        (Ns::A, b"effectLst") => &[EFFECT_LIST],
        (Ns::A, b"scene3d") => &[&[
            (Ns::A, "camera"),
            (Ns::A, "lightRig"),
            (Ns::A, "backdrop"),
            (Ns::A, "extLst"),
        ]],
        (Ns::A, b"camera" | b"lightRig") => &[&[(Ns::A, "rot")]],
        (Ns::A, b"backdrop") => &[&[
            (Ns::A, "anchor"),
            (Ns::A, "norm"),
            (Ns::A, "up"),
            (Ns::A, "extLst"),
        ]],
        (Ns::A, b"sp3d") => &[&[
            (Ns::A, "bevelT"),
            (Ns::A, "bevelB"),
            (Ns::A, "extrusionClr"),
            (Ns::A, "contourClr"),
            (Ns::A, "extLst"),
        ]],
        (Ns::A, b"graphic") => &[&[(Ns::A, "graphicData")]],
        (Ns::A, b"tbl") => &[&[(Ns::A, "tblPr"), (Ns::A, "tblGrid"), (Ns::A, "tr")]],
        (Ns::A, b"tblPr") => &[
            FILLS,
            EFFECTS,
            &[
                (Ns::A, "tableStyle"),
                (Ns::A, "tableStyleId"),
                (Ns::A, "extLst"),
            ],
        ],
        (Ns::A, b"tblGrid") => &[&[(Ns::A, "gridCol")]],
        (Ns::A, b"tr") => &[&[(Ns::A, "tc"), (Ns::A, "extLst")]],
        (Ns::A, b"tc") => &[&[(Ns::A, "txBody"), (Ns::A, "tcPr"), (Ns::A, "extLst")]],
        (Ns::A, b"tcPr") => &[CELL_PROPERTIES, FILLS],
        (
            Ns::A,
            b"gridCol"
            | b"spLocks"
            | b"picLocks"
            | b"grpSpLocks"
            | b"cxnSpLocks"
            | b"graphicFrameLocks"
            | b"overrideClrMapping",
        ) => &[EXTENSIONS],
        (Ns::A, b"extLst") => &[&[(Ns::A, "ext")]],
        (Ns::A, b"t") => &[],
        (Ns::A, leaf) if DRAWING_LEAVES.contains(&leaf) => &[],
        _ => return None,
    })
}

/// One open element of a vocabulary walk.
struct VocabularyFrame {
    /// The namespace and local name when the element is PresentationML or
    /// DrawingML.
    known: Option<(Ns, Vec<u8>)>,
    /// The namespaces mc:Ignorable lists in scope.
    ignorable: Vec<Vec<u8>>,
    /// Inside mc:AlternateContent, whose choices PowerPoint resolves itself.
    alternate: bool,
}

/// Refuse an element PowerPoint would refuse or repair in a replacement: a
/// `p:` or `a:` element where its parent does not allow it, or an element
/// of another namespace that mc:Ignorable does not cover.
fn check_vocabulary(xml: &[u8]) -> std::result::Result<(), String> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut stack: Vec<VocabularyFrame> = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| format!("the XML is malformed: {error}"))?;
        let namespace = match namespace {
            ResolveResult::Bound(uri) => uri.as_ref().to_vec(),
            _ => Vec::new(),
        };
        let kind = if namespace == P_NS.as_bytes() {
            Some(Ns::P)
        } else if namespace == A_NS.as_bytes() {
            Some(Ns::A)
        } else {
            None
        };
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let local = element.local_name().as_ref().to_vec();
                let parent = stack.last();
                let is_alternate = namespace == MC_NS.as_bytes() && local == b"AlternateContent";
                let alternate = parent.is_some_and(|frame| frame.alternate) || is_alternate;
                let mut ignorable = parent
                    .map(|frame| frame.ignorable.clone())
                    .unwrap_or_default();
                for attribute in element.attributes().flatten() {
                    let (attribute_namespace, attribute_local) =
                        reader.resolver().resolve_attribute(attribute.key);
                    if matches!(attribute_namespace, ResolveResult::Bound(uri) if uri.as_ref() == MC_NS.as_bytes())
                        && attribute_local.as_ref() == b"Ignorable"
                    {
                        for prefix in
                            String::from_utf8_lossy(attribute.value.as_ref()).split_whitespace()
                        {
                            let qualified = format!("{prefix}:x");
                            if let (ResolveResult::Bound(uri), _) = reader
                                .resolver()
                                .resolve_element(QName(qualified.as_bytes()))
                            {
                                ignorable.push(uri.as_ref().to_vec());
                            }
                        }
                    }
                }
                if let Some(frame) = parent
                    && !frame.alternate
                    && let Some((parent_kind, parent_local)) = &frame.known
                    && let Some(allowed) = presentation_children(*parent_kind, parent_local)
                {
                    let accepted = match kind {
                        Some(kind) => allowed.iter().any(|group| {
                            group.iter().any(|(child_kind, child)| {
                                *child_kind == kind && child.as_bytes() == local
                            })
                        }),
                        None => is_alternate || frame.ignorable.contains(&namespace),
                    };
                    if !accepted {
                        return Err(format!(
                            "{} cannot sit directly in {}:{}, where PowerPoint refuses or \
                             repairs the file, so the replacement was refused",
                            String::from_utf8_lossy(element.name().as_ref()),
                            if *parent_kind == Ns::P { "p" } else { "a" },
                            String::from_utf8_lossy(parent_local)
                        ));
                    }
                }
                if matches!(event, Event::Start(_)) {
                    stack.push(VocabularyFrame {
                        known: kind.map(|kind| (kind, local)),
                        ignorable,
                        alternate,
                    });
                }
            }
            Event::End(_) => {
                stack.pop();
            }
            Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

/// `xml` with `xmlns:prefix="uri"` on its root element unless the root
/// declares that prefix already.
fn declare(xml: &[u8], prefix: &str, uri: &str) -> Result<Vec<u8>> {
    let declarations = root_declarations(xml)?;
    if declarations.contains_key(prefix) {
        return Ok(xml.to_vec());
    }
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| refused(format!("the XML is malformed: {error}")))?
        {
            Event::Start(element) | Event::Empty(element) => {
                let name_end = before + 1 + element.name().as_ref().len();
                let mut declared = xml[..name_end].to_vec();
                declared.extend_from_slice(format!(" xmlns:{prefix}=\"{uri}\"").as_bytes());
                declared.extend_from_slice(&xml[name_end..]);
                return Ok(declared);
            }
            Event::Eof => return Err(refused("the XML holds no element")),
            _ => buffer.clear(),
        }
    }
}

/// The prefixes the root element of `xml` declares, with their namespaces.
fn root_declarations(xml: &[u8]) -> Result<BTreeMap<String, String>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| refused(format!("the XML is malformed: {error}")))?
        {
            Event::Start(element) | Event::Empty(element) => {
                let mut declarations = BTreeMap::new();
                for attribute in element.attributes().flatten() {
                    if let Some(prefix) = attribute.key.as_ref().strip_prefix(b"xmlns:") {
                        declarations.insert(
                            String::from_utf8_lossy(prefix).into_owned(),
                            String::from_utf8_lossy(attribute.value.as_ref()).into_owned(),
                        );
                    }
                }
                return Ok(declarations);
            }
            Event::Eof => return Ok(BTreeMap::new()),
            _ => buffer.clear(),
        }
    }
}

/// The span of the first child element of the root named `local`.
fn child_span(xml: &[u8], local: &str) -> Result<Option<(usize, usize)>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut start = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| refused(format!("the XML is malformed: {error}")))?;
        match event {
            Event::Start(element) => {
                depth += 1;
                if depth == 2
                    && start.is_none()
                    && element.local_name().as_ref() == local.as_bytes()
                {
                    start = Some(before);
                }
            }
            Event::Empty(element) => {
                if depth == 1
                    && start.is_none()
                    && element.local_name().as_ref() == local.as_bytes()
                {
                    return Ok(Some((before, reader.buffer_position() as usize)));
                }
            }
            Event::End(_) => {
                if depth == 2
                    && let Some(begin) = start
                {
                    return Ok(Some((begin, reader.buffer_position() as usize)));
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

/// Fail when the model's serialization drops an element the input named.
fn check_kept(input: &[u8], written: &[u8]) -> Result<()> {
    let kept = element_counts(written)?;
    let missing = element_counts(input)?
        .into_iter()
        .filter(|(name, count)| kept.get(name).copied().unwrap_or(0) < *count)
        .map(|((_, local), _)| String::from_utf8_lossy(&local).into_owned())
        .collect::<BTreeSet<_>>();
    if missing.is_empty() {
        return Ok(());
    }
    Err(refused(format!(
        "rpptx would not keep {}, so the replacement was refused",
        missing.into_iter().collect::<Vec<_>>().join(", ")
    )))
}

/// Element counts by namespace and local name.
type ElementCounts = BTreeMap<(Vec<u8>, Vec<u8>), usize>;

/// How many elements of each namespace and local name `xml` holds.
fn element_counts(xml: &[u8]) -> Result<ElementCounts> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut counts = BTreeMap::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| refused(format!("the XML is malformed: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let namespace = match namespace {
                    ResolveResult::Bound(uri) => uri.as_ref().to_vec(),
                    _ => Vec::new(),
                };
                *counts
                    .entry((namespace, element.local_name().as_ref().to_vec()))
                    .or_insert(0) += 1;
            }
            Event::Eof => return Ok(counts),
            _ => {}
        }
        buffer.clear();
    }
}

fn strip_xml_declaration(xml: &[u8]) -> &[u8] {
    let trimmed = xml.trim_ascii_start();
    if trimmed.starts_with(b"<?xml")
        && let Some(end) = trimmed.windows(2).position(|pair| pair == b"?>")
    {
        return &trimmed[end + 2..];
    }
    xml
}
