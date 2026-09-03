//! MathML and LaTeX conversion for the normalized OfficeMath tree.

use std::collections::HashSet;
use std::fmt::Write as _;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use rdocx_oxml::math::{
    FractionType, MathAccent, MathArgument, MathDelimiter, MathExpression, MathFraction, MathLimit,
    MathMatrix, MathMatrixRow, MathNary, MathPreSubSuperscript, MathRadical, MathRun, MathScript,
    MathSubSuperscript,
};

use crate::{Error, Result};

const MATHML_NS: &str = "http://www.w3.org/1998/Math/MathML";
const MAX_INPUT_BYTES: usize = 1024 * 1024;
const MAX_EVENTS: usize = 100_000;
const MAX_DEPTH: usize = 64;
const MAX_NODES: usize = 50_000;
const MAX_ROWS: usize = 256;
const MAX_COLUMNS: usize = 256;
const MAX_TEXT_BYTES: usize = 512 * 1024;
const MAX_DIAGNOSTICS: usize = 1024;

/// One stable source location and message for a lossy equation conversion.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MathConversionDiagnostic {
    /// Format-specific element path or byte offset.
    pub path: String,
    /// Stable description of the content that was not represented.
    pub message: String,
}

/// A converted value together with every ordered loss diagnostic.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MathConversionResult<T> {
    /// The normalized equation tree or canonical serialized text.
    pub value: T,
    /// Ordered diagnostics produced while converting the value.
    pub diagnostics: Vec<MathConversionDiagnostic>,
}

#[derive(Clone, Debug)]
struct XmlAttribute {
    namespace: Option<String>,
    local: String,
    value: String,
}

#[derive(Clone, Debug)]
enum XmlChild {
    Element(XmlNode),
    Text(String),
}

#[derive(Clone, Debug)]
struct XmlNode {
    namespace: Option<String>,
    local: String,
    attributes: Vec<XmlAttribute>,
    children: Vec<XmlChild>,
}

impl XmlNode {
    fn elements(&self) -> impl Iterator<Item = &XmlNode> {
        self.children.iter().filter_map(|child| match child {
            XmlChild::Element(element) => Some(element),
            XmlChild::Text(_) => None,
        })
    }

    fn attribute(&self, local: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| attribute.namespace.is_none() && attribute.local == local)
            .map(|attribute| attribute.value.as_str())
    }

    fn text(&self) -> String {
        let mut value = String::new();
        for child in &self.children {
            if let XmlChild::Text(text) = child {
                value.push_str(text);
            }
        }
        value
    }
}

struct Diagnostics {
    values: Vec<MathConversionDiagnostic>,
}

impl Diagnostics {
    fn new() -> Self {
        Self { values: Vec::new() }
    }

    fn push(&mut self, path: impl Into<String>, message: impl Into<String>) -> Result<()> {
        if self.values.len() >= MAX_DIAGNOSTICS {
            return Err(math_error("conversion exceeds the diagnostic limit"));
        }
        self.values.push(MathConversionDiagnostic {
            path: path.into(),
            message: message.into(),
        });
        Ok(())
    }
}

/// Import Presentation MathML into the normalized OfficeMath argument tree.
pub fn equation_from_mathml(input: &str) -> Result<MathConversionResult<MathArgument>> {
    if input.len() > MAX_INPUT_BYTES {
        return Err(math_error("MathML input exceeds the byte limit"));
    }
    let root = parse_mathml_xml(input.as_bytes())?;
    if root.namespace.as_deref() != Some(MATHML_NS) || root.local != "math" {
        return Err(math_error(
            "MathML root must be math in the W3C MathML namespace",
        ));
    }
    let mut diagnostics = Diagnostics::new();
    let mut nodes = 0_usize;
    diagnose_attributes(&root, "/math[1]", &[], &mut diagnostics)?;
    let mut value = mathml_children_to_argument(&root, "/math[1]", &mut diagnostics, &mut nodes)?;
    normalize_argument(&mut value);
    Ok(MathConversionResult {
        value,
        diagnostics: diagnostics.values,
    })
}

/// Export a normalized OfficeMath argument tree as canonical Presentation MathML.
pub fn equation_to_mathml(expression: &MathArgument) -> MathConversionResult<String> {
    let mut diagnostics = Diagnostics::new();
    if let Err(message) = validate_tree(expression) {
        return MathConversionResult {
            value: String::new(),
            diagnostics: vec![MathConversionDiagnostic {
                path: "/math[1]".to_owned(),
                message,
            }],
        };
    }
    let mut normalized = expression.clone();
    normalize_argument(&mut normalized);
    let mut output = format!(r#"<math xmlns="{MATHML_NS}">"#);
    let write_failed =
        write_mathml_argument(&normalized, "/math[1]", &mut output, &mut diagnostics).is_err();
    if write_failed {
        output.clear();
        diagnostics.values = vec![MathConversionDiagnostic {
            path: "/math[1]".to_owned(),
            message: "conversion exceeds the diagnostic limit".to_owned(),
        }];
    } else {
        output.push_str("</math>");
        if output.len() > MAX_INPUT_BYTES {
            output.clear();
            diagnostics.values = vec![MathConversionDiagnostic {
                path: "/math[1]".to_owned(),
                message: "serialized MathML exceeds the reader byte limit".to_owned(),
            }];
        } else if equation_from_mathml(&output).is_err() {
            output.clear();
            diagnostics.values = vec![MathConversionDiagnostic {
                path: "/math[1]".to_owned(),
                message: "serialized MathML exceeds the reader structural limits".to_owned(),
            }];
        }
    }
    MathConversionResult {
        value: output,
        diagnostics: diagnostics.values,
    }
}

/// Import the supported LaTeX subset into the normalized OfficeMath argument tree.
pub fn equation_from_latex(input: &str) -> Result<MathConversionResult<MathArgument>> {
    if input.len() > MAX_INPUT_BYTES {
        return Err(math_error("LaTeX input exceeds the byte limit"));
    }
    let mut parser = LatexParser::new(input);
    let mut value = parser.parse_argument(None)?;
    parser.skip_whitespace();
    if !parser.at_end() {
        return Err(parser.error("unexpected trailing LaTeX input"));
    }
    normalize_argument(&mut value);
    Ok(MathConversionResult {
        value,
        diagnostics: parser.diagnostics.values,
    })
}

/// Export a normalized OfficeMath argument tree as canonical LaTeX.
pub fn equation_to_latex(expression: &MathArgument) -> MathConversionResult<String> {
    let mut diagnostics = Diagnostics::new();
    if let Err(message) = validate_tree(expression) {
        return MathConversionResult {
            value: String::new(),
            diagnostics: vec![MathConversionDiagnostic {
                path: "byte:0".to_owned(),
                message,
            }],
        };
    }
    if diagnose_empty_latex_runs(expression, "byte:0", &mut diagnostics).is_err() {
        return MathConversionResult {
            value: String::new(),
            diagnostics: vec![MathConversionDiagnostic {
                path: "byte:0".to_owned(),
                message: "conversion exceeds the diagnostic limit".to_owned(),
            }],
        };
    }
    let mut normalized = expression.clone();
    normalize_argument(&mut normalized);
    let mut output = String::new();
    if write_latex_argument(&normalized, "byte:0", &mut output, &mut diagnostics).is_err() {
        output.clear();
        diagnostics.values = vec![MathConversionDiagnostic {
            path: "byte:0".to_owned(),
            message: "conversion exceeds the diagnostic limit".to_owned(),
        }];
    } else if output.len() > MAX_INPUT_BYTES {
        output.clear();
        diagnostics.values = vec![MathConversionDiagnostic {
            path: "byte:0".to_owned(),
            message: "serialized LaTeX exceeds the reader byte limit".to_owned(),
        }];
    } else if equation_from_latex(&output).is_err() {
        output.clear();
        diagnostics.values = vec![MathConversionDiagnostic {
            path: "byte:0".to_owned(),
            message: "serialized LaTeX exceeds the reader structural limits".to_owned(),
        }];
    }
    MathConversionResult {
        value: output,
        diagnostics: diagnostics.values,
    }
}

fn math_error(message: impl Into<String>) -> Error {
    Error::Other(format!("math conversion error: {}", message.into()))
}

fn parse_mathml_xml(xml: &[u8]) -> Result<XmlNode> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<XmlNode>::new();
    let mut root = None;
    let mut events = 0_usize;
    let mut text_bytes = 0_usize;
    loop {
        let offset = reader.buffer_position();
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| math_error(format!("malformed MathML at byte {offset}: {error}")))?;
        events = events
            .checked_add(1)
            .ok_or_else(|| math_error("MathML event count overflowed"))?;
        if events > MAX_EVENTS {
            return Err(math_error("MathML input exceeds the event limit"));
        }
        let namespace = resolved_namespace(namespace, offset)?;
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(math_error("MathML input exceeds the depth limit"));
                }
                stack.push(xml_node(&reader, namespace, &element, offset)?);
            }
            Event::Empty(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(math_error("MathML input exceeds the depth limit"));
                }
                attach_xml_node(
                    &mut stack,
                    &mut root,
                    xml_node(&reader, namespace, &element, offset)?,
                )?;
            }
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| math_error("unmatched MathML end tag"))?;
                attach_xml_node(&mut stack, &mut root, node)?;
            }
            Event::Text(text) => {
                let decoded = text
                    .xml10_content()
                    .map_err(|error| math_error(format!("invalid MathML text: {error}")))?;
                let value = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| math_error(format!("unresolved MathML entity: {error}")))?
                    .into_owned();
                append_xml_text(&mut stack, value, &mut text_bytes)?;
            }
            Event::CData(text) => {
                let value = text
                    .xml10_content()
                    .map_err(|error| math_error(format!("invalid MathML CDATA: {error}")))?
                    .into_owned();
                append_xml_text(&mut stack, value, &mut text_bytes)?;
            }
            Event::GeneralRef(reference) => {
                let value = if let Some(character) = reference
                    .resolve_char_ref()
                    .map_err(|error| math_error(format!("invalid MathML reference: {error}")))?
                {
                    character.to_string()
                } else {
                    let reference: &[u8] = &reference;
                    match reference {
                        b"amp" => "&",
                        b"lt" => "<",
                        b"gt" => ">",
                        b"apos" => "'",
                        b"quot" => "\"",
                        _ => return Err(math_error("unresolved MathML entity is unsupported")),
                    }
                    .to_owned()
                };
                append_xml_text(&mut stack, value, &mut text_bytes)?;
            }
            Event::DocType(_) => return Err(math_error("MathML DTD declarations are forbidden")),
            Event::Eof => break,
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {}
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(math_error("MathML input has an unclosed element"));
    }
    root.ok_or_else(|| math_error("MathML input has no root element"))
}

fn resolved_namespace(namespace: ResolveResult<'_>, offset: u64) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(namespace) => Ok(Some(
            std::str::from_utf8(namespace.as_ref())
                .map_err(|_| math_error(format!("namespace at byte {offset} is not UTF-8")))?
                .to_owned(),
        )),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(_) => Err(math_error(format!(
            "unresolved namespace prefix at byte {offset}"
        ))),
    }
}

fn xml_node(
    reader: &NsReader<&[u8]>,
    namespace: Option<String>,
    element: &BytesStart<'_>,
    offset: u64,
) -> Result<XmlNode> {
    let local = std::str::from_utf8(element.local_name().as_ref())
        .map_err(|_| math_error(format!("element name at byte {offset} is not UTF-8")))?
        .to_owned();
    let mut attributes = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| math_error(format!("invalid attribute at byte {offset}: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(namespace, offset)?;
        let local = std::str::from_utf8(local.as_ref())
            .map_err(|_| math_error(format!("attribute name at byte {offset} is not UTF-8")))?
            .to_owned();
        if !seen.insert((namespace.clone(), local.clone())) {
            return Err(math_error(format!(
                "duplicate expanded attribute at byte {offset}"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
            .map_err(|error| math_error(format!("invalid attribute at byte {offset}: {error}")))?
            .into_owned();
        attributes.push(XmlAttribute {
            namespace,
            local,
            value,
        });
    }
    Ok(XmlNode {
        namespace,
        local,
        attributes,
        children: Vec::new(),
    })
}

fn attach_xml_node(stack: &mut [XmlNode], root: &mut Option<XmlNode>, node: XmlNode) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlChild::Element(node));
    } else if root.replace(node).is_some() {
        return Err(math_error("MathML input has more than one root element"));
    }
    Ok(())
}

fn append_xml_text(stack: &mut [XmlNode], value: String, total: &mut usize) -> Result<()> {
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| math_error("MathML text size overflowed"))?;
    if *total > MAX_TEXT_BYTES {
        return Err(math_error("MathML input exceeds the text limit"));
    }
    let Some(parent) = stack.last_mut() else {
        if value.trim().is_empty() {
            return Ok(());
        }
        return Err(math_error("MathML text appears outside the root element"));
    };
    parent.children.push(XmlChild::Text(value));
    Ok(())
}

fn mathml_children_to_argument(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
    nodes: &mut usize,
) -> Result<MathArgument> {
    let mut expressions = Vec::new();
    let mut positions = std::collections::HashMap::<&str, usize>::new();
    for child in &node.children {
        match child {
            XmlChild::Text(text) if !text.trim().is_empty() => {
                diagnostics.push(path, "text outside a MathML token was discarded")?
            }
            XmlChild::Text(_) => {}
            XmlChild::Element(element) => {
                let position = positions.entry(element.local.as_str()).or_default();
                *position += 1;
                let child_path = format!("{path}/{}[{position}]", element.local);
                expressions.extend(mathml_element_to_expressions(
                    element,
                    &child_path,
                    diagnostics,
                    nodes,
                )?);
            }
        }
    }
    let mut rebuilt = Vec::with_capacity(expressions.len());
    let mut expressions = expressions.into_iter().peekable();
    while let Some(mut expression) = expressions.next() {
        if let MathExpression::Nary(nary) = &mut expression
            && nary.base.expressions.is_empty()
            && let Some(base) = expressions.next()
        {
            nary.base = MathArgument::new(vec![base]);
        }
        rebuilt.push(expression);
    }
    let mut argument = MathArgument::new(rebuilt);
    normalize_argument(&mut argument);
    Ok(argument)
}

fn mathml_element_to_expressions(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
    nodes: &mut usize,
) -> Result<Vec<MathExpression>> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| math_error("MathML node count overflowed"))?;
    if *nodes > MAX_NODES {
        return Err(math_error("MathML input exceeds the node limit"));
    }
    if node.namespace.as_deref() != Some(MATHML_NS) {
        diagnostics.push(path, "foreign-namespace element was discarded")?;
        return Ok(Vec::new());
    }
    let allowed_attributes = match node.local.as_str() {
        "mover" => &["accent"][..],
        "munder" => &["accentunder"][..],
        "munderover" => &["accent", "accentunder"][..],
        "mfenced" => &["open", "close", "separators"][..],
        _ => &[][..],
    };
    diagnose_attributes(node, path, allowed_attributes, diagnostics)?;
    if !matches!(
        node.local.as_str(),
        "math" | "mrow" | "mi" | "mn" | "mo" | "mtext" | "msqrt" | "mfenced"
    ) {
        diagnose_structural_text(node, path, diagnostics)?;
    }
    match node.local.as_str() {
        "math" | "mrow" => {
            if node.local == "mrow"
                && let Some(delimiter) = parse_explicit_fenced_row(node, path, diagnostics, nodes)?
            {
                return Ok(vec![MathExpression::Delimiter(delimiter)]);
            }
            if node.local == "mrow"
                && let Some(nary) = parse_explicit_nary_row(node, path, diagnostics, nodes)?
            {
                return Ok(vec![MathExpression::Nary(nary)]);
            }
            Ok(mathml_children_to_argument(node, path, diagnostics, nodes)?.expressions)
        }
        "mi" | "mn" | "mo" | "mtext" => {
            for (index, _) in node.elements().enumerate() {
                diagnostics.push(
                    format!("{path}/*[{}]", index + 1),
                    "nested MathML token content was discarded",
                )?;
            }
            Ok(vec![MathRun::new(node.text()).into()])
        }
        "mfrac" => {
            let children = mathml_element_children(node);
            require_child_count(path, &children, 2)?;
            Ok(vec![MathExpression::Fraction(MathFraction::new(
                mathml_node_to_argument(children[0], &format!("{path}/*[1]"), diagnostics, nodes)?,
                mathml_node_to_argument(children[1], &format!("{path}/*[2]"), diagnostics, nodes)?,
            ))])
        }
        "msub" | "msup" => {
            let children = mathml_element_children(node);
            require_child_count(path, &children, 2)?;
            if let Some(operator) =
                mathml_nary_operator(children[0], &format!("{path}/*[1]"), diagnostics)?
            {
                let mut nary = MathNary::new(operator, MathArgument::default());
                let script = mathml_node_to_argument(
                    children[1],
                    &format!("{path}/*[2]"),
                    diagnostics,
                    nodes,
                )?;
                if node.local == "msub" {
                    nary.subscript = script;
                    nary.hide_subscript = false;
                } else {
                    nary.superscript = script;
                    nary.hide_superscript = false;
                }
                return Ok(vec![MathExpression::Nary(nary)]);
            }
            let script = MathScript::new(
                mathml_node_to_argument(children[0], &format!("{path}/*[1]"), diagnostics, nodes)?,
                mathml_node_to_argument(children[1], &format!("{path}/*[2]"), diagnostics, nodes)?,
            );
            Ok(vec![if node.local == "msub" {
                MathExpression::Subscript(script)
            } else {
                MathExpression::Superscript(script)
            }])
        }
        "msubsup" => {
            let children = mathml_element_children(node);
            require_child_count(path, &children, 3)?;
            if let Some(operator) =
                mathml_nary_operator(children[0], &format!("{path}/*[1]"), diagnostics)?
            {
                let mut nary = MathNary::new(operator, MathArgument::default());
                nary.subscript = mathml_node_to_argument(
                    children[1],
                    &format!("{path}/*[2]"),
                    diagnostics,
                    nodes,
                )?;
                nary.superscript = mathml_node_to_argument(
                    children[2],
                    &format!("{path}/*[3]"),
                    diagnostics,
                    nodes,
                )?;
                nary.hide_subscript = false;
                nary.hide_superscript = false;
                return Ok(vec![MathExpression::Nary(nary)]);
            }
            Ok(vec![MathExpression::SubSuperscript(
                MathSubSuperscript::new(
                    mathml_node_to_argument(
                        children[0],
                        &format!("{path}/*[1]"),
                        diagnostics,
                        nodes,
                    )?,
                    mathml_node_to_argument(
                        children[1],
                        &format!("{path}/*[2]"),
                        diagnostics,
                        nodes,
                    )?,
                    mathml_node_to_argument(
                        children[2],
                        &format!("{path}/*[3]"),
                        diagnostics,
                        nodes,
                    )?,
                ),
            )])
        }
        "mmultiscripts" => parse_mathml_multiscripts(node, path, diagnostics, nodes),
        "msqrt" => Ok(vec![MathExpression::Radical(MathRadical::new(
            mathml_children_to_argument(node, path, diagnostics, nodes)?,
        ))]),
        "mroot" => {
            let children = mathml_element_children(node);
            require_child_count(path, &children, 2)?;
            Ok(vec![MathExpression::Radical(MathRadical::with_degree(
                mathml_node_to_argument(children[1], &format!("{path}/*[2]"), diagnostics, nodes)?,
                mathml_node_to_argument(children[0], &format!("{path}/*[1]"), diagnostics, nodes)?,
            ))])
        }
        "mtable" => parse_mathml_matrix(node, path, diagnostics, nodes),
        "munder" | "mover" | "munderover" => {
            parse_mathml_limits_or_accent(node, path, diagnostics, nodes)
        }
        "mfenced" => parse_mathml_fenced(node, path, diagnostics, nodes),
        "semantics" => {
            diagnostics.push(path, "MathML semantics metadata was discarded")?;
            let Some((index, content)) = node
                .elements()
                .enumerate()
                .find(|(_, child)| is_supported_mathml_element(child))
            else {
                return Ok(Vec::new());
            };
            mathml_element_to_expressions(
                content,
                &format!("{path}/*[{}]", index + 1),
                diagnostics,
                nodes,
            )
        }
        "none" | "mprescripts" => {
            diagnose_empty_mathml_marker(node, path, diagnostics)?;
            diagnostics.push(
                path,
                "MathML structural marker outside multiscripts was discarded",
            )?;
            Ok(Vec::new())
        }
        _ => {
            diagnostics.push(
                path,
                format!("unsupported MathML element {} was discarded", node.local),
            )?;
            Ok(Vec::new())
        }
    }
}

fn mathml_node_to_argument(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
    nodes: &mut usize,
) -> Result<MathArgument> {
    Ok(MathArgument::new(mathml_element_to_expressions(
        node,
        path,
        diagnostics,
        nodes,
    )?))
}

fn mathml_element_children(node: &XmlNode) -> Vec<&XmlNode> {
    node.elements().collect()
}

fn require_child_count(path: &str, children: &[&XmlNode], expected: usize) -> Result<()> {
    if children.len() != expected {
        return Err(math_error(format!(
            "{path} requires {expected} element children"
        )));
    }
    Ok(())
}

fn diagnose_attributes(
    node: &XmlNode,
    path: &str,
    allowed: &[&str],
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    for attribute in &node.attributes {
        if attribute.namespace.is_some() || !allowed.contains(&attribute.local.as_str()) {
            diagnostics.push(
                format!("{path}/@{}", attribute.local),
                "unsupported MathML attribute was discarded",
            )?;
        }
    }
    Ok(())
}

fn diagnose_structural_text(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    for child in &node.children {
        if let XmlChild::Text(text) = child
            && !text.trim().is_empty()
        {
            diagnostics.push(path, "text outside a MathML token was discarded")?;
        }
    }
    Ok(())
}

fn diagnose_token_children(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    for (index, _) in node.elements().enumerate() {
        diagnostics.push(
            format!("{path}/*[{}]", index + 1),
            "nested MathML token content was discarded",
        )?;
    }
    Ok(())
}

fn mathml_nary_operator(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
) -> Result<Option<String>> {
    if node.namespace.as_deref() != Some(MATHML_NS)
        || node.local != "mo"
        || !is_nary_operator(&node.text())
    {
        return Ok(None);
    }
    diagnose_attributes(node, path, &["largeop"], diagnostics)?;
    diagnose_boolean_attribute(node, path, "largeop", diagnostics)?;
    if node.attribute("largeop") == Some("false") {
        diagnostics.push(
            format!("{path}/@largeop"),
            "non-large MathML operator was normalized as n-ary",
        )?;
    }
    diagnose_token_children(node, path, diagnostics)?;
    Ok(Some(node.text()))
}

fn diagnose_empty_mathml_marker(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    diagnose_attributes(node, path, &[], diagnostics)?;
    diagnose_structural_text(node, path, diagnostics)?;
    if node.elements().next().is_some() {
        diagnostics.push(
            path,
            "content inside a MathML structural marker was discarded",
        )?;
    }
    Ok(())
}

fn diagnose_boolean_attribute(
    node: &XmlNode,
    path: &str,
    local: &str,
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    if let Some(value) = node.attribute(local)
        && !matches!(value, "true" | "false")
    {
        diagnostics.push(
            format!("{path}/@{local}"),
            "unsupported MathML attribute value was discarded",
        )?;
    }
    Ok(())
}

fn diagnose_enum_attribute(
    node: &XmlNode,
    path: &str,
    local: &str,
    allowed: &[&str],
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    if let Some(value) = node.attribute(local)
        && !allowed.contains(&value)
    {
        diagnostics.push(
            format!("{path}/@{local}"),
            "unsupported MathML attribute value was discarded",
        )?;
    }
    Ok(())
}

fn is_supported_mathml_element(node: &XmlNode) -> bool {
    node.namespace.as_deref() == Some(MATHML_NS)
        && matches!(
            node.local.as_str(),
            "math"
                | "mrow"
                | "mi"
                | "mn"
                | "mo"
                | "mtext"
                | "mfrac"
                | "msub"
                | "msup"
                | "msubsup"
                | "mmultiscripts"
                | "msqrt"
                | "mroot"
                | "mtable"
                | "munder"
                | "mover"
                | "munderover"
                | "mfenced"
                | "semantics"
        )
}

fn parse_mathml_multiscripts(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
    nodes: &mut usize,
) -> Result<Vec<MathExpression>> {
    let children = mathml_element_children(node);
    if children.is_empty() {
        return Err(math_error(format!("{path} requires a base")));
    }
    let base = mathml_node_to_argument(children[0], &format!("{path}/*[1]"), diagnostics, nodes)?;
    let marker = children.iter().position(|child| {
        child.namespace.as_deref() == Some(MATHML_NS) && child.local == "mprescripts"
    });
    if let Some(index) = marker {
        diagnose_empty_mathml_marker(
            children[index],
            &format!("{path}/*[{}]", index + 1),
            diagnostics,
        )?;
    }
    let (post, pre) = match marker {
        Some(index) => (&children[1..index], &children[index + 1..]),
        None => (&children[1..], &children[0..0]),
    };
    if post.len() > 2 || pre.len() > 2 || post.len() == 1 || pre.len() == 1 {
        diagnostics.push(path, "extra MathML multiscript pairs were discarded")?;
    }
    if !post.is_empty() && !pre.is_empty() {
        diagnostics.push(
            path,
            "mixed pre-scripts and post-scripts cannot be represented",
        )?;
        return Ok(Vec::new());
    }
    let pair = if pre.is_empty() { post } else { pre };
    let pair_start = if pre.is_empty() {
        1
    } else {
        marker.expect("pre-scripts require a marker") + 1
    };
    let subscript = pair
        .first()
        .map(|child| {
            mathml_optional_script(
                child,
                &format!("{path}/*[{}]", pair_start + 1),
                diagnostics,
                nodes,
            )
        })
        .transpose()?
        .unwrap_or_default();
    let superscript = pair
        .get(1)
        .map(|child| {
            mathml_optional_script(
                child,
                &format!("{path}/*[{}]", pair_start + 2),
                diagnostics,
                nodes,
            )
        })
        .transpose()?
        .unwrap_or_default();
    Ok(vec![if pre.is_empty() {
        MathExpression::SubSuperscript(MathSubSuperscript::new(base, subscript, superscript))
    } else {
        MathExpression::PreSubSuperscript(MathPreSubSuperscript::new(base, subscript, superscript))
    }])
}

fn mathml_optional_script(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
    nodes: &mut usize,
) -> Result<MathArgument> {
    if node.namespace.as_deref() == Some(MATHML_NS) && node.local == "none" {
        diagnose_empty_mathml_marker(node, path, diagnostics)?;
        return Ok(MathArgument::default());
    }
    mathml_node_to_argument(node, path, diagnostics, nodes)
}

fn parse_mathml_matrix(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
    nodes: &mut usize,
) -> Result<Vec<MathExpression>> {
    let row_nodes = mathml_element_children(node);
    if row_nodes.is_empty() || row_nodes.len() > MAX_ROWS {
        return Err(math_error(format!("{path} exceeds the matrix row limit")));
    }
    let mut rows = Vec::new();
    let mut width = None;
    for (row_index, row) in row_nodes.iter().enumerate() {
        if row.namespace.as_deref() != Some(MATHML_NS) || row.local != "mtr" {
            return Err(math_error(format!("{path} contains a non-row child")));
        }
        let row_path = format!("{path}/mtr[{}]", row_index + 1);
        diagnose_attributes(row, &row_path, &[], diagnostics)?;
        diagnose_structural_text(row, &row_path, diagnostics)?;
        let cells = mathml_element_children(row);
        if cells.is_empty() || cells.len() > MAX_COLUMNS {
            return Err(math_error(format!(
                "{path} exceeds the matrix column limit"
            )));
        }
        if width
            .replace(cells.len())
            .is_some_and(|value| value != cells.len())
        {
            return Err(math_error(format!("{path} has a ragged matrix")));
        }
        let mut converted = Vec::new();
        for (column_index, cell) in cells.iter().enumerate() {
            if cell.namespace.as_deref() != Some(MATHML_NS) || cell.local != "mtd" {
                return Err(math_error(format!("{path} contains a non-cell child")));
            }
            let cell_path = format!("{row_path}/mtd[{}]", column_index + 1);
            diagnose_attributes(cell, &cell_path, &[], diagnostics)?;
            converted.push(mathml_children_to_argument(
                cell,
                &cell_path,
                diagnostics,
                nodes,
            )?);
        }
        rows.push(MathMatrixRow::new(converted));
    }
    Ok(vec![MathExpression::Matrix(MathMatrix::new(rows))])
}

fn parse_mathml_limits_or_accent(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
    nodes: &mut usize,
) -> Result<Vec<MathExpression>> {
    diagnose_boolean_attribute(node, path, "accent", diagnostics)?;
    diagnose_boolean_attribute(node, path, "accentunder", diagnostics)?;
    let children = mathml_element_children(node);
    let expected = if node.local == "munderover" { 3 } else { 2 };
    require_child_count(path, &children, expected)?;
    let token_accent = children.get(1).is_some_and(|child| {
        child.namespace.as_deref() == Some(MATHML_NS)
            && child.local == "mo"
            && child.attribute("accent") == Some("true")
    });
    let accent = node.attribute("accent") == Some("true")
        || node.attribute("accentunder") == Some("true")
        || token_accent;
    if accent {
        if node.local != "mover" {
            diagnostics.push(path, "under-accent MathML cannot be represented")?;
            return Ok(Vec::new());
        }
        if children[1].namespace.as_deref() != Some(MATHML_NS) || children[1].local != "mo" {
            let _ = mathml_element_to_expressions(
                children[1],
                &format!("{path}/*[2]"),
                diagnostics,
                nodes,
            )?;
            diagnostics.push(path, "MathML accent requires an mo token")?;
            return Ok(Vec::new());
        }
        diagnose_attributes(
            children[1],
            &format!("{path}/*[2]"),
            &["accent"],
            diagnostics,
        )?;
        diagnose_boolean_attribute(children[1], &format!("{path}/*[2]"), "accent", diagnostics)?;
        diagnose_token_children(children[1], &format!("{path}/*[2]"), diagnostics)?;
        let character = match children[1].text().as_str() {
            "‾" | "¯" => "̄".to_owned(),
            "→" => "⃗".to_owned(),
            value => value.to_owned(),
        };
        if character.chars().count() != 1 {
            diagnostics.push(path, "MathML accent must contain one character")?;
            return Ok(Vec::new());
        }
        return Ok(vec![MathExpression::Accent(MathAccent::new(
            character,
            mathml_node_to_argument(children[0], &format!("{path}/*[1]"), diagnostics, nodes)?,
        ))]);
    }
    if let Some(operator) = mathml_nary_operator(children[0], &format!("{path}/*[1]"), diagnostics)?
    {
        let mut nary = MathNary::new(operator, MathArgument::default());
        nary.hide_subscript = node.local == "mover";
        nary.hide_superscript = node.local == "munder";
        if node.local == "munder" || node.local == "munderover" {
            nary.subscript =
                mathml_node_to_argument(children[1], &format!("{path}/*[2]"), diagnostics, nodes)?;
        }
        if node.local == "mover" {
            nary.superscript =
                mathml_node_to_argument(children[1], &format!("{path}/*[2]"), diagnostics, nodes)?;
        } else if node.local == "munderover" {
            nary.superscript =
                mathml_node_to_argument(children[2], &format!("{path}/*[3]"), diagnostics, nodes)?;
        }
        return Ok(vec![MathExpression::Nary(nary)]);
    }
    if node.local == "munder" {
        Ok(vec![MathExpression::LowerLimit(MathLimit::new(
            mathml_node_to_argument(children[0], &format!("{path}/*[1]"), diagnostics, nodes)?,
            mathml_node_to_argument(children[1], &format!("{path}/*[2]"), diagnostics, nodes)?,
        ))])
    } else if node.local == "mover" {
        Ok(vec![MathExpression::UpperLimit(MathLimit::new(
            mathml_node_to_argument(children[0], &format!("{path}/*[1]"), diagnostics, nodes)?,
            mathml_node_to_argument(children[1], &format!("{path}/*[2]"), diagnostics, nodes)?,
        ))])
    } else {
        diagnostics.push(path, "non-operator under-over pair cannot be represented")?;
        Ok(Vec::new())
    }
}

fn parse_mathml_fenced(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
    nodes: &mut usize,
) -> Result<Vec<MathExpression>> {
    diagnose_structural_text(node, path, diagnostics)?;
    let open = node.attribute("open").unwrap_or("(");
    let close = node.attribute("close").unwrap_or(")");
    let separators = node.attribute("separators").unwrap_or(",");
    if open.chars().count() > 1 || close.chars().count() > 1 || separators.chars().count() > 1 {
        return Err(math_error(format!("{path} has a non-scalar fence")));
    }
    let mut arguments = Vec::new();
    for (index, child) in node.elements().enumerate() {
        arguments.push(mathml_node_to_argument(
            child,
            &format!("{path}/*[{}]", index + 1),
            diagnostics,
            nodes,
        )?);
    }
    if arguments.is_empty() {
        return Err(math_error(format!("{path} requires fenced content")));
    }
    let mut delimiter = MathDelimiter::new(open, close, arguments);
    if delimiter.arguments.len() > 1 {
        delimiter.separator_character = separators.to_owned();
    }
    Ok(vec![MathExpression::Delimiter(delimiter)])
}

fn parse_explicit_fenced_row(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
    nodes: &mut usize,
) -> Result<Option<MathDelimiter>> {
    let children = mathml_element_children(node);
    if children.len() < 2 {
        return Ok(None);
    }
    let first = children[0];
    let last = children[children.len() - 1];
    if !is_mathml_fence(first, true) || !is_mathml_fence(last, false) {
        return Ok(None);
    }
    let mut arguments = Vec::new();
    let mut current = Vec::new();
    let mut separator = None::<String>;
    diagnose_attributes(
        first,
        &format!("{path}/*[1]"),
        &["fence", "stretchy", "form"],
        diagnostics,
    )?;
    diagnose_boolean_attribute(first, &format!("{path}/*[1]"), "fence", diagnostics)?;
    diagnose_boolean_attribute(first, &format!("{path}/*[1]"), "stretchy", diagnostics)?;
    diagnose_enum_attribute(
        first,
        &format!("{path}/*[1]"),
        "form",
        &["prefix", "infix", "postfix"],
        diagnostics,
    )?;
    diagnose_token_children(first, &format!("{path}/*[1]"), diagnostics)?;
    diagnose_attributes(
        last,
        &format!("{path}/*[{}]", children.len()),
        &["fence", "stretchy", "form"],
        diagnostics,
    )?;
    diagnose_boolean_attribute(
        last,
        &format!("{path}/*[{}]", children.len()),
        "fence",
        diagnostics,
    )?;
    diagnose_boolean_attribute(
        last,
        &format!("{path}/*[{}]", children.len()),
        "stretchy",
        diagnostics,
    )?;
    diagnose_enum_attribute(
        last,
        &format!("{path}/*[{}]", children.len()),
        "form",
        &["prefix", "infix", "postfix"],
        diagnostics,
    )?;
    diagnose_token_children(last, &format!("{path}/*[{}]", children.len()), diagnostics)?;
    for (index, child) in children[1..children.len() - 1].iter().enumerate() {
        if child.namespace.as_deref() == Some(MATHML_NS)
            && child.local == "mo"
            && child.attribute("separator") == Some("true")
            && child.text().chars().count() <= 1
        {
            diagnose_attributes(
                child,
                &format!("{path}/*[{}]", index + 2),
                &["separator"],
                diagnostics,
            )?;
            diagnose_boolean_attribute(
                child,
                &format!("{path}/*[{}]", index + 2),
                "separator",
                diagnostics,
            )?;
            diagnose_token_children(child, &format!("{path}/*[{}]", index + 2), diagnostics)?;
            let candidate = child.text();
            if separator
                .as_ref()
                .is_some_and(|current| current != &candidate)
            {
                diagnostics.push(
                    format!("{path}/*[{}]", index + 2),
                    "mixed MathML delimiter separators were normalized to the first character",
                )?;
            } else {
                separator.get_or_insert(candidate);
            }
            arguments.push(MathArgument::new(current));
            current = Vec::new();
        } else {
            current.extend(mathml_element_to_expressions(
                child,
                &format!("{path}/*[{}]", index + 2),
                diagnostics,
                nodes,
            )?);
        }
    }
    arguments.push(MathArgument::new(current));
    let mut delimiter = MathDelimiter::new(first.text(), last.text(), arguments);
    if delimiter.arguments.len() > 1 {
        delimiter.separator_character = separator.unwrap_or_else(|| ",".to_owned());
    }
    Ok(Some(delimiter))
}

fn is_mathml_fence(node: &XmlNode, leading: bool) -> bool {
    if node.namespace.as_deref() != Some(MATHML_NS) || node.local != "mo" {
        return false;
    }
    let character = node.text();
    let expected_form = if leading { "prefix" } else { "postfix" };
    let fence = node.attribute("fence");
    let stretchy = node.attribute("stretchy");
    let form = node.attribute("form");
    character.chars().count() <= 1
        && fence != Some("false")
        && stretchy != Some("false")
        && form.is_none_or(|value| value == expected_form)
        && (fence == Some("true") || (stretchy == Some("true") && form == Some(expected_form)))
}

fn parse_explicit_nary_row(
    node: &XmlNode,
    path: &str,
    diagnostics: &mut Diagnostics,
    nodes: &mut usize,
) -> Result<Option<MathNary>> {
    let children = mathml_element_children(node);
    if children.len() != 2 {
        return Ok(None);
    }
    let mut nary = if children[0].namespace.as_deref() == Some(MATHML_NS)
        && children[0].local == "mo"
        && children[0].attribute("largeop") == Some("true")
        && is_nary_operator(&children[0].text())
    {
        let operator = mathml_nary_operator(children[0], &format!("{path}/*[1]"), diagnostics)?
            .expect("the guarded MathML operator is n-ary");
        MathNary::new(operator, MathArgument::default())
    } else {
        let head_children = mathml_element_children(children[0]);
        if children[0].namespace.as_deref() != Some(MATHML_NS)
            || !matches!(
                children[0].local.as_str(),
                "msub" | "msup" | "msubsup" | "munder" | "mover" | "munderover"
            )
            || !head_children.first().is_some_and(|operator| {
                operator.namespace.as_deref() == Some(MATHML_NS)
                    && operator.local == "mo"
                    && is_nary_operator(&operator.text())
            })
        {
            return Ok(None);
        }
        let converted = mathml_element_to_expressions(
            children[0],
            &format!("{path}/*[1]"),
            diagnostics,
            nodes,
        )?;
        let [MathExpression::Nary(nary)] = converted.as_slice() else {
            return Ok(None);
        };
        if !nary.base.expressions.is_empty() {
            return Ok(None);
        }
        nary.clone()
    };
    nary.base = mathml_node_to_argument(children[1], &format!("{path}/*[2]"), diagnostics, nodes)?;
    Ok(Some(nary))
}

fn normalize_argument(argument: &mut MathArgument) {
    for expression in &mut argument.expressions {
        normalize_expression(expression);
    }
    let mut normalized: Vec<MathExpression> = Vec::with_capacity(argument.expressions.len());
    for expression in argument.expressions.drain(..) {
        let expression_is_fully_modelled = !expression.has_unsupported_content();
        let previous_is_fully_modelled = normalized
            .last()
            .is_some_and(|previous| !previous.has_unsupported_content());
        if let MathExpression::Run(run) = &expression
            && let Some(MathExpression::Run(previous)) = normalized.last_mut()
            && expression_is_fully_modelled
            && previous_is_fully_modelled
            && previous.properties == Default::default()
            && run.properties == Default::default()
        {
            previous.text.push_str(&run.text);
            continue;
        }
        normalized.push(expression);
    }
    argument.expressions = normalized;
}

fn normalize_expression(expression: &mut MathExpression) {
    match expression {
        MathExpression::Run(_) => {}
        MathExpression::Fraction(value) => {
            normalize_argument(&mut value.numerator);
            normalize_argument(&mut value.denominator);
        }
        MathExpression::Subscript(value) | MathExpression::Superscript(value) => {
            normalize_argument(&mut value.base);
            normalize_argument(&mut value.script);
        }
        MathExpression::SubSuperscript(value) => {
            normalize_argument(&mut value.base);
            normalize_argument(&mut value.subscript);
            normalize_argument(&mut value.superscript);
        }
        MathExpression::PreSubSuperscript(value) => {
            normalize_argument(&mut value.base);
            normalize_argument(&mut value.subscript);
            normalize_argument(&mut value.superscript);
        }
        MathExpression::Radical(value) => {
            normalize_argument(&mut value.degree);
            normalize_argument(&mut value.base);
        }
        MathExpression::Matrix(value) => {
            for row in &mut value.rows {
                for cell in &mut row.cells {
                    normalize_argument(cell);
                }
            }
        }
        MathExpression::LowerLimit(value) | MathExpression::UpperLimit(value) => {
            normalize_argument(&mut value.base);
            normalize_argument(&mut value.limit);
        }
        MathExpression::Nary(value) => {
            normalize_argument(&mut value.base);
            normalize_argument(&mut value.subscript);
            normalize_argument(&mut value.superscript);
        }
        MathExpression::Delimiter(value) => {
            for argument in &mut value.arguments {
                normalize_argument(argument);
            }
        }
        MathExpression::Accent(value) => normalize_argument(&mut value.base),
    }
}

fn validate_tree(argument: &MathArgument) -> std::result::Result<(), String> {
    fn visit(
        argument: &MathArgument,
        depth: usize,
        nodes: &mut usize,
        text: &mut usize,
    ) -> std::result::Result<(), String> {
        if depth > MAX_DEPTH {
            return Err("equation tree exceeds the depth limit".to_owned());
        }
        for expression in &argument.expressions {
            *nodes = nodes
                .checked_add(1)
                .ok_or_else(|| "equation node count overflowed".to_owned())?;
            if *nodes > MAX_NODES {
                return Err("equation tree exceeds the node limit".to_owned());
            }
            match expression {
                MathExpression::Run(run) => {
                    *text = text
                        .checked_add(run.text.len())
                        .ok_or_else(|| "equation text size overflowed".to_owned())?;
                    if *text > MAX_TEXT_BYTES {
                        return Err("equation tree exceeds the text limit".to_owned());
                    }
                }
                MathExpression::Fraction(value) => {
                    visit(&value.numerator, depth + 1, nodes, text)?;
                    visit(&value.denominator, depth + 1, nodes, text)?;
                }
                MathExpression::Subscript(value) | MathExpression::Superscript(value) => {
                    visit(&value.base, depth + 1, nodes, text)?;
                    visit(&value.script, depth + 1, nodes, text)?;
                }
                MathExpression::SubSuperscript(value) => {
                    visit(&value.base, depth + 1, nodes, text)?;
                    visit(&value.subscript, depth + 1, nodes, text)?;
                    visit(&value.superscript, depth + 1, nodes, text)?;
                }
                MathExpression::PreSubSuperscript(value) => {
                    visit(&value.base, depth + 1, nodes, text)?;
                    visit(&value.subscript, depth + 1, nodes, text)?;
                    visit(&value.superscript, depth + 1, nodes, text)?;
                }
                MathExpression::Radical(value) => {
                    visit(&value.degree, depth + 1, nodes, text)?;
                    visit(&value.base, depth + 1, nodes, text)?;
                }
                MathExpression::Matrix(value) => {
                    if value.rows.is_empty() || value.rows.len() > MAX_ROWS {
                        return Err("equation matrix exceeds the row limit".to_owned());
                    }
                    let mut width = None;
                    for row in &value.rows {
                        if row.cells.is_empty() || row.cells.len() > MAX_COLUMNS {
                            return Err("equation matrix exceeds the column limit".to_owned());
                        }
                        if width
                            .replace(row.cells.len())
                            .is_some_and(|previous| previous != row.cells.len())
                        {
                            return Err("equation matrix is ragged".to_owned());
                        }
                        for cell in &row.cells {
                            visit(cell, depth + 1, nodes, text)?;
                        }
                    }
                }
                MathExpression::LowerLimit(value) | MathExpression::UpperLimit(value) => {
                    visit(&value.base, depth + 1, nodes, text)?;
                    visit(&value.limit, depth + 1, nodes, text)?;
                }
                MathExpression::Nary(value) => {
                    visit(&value.base, depth + 1, nodes, text)?;
                    visit(&value.subscript, depth + 1, nodes, text)?;
                    visit(&value.superscript, depth + 1, nodes, text)?;
                }
                MathExpression::Delimiter(value) => {
                    if value.arguments.is_empty() {
                        return Err("equation delimiter has no arguments".to_owned());
                    }
                    for child in &value.arguments {
                        visit(child, depth + 1, nodes, text)?;
                    }
                }
                MathExpression::Accent(value) => visit(&value.base, depth + 1, nodes, text)?,
            }
        }
        Ok(())
    }
    visit(argument, 0, &mut 0, &mut 0)
}

fn argument_has_direct_unsupported_content(argument: &MathArgument) -> bool {
    let mut direct = argument.clone();
    direct.expressions.clear();
    direct.has_unsupported_content()
}

fn expression_has_direct_unsupported_content(expression: &MathExpression) -> bool {
    let mut direct = expression.clone();
    match &mut direct {
        MathExpression::Run(_) => {}
        MathExpression::Fraction(value) => {
            value.numerator = MathArgument::default();
            value.denominator = MathArgument::default();
        }
        MathExpression::Subscript(value) | MathExpression::Superscript(value) => {
            value.base = MathArgument::default();
            value.script = MathArgument::default();
        }
        MathExpression::SubSuperscript(value) => {
            value.base = MathArgument::default();
            value.subscript = MathArgument::default();
            value.superscript = MathArgument::default();
        }
        MathExpression::PreSubSuperscript(value) => {
            value.base = MathArgument::default();
            value.subscript = MathArgument::default();
            value.superscript = MathArgument::default();
        }
        MathExpression::Radical(value) => {
            value.degree = MathArgument::default();
            value.base = MathArgument::default();
        }
        MathExpression::Matrix(value) => {
            for row in &mut value.rows {
                row.cells.clear();
            }
        }
        MathExpression::LowerLimit(value) | MathExpression::UpperLimit(value) => {
            value.base = MathArgument::default();
            value.limit = MathArgument::default();
        }
        MathExpression::Nary(value) => {
            value.base = MathArgument::default();
            value.subscript = MathArgument::default();
            value.superscript = MathArgument::default();
        }
        MathExpression::Delimiter(value) => value.arguments.clear(),
        MathExpression::Accent(value) => value.base = MathArgument::default(),
    }
    direct.has_unsupported_content()
}

fn write_mathml_argument(
    argument: &MathArgument,
    path: &str,
    output: &mut String,
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    if argument_has_direct_unsupported_content(argument) {
        diagnostics.push(path, "unmodelled OfficeMath argument content was discarded")?;
    }
    for (index, expression) in argument.expressions.iter().enumerate() {
        write_mathml_expression(
            expression,
            &format!("{path}/*[{}]", index + 1),
            output,
            diagnostics,
        )?;
    }
    Ok(())
}

fn write_mathml_expression(
    expression: &MathExpression,
    path: &str,
    output: &mut String,
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    if expression_has_direct_unsupported_content(expression) {
        diagnostics.push(path, "unmodelled OfficeMath content was discarded")?;
    }
    match expression {
        MathExpression::Run(run) => {
            if run.properties != Default::default() {
                diagnostics.push(
                    format!("{path}/rPr"),
                    "OfficeMath run properties were discarded",
                )?;
            }
            if !is_xml_1_0_text(&run.text) {
                diagnostics.push(
                    path,
                    "OfficeMath run contains a forbidden XML 1.0 character",
                )?;
                return Ok(());
            }
            output.push_str("<mi>");
            output.push_str(&xml_escape(&run.text));
            output.push_str("</mi>");
        }
        MathExpression::Fraction(value) => {
            if value.fraction_type != FractionType::Bar {
                diagnostics.push(path, "non-bar OfficeMath fraction type was discarded")?;
            }
            output.push_str("<mfrac><mrow>");
            write_mathml_argument(&value.numerator, path, output, diagnostics)?;
            output.push_str("</mrow><mrow>");
            write_mathml_argument(&value.denominator, path, output, diagnostics)?;
            output.push_str("</mrow></mfrac>");
        }
        MathExpression::Subscript(value) | MathExpression::Superscript(value) => {
            if value.alignment.is_some() {
                diagnostics.push(path, "OfficeMath script alignment was discarded")?;
            }
            output.push_str(if matches!(expression, MathExpression::Subscript(_)) {
                "<msub><mrow>"
            } else {
                "<msup><mrow>"
            });
            write_mathml_argument(&value.base, path, output, diagnostics)?;
            output.push_str("</mrow><mrow>");
            write_mathml_argument(&value.script, path, output, diagnostics)?;
            output.push_str(if matches!(expression, MathExpression::Subscript(_)) {
                "</mrow></msub>"
            } else {
                "</mrow></msup>"
            });
        }
        MathExpression::SubSuperscript(value) => {
            if value.alignment.is_some() {
                diagnostics.push(path, "OfficeMath script alignment was discarded")?;
            }
            output.push_str("<msubsup><mrow>");
            write_mathml_argument(&value.base, path, output, diagnostics)?;
            output.push_str("</mrow><mrow>");
            write_mathml_argument(&value.subscript, path, output, diagnostics)?;
            output.push_str("</mrow><mrow>");
            write_mathml_argument(&value.superscript, path, output, diagnostics)?;
            output.push_str("</mrow></msubsup>");
        }
        MathExpression::PreSubSuperscript(value) => {
            if value.alignment.is_some() {
                diagnostics.push(path, "OfficeMath script alignment was discarded")?;
            }
            output.push_str("<mmultiscripts><mrow>");
            write_mathml_argument(&value.base, path, output, diagnostics)?;
            output.push_str("</mrow><mprescripts/><mrow>");
            write_mathml_argument(&value.subscript, path, output, diagnostics)?;
            output.push_str("</mrow><mrow>");
            write_mathml_argument(&value.superscript, path, output, diagnostics)?;
            output.push_str("</mrow></mmultiscripts>");
        }
        MathExpression::Radical(value) => {
            if value.hide_degree != value.degree.expressions.is_empty() {
                diagnostics.push(path, "OfficeMath radical degree visibility was discarded")?;
            }
            if value.degree.expressions.is_empty() {
                output.push_str("<msqrt>");
                write_mathml_argument(&value.base, path, output, diagnostics)?;
                output.push_str("</msqrt>");
            } else {
                output.push_str("<mroot><mrow>");
                write_mathml_argument(&value.base, path, output, diagnostics)?;
                output.push_str("</mrow><mrow>");
                write_mathml_argument(&value.degree, path, output, diagnostics)?;
                output.push_str("</mrow></mroot>");
            }
        }
        MathExpression::Matrix(value) => {
            if value.properties != Default::default() {
                diagnostics.push(path, "OfficeMath matrix properties were discarded")?;
            }
            output.push_str("<mtable>");
            for row in &value.rows {
                output.push_str("<mtr>");
                for cell in &row.cells {
                    output.push_str("<mtd>");
                    write_mathml_argument(cell, path, output, diagnostics)?;
                    output.push_str("</mtd>");
                }
                output.push_str("</mtr>");
            }
            output.push_str("</mtable>");
        }
        MathExpression::LowerLimit(value) | MathExpression::UpperLimit(value) => {
            output.push_str(if matches!(expression, MathExpression::LowerLimit(_)) {
                "<munder><mrow>"
            } else {
                "<mover><mrow>"
            });
            write_mathml_argument(&value.base, path, output, diagnostics)?;
            output.push_str("</mrow><mrow>");
            write_mathml_argument(&value.limit, path, output, diagnostics)?;
            output.push_str(if matches!(expression, MathExpression::LowerLimit(_)) {
                "</mrow></munder>"
            } else {
                "</mrow></mover>"
            });
        }
        MathExpression::Nary(value) => write_mathml_nary(value, path, output, diagnostics)?,
        MathExpression::Delimiter(value) => {
            if value.grow.is_some() {
                diagnostics.push(path, "OfficeMath delimiter growth setting was discarded")?;
            }
            if value.begin_character.chars().count() > 1
                || value.end_character.chars().count() > 1
                || value.separator_character.chars().count() > 1
            {
                diagnostics.push(path, "non-scalar OfficeMath delimiter was discarded")?;
                return Ok(());
            }
            if !is_xml_1_0_text(&value.begin_character)
                || !is_xml_1_0_text(&value.end_character)
                || !is_xml_1_0_text(&value.separator_character)
            {
                diagnostics.push(
                    path,
                    "OfficeMath delimiter contains a forbidden XML 1.0 character",
                )?;
                return Ok(());
            }
            write!(
                output,
                "<mrow><mo fence=\"true\" stretchy=\"true\">{}</mo>",
                xml_escape(&value.begin_character)
            )
            .expect("String writes are infallible");
            for (index, argument) in value.arguments.iter().enumerate() {
                if index != 0 {
                    write!(
                        output,
                        "<mo separator=\"true\">{}</mo>",
                        xml_escape(&value.separator_character)
                    )
                    .expect("String writes are infallible");
                }
                write_mathml_argument(argument, path, output, diagnostics)?;
            }
            write!(
                output,
                "<mo fence=\"true\" stretchy=\"true\">{}</mo></mrow>",
                xml_escape(&value.end_character)
            )
            .expect("String writes are infallible");
        }
        MathExpression::Accent(value) => {
            if value.character.chars().count() != 1 {
                diagnostics.push(path, "non-scalar OfficeMath accent was discarded")?;
                return Ok(());
            }
            if !is_xml_1_0_text(&value.character) {
                diagnostics.push(
                    path,
                    "OfficeMath accent contains a forbidden XML 1.0 character",
                )?;
                return Ok(());
            }
            output.push_str("<mover accent=\"true\"><mrow>");
            write_mathml_argument(&value.base, path, output, diagnostics)?;
            write!(
                output,
                "</mrow><mo>{}</mo></mover>",
                xml_escape(&value.character)
            )
            .expect("String writes are infallible");
        }
    }
    Ok(())
}

fn write_mathml_nary(
    value: &MathNary,
    path: &str,
    output: &mut String,
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    if !is_nary_operator(&value.character) {
        diagnostics.push(path, "unsupported OfficeMath n-ary operator was discarded")?;
        return Ok(());
    }
    if value.grow.is_some() || value.limit_location.is_some() {
        diagnostics.push(path, "OfficeMath n-ary layout properties were discarded")?;
    }
    let sub = !value.hide_subscript && !value.subscript.expressions.is_empty();
    let sup = !value.hide_superscript && !value.superscript.expressions.is_empty();
    if value.hide_subscript && !value.subscript.expressions.is_empty() {
        diagnostics.push(path, "hidden OfficeMath n-ary subscript was discarded")?;
    } else if !value.hide_subscript && value.subscript.expressions.is_empty() {
        diagnostics.push(path, "OfficeMath n-ary subscript visibility was discarded")?;
    }
    if value.hide_superscript && !value.superscript.expressions.is_empty() {
        diagnostics.push(path, "hidden OfficeMath n-ary superscript was discarded")?;
    } else if !value.hide_superscript && value.superscript.expressions.is_empty() {
        diagnostics.push(
            path,
            "OfficeMath n-ary superscript visibility was discarded",
        )?;
    }
    output.push_str("<mrow>");
    if sub && sup {
        output.push_str("<munderover><mo largeop=\"true\">");
    } else if sub {
        output.push_str("<munder><mo largeop=\"true\">");
    } else if sup {
        output.push_str("<mover><mo largeop=\"true\">");
    }
    if sub || sup {
        output.push_str(&xml_escape(&value.character));
        output.push_str("</mo><mrow>");
        if sub {
            write_mathml_argument(&value.subscript, path, output, diagnostics)?;
        } else {
            write_mathml_argument(&value.superscript, path, output, diagnostics)?;
        }
        output.push_str("</mrow>");
        if sub && sup {
            output.push_str("<mrow>");
            write_mathml_argument(&value.superscript, path, output, diagnostics)?;
            output.push_str("</mrow></munderover>");
        } else if sub {
            output.push_str("</munder>");
        } else {
            output.push_str("</mover>");
        }
    } else {
        write!(
            output,
            "<mo largeop=\"true\">{}</mo>",
            xml_escape(&value.character)
        )
        .expect("String writes are infallible");
    }
    output.push_str("<mrow>");
    write_mathml_argument(&value.base, path, output, diagnostics)?;
    output.push_str("</mrow></mrow>");
    Ok(())
}

fn xml_escape(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn is_xml_1_0_text(value: &str) -> bool {
    value.chars().all(|character| {
        matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
            || ('\u{20}'..='\u{d7ff}').contains(&character)
            || ('\u{e000}'..='\u{fffd}').contains(&character)
            || ('\u{10000}'..='\u{10ffff}').contains(&character)
    })
}

fn is_nary_operator(value: &str) -> bool {
    matches!(value, "∑" | "∏" | "∐" | "∫" | "∬" | "∭" | "∮")
}

struct LatexParser<'a> {
    input: &'a str,
    offset: usize,
    depth: usize,
    nodes: usize,
    text_bytes: usize,
    tokens: usize,
    diagnostics: Diagnostics,
}

impl<'a> LatexParser<'a> {
    fn new(input: &'a str) -> Self {
        Self {
            input,
            offset: 0,
            depth: 0,
            nodes: 0,
            text_bytes: 0,
            tokens: 0,
            diagnostics: Diagnostics::new(),
        }
    }

    fn nested(input: &'a str, depth: usize) -> Self {
        let mut parser = Self::new(input);
        parser.depth = depth;
        parser
    }

    fn parse_argument(&mut self, terminator: Option<char>) -> Result<MathArgument> {
        self.depth += 1;
        if self.depth > MAX_DEPTH {
            return Err(self.error("LaTeX input exceeds the depth limit"));
        }
        let mut expressions = Vec::new();
        loop {
            self.skip_whitespace();
            if self.at_end() || terminator.is_some_and(|value| self.peek() == Some(value)) {
                break;
            }
            if latex_command_at(self.input, self.offset)
                .is_some_and(|(command, _)| matches!(command, "right" | "end"))
            {
                break;
            }
            let atom = self.parse_complete_atom()?;
            expressions.extend(atom.expressions);
        }
        if let Some(value) = terminator {
            if self.peek() != Some(value) {
                return Err(self.error(format!("missing closing {value}")));
            }
            self.bump();
        }
        self.depth -= 1;
        let mut argument = MathArgument::new(expressions);
        normalize_argument(&mut argument);
        Ok(argument)
    }

    fn parse_atom(&mut self) -> Result<MathArgument> {
        self.bump_token()?;
        let start = self.offset;
        match self.peek() {
            Some('{') => {
                self.bump();
                self.parse_argument(Some('}'))
            }
            Some('\\') => self.parse_command(),
            Some('&') => {
                self.bump();
                Err(self.error("matrix separator outside a matrix"))
            }
            Some('}') | Some('[') | Some(']') => Err(self.error("unexpected LaTeX delimiter")),
            Some(_) => {
                let mut text = String::new();
                while let Some(character) = self.peek() {
                    if character.is_whitespace()
                        || matches!(character, '{' | '}' | '[' | ']' | '\\' | '_' | '^' | '&')
                    {
                        break;
                    }
                    text.push(character);
                    self.bump();
                }
                self.add_text(text.len())?;
                self.add_node()?;
                Ok(MathArgument::text(text))
            }
            None => Err(math_error(format!("missing LaTeX atom at byte {start}"))),
        }
    }

    fn parse_complete_atom(&mut self) -> Result<MathArgument> {
        let atom_offset = self.offset;
        let mut atom = self.parse_atom()?;
        self.skip_whitespace();
        if matches!(atom.expressions.as_slice(), [MathExpression::Nary(_)])
            && let Some((command, end)) = latex_command_at(self.input, self.offset)
            && matches!(command, "limits" | "nolimits")
        {
            self.diagnostics.push(
                format!("byte:{}", self.offset),
                format!("LaTeX \\{command} n-ary placement was discarded"),
            )?;
            self.bump_token()?;
            self.offset = end;
            self.skip_whitespace();
        }
        if matches!(atom.expressions.as_slice(), [MathExpression::Nary(_)])
            && self.input[self.offset..].starts_with("{}")
        {
            self.bump_token()?;
            self.bump();
            self.bump();
        }
        let (subscript, superscript) = self.parse_following_scripts()?;
        atom = self.attach_following_scripts(atom, subscript, superscript, atom_offset)?;
        if let Some(MathExpression::Nary(nary)) = atom.expressions.first_mut()
            && nary.base.expressions.is_empty()
            && !self.at_end()
        {
            nary.base = self.parse_complete_atom()?;
        }
        Ok(atom)
    }

    fn parse_following_scripts(&mut self) -> Result<(Option<MathArgument>, Option<MathArgument>)> {
        let mut subscript = None;
        let mut superscript = None;
        loop {
            self.skip_whitespace();
            match self.peek() {
                Some('_') => {
                    self.bump();
                    if subscript.is_some() {
                        return Err(self.error("duplicate subscript"));
                    }
                    subscript = Some(self.parse_required_argument()?);
                }
                Some('^') => {
                    self.bump();
                    if superscript.is_some() {
                        return Err(self.error("duplicate superscript"));
                    }
                    superscript = Some(self.parse_required_argument()?);
                }
                _ => return Ok((subscript, superscript)),
            }
        }
    }

    fn attach_following_scripts(
        &mut self,
        atom: MathArgument,
        subscript: Option<MathArgument>,
        superscript: Option<MathArgument>,
        offset: usize,
    ) -> Result<MathArgument> {
        if subscript.is_none() && superscript.is_none() {
            return Ok(atom);
        }
        if !atom.expressions.is_empty() {
            return self.apply_scripts(atom, subscript, superscript, offset);
        }
        let base = self.parse_complete_atom()?;
        self.add_node()?;
        Ok(MathArgument::new(vec![MathExpression::PreSubSuperscript(
            MathPreSubSuperscript::new(
                base,
                subscript.unwrap_or_default(),
                superscript.unwrap_or_default(),
            ),
        )]))
    }

    fn parse_command(&mut self) -> Result<MathArgument> {
        let start = self.offset;
        self.bump();
        let command = self.read_command_name();
        match command.as_str() {
            "frac" => {
                let numerator = self.parse_required_group()?;
                let denominator = self.parse_required_group()?;
                self.add_node()?;
                Ok(MathArgument::new(vec![MathExpression::Fraction(
                    MathFraction::new(numerator, denominator),
                )]))
            }
            "sqrt" => {
                self.skip_whitespace();
                let degree = if self.peek() == Some('[') {
                    self.bump();
                    Some(self.parse_argument(Some(']'))?)
                } else {
                    None
                };
                let base = self.parse_required_group()?;
                self.add_node()?;
                Ok(MathArgument::new(vec![MathExpression::Radical(
                    match degree {
                        Some(degree) => MathRadical::with_degree(degree, base),
                        None => MathRadical::new(base),
                    },
                )]))
            }
            "sum" | "prod" | "coprod" | "int" | "iint" | "iiint" | "oint" => {
                self.add_node()?;
                Ok(MathArgument::new(vec![MathExpression::Nary(
                    MathNary::new(latex_nary_character(&command), MathArgument::default()),
                )]))
            }
            "left" => self.parse_left_right(start),
            "hat" | "widehat" | "bar" | "overline" | "vec" | "overrightarrow" | "tilde"
            | "widetilde" | "dot" | "ddot" => {
                let base = self.parse_required_argument()?;
                self.add_node()?;
                Ok(MathArgument::new(vec![MathExpression::Accent(
                    MathAccent::new(latex_accent_character(&command), base),
                )]))
            }
            "underset" | "overset" => {
                let limit = self.parse_required_group()?;
                let base = self.parse_required_group()?;
                self.add_node()?;
                Ok(MathArgument::new(vec![if command == "underset" {
                    MathExpression::LowerLimit(MathLimit::new(base, limit))
                } else {
                    MathExpression::UpperLimit(MathLimit::new(base, limit))
                }]))
            }
            "begin" => self.parse_matrix_environment(start),
            "," | ";" | "!" | " " => Ok(MathArgument::default()),
            "{" | "}" | "_" | "^" | "#" | "$" | "%" | "&" | "\\" => {
                self.add_text(command.len())?;
                self.add_node()?;
                Ok(MathArgument::text(command))
            }
            "lbrack" => {
                self.add_text(1)?;
                self.add_node()?;
                Ok(MathArgument::text("["))
            }
            "rbrack" => {
                self.add_text(1)?;
                self.add_node()?;
                Ok(MathArgument::text("]"))
            }
            "backslash" => {
                self.add_text(1)?;
                self.add_node()?;
                Ok(MathArgument::text("\\"))
            }
            _ => {
                self.diagnostics.push(
                    format!("byte:{start}"),
                    format!("unsupported LaTeX command \\{command} was discarded"),
                )?;
                self.skip_optional_lossy_argument()?;
                Ok(MathArgument::default())
            }
        }
    }

    fn parse_left_right(&mut self, start: usize) -> Result<MathArgument> {
        self.skip_whitespace();
        let begin = self.read_delimiter()?;
        let content_start = self.offset;
        let Some(relative_end) = find_matching_right(&self.input[content_start..]) else {
            return Err(math_error(format!("missing \\right for byte {start}")));
        };
        let content_end = content_start + relative_end;
        let content = &self.input[content_start..content_end];
        let (parts, separator) = split_latex_fenced(content)?;
        let mut arguments = Vec::new();
        for (part_offset, part) in parts {
            let mut nested = LatexParser::nested(part, self.depth);
            let argument = nested.parse_argument(None)?;
            nested.skip_whitespace();
            if !nested.at_end() {
                return Err(math_error(format!(
                    "unexpected trailing nested LaTeX input at byte {}",
                    content_start + part_offset + nested.offset
                )));
            }
            arguments.push(argument);
            self.absorb_nested(nested, content_start + part_offset)?;
        }
        self.offset = content_end + "\\right".len();
        self.skip_whitespace();
        let end = self.read_delimiter()?;
        self.add_node()?;
        let mut delimiter = MathDelimiter::new(begin, end, arguments);
        if delimiter.arguments.len() > 1 {
            delimiter.separator_character = separator;
        }
        Ok(MathArgument::new(vec![MathExpression::Delimiter(
            delimiter,
        )]))
    }

    fn parse_matrix_environment(&mut self, start: usize) -> Result<MathArgument> {
        let environment = self.read_raw_group()?;
        let body_start = self.offset;
        let Some((relative_end, relative_after)) =
            find_matching_environment(&self.input[body_start..], &environment)?
        else {
            return Err(math_error(format!(
                "missing \\end{{{environment}}} for byte {start}"
            )));
        };
        let body_end = body_start + relative_end;
        let body = &self.input[body_start..body_end];
        self.offset = body_start + relative_after;
        if !matches!(environment.as_str(), "matrix" | "pmatrix" | "bmatrix") {
            self.diagnostics.push(
                format!("byte:{start}"),
                format!("unsupported LaTeX environment {environment} was discarded"),
            )?;
            return Ok(MathArgument::default());
        }
        let mut rows = Vec::new();
        let mut width = None;
        for row in split_latex_matrix(body)? {
            if rows.len() >= MAX_ROWS {
                return Err(self.error("LaTeX matrix exceeds the row limit"));
            }
            let mut cells = Vec::new();
            for (cell_offset, cell_text) in row {
                if cells.len() >= MAX_COLUMNS {
                    return Err(self.error("LaTeX matrix exceeds the column limit"));
                }
                let mut nested = LatexParser::nested(cell_text, self.depth);
                let cell = nested.parse_argument(None)?;
                nested.skip_whitespace();
                if !nested.at_end() {
                    return Err(math_error(format!(
                        "unexpected trailing nested LaTeX input at byte {}",
                        body_start + cell_offset + nested.offset
                    )));
                }
                cells.push(cell);
                self.absorb_nested(nested, body_start + cell_offset)?;
            }
            if width
                .replace(cells.len())
                .is_some_and(|value| value != cells.len())
            {
                return Err(self.error("LaTeX matrix is ragged"));
            }
            rows.push(MathMatrixRow::new(cells));
        }
        if rows.is_empty() || rows.iter().any(|row| row.cells.is_empty()) {
            return Err(self.error("LaTeX matrix must contain cells"));
        }
        self.add_node()?;
        let matrix = MathExpression::Matrix(MathMatrix::new(rows));
        let expression = match environment.as_str() {
            "pmatrix" => MathExpression::Delimiter(MathDelimiter::new(
                "(",
                ")",
                vec![MathArgument::new(vec![matrix])],
            )),
            "bmatrix" => MathExpression::Delimiter(MathDelimiter::new(
                "[",
                "]",
                vec![MathArgument::new(vec![matrix])],
            )),
            _ => matrix,
        };
        Ok(MathArgument::new(vec![expression]))
    }

    fn apply_scripts(
        &mut self,
        base: MathArgument,
        subscript: Option<MathArgument>,
        superscript: Option<MathArgument>,
        offset: usize,
    ) -> Result<MathArgument> {
        self.add_node()?;
        let expression = if let [MathExpression::Nary(nary)] = base.expressions.as_slice() {
            let mut nary = nary.clone();
            if let Some(subscript) = subscript {
                nary.subscript = subscript;
                nary.hide_subscript = false;
            }
            if let Some(superscript) = superscript {
                nary.superscript = superscript;
                nary.hide_superscript = false;
            }
            MathExpression::Nary(nary)
        } else {
            match (subscript, superscript) {
                (Some(subscript), Some(superscript)) => MathExpression::SubSuperscript(
                    MathSubSuperscript::new(base, subscript, superscript),
                ),
                (Some(script), None) => MathExpression::Subscript(MathScript::new(base, script)),
                (None, Some(script)) => MathExpression::Superscript(MathScript::new(base, script)),
                (None, None) => return Err(math_error(format!("empty scripts at byte {offset}"))),
            }
        };
        Ok(MathArgument::new(vec![expression]))
    }

    fn parse_required_group(&mut self) -> Result<MathArgument> {
        self.skip_whitespace();
        if self.peek() != Some('{') {
            return Err(self.error("expected a braced argument"));
        }
        self.bump();
        self.parse_argument(Some('}'))
    }

    fn parse_required_argument(&mut self) -> Result<MathArgument> {
        self.skip_whitespace();
        self.parse_atom()
    }

    fn read_raw_group(&mut self) -> Result<String> {
        self.skip_whitespace();
        if self.peek() != Some('{') {
            return Err(self.error("expected a braced name"));
        }
        self.bump();
        let start = self.offset;
        while !self.at_end() && self.peek() != Some('}') {
            self.bump();
        }
        if self.peek() != Some('}') {
            return Err(self.error("unterminated braced name"));
        }
        let value = self.input[start..self.offset].to_owned();
        self.bump();
        Ok(value)
    }

    fn read_delimiter(&mut self) -> Result<String> {
        if self.peek() == Some('\\') {
            self.bump();
            let name = self.read_command_name();
            let value = match name.as_str() {
                "langle" => "⟨",
                "rangle" => "⟩",
                "lbrace" => "{",
                "rbrace" => "}",
                "vert" | "|" => "|",
                "backslash" => "\\",
                "." => "",
                _ => return Err(self.error("unsupported LaTeX delimiter")),
            };
            if self.input[self.offset..].starts_with("{}") {
                self.bump();
                self.bump();
            }
            return Ok(value.to_owned());
        }
        self.bump()
            .map(|value| {
                if value == '.' {
                    String::new()
                } else {
                    value.to_string()
                }
            })
            .ok_or_else(|| self.error("missing LaTeX delimiter"))
    }

    fn read_command_name(&mut self) -> String {
        let start = self.offset;
        if self.peek().is_some_and(char::is_alphabetic) {
            while self.peek().is_some_and(char::is_alphabetic) {
                self.bump();
            }
        } else {
            self.bump();
        }
        self.input[start..self.offset].to_owned()
    }

    fn skip_optional_lossy_argument(&mut self) -> Result<()> {
        self.skip_whitespace();
        if self.peek() == Some('[') {
            self.bump();
            let _ = self.parse_argument(Some(']'))?;
        }
        self.skip_whitespace();
        if self.peek() == Some('{') {
            self.bump();
            let _ = self.parse_argument(Some('}'))?;
        }
        Ok(())
    }

    fn add_node(&mut self) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .ok_or_else(|| self.error("LaTeX node count overflowed"))?;
        if self.nodes > MAX_NODES {
            return Err(self.error("LaTeX input exceeds the node limit"));
        }
        Ok(())
    }

    fn absorb_nested(&mut self, nested: LatexParser<'_>, base: usize) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(nested.nodes)
            .ok_or_else(|| self.error("LaTeX node count overflowed"))?;
        self.tokens = self
            .tokens
            .checked_add(nested.tokens)
            .ok_or_else(|| self.error("LaTeX token count overflowed"))?;
        self.text_bytes = self
            .text_bytes
            .checked_add(nested.text_bytes)
            .ok_or_else(|| self.error("LaTeX text size overflowed"))?;
        if self.nodes > MAX_NODES || self.tokens > MAX_EVENTS || self.text_bytes > MAX_TEXT_BYTES {
            return Err(self.error("nested LaTeX input exceeds a conversion limit"));
        }
        for diagnostic in nested.diagnostics.values {
            self.diagnostics.push(
                format!("byte:{}", base + diagnostic_byte(&diagnostic.path)),
                diagnostic.message,
            )?;
        }
        Ok(())
    }

    fn add_text(&mut self, bytes: usize) -> Result<()> {
        self.text_bytes = self
            .text_bytes
            .checked_add(bytes)
            .ok_or_else(|| self.error("LaTeX text size overflowed"))?;
        if self.text_bytes > MAX_TEXT_BYTES {
            return Err(self.error("LaTeX input exceeds the text limit"));
        }
        Ok(())
    }

    fn bump_token(&mut self) -> Result<()> {
        self.tokens = self
            .tokens
            .checked_add(1)
            .ok_or_else(|| self.error("LaTeX token count overflowed"))?;
        if self.tokens > MAX_EVENTS {
            return Err(self.error("LaTeX input exceeds the token limit"));
        }
        Ok(())
    }

    fn skip_whitespace(&mut self) {
        while self.peek().is_some_and(char::is_whitespace) {
            self.bump();
        }
    }

    fn peek(&self) -> Option<char> {
        self.input[self.offset..].chars().next()
    }

    fn bump(&mut self) -> Option<char> {
        let value = self.peek()?;
        self.offset += value.len_utf8();
        Some(value)
    }

    fn at_end(&self) -> bool {
        self.offset == self.input.len()
    }

    fn error(&self, message: impl Into<String>) -> Error {
        math_error(format!("{} at byte {}", message.into(), self.offset))
    }
}

fn diagnostic_byte(path: &str) -> usize {
    path.strip_prefix("byte:")
        .and_then(|value| value.parse().ok())
        .unwrap_or(0)
}

fn latex_command_at(input: &str, offset: usize) -> Option<(&str, usize)> {
    if input.get(offset..)?.chars().next()? != '\\' {
        return None;
    }
    let start = offset + 1;
    let first = input.get(start..)?.chars().next()?;
    let mut end = start + first.len_utf8();
    if first.is_alphabetic() {
        while let Some(character) = input.get(end..)?.chars().next() {
            if !character.is_alphabetic() {
                break;
            }
            end += character.len_utf8();
        }
    }
    Some((&input[start..end], end))
}

fn raw_group_at(input: &str, offset: usize) -> Option<(&str, usize)> {
    if input.get(offset..)?.chars().next()? != '{' {
        return None;
    }
    let start = offset + 1;
    let relative_end = input.get(start..)?.find('}')?;
    let end = start + relative_end;
    Some((&input[start..end], end + 1))
}

fn delimiter_at(input: &str, offset: usize) -> Option<(String, usize)> {
    let character = input.get(offset..)?.chars().next()?;
    if character != '\\' {
        return Some((
            if character == '.' {
                String::new()
            } else {
                character.to_string()
            },
            offset + character.len_utf8(),
        ));
    }
    let (command, end) = latex_command_at(input, offset)?;
    let value = match command {
        "langle" => "⟨",
        "rangle" => "⟩",
        "lbrace" => "{",
        "rbrace" => "}",
        "vert" | "|" => "|",
        "backslash" => "\\",
        "." => "",
        _ => return None,
    };
    let after = if input[end..].starts_with("{}") {
        end + 2
    } else {
        end
    };
    Some((value.to_owned(), after))
}

fn find_matching_right(input: &str) -> Option<usize> {
    let mut depth = 0_usize;
    let mut offset = 0_usize;
    while offset < input.len() {
        if let Some((command, end)) = latex_command_at(input, offset) {
            match command {
                "left" => {
                    depth += 1;
                    offset = delimiter_at(input, end)?.1;
                }
                "right" => {
                    if depth == 0 {
                        return Some(offset);
                    }
                    depth -= 1;
                    offset = delimiter_at(input, end)?.1;
                }
                _ => offset = end,
            }
        } else {
            offset += input[offset..].chars().next()?.len_utf8();
        }
    }
    None
}

fn split_latex_fenced(input: &str) -> Result<(Vec<(usize, &str)>, String)> {
    let mut brace_depth = 0_usize;
    let mut delimiter_depth = 0_usize;
    let mut environment_depth = 0_usize;
    let mut parts = Vec::new();
    let mut start = 0_usize;
    let mut separator = None::<String>;
    let mut offset = 0_usize;
    while offset < input.len() {
        if let Some((command, end)) = latex_command_at(input, offset) {
            match command {
                "begin" | "end" => {
                    let Some((_, after)) = raw_group_at(input, end) else {
                        return Err(math_error(
                            "LaTeX environment command requires a braced name",
                        ));
                    };
                    if command == "begin" {
                        environment_depth += 1;
                    } else {
                        environment_depth = environment_depth.saturating_sub(1);
                    }
                    offset = after;
                }
                "left" => {
                    delimiter_depth += 1;
                    offset = delimiter_at(input, end)
                        .map(|(_, after)| after)
                        .ok_or_else(|| math_error("missing nested LaTeX left delimiter"))?;
                }
                "right" => {
                    delimiter_depth = delimiter_depth.saturating_sub(1);
                    offset = delimiter_at(input, end)
                        .map(|(_, after)| after)
                        .ok_or_else(|| math_error("missing nested LaTeX right delimiter"))?;
                }
                "middle" if brace_depth == 0 && delimiter_depth == 0 && environment_depth == 0 => {
                    let (value, after) = delimiter_at(input, end)
                        .ok_or_else(|| math_error("missing LaTeX middle delimiter"))?;
                    if separator.as_ref().is_some_and(|current| current != &value) {
                        return Err(math_error(
                            "mixed LaTeX delimiter separators are unsupported",
                        ));
                    }
                    separator.get_or_insert(value);
                    parts.push((start, &input[start..offset]));
                    start = after;
                    offset = after;
                }
                _ => offset = end,
            }
            continue;
        }
        let character = input[offset..]
            .chars()
            .next()
            .ok_or_else(|| math_error("invalid LaTeX delimiter scan offset"))?;
        match character {
            '{' => brace_depth += 1,
            '}' => brace_depth = brace_depth.saturating_sub(1),
            ',' | '|' if brace_depth == 0 && delimiter_depth == 0 && environment_depth == 0 => {
                let value = character.to_string();
                if separator.as_ref().is_some_and(|current| current != &value) {
                    return Err(math_error(
                        "mixed LaTeX delimiter separators are unsupported",
                    ));
                }
                separator.get_or_insert(value);
                parts.push((start, &input[start..offset]));
                start = offset + character.len_utf8();
            }
            _ => {}
        }
        offset += character.len_utf8();
    }
    parts.push((start, &input[start..]));
    Ok((parts, separator.unwrap_or_else(|| "|".to_owned())))
}

fn find_matching_environment(input: &str, outer: &str) -> Result<Option<(usize, usize)>> {
    let mut stack = Vec::<&str>::new();
    let mut offset = 0_usize;
    while offset < input.len() {
        let Some((command, end)) = latex_command_at(input, offset) else {
            offset += input[offset..]
                .chars()
                .next()
                .ok_or_else(|| math_error("invalid LaTeX environment scan offset"))?
                .len_utf8();
            continue;
        };
        if matches!(command, "begin" | "end") {
            let Some((environment, after)) = raw_group_at(input, end) else {
                return Err(math_error(
                    "LaTeX environment command requires a braced name",
                ));
            };
            if command == "begin" {
                stack.push(environment);
            } else if let Some(expected) = stack.pop() {
                if environment != expected {
                    return Err(math_error("mismatched nested LaTeX environment"));
                }
            } else if environment == outer {
                return Ok(Some((offset, after)));
            } else {
                return Err(math_error("mismatched LaTeX environment close"));
            }
            offset = after;
        } else {
            offset = end;
        }
    }
    Ok(None)
}

fn split_latex_matrix(input: &str) -> Result<Vec<Vec<(usize, &str)>>> {
    let mut rows = vec![Vec::new()];
    let mut cell_start = 0_usize;
    let mut brace_depth = 0_usize;
    let mut delimiter_depth = 0_usize;
    let mut environment_depth = 0_usize;
    let mut offset = 0_usize;
    while offset < input.len() {
        if let Some((command, end)) = latex_command_at(input, offset) {
            match command {
                "begin" | "end" => {
                    let (_, after) = raw_group_at(input, end)
                        .ok_or_else(|| math_error("matrix environment name is not braced"))?;
                    if command == "begin" {
                        environment_depth += 1;
                    } else {
                        environment_depth = environment_depth.saturating_sub(1);
                    }
                    offset = after;
                }
                "left" => {
                    delimiter_depth += 1;
                    offset = delimiter_at(input, end)
                        .map(|(_, after)| after)
                        .ok_or_else(|| math_error("matrix left delimiter is missing"))?;
                }
                "right" => {
                    delimiter_depth = delimiter_depth.saturating_sub(1);
                    offset = delimiter_at(input, end)
                        .map(|(_, after)| after)
                        .ok_or_else(|| math_error("matrix right delimiter is missing"))?;
                }
                "\\" if brace_depth == 0 && delimiter_depth == 0 && environment_depth == 0 => {
                    rows.last_mut()
                        .ok_or_else(|| math_error("matrix row state is empty"))?
                        .push((cell_start, &input[cell_start..offset]));
                    rows.push(Vec::new());
                    cell_start = end;
                    offset = end;
                }
                _ => offset = end,
            }
            continue;
        }
        let character = input[offset..]
            .chars()
            .next()
            .ok_or_else(|| math_error("invalid LaTeX matrix scan offset"))?;
        match character {
            '{' => brace_depth += 1,
            '}' => brace_depth = brace_depth.saturating_sub(1),
            '&' if brace_depth == 0 && delimiter_depth == 0 && environment_depth == 0 => {
                rows.last_mut()
                    .ok_or_else(|| math_error("matrix row state is empty"))?
                    .push((cell_start, &input[cell_start..offset]));
                cell_start = offset + 1;
            }
            _ => {}
        }
        offset += character.len_utf8();
    }
    rows.last_mut()
        .ok_or_else(|| math_error("matrix row state is empty"))?
        .push((cell_start, &input[cell_start..]));
    Ok(rows)
}

fn latex_nary_character(command: &str) -> &'static str {
    match command {
        "sum" => "∑",
        "prod" => "∏",
        "coprod" => "∐",
        "int" => "∫",
        "iint" => "∬",
        "iiint" => "∭",
        "oint" => "∮",
        _ => "∫",
    }
}

fn latex_nary_command(character: &str) -> Option<&'static str> {
    match character {
        "∑" => Some("sum"),
        "∏" => Some("prod"),
        "∐" => Some("coprod"),
        "∫" => Some("int"),
        "∬" => Some("iint"),
        "∭" => Some("iiint"),
        "∮" => Some("oint"),
        _ => None,
    }
}

fn latex_accent_character(command: &str) -> &'static str {
    match command {
        "hat" | "widehat" => "̂",
        "bar" | "overline" => "̄",
        "vec" | "overrightarrow" => "⃗",
        "tilde" | "widetilde" => "̃",
        "dot" => "̇",
        "ddot" => "̈",
        _ => "̂",
    }
}

fn latex_accent_command(character: &str) -> Option<&'static str> {
    match character {
        "̂" => Some("hat"),
        "̄" => Some("bar"),
        "⃗" => Some("vec"),
        "̃" => Some("tilde"),
        "̇" => Some("dot"),
        "̈" => Some("ddot"),
        _ => None,
    }
}

fn diagnose_empty_latex_runs(
    argument: &MathArgument,
    path: &str,
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    for (index, expression) in argument.expressions.iter().enumerate() {
        let expression_path = format!("{path}/{}", index + 1);
        match expression {
            MathExpression::Run(run) if run.text.is_empty() => diagnostics.push(
                expression_path,
                "empty OfficeMath run was discarded by LaTeX",
            )?,
            MathExpression::Run(_) => {}
            MathExpression::Fraction(value) => {
                diagnose_empty_latex_runs(&value.numerator, &expression_path, diagnostics)?;
                diagnose_empty_latex_runs(&value.denominator, &expression_path, diagnostics)?;
            }
            MathExpression::Subscript(value) | MathExpression::Superscript(value) => {
                diagnose_empty_latex_runs(&value.base, &expression_path, diagnostics)?;
                diagnose_empty_latex_runs(&value.script, &expression_path, diagnostics)?;
            }
            MathExpression::SubSuperscript(value) => {
                diagnose_empty_latex_runs(&value.base, &expression_path, diagnostics)?;
                diagnose_empty_latex_runs(&value.subscript, &expression_path, diagnostics)?;
                diagnose_empty_latex_runs(&value.superscript, &expression_path, diagnostics)?;
            }
            MathExpression::PreSubSuperscript(value) => {
                diagnose_empty_latex_runs(&value.base, &expression_path, diagnostics)?;
                diagnose_empty_latex_runs(&value.subscript, &expression_path, diagnostics)?;
                diagnose_empty_latex_runs(&value.superscript, &expression_path, diagnostics)?;
            }
            MathExpression::Radical(value) => {
                diagnose_empty_latex_runs(&value.degree, &expression_path, diagnostics)?;
                diagnose_empty_latex_runs(&value.base, &expression_path, diagnostics)?;
            }
            MathExpression::Matrix(value) => {
                for row in &value.rows {
                    for cell in &row.cells {
                        diagnose_empty_latex_runs(cell, &expression_path, diagnostics)?;
                    }
                }
            }
            MathExpression::LowerLimit(value) | MathExpression::UpperLimit(value) => {
                diagnose_empty_latex_runs(&value.base, &expression_path, diagnostics)?;
                diagnose_empty_latex_runs(&value.limit, &expression_path, diagnostics)?;
            }
            MathExpression::Nary(value) => {
                diagnose_empty_latex_runs(&value.base, &expression_path, diagnostics)?;
                diagnose_empty_latex_runs(&value.subscript, &expression_path, diagnostics)?;
                diagnose_empty_latex_runs(&value.superscript, &expression_path, diagnostics)?;
            }
            MathExpression::Delimiter(value) => {
                for child in &value.arguments {
                    diagnose_empty_latex_runs(child, &expression_path, diagnostics)?;
                }
            }
            MathExpression::Accent(value) => {
                diagnose_empty_latex_runs(&value.base, &expression_path, diagnostics)?;
            }
        }
    }
    Ok(())
}

fn write_latex_argument(
    argument: &MathArgument,
    path: &str,
    output: &mut String,
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    if argument_has_direct_unsupported_content(argument) {
        diagnostics.push(path, "unmodelled OfficeMath argument content was discarded")?;
    }
    for (index, expression) in argument.expressions.iter().enumerate() {
        write_latex_expression(
            expression,
            &format!("{path}/{}", index + 1),
            output,
            diagnostics,
        )?;
    }
    Ok(())
}

fn write_latex_group(
    argument: &MathArgument,
    path: &str,
    output: &mut String,
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    output.push('{');
    write_latex_argument(argument, path, output, diagnostics)?;
    output.push('}');
    Ok(())
}

fn write_latex_expression(
    expression: &MathExpression,
    path: &str,
    output: &mut String,
    diagnostics: &mut Diagnostics,
) -> Result<()> {
    if expression_has_direct_unsupported_content(expression) {
        diagnostics.push(path, "unmodelled OfficeMath content was discarded")?;
    }
    match expression {
        MathExpression::Run(run) => {
            if run.properties != Default::default() {
                diagnostics.push(path, "OfficeMath run properties were discarded")?;
            }
            if run.text.chars().any(char::is_whitespace) {
                diagnostics.push(path, "OfficeMath run whitespace was discarded by LaTeX")?;
            }
            output.push_str(&latex_escape(&run.text));
        }
        MathExpression::Fraction(value) => {
            if value.fraction_type != FractionType::Bar {
                diagnostics.push(path, "non-bar OfficeMath fraction type was discarded")?;
            }
            output.push_str("\\frac");
            write_latex_group(&value.numerator, path, output, diagnostics)?;
            write_latex_group(&value.denominator, path, output, diagnostics)?;
        }
        MathExpression::Subscript(value) | MathExpression::Superscript(value) => {
            if value.alignment.is_some() {
                diagnostics.push(path, "OfficeMath script alignment was discarded")?;
            }
            write_latex_group(&value.base, path, output, diagnostics)?;
            output.push(if matches!(expression, MathExpression::Subscript(_)) {
                '_'
            } else {
                '^'
            });
            write_latex_group(&value.script, path, output, diagnostics)?;
        }
        MathExpression::SubSuperscript(value) => {
            if value.alignment.is_some() {
                diagnostics.push(path, "OfficeMath script alignment was discarded")?;
            }
            write_latex_group(&value.base, path, output, diagnostics)?;
            output.push('_');
            write_latex_group(&value.subscript, path, output, diagnostics)?;
            output.push('^');
            write_latex_group(&value.superscript, path, output, diagnostics)?;
        }
        MathExpression::PreSubSuperscript(value) => {
            if value.alignment.is_some() {
                diagnostics.push(path, "OfficeMath script alignment was discarded")?;
            }
            output.push_str("{}");
            output.push('_');
            write_latex_group(&value.subscript, path, output, diagnostics)?;
            output.push('^');
            write_latex_group(&value.superscript, path, output, diagnostics)?;
            write_latex_group(&value.base, path, output, diagnostics)?;
        }
        MathExpression::Radical(value) => {
            if value.hide_degree != value.degree.expressions.is_empty() {
                diagnostics.push(path, "OfficeMath radical degree visibility was discarded")?;
            }
            output.push_str("\\sqrt");
            if !value.degree.expressions.is_empty() {
                output.push('[');
                write_latex_argument(&value.degree, path, output, diagnostics)?;
                output.push(']');
            }
            write_latex_group(&value.base, path, output, diagnostics)?;
        }
        MathExpression::Matrix(value) => {
            if value.properties != Default::default() {
                diagnostics.push(path, "OfficeMath matrix properties were discarded")?;
            }
            output.push_str("\\begin{matrix}");
            for (row_index, row) in value.rows.iter().enumerate() {
                if row_index != 0 {
                    output.push_str("\\\\");
                }
                for (cell_index, cell) in row.cells.iter().enumerate() {
                    if cell_index != 0 {
                        output.push('&');
                    }
                    write_latex_argument(cell, path, output, diagnostics)?;
                }
            }
            output.push_str("\\end{matrix}");
        }
        MathExpression::LowerLimit(value) | MathExpression::UpperLimit(value) => {
            output.push_str(if matches!(expression, MathExpression::LowerLimit(_)) {
                "\\underset"
            } else {
                "\\overset"
            });
            write_latex_group(&value.limit, path, output, diagnostics)?;
            write_latex_group(&value.base, path, output, diagnostics)?;
        }
        MathExpression::Nary(value) => {
            let Some(command) = latex_nary_command(&value.character) else {
                diagnostics.push(path, "unsupported OfficeMath n-ary operator was discarded")?;
                return Ok(());
            };
            if value.grow.is_some() || value.limit_location.is_some() {
                diagnostics.push(path, "OfficeMath n-ary layout properties were discarded")?;
            }
            if value.hide_subscript && !value.subscript.expressions.is_empty() {
                diagnostics.push(path, "hidden OfficeMath n-ary subscript was discarded")?;
            } else if !value.hide_subscript && value.subscript.expressions.is_empty() {
                diagnostics.push(path, "OfficeMath n-ary subscript visibility was discarded")?;
            }
            if value.hide_superscript && !value.superscript.expressions.is_empty() {
                diagnostics.push(path, "hidden OfficeMath n-ary superscript was discarded")?;
            } else if !value.hide_superscript && value.superscript.expressions.is_empty() {
                diagnostics.push(
                    path,
                    "OfficeMath n-ary superscript visibility was discarded",
                )?;
            }
            write!(output, "\\{command}").expect("String writes are infallible");
            if !value.hide_subscript && !value.subscript.expressions.is_empty() {
                output.push('_');
                write_latex_group(&value.subscript, path, output, diagnostics)?;
            }
            if !value.hide_superscript && !value.superscript.expressions.is_empty() {
                output.push('^');
                write_latex_group(&value.superscript, path, output, diagnostics)?;
            }
            write_latex_group(&value.base, path, output, diagnostics)?;
        }
        MathExpression::Delimiter(value) => {
            if value.grow.is_some() {
                diagnostics.push(path, "OfficeMath delimiter growth setting was discarded")?;
            }
            if value.begin_character.chars().count() > 1
                || value.end_character.chars().count() > 1
                || value.separator_character.chars().count() > 1
            {
                diagnostics.push(path, "non-scalar OfficeMath delimiter was discarded")?;
                return Ok(());
            }
            let (Some(begin), Some(separator), Some(end)) = (
                latex_delimiter(&value.begin_character),
                latex_delimiter(&value.separator_character),
                latex_delimiter(&value.end_character),
            ) else {
                diagnostics.push(
                    path,
                    "unsupported OfficeMath LaTeX delimiter character was discarded",
                )?;
                return Ok(());
            };
            output.push_str("\\left");
            output.push_str(&begin);
            for (index, argument) in value.arguments.iter().enumerate() {
                if index != 0 {
                    output.push_str("\\middle");
                    output.push_str(&separator);
                }
                write_latex_group(argument, path, output, diagnostics)?;
            }
            output.push_str("\\right");
            output.push_str(&end);
        }
        MathExpression::Accent(value) => {
            let Some(command) = latex_accent_command(&value.character) else {
                diagnostics.push(path, "unsupported OfficeMath accent was discarded")?;
                return Ok(());
            };
            write!(output, "\\{command}").expect("String writes are infallible");
            write_latex_group(&value.base, path, output, diagnostics)?;
        }
    }
    Ok(())
}

fn latex_escape(value: &str) -> String {
    let mut escaped = String::new();
    for character in value.chars() {
        match character {
            '\\' => escaped.push_str("\\backslash{}"),
            '[' => escaped.push_str("\\lbrack{}"),
            ']' => escaped.push_str("\\rbrack{}"),
            '{' | '}' | '_' | '^' | '#' | '$' | '%' | '&' => {
                escaped.push('\\');
                escaped.push(character);
            }
            _ => escaped.push(character),
        }
    }
    escaped
}

fn latex_delimiter(value: &str) -> Option<String> {
    Some(match value {
        "" => ".".to_owned(),
        "{" => "\\lbrace{}".to_owned(),
        "}" => "\\rbrace{}".to_owned(),
        "⟨" => "\\langle{}".to_owned(),
        "⟩" => "\\rangle{}".to_owned(),
        "|" => "\\vert{}".to_owned(),
        "\\" => "\\backslash{}".to_owned(),
        value
            if value
                .chars()
                .any(|character| character.is_alphabetic() || character.is_whitespace()) =>
        {
            return None;
        }
        value => value.to_owned(),
    })
}

#[cfg(test)]
mod tests {
    use std::io::Write;
    use std::process::{Command, Stdio};

    use rdocx_oxml::math::{CT_OMath, LimitLocation, MathStyle};

    use super::*;

    fn supported_tree() -> MathArgument {
        MathArgument::new(vec![
            MathExpression::Fraction(MathFraction::new(
                MathArgument::text("a"),
                MathArgument::text("2"),
            )),
            MathExpression::SubSuperscript(MathSubSuperscript::new(
                MathArgument::text("x"),
                MathArgument::text("i"),
                MathArgument::text("n"),
            )),
            MathExpression::Radical(MathRadical::with_degree(
                MathArgument::text("3"),
                MathArgument::text("y"),
            )),
            MathExpression::Delimiter(MathDelimiter::new("(", ")", vec![MathArgument::text("z")])),
            MathExpression::Accent(MathAccent::new("̂", MathArgument::text("q"))),
            MathExpression::Matrix(MathMatrix::new(vec![
                MathMatrixRow::new(vec![MathArgument::text("1"), MathArgument::text("2")]),
                MathMatrixRow::new(vec![MathArgument::text("3"), MathArgument::text("4")]),
            ])),
        ])
    }

    fn complete_round_trip_tree() -> MathArgument {
        let mut nary = MathNary::new(
            "∑",
            MathArgument::new(vec![
                MathRun::new("x").into(),
                MathExpression::Fraction(MathFraction::new(
                    MathArgument::text("1"),
                    MathArgument::text("2"),
                )),
            ]),
        );
        nary.subscript = MathArgument::text("i");
        nary.superscript = MathArgument::text("n");
        nary.hide_subscript = false;
        nary.hide_superscript = false;

        let mut delimiter = MathDelimiter::new(
            "+",
            "-",
            vec![MathArgument::text("a"), MathArgument::text("b")],
        );
        delimiter.separator_character = ";".to_owned();

        MathArgument::new(vec![
            MathRun::new("r").into(),
            MathExpression::Fraction(MathFraction::new(
                MathArgument::text("a"),
                MathArgument::text("2"),
            )),
            MathExpression::Subscript(MathScript::new(
                MathArgument::text("x"),
                MathArgument::text("i"),
            )),
            MathExpression::Superscript(MathScript::new(
                MathArgument::text("y"),
                MathArgument::text("2"),
            )),
            MathExpression::SubSuperscript(MathSubSuperscript::new(
                MathArgument::text("z"),
                MathArgument::text("j"),
                MathArgument::text("m"),
            )),
            MathExpression::PreSubSuperscript(MathPreSubSuperscript::new(
                MathArgument::text("p"),
                MathArgument::text("k"),
                MathArgument::text("q"),
            )),
            MathExpression::Radical(MathRadical::new(MathArgument::text("s"))),
            MathExpression::Radical(MathRadical::with_degree(
                MathArgument::text("3"),
                MathArgument::text("t"),
            )),
            MathExpression::LowerLimit(MathLimit::new(
                MathArgument::text("lim"),
                MathArgument::text("0"),
            )),
            MathExpression::UpperLimit(MathLimit::new(
                MathArgument::text("max"),
                MathArgument::text("n"),
            )),
            MathExpression::Nary(nary),
            MathExpression::Nary(MathNary::new(
                "∫",
                MathArgument::new(vec![
                    MathRun::new("u").into(),
                    MathExpression::Radical(MathRadical::new(MathArgument::text("v"))),
                ]),
            )),
            MathExpression::Delimiter(delimiter),
            MathExpression::Accent(MathAccent::new("̂", MathArgument::text("q"))),
            MathExpression::Matrix(MathMatrix::new(vec![
                MathMatrixRow::new(vec![MathArgument::text("1"), MathArgument::text("2")]),
                MathMatrixRow::new(vec![MathArgument::text("3"), MathArgument::text("4")]),
            ])),
        ])
    }

    fn pandoc_convert(pandoc: &str, arguments: &[&str], input: &str) -> String {
        let mut child = Command::new(pandoc)
            .args(arguments)
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .spawn()
            .expect("start Pandoc texmath conversion");
        child
            .stdin
            .take()
            .expect("Pandoc stdin")
            .write_all(input.as_bytes())
            .expect("write source-built equation");
        let output = child.wait_with_output().expect("wait for Pandoc");
        assert!(output.status.success());
        String::from_utf8(output.stdout).expect("Pandoc UTF-8")
    }

    fn differential_accepts(
        expected: &MathConversionResult<MathArgument>,
        actual: &MathConversionResult<MathArgument>,
    ) -> bool {
        expected == actual
    }

    #[test]
    fn mathml_supported_subset_maps_to_one_normalized_expression_tree() {
        let input = concat!(
            r#"<m:math xmlns:m="http://www.w3.org/1998/Math/MathML">"#,
            "<m:mfrac><m:mi>a</m:mi><m:mn>2</m:mn></m:mfrac>",
            "<m:msubsup><m:mi>x</m:mi><m:mi>i</m:mi><m:mi>n</m:mi></m:msubsup>",
            "<m:mroot><m:mi>y</m:mi><m:mn>3</m:mn></m:mroot>",
            "<m:mfenced><m:mi>z</m:mi></m:mfenced>",
            r#"<m:mover accent="true"><m:mi>q</m:mi><m:mo>̂</m:mo></m:mover>"#,
            "<m:mtable><m:mtr><m:mtd><m:mn>1</m:mn></m:mtd><m:mtd><m:mn>2</m:mn></m:mtd></m:mtr>",
            "<m:mtr><m:mtd><m:mn>3</m:mn></m:mtd><m:mtd><m:mn>4</m:mn></m:mtd></m:mtr></m:mtable>",
            "</m:math>"
        );
        let converted = equation_from_mathml(input).expect("supported MathML must convert");
        assert_eq!(converted.value, supported_tree());
        assert!(converted.diagnostics.is_empty());

        let coverage = equation_from_mathml(&format!(
            concat!(
                r#"<math xmlns="{}"><mrow><mi>x</mi><mn>2</mn><mo>+</mo><mtext>word</mtext></mrow>"#,
                "<msub><mi>x</mi><mi>i</mi></msub>",
                "<msup><mi>x</mi><mn>2</mn></msup>",
                "<mmultiscripts><mi>x</mi><mprescripts/><mi>a</mi><mi>b</mi></mmultiscripts>",
                "<msqrt><mi>z</mi></msqrt>",
                "<munder><mi>lim</mi><mn>0</mn></munder>",
                "<mover><mi>max</mi><mi>n</mi></mover>",
                "<munderover><mo>∑</mo><mi>i</mi><mi>n</mi></munderover><mi>x</mi>",
                r#"<mfenced open="[" close="]" separators=";"><mi>a</mi><mi>b</mi></mfenced></math>"#
            ),
            MATHML_NS
        ))
        .expect("complete supported MathML subset");
        assert!(coverage.diagnostics.is_empty());
        assert_eq!(coverage.value.expressions.len(), 9);
        assert!(matches!(
            &coverage.value.expressions[0],
            MathExpression::Run(run) if run.text == "x2+word"
        ));
        assert!(matches!(
            coverage.value.expressions[1],
            MathExpression::Subscript(_)
        ));
        assert!(matches!(
            coverage.value.expressions[2],
            MathExpression::Superscript(_)
        ));
        assert!(matches!(
            coverage.value.expressions[3],
            MathExpression::PreSubSuperscript(_)
        ));
        assert!(matches!(
            coverage.value.expressions[4],
            MathExpression::Radical(_)
        ));
        assert!(matches!(
            coverage.value.expressions[5],
            MathExpression::LowerLimit(_)
        ));
        assert!(matches!(
            coverage.value.expressions[6],
            MathExpression::UpperLimit(_)
        ));
        assert!(matches!(
            coverage.value.expressions[7],
            MathExpression::Nary(_)
        ));
        assert!(matches!(
            &coverage.value.expressions[8],
            MathExpression::Delimiter(value) if value.separator_character == ";"
        ));

        let foreign = equation_from_mathml(&format!(
            concat!(
                r#"<math xmlns="{}" xmlns:f="urn:foreign">"#,
                r#"<mmultiscripts><mi>x</mi><f:mprescripts/><mi>a</mi><mi>b</mi></mmultiscripts>"#,
                r#"<mrow><f:mo fence="true">(</f:mo><mi>y</mi><mo fence="true">)</mo></mrow>"#,
                "</math>"
            ),
            MATHML_NS
        ))
        .expect("foreign lookalikes are safe losses");
        assert!(foreign.diagnostics.iter().any(|value| {
            value.path == "/math[1]/mmultiscripts[1]"
                && value.message == "extra MathML multiscript pairs were discarded"
        }));
        assert!(
            foreign
                .diagnostics
                .iter()
                .any(|value| value.message == "foreign-namespace element was discarded")
        );
        assert!(
            !foreign
                .value
                .expressions
                .iter()
                .any(|value| matches!(value, MathExpression::PreSubSuperscript(_)))
        );

        let foreign_nary = equation_from_mathml(&format!(
            concat!(
                r#"<math xmlns="{}" xmlns:f="urn:foreign">"#,
                "<munder><f:mo>∑</f:mo><mi>i</mi></munder><mi>x</mi>",
                "<mover><mi>∑</mi><mi>n</mi></mover>",
                "</math>"
            ),
            MATHML_NS
        ))
        .expect("n-ary recognition requires an expanded-name mo");
        assert!(foreign_nary.diagnostics.iter().any(|value| {
            value.path.ends_with("munder[1]/*[1]")
                && value.message == "foreign-namespace element was discarded"
        }));
        assert!(
            !foreign_nary
                .value
                .expressions
                .iter()
                .any(|value| matches!(value, MathExpression::Nary(_)))
        );

        let attribute_losses = equation_from_mathml(&format!(
            concat!(
                r#"<math xmlns="{}"><mo largeop="true">x</mo>"#,
                r#"<mover accent="invalid"><mi>x</mi><mi>n</mi></mover>"#,
                r#"<mtable><mtr row="lost">row text<mtd cell="lost"><mi>x</mi></mtd></mtr></mtable>"#,
                r#"<mrow><mo fence="true" stretchy="invalid">(</mo><mi>z</mi><mo fence="true" form="invalid">)</mo></mrow>"#,
                "</math>"
            ),
            MATHML_NS
        ))
        .expect("unsupported safe attributes are diagnosed");
        assert!(attribute_losses.diagnostics.iter().any(|value| {
            value.path == "/math[1]/mo[1]/@largeop"
                && value.message == "unsupported MathML attribute was discarded"
        }));
        assert!(attribute_losses.diagnostics.iter().any(|value| {
            value.path == "/math[1]/mover[1]/@accent"
                && value.message == "unsupported MathML attribute value was discarded"
        }));
        assert!(attribute_losses.diagnostics.iter().any(|value| {
            value.path.ends_with("mrow[1]/mo[1]/@stretchy")
                && value.message == "unsupported MathML attribute was discarded"
        }));
        assert!(attribute_losses.diagnostics.iter().any(|value| {
            value.path.ends_with("mrow[1]/mo[2]/@form")
                && value.message == "unsupported MathML attribute was discarded"
        }));
        for path in [
            "/math[1]/mrow[1]/mo[1]/@stretchy",
            "/math[1]/mrow[1]/mo[2]/@form",
        ] {
            assert_eq!(
                attribute_losses
                    .diagnostics
                    .iter()
                    .filter(|value| value.path == path)
                    .count(),
                1,
                "one discarded attribute produces one diagnostic"
            );
        }
        assert!(
            attribute_losses
                .diagnostics
                .iter()
                .any(|value| value.path.ends_with("mtr[1]/@row"))
        );
        assert!(
            attribute_losses
                .diagnostics
                .iter()
                .any(|value| value.path.ends_with("mtr[1]/mtd[1]/@cell"))
        );
        assert!(attribute_losses.diagnostics.iter().any(|value| {
            value.path.ends_with("mtr[1]")
                && value.message == "text outside a MathML token was discarded"
        }));

        let special_token_losses = equation_from_mathml(&format!(
            concat!(
                r#"<math xmlns="{}" xmlns:f="urn:foreign">"#,
                r#"<msub><mo data-loss="yes">∑<mtext>nested</mtext></mo><mi>i</mi></msub><mi>x</mi>"#,
                r#"<mover accent="true"><mi>y</mi><f:mo>̂</f:mo></mover></math>"#
            ),
            MATHML_NS
        ))
        .expect("special MathML tokens diagnose discarded content");
        assert!(special_token_losses.diagnostics.iter().any(|value| {
            value.path.ends_with("msub[1]/*[1]/@data-loss")
                && value.message == "unsupported MathML attribute was discarded"
        }));
        assert!(special_token_losses.diagnostics.iter().any(|value| {
            value.path.ends_with("msub[1]/*[1]/*[1]")
                && value.message == "nested MathML token content was discarded"
        }));
        assert!(
            special_token_losses
                .diagnostics
                .iter()
                .any(|value| value.message == "foreign-namespace element was discarded")
        );
        assert!(matches!(
            special_token_losses.value.expressions.first(),
            Some(MathExpression::Nary(_))
        ));
        assert!(
            !special_token_losses
                .value
                .expressions
                .iter()
                .any(|value| matches!(value, MathExpression::Accent(_)))
        );

        for contradictory in [
            r#"<mrow><mo fence="false" stretchy="true" form="prefix">(</mo><mi>x</mi><mo fence="true">)</mo></mrow>"#,
            r#"<mrow><mo fence="true" stretchy="false" form="prefix">(</mo><mi>x</mi><mo fence="true">)</mo></mrow>"#,
            r#"<mrow><mo fence="true" form="postfix">(</mo><mi>x</mi><mo fence="true">)</mo></mrow>"#,
        ] {
            let converted = equation_from_mathml(&format!(
                r#"<math xmlns="{MATHML_NS}">{contradictory}</math>"#
            ))
            .expect("contradictory fence attributes are safe losses");
            assert!(
                !converted
                    .value
                    .expressions
                    .iter()
                    .any(|value| matches!(value, MathExpression::Delimiter(_)))
            );
            assert!(!converted.diagnostics.is_empty());
        }

        let mixed_separators = equation_from_mathml(&format!(
            concat!(
                r#"<math xmlns="{}"><mrow><mo fence="true">(</mo><mi>a</mi>"#,
                r#"<mo separator="true">,</mo><mi>b</mi><mo separator="true">|</mo>"#,
                r#"<mi>c</mi><mo fence="true">)</mo></mrow></math>"#
            ),
            MATHML_NS
        ))
        .expect("mixed separators remain a diagnosed delimiter");
        assert!(matches!(
            mixed_separators.value.expressions.as_slice(),
            [MathExpression::Delimiter(value)]
                if value.separator_character == "," && value.arguments.len() == 3
        ));
        assert!(mixed_separators.diagnostics.iter().any(|value| {
            value.message
                == "mixed MathML delimiter separators were normalized to the first character"
        }));

        let fenced_text = equation_from_mathml(&format!(
            r#"<math xmlns="{MATHML_NS}"><mfenced>x<mi>y</mi></mfenced></math>"#
        ))
        .expect("direct mfenced text is a safe loss");
        assert!(fenced_text.diagnostics.iter().any(|value| {
            value.path == "/math[1]/mfenced[1]"
                && value.message == "text outside a MathML token was discarded"
        }));
    }

    #[test]
    fn latex_supported_subset_maps_to_one_normalized_expression_tree() {
        let input = r"\frac{a}{2}{x}_{i}^{n}\sqrt[3]{y}\left(z\right)\hat{q}\begin{matrix}1&2\\3&4\end{matrix}";
        let converted = equation_from_latex(input).expect("supported LaTeX must convert");
        assert_eq!(converted.value, supported_tree());
        assert!(converted.diagnostics.is_empty());

        let coverage = equation_from_latex(
            r"x_i x^2 {}_{a}^{b}{x} \sum_{i}^{n}{x} \bar{x}\vec{x}\tilde{x}\dot{x}\ddot{x} \begin{pmatrix}1&2\\3&4\end{pmatrix} \begin{bmatrix}1\end{bmatrix}",
        )
        .expect("complete supported LaTeX subset");
        assert!(coverage.diagnostics.is_empty());
        assert!(matches!(
            coverage.value.expressions[0],
            MathExpression::Subscript(_)
        ));
        assert!(matches!(
            coverage.value.expressions[1],
            MathExpression::Superscript(_)
        ));
        assert!(matches!(
            coverage.value.expressions[2],
            MathExpression::PreSubSuperscript(_)
        ));
        assert!(matches!(
            coverage.value.expressions[3],
            MathExpression::Nary(_)
        ));
        assert!(
            coverage
                .value
                .expressions
                .iter()
                .filter(|value| matches!(value, MathExpression::Accent(_)))
                .count()
                >= 5
        );
        assert!(
            coverage
                .value
                .expressions
                .iter()
                .filter(|value| matches!(value, MathExpression::Delimiter(_)))
                .count()
                >= 2
        );

        let nested = equation_from_latex(r"\left(\left[a,b\right],c\right)")
            .expect("nested delimiters use command-token scope");
        let [MathExpression::Delimiter(outer)] = nested.value.expressions.as_slice() else {
            panic!("outer delimiter");
        };
        assert_eq!(outer.arguments.len(), 2);
        assert!(matches!(
            outer.arguments[0].expressions.as_slice(),
            [MathExpression::Delimiter(inner)] if inner.arguments.len() == 2
        ));

        let custom_separator =
            equation_from_latex(r"\left+a\middle;b\right-").expect("canonical middle separator");
        assert!(matches!(
            custom_separator.value.expressions.as_slice(),
            [MathExpression::Delimiter(value)]
                if value.begin_character == "+"
                    && value.end_character == "-"
                    && value.separator_character == ";"
                    && value.arguments.len() == 2
        ));

        let escaped_and_nested_matrix = equation_from_latex(
            r"\begin{matrix}a\&b&\begin{matrix}1&2\end{matrix}\\c&d\end{matrix}",
        )
        .expect("matrix scanning follows grammar scope");
        let [MathExpression::Matrix(matrix)] =
            escaped_and_nested_matrix.value.expressions.as_slice()
        else {
            panic!("outer matrix");
        };
        assert_eq!(matrix.rows.len(), 2);
        assert_eq!(matrix.rows[0].cells.len(), 2);
        assert_eq!(matrix.rows[0].cells[0], MathArgument::text("a&b"));
        assert!(matches!(
            matrix.rows[0].cells[1].expressions.as_slice(),
            [MathExpression::Matrix(_)]
        ));

        let matrix_inside_fence =
            equation_from_latex(r"\left(\begin{matrix}a,b&c\end{matrix}\middle;d\right)")
                .expect("fence separators ignore nested environment content");
        let [MathExpression::Delimiter(delimiter)] =
            matrix_inside_fence.value.expressions.as_slice()
        else {
            panic!("delimiter around matrix");
        };
        assert_eq!(delimiter.arguments.len(), 2);
        assert_eq!(delimiter.separator_character, ";");
        assert!(matches!(
            delimiter.arguments[0].expressions.as_slice(),
            [MathExpression::Matrix(_)]
        ));

        for malformed in [
            r"\begin{matrix}a\right)b\end{matrix}",
            r"\left(a\end{matrix}b\right)",
        ] {
            assert!(
                equation_from_latex(malformed).is_err(),
                "nested parser must reject an unconsumed suffix: {malformed}"
            );
        }

        let scripted_nary_operand =
            equation_from_latex(r"\sum_i^n x_j").expect("scripted n-ary operand");
        let [MathExpression::Nary(nary)] = scripted_nary_operand.value.expressions.as_slice()
        else {
            panic!("n-ary expression");
        };
        assert!(matches!(
            nary.base.expressions.as_slice(),
            [MathExpression::Subscript(value)]
                if value.base == MathArgument::text("x")
                    && value.script == MathArgument::text("j")
        ));

        let prescripted_nary_operand =
            equation_from_latex(r"\sum_i^n {}_j^k x").expect("pre-scripted n-ary operand");
        let [MathExpression::Nary(nary)] = prescripted_nary_operand.value.expressions.as_slice()
        else {
            panic!("n-ary expression with a pre-scripted base");
        };
        assert!(matches!(
            nary.base.expressions.as_slice(),
            [MathExpression::PreSubSuperscript(value)]
                if value.base == MathArgument::text("x")
                    && value.subscript == MathArgument::text("j")
                    && value.superscript == MathArgument::text("k")
        ));

        let nested_nary =
            equation_from_latex(r"\sum_i^n \prod\limits_j^m x").expect("nested n-ary operand");
        let [MathExpression::Nary(outer)] = nested_nary.value.expressions.as_slice() else {
            panic!("outer n-ary expression");
        };
        let [MathExpression::Nary(inner)] = outer.base.expressions.as_slice() else {
            panic!("inner n-ary operand");
        };
        assert_eq!(inner.base, MathArgument::text("x"));
        assert!(
            nested_nary
                .diagnostics
                .iter()
                .any(|value| { value.message == "LaTeX \\limits n-ary placement was discarded" })
        );
    }

    #[test]
    fn unsupported_conversion_constructs_report_stable_paths_without_semantic_substitution() {
        let mathml = equation_from_mathml(
            r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><semantics><mi>x</mi><annotation>guess</annotation></semantics><maction><mi>y</mi></maction></math>"#,
        )
        .expect("safe unsupported MathML must be diagnosed");
        assert_eq!(mathml.value, MathArgument::text("x"));
        assert_eq!(
            mathml.diagnostics,
            vec![
                MathConversionDiagnostic {
                    path: "/math[1]/semantics[1]".to_owned(),
                    message: "MathML semantics metadata was discarded".to_owned(),
                },
                MathConversionDiagnostic {
                    path: "/math[1]/maction[1]".to_owned(),
                    message: "unsupported MathML element maction was discarded".to_owned(),
                },
            ]
        );
        let latex = equation_from_latex(r"x\unknown{y}z").expect("lossy LaTeX is safe");
        assert_eq!(latex.value, MathArgument::text("xz"));
        assert_eq!(latex.diagnostics[0].path, "byte:1");

        let metadata_first = equation_from_mathml(&format!(
            r#"<math xmlns="{MATHML_NS}"><semantics><annotation>meta</annotation><mi>kept</mi></semantics></math>"#
        ))
        .expect("semantics retains its first supported descendant");
        assert_eq!(metadata_first.value, MathArgument::text("kept"));
        assert_eq!(metadata_first.diagnostics.len(), 1);

        let recovered_environment = equation_from_latex(r"a\begin{array}x\end{array}b")
            .expect("unsupported environments are consumed as one safe loss");
        assert_eq!(recovered_environment.value, MathArgument::text("ab"));
        assert_eq!(recovered_environment.diagnostics.len(), 1);
        assert_eq!(recovered_environment.diagnostics[0].path, "byte:1");

        for (source, offset, command) in [
            (r"\sum\limits_{i}^{n}{x}", 4, "limits"),
            (r"\sum\nolimits_{i}^{n}{x}", 4, "nolimits"),
        ] {
            let converted = equation_from_latex(source).expect("lossy n-ary placement");
            assert_eq!(converted.diagnostics.len(), 1);
            assert_eq!(converted.diagnostics[0].path, format!("byte:{offset}"));
            assert_eq!(
                converted.diagnostics[0].message,
                format!("LaTeX \\{command} n-ary placement was discarded")
            );
        }

        let prefixed_command = equation_from_latex(r"a\leftover{x}b")
            .expect("command prefixes do not alter delimiter scope");
        assert_eq!(prefixed_command.value, MathArgument::text("ab"));
        assert_eq!(prefixed_command.diagnostics.len(), 1);

        let preserved = CT_OMath::from_xml(
            format!(
                concat!(
                    r#"<m:oMath xmlns:m="{}"><m:f><m:num>"#,
                    "<m:unsupported/><m:r><m:t>x</m:t></m:r>",
                    "</m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>"
                ),
                "http://schemas.openxmlformats.org/officeDocument/2006/math"
            )
            .as_bytes(),
        )
        .expect("OfficeMath preservation fixture");
        let MathExpression::Fraction(fraction) = &preserved.expressions[0] else {
            panic!("preserved fraction");
        };
        for converted in [
            equation_to_mathml(&fraction.numerator),
            equation_to_latex(&fraction.numerator),
        ] {
            assert!(converted.diagnostics.iter().any(|value| {
                value.message == "unmodelled OfficeMath argument content was discarded"
            }));
        }
        for converted in [
            equation_to_mathml(&MathArgument::new(preserved.expressions.clone())),
            equation_to_latex(&MathArgument::new(preserved.expressions.clone())),
        ] {
            assert_eq!(
                converted
                    .diagnostics
                    .iter()
                    .filter(|value| value.message.contains("unmodelled OfficeMath"))
                    .count(),
                1,
                "one preserved child produces one preservation diagnostic"
            );
        }

        let preserved_matrix = CT_OMath::from_xml(
            format!(
                concat!(
                    r#"<m:oMath xmlns:m="{}" xmlns:x="urn:producer"><m:m><m:mr><x:unsupported/>"#,
                    "<m:e><m:r><m:t>x</m:t></m:r></m:e>",
                    "</m:mr></m:m></m:oMath>"
                ),
                "http://schemas.openxmlformats.org/officeDocument/2006/math"
            )
            .as_bytes(),
        )
        .expect("OfficeMath matrix-row preservation fixture");
        for converted in [
            equation_to_mathml(&MathArgument::new(preserved_matrix.expressions.clone())),
            equation_to_latex(&MathArgument::new(preserved_matrix.expressions)),
        ] {
            assert!(
                converted.diagnostics.iter().any(|value| {
                    value.message == "unmodelled OfficeMath content was discarded"
                }),
                "matrix-row loss diagnostics: {:?}",
                converted.diagnostics
            );
        }

        let malformed_markers = equation_from_mathml(&format!(
            concat!(
                r#"<math xmlns="{}"><mmultiscripts><mi>x</mi>"#,
                r#"<mprescripts marker="lost"><mi>nested</mi></mprescripts>"#,
                r#"<none missing="lost">text<mi>nested</mi></none><mi>n</mi>"#,
                "</mmultiscripts></math>"
            ),
            MATHML_NS
        ))
        .expect("malformed structural markers are safe losses");
        for suffix in ["/*[2]/@marker", "/*[3]/@missing"] {
            assert!(
                malformed_markers
                    .diagnostics
                    .iter()
                    .any(|value| value.path.ends_with(suffix))
            );
        }
        assert!(malformed_markers.diagnostics.iter().any(|value| {
            value.path.ends_with("/*[2]")
                && value.message == "content inside a MathML structural marker was discarded"
        }));
        assert!(malformed_markers.diagnostics.iter().any(|value| {
            value.path.ends_with("/*[3]")
                && value.message == "text outside a MathML token was discarded"
        }));

        let unsupported_nary = MathArgument::new(vec![MathExpression::Nary(MathNary::new(
            "⊕",
            MathArgument::text("x"),
        ))]);
        for converted in [
            equation_to_mathml(&unsupported_nary),
            equation_to_latex(&unsupported_nary),
        ] {
            assert!(!converted.value.contains('⊕'));
            assert_eq!(
                converted.diagnostics[0].message,
                "unsupported OfficeMath n-ary operator was discarded"
            );
        }

        let alphabetic_fence = MathArgument::new(vec![MathExpression::Delimiter(
            MathDelimiter::new("a", ")", vec![MathArgument::text("x")]),
        )]);
        let alphabetic_fence = equation_to_latex(&alphabetic_fence);
        assert!(alphabetic_fence.value.is_empty());
        assert_eq!(
            alphabetic_fence.diagnostics[0].message,
            "unsupported OfficeMath LaTeX delimiter character was discarded"
        );
        let backslash_fence = MathArgument::new(vec![MathExpression::Delimiter(
            MathDelimiter::new("\\", ")", vec![MathArgument::text("x")]),
        )]);
        let backslash_latex = equation_to_latex(&backslash_fence);
        assert!(backslash_latex.diagnostics.is_empty());
        assert_eq!(
            equation_from_latex(&backslash_latex.value)
                .expect("escaped backslash fence")
                .value,
            backslash_fence
        );

        let bracket_run = MathArgument::text("[x]");
        let bracket_latex = equation_to_latex(&bracket_run);
        assert!(bracket_latex.diagnostics.is_empty());
        assert_eq!(
            equation_from_latex(&bracket_latex.value)
                .expect("escaped bracket run")
                .value,
            bracket_run
        );
        let spaced_run = equation_to_latex(&MathArgument::text("a b"));
        assert_eq!(
            spaced_run.diagnostics[0].message,
            "OfficeMath run whitespace was discarded by LaTeX"
        );
        let empty_run = equation_to_latex(&MathArgument::text(""));
        assert_eq!(
            empty_run.diagnostics[0].message,
            "empty OfficeMath run was discarded by LaTeX"
        );
        let adjacent_empty_run = equation_to_latex(&MathArgument::new(vec![
            MathRun::new("").into(),
            MathRun::new("x").into(),
        ]));
        assert_eq!(adjacent_empty_run.value, "x");
        assert!(
            adjacent_empty_run
                .diagnostics
                .iter()
                .any(|value| { value.message == "empty OfficeMath run was discarded by LaTeX" })
        );

        let invalid_accent = MathArgument::new(vec![MathExpression::Accent(MathAccent::new(
            "ab",
            MathArgument::text("x"),
        ))]);
        let invalid_accent = equation_to_mathml(&invalid_accent);
        assert!(!invalid_accent.value.contains("ab"));
        assert_eq!(
            invalid_accent.diagnostics[0].message,
            "non-scalar OfficeMath accent was discarded"
        );

        let forbidden_xml = equation_to_mathml(&MathArgument::text("a\0b"));
        assert!(!forbidden_xml.value.contains('\0'));
        assert_eq!(
            forbidden_xml.diagnostics[0].message,
            "OfficeMath run contains a forbidden XML 1.0 character"
        );

        let mut sub_sup = MathSubSuperscript::new(
            MathArgument::text("x"),
            MathArgument::text("i"),
            MathArgument::text("n"),
        );
        sub_sup.alignment = Some(true);
        let mut pre = MathPreSubSuperscript::new(
            MathArgument::text("x"),
            MathArgument::text("i"),
            MathArgument::text("n"),
        );
        pre.alignment = Some(false);
        let mut radical =
            MathRadical::with_degree(MathArgument::text("3"), MathArgument::text("x"));
        radical.hide_degree = true;
        let mut delimiter = MathDelimiter::new("(", ")", vec![MathArgument::text("x")]);
        delimiter.grow = Some(false);
        let mut nary = MathNary::new("∑", MathArgument::text("x"));
        nary.subscript = MathArgument::text("i");
        nary.superscript = MathArgument::text("n");
        nary.grow = Some(false);
        nary.limit_location = Some(LimitLocation::UnderOver);
        let typed_losses = MathArgument::new(vec![
            MathExpression::SubSuperscript(sub_sup),
            MathExpression::PreSubSuperscript(pre),
            MathExpression::Radical(radical),
            MathExpression::Delimiter(delimiter),
            MathExpression::Nary(nary),
        ]);
        for converted in [
            equation_to_mathml(&typed_losses),
            equation_to_latex(&typed_losses),
        ] {
            for message in [
                "OfficeMath script alignment was discarded",
                "OfficeMath radical degree visibility was discarded",
                "OfficeMath delimiter growth setting was discarded",
                "OfficeMath n-ary layout properties were discarded",
                "hidden OfficeMath n-ary subscript was discarded",
                "hidden OfficeMath n-ary superscript was discarded",
            ] {
                assert!(
                    converted
                        .diagnostics
                        .iter()
                        .any(|value| value.message == message),
                    "missing diagnostic: {message}"
                );
            }
        }
    }

    #[test]
    fn math_converters_reject_every_declared_input_tree_and_output_limit() {
        assert!(equation_from_latex(&"x".repeat(MAX_INPUT_BYTES + 1)).is_err());
        assert!(equation_from_latex(&"x".repeat(MAX_TEXT_BYTES + 1)).is_err());
        assert!(equation_from_latex(&"\\unknown{}".repeat(MAX_DIAGNOSTICS + 1)).is_err());
        assert!(
            equation_from_mathml(&format!(
                "<math xmlns=\"{MATHML_NS}\">{}</math>",
                " ".repeat(MAX_INPUT_BYTES)
            ))
            .is_err()
        );
        assert!(
            equation_from_mathml(&format!(
                "<math xmlns=\"{MATHML_NS}\">{}</math>",
                "x".repeat(MAX_TEXT_BYTES + 1)
            ))
            .is_err()
        );
        assert!(
            equation_from_latex(&format!(
                "{}x{}",
                "{".repeat(MAX_DEPTH + 1),
                "}".repeat(MAX_DEPTH + 1)
            ))
            .is_err()
        );
        let empty_matrix = MathArgument::new(vec![MathExpression::Matrix(MathMatrix::new(vec![]))]);
        assert!(equation_to_mathml(&empty_matrix).value.is_empty());
        assert!(equation_to_latex(&empty_matrix).value.is_empty());

        let event_bomb = format!(
            "<math xmlns=\"{MATHML_NS}\">{}</math>",
            "<!---->".repeat(MAX_EVENTS + 1)
        );
        assert!(equation_from_mathml(&event_bomb).is_err());
        let diagnostic_bomb = format!(
            "<math xmlns=\"{MATHML_NS}\">{}</math>",
            "<maction/>".repeat(MAX_DIAGNOSTICS + 1)
        );
        assert!(equation_from_mathml(&diagnostic_bomb).is_err());
        assert!(equation_from_latex(&"{}".repeat(MAX_EVENTS + 1)).is_err());
        assert!(equation_from_latex(&"x ".repeat(MAX_NODES + 1)).is_err());

        let too_many_rows = format!(
            "<math xmlns=\"{MATHML_NS}\"><mtable>{}</mtable></math>",
            "<mtr><mtd><mn>1</mn></mtd></mtr>".repeat(MAX_ROWS + 1)
        );
        assert!(equation_from_mathml(&too_many_rows).is_err());
        let too_many_columns = format!(
            "<math xmlns=\"{MATHML_NS}\"><mtable><mtr>{}</mtr></mtable></math>",
            "<mtd><mn>1</mn></mtd>".repeat(MAX_COLUMNS + 1)
        );
        assert!(equation_from_mathml(&too_many_columns).is_err());

        let too_deep_mathml = format!(
            "<math xmlns=\"{MATHML_NS}\">{}<mi>x</mi>{}</math>",
            "<mrow>".repeat(MAX_DEPTH),
            "</mrow>".repeat(MAX_DEPTH)
        );
        assert!(equation_from_mathml(&too_deep_mathml).is_err());
        let too_many_mathml_nodes = format!(
            "<math xmlns=\"{MATHML_NS}\">{}</math>",
            "<mi/>".repeat(MAX_NODES + 1)
        );
        assert!(equation_from_mathml(&too_many_mathml_nodes).is_err());

        let too_many_latex_rows = format!(
            "\\begin{{matrix}}{}\\end{{matrix}}",
            std::iter::repeat_n("x", MAX_ROWS + 1)
                .collect::<Vec<_>>()
                .join("\\\\")
        );
        assert!(equation_from_latex(&too_many_latex_rows).is_err());
        let too_many_latex_columns = format!(
            "\\begin{{matrix}}{}\\end{{matrix}}",
            std::iter::repeat_n("x", MAX_COLUMNS + 1)
                .collect::<Vec<_>>()
                .join("&")
        );
        assert!(equation_from_latex(&too_many_latex_columns).is_err());

        let too_many_output_nodes = MathArgument::new(
            (0..=MAX_NODES)
                .map(|_| MathExpression::Run(MathRun::new("x")))
                .collect(),
        );
        assert!(equation_to_mathml(&too_many_output_nodes).value.is_empty());
        assert!(equation_to_latex(&too_many_output_nodes).value.is_empty());
        let too_much_output_text = MathArgument::text("x".repeat(MAX_TEXT_BYTES + 1));
        assert!(equation_to_mathml(&too_much_output_text).value.is_empty());
        assert!(equation_to_latex(&too_much_output_text).value.is_empty());
        let escaped_mathml = equation_to_mathml(&MathArgument::text("&".repeat(MAX_TEXT_BYTES)));
        assert!(escaped_mathml.value.is_empty());
        assert_eq!(
            escaped_mathml.diagnostics[0].message,
            "serialized MathML exceeds the reader byte limit"
        );
        let escaped_latex = equation_to_latex(&MathArgument::text("\\".repeat(MAX_TEXT_BYTES)));
        assert!(escaped_latex.value.is_empty());
        assert_eq!(
            escaped_latex.diagnostics[0].message,
            "serialized LaTeX exceeds the reader byte limit"
        );
        let mut too_deep = MathArgument::text("x");
        for _ in 0..=MAX_DEPTH {
            too_deep = MathArgument::new(vec![MathExpression::Radical(MathRadical::new(too_deep))]);
        }
        assert!(equation_to_mathml(&too_deep).value.is_empty());
        assert!(equation_to_latex(&too_deep).value.is_empty());

        let mut reader_depth_boundary = MathArgument::text("x");
        for _ in 0..MAX_DEPTH {
            reader_depth_boundary = MathArgument::new(vec![MathExpression::Radical(
                MathRadical::new(reader_depth_boundary),
            )]);
        }
        assert!(equation_to_mathml(&reader_depth_boundary).value.is_empty());
        assert!(equation_to_latex(&reader_depth_boundary).value.is_empty());

        let mut nested_matrix_boundary = MathArgument::text("x");
        for _ in 0..(MAX_DEPTH / 3) {
            nested_matrix_boundary =
                MathArgument::new(vec![MathExpression::Matrix(MathMatrix::new(vec![
                    MathMatrixRow::new(vec![nested_matrix_boundary]),
                ]))]);
        }
        assert!(equation_to_mathml(&nested_matrix_boundary).value.is_empty());

        let too_many_output_rows =
            MathArgument::new(vec![MathExpression::Matrix(MathMatrix::new(
                (0..=MAX_ROWS)
                    .map(|_| MathMatrixRow::new(vec![MathArgument::text("x")]))
                    .collect(),
            ))]);
        assert!(equation_to_mathml(&too_many_output_rows).value.is_empty());
        assert!(equation_to_latex(&too_many_output_rows).value.is_empty());
        let too_many_output_columns =
            MathArgument::new(vec![MathExpression::Matrix(MathMatrix::new(vec![
                MathMatrixRow::new((0..=MAX_COLUMNS).map(|_| MathArgument::text("x")).collect()),
            ]))]);
        assert!(
            equation_to_mathml(&too_many_output_columns)
                .value
                .is_empty()
        );
        assert!(equation_to_latex(&too_many_output_columns).value.is_empty());

        let mut diagnostic_runs = Vec::new();
        for _ in 0..=MAX_DIAGNOSTICS {
            let mut run = MathRun::new("x");
            run.properties.style = Some(MathStyle::Bold);
            diagnostic_runs.push(MathExpression::Run(run));
        }
        let diagnostic_tree = MathArgument::new(diagnostic_runs);
        let mathml = equation_to_mathml(&diagnostic_tree);
        let latex = equation_to_latex(&diagnostic_tree);
        assert!(mathml.value.is_empty());
        assert!(latex.value.is_empty());
        assert_eq!(mathml.diagnostics.len(), 1);
        assert_eq!(latex.diagnostics.len(), 1);
    }

    #[test]
    fn supported_equations_preserve_their_normalized_tree_through_all_four_conversion_directions() {
        let source = complete_round_trip_tree();
        let mathml = equation_to_mathml(&source);
        assert!(mathml.diagnostics.is_empty());
        assert_eq!(
            equation_from_mathml(&mathml.value)
                .expect("emitted MathML")
                .value,
            source
        );
        let latex = equation_to_latex(&source);
        assert!(latex.diagnostics.is_empty());
        assert_eq!(
            equation_from_latex(&latex.value)
                .expect("emitted LaTeX")
                .value,
            source
        );

        let delimiter_with_separator_text = MathArgument::new(vec![MathExpression::Delimiter(
            MathDelimiter::new("(", ")", vec![MathArgument::text("a,b|c")]),
        )]);
        let delimiter_latex = equation_to_latex(&delimiter_with_separator_text);
        assert!(delimiter_latex.diagnostics.is_empty());
        assert_eq!(
            equation_from_latex(&delimiter_latex.value)
                .expect("grouped delimiter argument")
                .value,
            delimiter_with_separator_text
        );

        let adjacent_runs =
            MathArgument::new(vec![MathRun::new("a").into(), MathRun::new("b").into()]);
        assert_eq!(
            equation_to_mathml(&adjacent_runs).value,
            equation_to_mathml(&MathArgument::text("ab")).value,
            "canonical MathML ignores adjacent compatible run boundaries"
        );
    }

    #[test]
    #[ignore = "requires the exact pinned Pandoc 3.10 executable"]
    fn mathml_and_latex_conversion_matches_pinned_pandoc_texmath_trees() {
        let pandoc = std::env::var("RDOCX_PANDOC").unwrap_or_else(|_| "pandoc".to_owned());
        let identity = Command::new(&pandoc)
            .arg("--version")
            .output()
            .expect("run pinned Pandoc");
        assert!(identity.status.success());
        assert_eq!(
            identity.stdout.split(|byte| *byte == b'\n').next(),
            Some(b"pandoc 3.10".as_slice())
        );
        let cases = [
            r"x+2",
            r"\frac{x_1}{2}",
            r"x_i",
            r"x^n",
            r"x_i^n",
            r"{}_{i}^{n}{x}",
            r"\sqrt{x}",
            r"\sqrt[3]{y}",
            r"\underset{i}{x}",
            r"\overset{n}{x}",
            r"\sum_{i}^{n}{x}",
            r"\prod_{i}^{n}{x}",
            r"\coprod_{i}^{n}{x}",
            r"\int_{i}^{n}{x}",
            r"\iint_{i}^{n}{x}",
            r"\iiint_{i}^{n}{x}",
            r"\oint_{i}^{n}{x}",
            r"\left(a+b\right)",
            r"\hat{x}",
            r"\bar{x}",
            r"\overline{x}",
            r"\vec{x}",
            r"\tilde{x}",
            r"\dot{x}",
            r"\ddot{x}",
            r"\begin{matrix}a&b\\c&d\end{matrix}",
            r"\begin{pmatrix}a&b\\c&d\end{pmatrix}",
            r"\begin{bmatrix}a&b\\c&d\end{bmatrix}",
        ];
        for latex in cases {
            let html = pandoc_convert(
                &pandoc,
                &["--from=latex", "--to=html", "--mathml"],
                &format!("${latex}$"),
            );
            let start = html.find("<math").expect("Pandoc MathML start");
            let end = html.find("</math>").expect("Pandoc MathML end") + "</math>".len();
            let ours = equation_from_latex(latex)
                .expect("our LaTeX conversion")
                .value;
            let oracle = equation_from_mathml(&html[start..end]).expect("Pandoc MathML conversion");
            if latex.starts_with("{}_") {
                assert_ne!(ours, oracle.value, "Pandoc pre-script divergence");
                assert!(matches!(
                    oracle.value.expressions.as_slice(),
                    [MathExpression::SubSuperscript(value), MathExpression::Run(run)]
                        if value.base.expressions.is_empty() && run.text == "x"
                ));
            } else {
                assert_eq!(ours, oracle.value, "Pandoc structure for {latex}");
            }
            assert_eq!(oracle.diagnostics[0].path, "/math[1]/@display");
            assert_eq!(
                oracle.diagnostics[1],
                MathConversionDiagnostic {
                    path: "/math[1]/semantics[1]".to_owned(),
                    message: "MathML semantics metadata was discarded".to_owned(),
                }
            );
            if latex.contains("matrix") {
                assert_eq!(oracle.diagnostics.len(), 10);
                assert!(oracle.diagnostics[2..].iter().all(|value| {
                    value.path.contains("/mtd[")
                        && value.message == "unsupported MathML attribute was discarded"
                }));
            } else {
                assert_eq!(oracle.diagnostics.len(), 2);
            }

            let mathml = equation_to_mathml(&ours);
            assert!(mathml.diagnostics.is_empty());
            let oracle_latex =
                pandoc_convert(&pandoc, &["--from=html", "--to=latex"], &mathml.value);
            let oracle_latex = oracle_latex.trim();
            let oracle_latex = oracle_latex
                .strip_prefix("\\(")
                .and_then(|value| value.strip_suffix("\\)"))
                .expect("Pandoc inline LaTeX wrapper");
            let reopened = equation_from_latex(oracle_latex).expect("Pandoc LaTeX conversion");
            if latex.starts_with("\\left") {
                assert_eq!(
                    reopened.value,
                    MathArgument::text("(a+b)"),
                    "Pandoc intentionally removes explicit delimiter scope"
                );
            } else {
                assert_eq!(ours, reopened.value, "Pandoc reverse structure for {latex}");
            }
            if ours
                .expressions
                .iter()
                .any(|value| matches!(value, MathExpression::Nary(_)))
            {
                assert_eq!(reopened.diagnostics.len(), 1);
                assert_eq!(
                    reopened.diagnostics[0].message,
                    "LaTeX \\limits n-ary placement was discarded"
                );
            } else {
                assert!(reopened.diagnostics.is_empty());
            }
        }
    }

    #[test]
    fn conversion_differential_rejects_structure_scope_order_and_diagnostic_perturbations() {
        let expected = equation_from_latex(r"\frac{a_1}{b}").expect("supported LaTeX");
        let swapped = equation_from_latex(r"\frac{b}{a_1}").expect("supported LaTeX");
        let detached = equation_from_latex(r"\frac{a}{b}_1").expect("supported LaTeX");
        let scoped = equation_from_latex(r"\left(a_1\right)").expect("supported delimiter");
        let unscoped = equation_from_latex(r"a_1").expect("supported unscoped expression");
        let matrix =
            equation_from_latex(r"\begin{matrix}1&2\\3&4\end{matrix}").expect("supported matrix");
        let reordered_matrix = equation_from_latex(r"\begin{matrix}2&1\\3&4\end{matrix}")
            .expect("supported reordered matrix");
        let lossy = equation_from_latex(r"a\unsupported{b}").expect("lossy LaTeX");
        let dropped_diagnostic = MathConversionResult {
            value: lossy.value.clone(),
            diagnostics: Vec::new(),
        };
        assert!(!differential_accepts(&expected, &swapped));
        assert!(!differential_accepts(&expected, &detached));
        assert!(!differential_accepts(&scoped, &unscoped));
        assert!(!differential_accepts(&matrix, &reordered_matrix));
        assert!(!differential_accepts(&lossy, &dropped_diagnostic));
    }
}
