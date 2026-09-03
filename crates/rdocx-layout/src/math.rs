use oxml_layout::{
    Color, Diagnostic, FontManager, GlyphRun, GroupElement, Point, PositionedElement, Result,
    Transform,
};
use rdocx_oxml::math::{
    FractionType, LimitLocation, MathAccent, MathArgument, MathDelimiter, MathExpression,
    MathFraction, MathJustification, MathLimit, MathMatrix, MathNary, MathPreSubSuperscript,
    MathProperties, MathRadical, MathRun, MathScript, MathStyle, MathSubSuperscript,
    MatrixBaseJustification, OfficeMath,
};

const DEFAULT_MATH_FONT: &str = "Caladea";
const DEFAULT_MATH_SIZE: f64 = 11.0;
const SCRIPT_SCALE: f64 = 0.7;
const SUBSCRIPT_SHIFT_EM: f64 = 0.18;
const SUPERSCRIPT_SHIFT_EM: f64 = 0.43;
const GAP_EM: f64 = 0.12;
const RULE_EM: f64 = 0.055;
const MAX_DELIMITER_SCALE: f64 = 4.0;
const EQUATION_SPACING_EM: f64 = 0.107;
const LIMIT_GAP_EM: f64 = -0.29;
const MATH_TEXT_X_SCALE: f64 = 1.1;
const NARY_OPERATOR_X_SCALE: f64 = 1.89;
const NARY_UPPER_GAP_EM: f64 = 0.482;
const NARY_LOWER_GAP_EM: f64 = 0.3;
const DEFAULT_DISPLAY_PRE_SPACING: f64 = 8.36;

#[derive(Debug, Clone)]
pub(crate) struct MeasuredMath {
    pub(crate) width: f64,
    pub(crate) ascent: f64,
    pub(crate) descent: f64,
    pub(crate) group: GroupElement,
}

impl MeasuredMath {
    pub(crate) fn height(&self) -> f64 {
        self.ascent + self.descent
    }

    fn empty() -> Self {
        Self {
            width: 0.0,
            ascent: 0.0,
            descent: 0.0,
            group: group(Vec::new()),
        }
    }

    #[cfg(test)]
    fn is_bounded(&self) -> bool {
        self.width.is_finite()
            && self.ascent.is_finite()
            && self.descent.is_finite()
            && self.width >= 0.0
            && self.ascent >= 0.0
            && self.descent >= 0.0
    }
}

pub(crate) fn layout_officemath(
    equation: &OfficeMath,
    fm: &mut FontManager,
    font_size: f64,
    color: Color,
    available_width: f64,
    math_properties: Option<&MathProperties>,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<(MeasuredMath, bool)> {
    let size = finite_positive(font_size).unwrap_or(DEFAULT_MATH_SIZE);
    let font_family = math_properties
        .and_then(|properties| properties.math_font.as_deref())
        .filter(|family| !family.is_empty())
        .or(Some(DEFAULT_MATH_FONT));
    let mut equation = equation.clone();
    apply_limit_defaults(&mut equation, math_properties);
    if equation.has_unsupported_content() {
        diagnostics.push(Diagnostic {
            message: format!(
                "OfficeMath content at {source_path} was preserved but could not be rendered"
            ),
        });
    }

    match &equation {
        OfficeMath::Inline(value) => Ok((
            layout_expressions(
                &value.expressions,
                fm,
                font_family,
                size,
                color,
                size * EQUATION_SPACING_EM,
                source_path,
                diagnostics,
            )?,
            false,
        )),
        OfficeMath::Display(value) => {
            let mut equations = Vec::with_capacity(value.equations.len());
            for (index, equation) in value.equations.iter().enumerate() {
                equations.push(layout_expressions(
                    &equation.expressions,
                    fm,
                    font_family,
                    size,
                    color,
                    size * EQUATION_SPACING_EM,
                    &format!("{source_path}/equation/{index}"),
                    diagnostics,
                )?);
            }
            let inter_spacing = math_properties
                .and_then(|properties| properties.inter_spacing)
                .map_or(0.0, twips_to_points);
            let measured = layout_spaced_sequence(equations, inter_spacing);
            let left_margin = math_properties
                .and_then(|properties| properties.left_margin)
                .map_or(0.0, twips_to_points);
            let right_margin = math_properties
                .and_then(|properties| properties.right_margin)
                .map_or(0.0, twips_to_points);
            let intra_spacing = math_properties
                .and_then(|properties| properties.intra_spacing)
                .map_or(0.0, twips_to_points);
            let pre_spacing = math_properties
                .and_then(|properties| properties.pre_spacing)
                .map_or(DEFAULT_DISPLAY_PRE_SPACING, twips_to_points);
            let post_spacing = math_properties
                .and_then(|properties| properties.post_spacing)
                .map_or(0.0, twips_to_points);
            let content_width = (available_width - left_margin - right_margin).max(0.0);
            let x = match value
                .properties
                .justification
                .or_else(|| math_properties.and_then(|properties| properties.justification))
                .unwrap_or(MathJustification::CenterGroup)
            {
                MathJustification::Left => left_margin + intra_spacing,
                MathJustification::Right => {
                    left_margin + (content_width - measured.width - intra_spacing).max(0.0)
                }
                MathJustification::Center | MathJustification::CenterGroup => {
                    left_margin + ((content_width - measured.width) / 2.0).max(0.0)
                }
            };
            let width = available_width.max(measured.width);
            Ok((
                MeasuredMath {
                    width,
                    ascent: measured.ascent + pre_spacing,
                    descent: measured.descent + post_spacing,
                    group: group(vec![translated(measured.group, x, pre_spacing)]),
                },
                true,
            ))
        }
    }
}

fn twips_to_points(value: u32) -> f64 {
    value as f64 / 20.0
}

fn apply_limit_defaults(equation: &mut OfficeMath, properties: Option<&MathProperties>) {
    let Some(properties) = properties else {
        return;
    };
    let equations = match equation {
        OfficeMath::Inline(equation) => std::slice::from_mut(equation),
        OfficeMath::Display(display) => display.equations.as_mut_slice(),
    };
    for equation in equations {
        apply_expression_limit_defaults(&mut equation.expressions, properties);
    }
}

fn apply_expression_limit_defaults(
    expressions: &mut [MathExpression],
    properties: &MathProperties,
) {
    fn argument(argument: &mut MathArgument, properties: &MathProperties) {
        apply_expression_limit_defaults(&mut argument.expressions, properties);
    }

    for expression in expressions {
        match expression {
            MathExpression::Run(_) => {}
            MathExpression::Fraction(value) => {
                argument(&mut value.numerator, properties);
                argument(&mut value.denominator, properties);
            }
            MathExpression::Subscript(value) | MathExpression::Superscript(value) => {
                argument(&mut value.base, properties);
                argument(&mut value.script, properties);
            }
            MathExpression::SubSuperscript(value) => {
                argument(&mut value.base, properties);
                argument(&mut value.subscript, properties);
                argument(&mut value.superscript, properties);
            }
            MathExpression::PreSubSuperscript(value) => {
                argument(&mut value.base, properties);
                argument(&mut value.subscript, properties);
                argument(&mut value.superscript, properties);
            }
            MathExpression::Radical(value) => {
                argument(&mut value.degree, properties);
                argument(&mut value.base, properties);
            }
            MathExpression::Matrix(value) => {
                for row in &mut value.rows {
                    for cell in &mut row.cells {
                        argument(cell, properties);
                    }
                }
            }
            MathExpression::LowerLimit(value) | MathExpression::UpperLimit(value) => {
                argument(&mut value.base, properties);
                argument(&mut value.limit, properties);
            }
            MathExpression::Nary(value) => {
                if value.limit_location.is_none() {
                    let is_integral = value.character.is_empty()
                        || value.character.chars().any(|character| {
                            matches!(character, '\u{222b}' | '\u{222c}' | '\u{222d}' | '\u{222e}')
                        });
                    value.limit_location = if is_integral {
                        properties.integral_limit_location
                    } else {
                        properties.nary_limit_location
                    };
                }
                argument(&mut value.base, properties);
                argument(&mut value.subscript, properties);
                argument(&mut value.superscript, properties);
            }
            MathExpression::Delimiter(value) => {
                for argument_value in &mut value.arguments {
                    argument(argument_value, properties);
                }
            }
            MathExpression::Accent(value) => argument(&mut value.base, properties),
        }
    }
}

fn layout_expressions(
    expressions: &[MathExpression],
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    spacing: f64,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    let mut measured = Vec::with_capacity(expressions.len());
    for (index, expression) in expressions.iter().enumerate() {
        measured.push(layout_expression(
            expression,
            fm,
            font_family,
            font_size,
            color,
            &format!("{source_path}/expression/{index}"),
            diagnostics,
        )?);
    }
    Ok(layout_spaced_sequence(measured, spacing))
}

fn layout_argument(
    argument: &MathArgument,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    layout_expressions(
        &argument.expressions,
        fm,
        font_family,
        font_size,
        color,
        0.0,
        source_path,
        diagnostics,
    )
}

fn layout_expression(
    expression: &MathExpression,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    match expression {
        MathExpression::Run(value) => layout_run(
            value,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::Fraction(value) => layout_fraction(
            value,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::Subscript(value) => layout_script(
            value,
            false,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::Superscript(value) => layout_script(
            value,
            true,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::SubSuperscript(value) => layout_sub_superscript(
            value,
            false,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::PreSubSuperscript(value) => layout_pre_sub_superscript(
            value,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::Radical(value) => layout_radical(
            value,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::Matrix(value) => layout_matrix(
            value,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::LowerLimit(value) => layout_limit(
            value,
            false,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::UpperLimit(value) => layout_limit(
            value,
            true,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::Nary(value) => layout_nary(
            value,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::Delimiter(value) => layout_delimiter(
            value,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
        MathExpression::Accent(value) => layout_accent(
            value,
            fm,
            font_family,
            font_size,
            color,
            source_path,
            diagnostics,
        ),
    }
}

fn layout_run(
    run: &MathRun,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    let style = if run.properties.normal == Some(true) || run.properties.literal == Some(true) {
        MathStyle::Plain
    } else {
        run.properties.style.unwrap_or(MathStyle::Italic)
    };
    let bold = matches!(style, MathStyle::Bold | MathStyle::BoldItalic);
    let italic = matches!(style, MathStyle::Italic | MathStyle::BoldItalic);
    layout_text(
        &run.text,
        fm,
        font_family,
        font_size,
        color,
        bold,
        italic,
        source_path,
        diagnostics,
    )
}

#[allow(clippy::too_many_arguments)]
fn layout_text(
    text: &str,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    bold: bool,
    italic: bool,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    let family = font_family
        .filter(|value| !value.is_empty())
        .unwrap_or(DEFAULT_MATH_FONT);
    let font_id = fm.resolve_font_for_text(Some(family), bold, italic, text)?;
    let shaped = fm.shape_text(font_id, text, font_size)?;
    let metrics = fm.metrics(font_id, font_size)?;
    if shaped.glyph_ids.contains(&0) {
        diagnostics.push(Diagnostic {
            message: format!("OfficeMath text at {source_path} contains an unavailable glyph"),
        });
    }
    Ok(horizontally_scaled(
        MeasuredMath {
            width: shaped.width.max(0.0),
            ascent: metrics.ascent.max(0.0),
            descent: metrics.descent.max(0.0),
            group: group(vec![PositionedElement::Text(GlyphRun {
                origin: Point {
                    x: 0.0,
                    y: metrics.ascent.max(0.0),
                },
                font_id,
                font_size,
                glyph_ids: shaped.glyph_ids,
                advances: shaped.advances,
                text: text.to_owned(),
                source: None,
                color,
                bold,
                italic,
                field_kind: None,
                note: None,
            })]),
        },
        MATH_TEXT_X_SCALE,
    ))
}

#[allow(clippy::too_many_arguments)]
fn layout_fraction(
    value: &MathFraction,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    let child_size = font_size;
    let numerator = layout_argument(
        &value.numerator,
        fm,
        font_family,
        child_size,
        color,
        &format!("{source_path}/numerator"),
        diagnostics,
    )?;
    let denominator = layout_argument(
        &value.denominator,
        fm,
        font_family,
        child_size,
        color,
        &format!("{source_path}/denominator"),
        diagnostics,
    )?;
    if matches!(
        value.fraction_type,
        FractionType::Linear | FractionType::Skewed
    ) {
        let slash = layout_text(
            "/",
            fm,
            font_family,
            font_size,
            color,
            false,
            false,
            source_path,
            diagnostics,
        )?;
        return Ok(layout_sequence(vec![numerator, slash, denominator]));
    }
    let gap = font_size * 0.55;
    let rule = font_size * RULE_EM;
    let padding = font_size * 0.06;
    let width = numerator.width.max(denominator.width) + padding * 2.0;
    let ascent = numerator.height() + gap + rule / 2.0;
    let numerator_x = (width - numerator.width) / 2.0;
    let denominator_x = (width - denominator.width) / 2.0;
    let denominator_y = ascent - font_size * 0.36;
    let descent = (denominator_y + denominator.height() - ascent).max(0.0);
    let mut children = vec![
        translated(numerator.group, numerator_x, 0.0),
        translated(denominator.group, denominator_x, denominator_y),
    ];
    if value.fraction_type != FractionType::NoBar {
        children.push(PositionedElement::Line {
            start: Point {
                x: padding / 2.0,
                y: ascent,
            },
            end: Point {
                x: width - padding / 2.0,
                y: ascent,
            },
            width: rule,
            color,
            dash_pattern: None,
        });
    }
    Ok(MeasuredMath {
        width,
        ascent,
        descent,
        group: group(children),
    })
}

#[allow(clippy::too_many_arguments)]
fn layout_script(
    value: &MathScript,
    superscript: bool,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    let base = layout_argument(
        &value.base,
        fm,
        font_family,
        font_size,
        color,
        &format!("{source_path}/base"),
        diagnostics,
    )?;
    let script = layout_argument(
        &value.script,
        fm,
        font_family,
        font_size * SCRIPT_SCALE,
        color,
        &format!("{source_path}/script"),
        diagnostics,
    )?;
    layout_script_pair(base, script, superscript, font_size)
}

fn layout_script_pair(
    base: MeasuredMath,
    script: MeasuredMath,
    superscript: bool,
    font_size: f64,
) -> Result<MeasuredMath> {
    let shift = font_size
        * if superscript {
            SUPERSCRIPT_SHIFT_EM
        } else {
            SUBSCRIPT_SHIFT_EM
        };
    let (ascent, descent, script_baseline) = if superscript {
        (
            base.ascent.max(shift + script.ascent),
            base.descent.max((script.descent - shift).max(0.0)),
            -shift,
        )
    } else {
        (
            base.ascent.max((script.ascent - shift).max(0.0)),
            base.descent.max(shift + script.descent),
            shift,
        )
    };
    let base_y = ascent - base.ascent;
    let script_y = ascent + script_baseline - script.ascent;
    let gap = font_size * GAP_EM / 2.0;
    Ok(MeasuredMath {
        width: base.width + gap + script.width,
        ascent,
        descent,
        group: group(vec![
            translated(base.group, 0.0, base_y),
            translated(script.group, base.width + gap, script_y),
        ]),
    })
}

#[allow(clippy::too_many_arguments)]
fn layout_sub_superscript(
    value: &MathSubSuperscript,
    pre: bool,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    layout_three_part_script(
        &value.base,
        &value.subscript,
        &value.superscript,
        pre,
        fm,
        font_family,
        font_size,
        color,
        source_path,
        diagnostics,
    )
}

#[allow(clippy::too_many_arguments)]
fn layout_pre_sub_superscript(
    value: &MathPreSubSuperscript,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    layout_three_part_script(
        &value.base,
        &value.subscript,
        &value.superscript,
        true,
        fm,
        font_family,
        font_size,
        color,
        source_path,
        diagnostics,
    )
}

#[allow(clippy::too_many_arguments)]
fn layout_three_part_script(
    base_value: &MathArgument,
    subscript_value: &MathArgument,
    superscript_value: &MathArgument,
    pre: bool,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    let base = layout_argument(
        base_value,
        fm,
        font_family,
        font_size,
        color,
        &format!("{source_path}/base"),
        diagnostics,
    )?;
    let subscript = layout_argument(
        subscript_value,
        fm,
        font_family,
        font_size * SCRIPT_SCALE,
        color,
        &format!("{source_path}/subscript"),
        diagnostics,
    )?;
    let superscript = layout_argument(
        superscript_value,
        fm,
        font_family,
        font_size * SCRIPT_SCALE,
        color,
        &format!("{source_path}/superscript"),
        diagnostics,
    )?;
    let superscript_shift = font_size * SUPERSCRIPT_SHIFT_EM;
    let subscript_shift = font_size * SUBSCRIPT_SHIFT_EM;
    let ascent = base.ascent.max(superscript_shift + superscript.ascent);
    let descent = base.descent.max(subscript_shift + subscript.descent);
    let script_width = subscript.width.max(superscript.width);
    let gap = font_size * GAP_EM / 2.0;
    let (script_x, base_x) = if pre {
        (0.0, script_width + gap)
    } else {
        (base.width + gap, 0.0)
    };
    Ok(MeasuredMath {
        width: base.width + gap + script_width,
        ascent,
        descent,
        group: group(vec![
            translated(base.group, base_x, ascent - base.ascent),
            translated(
                superscript.group,
                script_x,
                ascent - superscript_shift - superscript.ascent,
            ),
            translated(
                subscript.group,
                script_x,
                ascent + subscript_shift - subscript.ascent,
            ),
        ]),
    })
}

#[allow(clippy::too_many_arguments)]
fn layout_radical(
    value: &MathRadical,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    let base = layout_argument(
        &value.base,
        fm,
        font_family,
        font_size,
        color,
        &format!("{source_path}/base"),
        diagnostics,
    )?;
    let root = layout_text(
        "√",
        fm,
        font_family,
        font_size,
        color,
        false,
        false,
        source_path,
        diagnostics,
    )?;
    let gap = font_size * GAP_EM;
    let rule = font_size * RULE_EM;
    let target_height = base.height() + gap + rule;
    let root = vertically_scaled(root, target_height);
    let degree = if value.hide_degree {
        MeasuredMath::empty()
    } else {
        layout_argument(
            &value.degree,
            fm,
            font_family,
            font_size * SCRIPT_SCALE * SCRIPT_SCALE,
            color,
            &format!("{source_path}/degree"),
            diagnostics,
        )?
    };
    let degree_overlap = degree.width * 0.35;
    let root_x = (degree.width - degree_overlap).max(0.0);
    let base_x = root_x + root.width;
    let ascent = root.ascent.max(base.ascent + gap + rule);
    let descent = root.descent.max(base.descent);
    let base_y = ascent - base.ascent;
    let mut children = vec![
        translated(root.group, root_x, ascent - root.ascent),
        translated(base.group, base_x, base_y),
        PositionedElement::Line {
            start: Point {
                x: base_x,
                y: base_y - gap / 2.0,
            },
            end: Point {
                x: base_x + base.width,
                y: base_y - gap / 2.0,
            },
            width: rule,
            color,
            dash_pattern: None,
        },
    ];
    if degree.width > 0.0 {
        children.push(translated(degree.group, 0.0, 0.0));
    }
    Ok(horizontally_scaled(
        MeasuredMath {
            width: base_x + base.width,
            ascent,
            descent,
            group: group(children),
        },
        0.9,
    ))
}

#[allow(clippy::too_many_arguments)]
fn layout_matrix(
    value: &MathMatrix,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    if value.rows.is_empty() {
        return Ok(MeasuredMath::empty());
    }
    let column_count = value
        .rows
        .iter()
        .map(|row| row.cells.len())
        .max()
        .unwrap_or(0);
    let mut rows = Vec::with_capacity(value.rows.len());
    let mut column_widths = vec![0.0_f64; column_count];
    for (row_index, row) in value.rows.iter().enumerate() {
        let mut cells = Vec::with_capacity(row.cells.len());
        for (column_index, cell) in row.cells.iter().enumerate() {
            let measured = layout_argument(
                cell,
                fm,
                font_family,
                font_size,
                color,
                &format!("{source_path}/row/{row_index}/cell/{column_index}"),
                diagnostics,
            )?;
            column_widths[column_index] = column_widths[column_index].max(measured.width);
            cells.push(measured);
        }
        rows.push(cells);
    }
    let column_gap = value
        .properties
        .column_spacing
        .map_or(font_size * 1.06, |value| value as f64 / 20.0);
    let row_gap = value
        .properties
        .row_spacing
        .map_or(0.0, |value| value as f64 / 20.0);
    let width =
        column_widths.iter().sum::<f64>() + column_gap * column_count.saturating_sub(1) as f64;
    let row_metrics = rows
        .iter()
        .map(|row| {
            (
                row.iter().map(|cell| cell.ascent).fold(0.0, f64::max),
                row.iter().map(|cell| cell.descent).fold(0.0, f64::max),
            )
        })
        .collect::<Vec<_>>();
    let height = row_metrics
        .iter()
        .map(|(ascent, descent)| ascent + descent)
        .sum::<f64>()
        + row_gap * rows.len().saturating_sub(1) as f64;
    let ascent = match value.properties.base_justification {
        Some(MatrixBaseJustification::Top) => row_metrics[0].0,
        Some(MatrixBaseJustification::Bottom) => {
            height - row_metrics.last().map_or(0.0, |(_, descent)| *descent)
        }
        Some(MatrixBaseJustification::Center) | None => {
            let (row_ascent, row_descent) = row_metrics[0];
            height / 2.0 + (row_ascent - row_descent) / 2.0
        }
    };
    let mut children = Vec::new();
    let mut y = 0.0;
    for (row, (row_ascent, row_descent)) in rows.into_iter().zip(row_metrics) {
        let mut x = 0.0;
        for (column_index, cell) in row.into_iter().enumerate() {
            let cell_x = x + (column_widths[column_index] - cell.width) / 2.0;
            children.push(translated(cell.group, cell_x, y + row_ascent - cell.ascent));
            x += column_widths[column_index] + column_gap;
        }
        y += row_ascent + row_descent + row_gap;
    }
    Ok(MeasuredMath {
        width,
        ascent,
        descent: height - ascent,
        group: group(children),
    })
}

#[allow(clippy::too_many_arguments)]
fn layout_limit(
    value: &MathLimit,
    upper: bool,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    let base = layout_argument(
        &value.base,
        fm,
        font_family,
        font_size,
        color,
        &format!("{source_path}/base"),
        diagnostics,
    )?;
    let limit = layout_argument(
        &value.limit,
        fm,
        font_family,
        font_size * SCRIPT_SCALE,
        color,
        &format!("{source_path}/limit"),
        diagnostics,
    )?;
    let gap = font_size * LIMIT_GAP_EM;
    let width = base.width.max(limit.width);
    if upper {
        let limit_height = limit.height();
        let ascent = limit.height() + gap + base.ascent;
        Ok(MeasuredMath {
            width,
            ascent,
            descent: base.descent,
            group: group(vec![
                translated(limit.group, (width - limit.width) / 2.0, 0.0),
                translated(base.group, (width - base.width) / 2.0, limit_height + gap),
            ]),
        })
    } else {
        let ascent = base.ascent;
        let base_height = base.height();
        Ok(MeasuredMath {
            width,
            ascent,
            descent: base.descent + gap + limit.height(),
            group: group(vec![
                translated(base.group, (width - base.width) / 2.0, 0.0),
                translated(limit.group, (width - limit.width) / 2.0, base_height + gap),
            ]),
        })
    }
}

#[allow(clippy::too_many_arguments)]
fn layout_nary(
    value: &MathNary,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    let operator_text = if value.character.is_empty() {
        "∫"
    } else {
        value.character.as_str()
    };
    let operator = layout_text(
        operator_text,
        fm,
        font_family,
        font_size,
        color,
        false,
        false,
        source_path,
        diagnostics,
    )?;
    let operator = if value.grow == Some(false) {
        operator
    } else {
        horizontally_scaled(operator, NARY_OPERATOR_X_SCALE)
    };
    let base = layout_argument(
        &value.base,
        fm,
        font_family,
        font_size,
        color,
        &format!("{source_path}/base"),
        diagnostics,
    )?;
    let mut decorated = operator;
    let under_over = value.limit_location == Some(LimitLocation::UnderOver);
    if !value.hide_subscript && !value.subscript.expressions.is_empty() {
        let subscript = layout_argument(
            &value.subscript,
            fm,
            font_family,
            font_size * SCRIPT_SCALE,
            color,
            &format!("{source_path}/subscript"),
            diagnostics,
        )?;
        decorated = if under_over {
            stack_below(decorated, subscript, font_size)
        } else {
            layout_script_pair(decorated, subscript, false, font_size)?
        };
    }
    if !value.hide_superscript && !value.superscript.expressions.is_empty() {
        let superscript = layout_argument(
            &value.superscript,
            fm,
            font_family,
            font_size * SCRIPT_SCALE,
            color,
            &format!("{source_path}/superscript"),
            diagnostics,
        )?;
        decorated = if under_over {
            stack_above(decorated, superscript, font_size)
        } else {
            layout_script_pair(decorated, superscript, true, font_size)?
        };
    }
    Ok(layout_sequence(vec![decorated, base]))
}

fn stack_above(base: MeasuredMath, limit: MeasuredMath, font_size: f64) -> MeasuredMath {
    let gap = font_size * NARY_UPPER_GAP_EM;
    let width = base.width.max(limit.width);
    let limit_height = limit.height();
    let ascent = limit_height + gap + base.ascent;
    MeasuredMath {
        width,
        ascent,
        descent: base.descent,
        group: group(vec![
            translated(limit.group, (width - limit.width) / 2.0, 0.0),
            translated(base.group, (width - base.width) / 2.0, limit_height + gap),
        ]),
    }
}

fn stack_below(base: MeasuredMath, limit: MeasuredMath, font_size: f64) -> MeasuredMath {
    let gap = font_size * NARY_LOWER_GAP_EM;
    let width = base.width.max(limit.width);
    let base_height = base.height();
    MeasuredMath {
        width,
        ascent: base.ascent,
        descent: base.descent + gap + limit.height(),
        group: group(vec![
            translated(base.group, (width - base.width) / 2.0, 0.0),
            translated(limit.group, (width - limit.width) / 2.0, base_height + gap),
        ]),
    }
}

#[allow(clippy::too_many_arguments)]
fn layout_delimiter(
    value: &MathDelimiter,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    let mut arguments = Vec::new();
    for (index, argument) in value.arguments.iter().enumerate() {
        if index > 0 && !value.separator_character.is_empty() {
            arguments.push(layout_text(
                &value.separator_character,
                fm,
                font_family,
                font_size,
                color,
                false,
                false,
                source_path,
                diagnostics,
            )?);
        }
        arguments.push(layout_argument(
            argument,
            fm,
            font_family,
            font_size,
            color,
            &format!("{source_path}/argument/{index}"),
            diagnostics,
        )?);
    }
    let contents = layout_sequence(arguments);
    let target_height = if value.grow == Some(false) {
        font_size
    } else {
        contents.height() * 1.08
    };
    let mut pieces = Vec::new();
    if !value.begin_character.is_empty() {
        pieces.push(vertically_scaled(
            layout_text(
                &value.begin_character,
                fm,
                font_family,
                font_size,
                color,
                false,
                false,
                source_path,
                diagnostics,
            )?,
            target_height,
        ));
    }
    pieces.push(contents);
    if !value.end_character.is_empty() {
        pieces.push(vertically_scaled(
            layout_text(
                &value.end_character,
                fm,
                font_family,
                font_size,
                color,
                false,
                false,
                source_path,
                diagnostics,
            )?,
            target_height,
        ));
    }
    Ok(layout_sequence(pieces))
}

#[allow(clippy::too_many_arguments)]
fn layout_accent(
    value: &MathAccent,
    fm: &mut FontManager,
    font_family: Option<&str>,
    font_size: f64,
    color: Color,
    source_path: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<MeasuredMath> {
    let base = layout_argument(
        &value.base,
        fm,
        font_family,
        font_size,
        color,
        &format!("{source_path}/base"),
        diagnostics,
    )?;
    let accent_text = if value.character.is_empty() {
        "̂"
    } else {
        value.character.as_str()
    };
    let gap = font_size * GAP_EM / 2.0;
    if matches!(accent_text, "¯" | "̅" | "̲") {
        let rule = font_size * RULE_EM;
        return Ok(MeasuredMath {
            width: base.width,
            ascent: base.ascent + gap + rule,
            descent: base.descent,
            group: group(vec![
                translated(base.group, 0.0, gap + rule),
                PositionedElement::Line {
                    start: Point {
                        x: 0.0,
                        y: rule / 2.0,
                    },
                    end: Point {
                        x: base.width,
                        y: rule / 2.0,
                    },
                    width: rule,
                    color,
                    dash_pattern: None,
                },
            ]),
        });
    }
    let accent = layout_text(
        accent_text,
        fm,
        font_family,
        font_size * SCRIPT_SCALE,
        color,
        false,
        false,
        source_path,
        diagnostics,
    )?;
    let width = base.width.max(accent.width);
    let accent_raise = font_size * 0.262;
    let ascent = base.ascent + accent_raise;
    Ok(MeasuredMath {
        width,
        ascent,
        descent: base.descent.max(accent.descent),
        group: group(vec![
            translated(accent.group, (width - accent.width) / 2.0, 0.0),
            translated(base.group, (width - base.width) / 2.0, accent_raise),
        ]),
    })
}

fn vertically_scaled(value: MeasuredMath, target_height: f64) -> MeasuredMath {
    let height = value.height();
    if height <= 0.0 || target_height <= height {
        return value;
    }
    let scale = (target_height / height).min(MAX_DELIMITER_SCALE);
    MeasuredMath {
        width: value.width,
        ascent: value.ascent * scale,
        descent: value.descent * scale,
        group: GroupElement {
            transform: Transform {
                d: scale,
                ..Transform::IDENTITY
            },
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: value.group.children,
        },
    }
}

fn horizontally_scaled(value: MeasuredMath, scale: f64) -> MeasuredMath {
    if scale <= 0.0 || (scale - 1.0).abs() <= f64::EPSILON {
        return value;
    }
    MeasuredMath {
        width: value.width * scale,
        ascent: value.ascent,
        descent: value.descent,
        group: GroupElement {
            transform: Transform {
                a: scale,
                ..Transform::IDENTITY
            },
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: value.group.children,
        },
    }
}

fn layout_sequence(values: Vec<MeasuredMath>) -> MeasuredMath {
    layout_spaced_sequence(values, 0.0)
}

fn layout_spaced_sequence(values: Vec<MeasuredMath>, spacing: f64) -> MeasuredMath {
    if values.is_empty() {
        return MeasuredMath::empty();
    }
    let ascent = values.iter().map(|value| value.ascent).fold(0.0, f64::max);
    let descent = values.iter().map(|value| value.descent).fold(0.0, f64::max);
    let width = values.iter().map(|value| value.width).sum::<f64>()
        + spacing.max(0.0) * values.len().saturating_sub(1) as f64;
    let mut x = 0.0;
    let mut children = Vec::with_capacity(values.len());
    for value in values {
        children.push(translated(value.group, x, ascent - value.ascent));
        x += value.width + spacing.max(0.0);
    }
    MeasuredMath {
        width,
        ascent,
        descent,
        group: group(children),
    }
}

fn group(children: Vec<PositionedElement>) -> GroupElement {
    GroupElement {
        transform: Transform::IDENTITY,
        clip: None,
        opacity: 1.0,
        effects: Vec::new(),
        children,
    }
}

fn translated(mut group: GroupElement, x: f64, y: f64) -> PositionedElement {
    group.transform = group.transform.then(Transform {
        e: x,
        f: y,
        ..Transform::IDENTITY
    });
    PositionedElement::Group(group)
}

fn finite_positive(value: f64) -> Option<f64> {
    (value.is_finite() && value > 0.0).then_some(value)
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;

    use oxml_layout::{InlineItem, LineBreakParams, LineItem, break_into_lines, walk};
    use rdocx_oxml::document::CT_Document;
    use rdocx_oxml::math::{
        CT_OMath, MathMatrixRow, MathNary, MathPreSubSuperscript, MathSubSuperscript,
    };
    use rdocx_oxml::styles::CT_Styles;

    use super::*;

    type Geometry = (f64, f64, f64);

    fn text(value: &str) -> MathArgument {
        MathArgument::text(value)
    }

    fn supported_expressions() -> Vec<MathExpression> {
        let mut nary = MathNary::new("∑", text("x"));
        nary.hide_subscript = false;
        nary.hide_superscript = false;
        nary.subscript = text("i=0");
        nary.superscript = text("n");
        nary.limit_location = Some(LimitLocation::UnderOver);
        vec![
            MathRun::new("x").into(),
            MathExpression::Fraction(MathFraction::new(text("1"), text("2"))),
            MathExpression::Subscript(MathScript::new(text("x"), text("i"))),
            MathExpression::Superscript(MathScript::new(text("x"), text("2"))),
            MathExpression::SubSuperscript(MathSubSuperscript::new(
                text("x"),
                text("i"),
                text("2"),
            )),
            MathExpression::PreSubSuperscript(MathPreSubSuperscript::new(
                text("X"),
                text("a"),
                text("b"),
            )),
            MathExpression::Radical(MathRadical::with_degree(text("3"), text("x"))),
            MathExpression::Matrix(MathMatrix::new(vec![
                MathMatrixRow::new(vec![text("a"), text("b")]),
                MathMatrixRow::new(vec![text("c"), text("d")]),
            ])),
            MathExpression::LowerLimit(MathLimit::new(text("lim"), text("0"))),
            MathExpression::UpperLimit(MathLimit::new(text("max"), text("n"))),
            MathExpression::Nary(nary),
            MathExpression::Delimiter(MathDelimiter::new("(", ")", vec![text("x"), text("y")])),
            MathExpression::Accent(MathAccent::new("̂", text("x"))),
        ]
    }

    fn deterministic_font_manager() -> FontManager {
        FontManager::new_deterministic().expect("bundled fonts")
    }

    fn layout_input(document: CT_Document) -> crate::LayoutInput {
        crate::LayoutInput {
            document,
            automatic_hyphenation: false,
            math_properties: None,
            revision_view: crate::RevisionView::Accepted,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images: HashMap::new(),
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: None,
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
        }
    }

    #[test]
    fn supported_math_expressions_measure_and_lower_to_bounded_groups() {
        let mut fm = deterministic_font_manager();
        let mut diagnostics = Vec::new();
        for (index, expression) in supported_expressions().iter().enumerate() {
            let measured = layout_expression(
                expression,
                &mut fm,
                None,
                11.0,
                Color::BLACK,
                &format!("test/{index}"),
                &mut diagnostics,
            )
            .expect("supported expression layout");
            assert!(measured.is_bounded(), "expression {index}: {measured:?}");
            assert!(measured.width > 0.0, "expression {index} has no width");
            assert!(measured.height() > 0.0, "expression {index} has no height");
            assert!(!measured.group.children.is_empty(), "expression {index}");
        }
    }

    #[test]
    fn equation_descent_does_not_move_the_surrounding_text_baseline() {
        let mut fm = deterministic_font_manager();
        let mut diagnostics = Vec::new();
        let ordinary = layout_text(
            "before",
            &mut fm,
            None,
            11.0,
            Color::BLACK,
            false,
            false,
            "text",
            &mut diagnostics,
        )
        .expect("text layout");
        let subscript = layout_expression(
            &MathExpression::Subscript(MathScript::new(text("x"), text("deep"))),
            &mut fm,
            None,
            11.0,
            Color::BLACK,
            "equation",
            &mut diagnostics,
        )
        .expect("subscript layout");
        assert!(subscript.descent > ordinary.descent);
        let lines = break_into_lines(
            &[
                InlineItem::Text(oxml_layout::TextSegment {
                    text: "before".to_owned(),
                    direction: oxml_layout::TextDirection::Auto,
                    source: None,
                    font_id: match &ordinary.group.children[0] {
                        PositionedElement::Text(run) => run.font_id,
                        _ => panic!("text glyph run"),
                    },
                    font_size: 11.0,
                    glyph_ids: vec![1],
                    advances: vec![ordinary.width],
                    width: ordinary.width,
                    ascent: ordinary.ascent,
                    descent: ordinary.descent,
                    line_gap: 0.0,
                    color: Color::BLACK,
                    bold: false,
                    italic: false,
                    underline: None,
                    strike: false,
                    dstrike: false,
                    highlight: None,
                    baseline_offset: 0.0,
                    hyperlink_url: None,
                    field_kind: None,
                    note: None,
                }),
                InlineItem::Group {
                    width: subscript.width,
                    height: subscript.height(),
                    baseline: Some(subscript.ascent),
                    group: subscript.group,
                },
            ],
            &LineBreakParams {
                available_width: 300.0,
                ..Default::default()
            },
            &fm,
        )
        .expect("line layout");
        assert!((lines[0].ascent - ordinary.ascent.max(subscript.ascent)).abs() < 1.0e-9);
        assert!((lines[0].descent - subscript.descent).abs() < 1.0e-9);
        let LineItem::Text(adjacent) = &lines[0].items[0] else {
            panic!("adjacent text");
        };
        assert_eq!(adjacent.baseline_offset, 0.0);
    }

    #[test]
    fn unsupported_math_content_is_diagnostic_and_remains_preserved() {
        let namespace = rdocx_oxml::namespace::M_NS;
        let source = format!(
            r#"<m:oMath xmlns:m="{namespace}" xmlns:x="urn:producer"><m:r><m:t>x</m:t></m:r><x:visible data="kept">raw</x:visible></m:oMath>"#
        );
        let equation = CT_OMath::from_xml(source.as_bytes()).expect("math parses");
        let original = equation.to_xml().expect("math saves");
        let mut fm = deterministic_font_manager();
        let mut diagnostics = Vec::new();
        let (measured, display) = layout_officemath(
            &OfficeMath::Inline(equation.clone()),
            &mut fm,
            11.0,
            Color::BLACK,
            300.0,
            None,
            "paragraph/run-boundary/0/raw-child/0",
            &mut diagnostics,
        )
        .expect("math layout");
        assert!(!display);
        assert!(measured.width > 0.0);
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(
            diagnostics[0].message,
            "OfficeMath content at paragraph/run-boundary/0/raw-child/0 was preserved but could not be rendered"
        );
        assert_eq!(equation.to_xml().expect("math resaves"), original);
        assert!(
            std::str::from_utf8(&original)
                .expect("UTF-8 XML")
                .contains(r#"<x:visible data="kept">raw</x:visible>"#)
        );
    }

    #[test]
    fn document_math_properties_control_display_geometry_font_and_nary_limits() {
        let mut nary = MathNary::new("∑", text("x"));
        nary.hide_subscript = false;
        nary.hide_superscript = false;
        nary.subscript = text("i=0");
        nary.superscript = text("n");
        let equation = OfficeMath::display(vec![CT_OMath::new(vec![MathExpression::Nary(nary)])]);
        let mut properties = MathProperties::new();
        properties.math_font = Some("Caladea".to_owned());
        properties.justification = Some(MathJustification::Right);
        properties.left_margin = Some(200);
        properties.right_margin = Some(400);
        properties.pre_spacing = Some(80);
        properties.post_spacing = Some(120);
        properties.intra_spacing = Some(60);
        properties.nary_limit_location = Some(LimitLocation::UnderOver);

        let mut fm = deterministic_font_manager();
        let (measured, display) = layout_officemath(
            &equation,
            &mut fm,
            11.0,
            Color::BLACK,
            300.0,
            Some(&properties),
            "display",
            &mut Vec::new(),
        )
        .expect("configured display layout");
        assert!(display);
        assert_eq!(measured.width, 300.0);
        assert!(measured.ascent > 27.0);
        assert!(measured.descent > 18.0);
        let PositionedElement::Group(content) = &measured.group.children[0] else {
            panic!("display content group");
        };
        assert!(content.transform.e > 10.0);
        assert_eq!(content.transform.f, 4.0);
        let expected_font = fm
            .resolve_font_for_text(Some("Caladea"), false, false, "∑")
            .expect("Caladea font");
        let mut actual_fonts = Vec::new();
        walk(&measured.group.children, &mut |element, _| {
            if let PositionedElement::Text(run) = element {
                actual_fonts.push(run.font_id);
            }
        });
        assert!(actual_fonts.contains(&expected_font));
    }

    #[test]
    fn matrix_baselines_and_delimiter_growth_follow_typed_properties() {
        let rows = vec![
            MathMatrixRow::new(vec![text("top")]),
            MathMatrixRow::new(vec![text("bottom")]),
        ];
        let mut top = MathMatrix::new(rows.clone());
        top.properties.base_justification = Some(MatrixBaseJustification::Top);
        let mut bottom = MathMatrix::new(rows);
        bottom.properties.base_justification = Some(MatrixBaseJustification::Bottom);
        let mut fm = deterministic_font_manager();
        let top = layout_matrix(
            &top,
            &mut fm,
            None,
            11.0,
            Color::BLACK,
            "matrix/top",
            &mut Vec::new(),
        )
        .expect("top matrix");
        let bottom = layout_matrix(
            &bottom,
            &mut fm,
            None,
            11.0,
            Color::BLACK,
            "matrix/bottom",
            &mut Vec::new(),
        )
        .expect("bottom matrix");
        assert!(top.ascent < bottom.ascent);
        assert!(top.descent > bottom.descent);
        assert_eq!(top.height(), bottom.height());

        let fraction = MathArgument::new(vec![MathExpression::Fraction(MathFraction::new(
            text("1"),
            text("2"),
        ))]);
        let mut fixed = MathDelimiter::new("(", ")", vec![fraction.clone()]);
        fixed.grow = Some(false);
        let mut growing = MathDelimiter::new("(", ")", vec![fraction]);
        growing.grow = Some(true);
        let fixed = layout_delimiter(
            &fixed,
            &mut fm,
            None,
            11.0,
            Color::BLACK,
            "delimiter/fixed",
            &mut Vec::new(),
        )
        .expect("fixed delimiter");
        let growing = layout_delimiter(
            &growing,
            &mut fm,
            None,
            11.0,
            Color::BLACK,
            "delimiter/growing",
            &mut Vec::new(),
        )
        .expect("growing delimiter");
        let delimiter_scale = |value: &MeasuredMath| {
            let PositionedElement::Group(delimiter) = &value.group.children[0] else {
                panic!("delimiter group");
            };
            delimiter.transform.d
        };
        assert_eq!(delimiter_scale(&fixed), 1.0);
        assert!(delimiter_scale(&growing) > 1.0);
    }

    #[test]
    fn inline_and_display_equations_share_the_document_page_frame_and_pdf_backends() {
        let word = rdocx_oxml::namespace::W_NS;
        let math = rdocx_oxml::namespace::M_NS;
        let xml = format!(
            r#"<w:document xmlns:w="{word}" xmlns:m="{math}"><w:body><w:p><w:r><w:t>before</w:t></w:r><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath><w:r><w:t>after</w:t></w:r></w:p><w:p><m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr><m:oMath><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara></w:p><w:sectPr/></w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let result =
            crate::layout_document_deterministic(&layout_input(document)).expect("document layout");
        assert_eq!(result.pages.len(), 1);
        let mut tokens = Vec::new();
        walk(&result.pages[0].elements, &mut |element, _| {
            if let PositionedElement::Text(run) = element {
                tokens.push(run.text.clone());
            }
        });
        assert!(tokens.iter().any(|token| token == "before"));
        assert!(tokens.iter().any(|token| token == "x"));
        assert!(tokens.iter().any(|token| token == "after"));
        assert!(tokens.iter().any(|token| token == "1"));
        assert!(tokens.iter().any(|token| token == "2"));
        let token_position = |expected: &str| {
            tokens
                .iter()
                .position(|token| token == expected)
                .unwrap_or_else(|| panic!("missing {expected}"))
        };
        assert!(token_position("before") < token_position("x"));
        assert!(token_position("x") < token_position("after"));
        assert!(token_position("after") < token_position("1"));
        fn contains_group(element: &PositionedElement) -> bool {
            match element {
                PositionedElement::Group(_) => true,
                PositionedElement::MarkedContent { children, .. } => {
                    children.iter().any(contains_group)
                }
                _ => false,
            }
        }
        assert!(result.pages[0].elements.iter().any(contains_group));
    }

    #[test]
    fn officemath_baselines_and_glyph_geometry_match_the_pinned_word_pdf_oracle() {
        const TOLERANCE: f64 = 1.0;
        const EXPECTED: [(f64, f64, f64); 13] = [
            (5.43, 10.45, 2.44),
            (7.71, 19.24, 8.93),
            (8.39, 10.45, 3.69),
            (10.57, 12.04, 2.44),
            (10.57, 12.04, 3.69),
            (11.71, 12.04, 3.69),
            (13.88, 12.38, 2.81),
            (24.37, 16.90, 8.89),
            (16.18, 10.45, 8.28),
            (21.47, 16.28, 2.44),
            (21.62, 24.78, 14.77),
            (24.07, 11.29, 2.64),
            (5.43, 13.33, 5.68),
        ];
        let (combined, actual) = golden_geometry();
        assert!(
            geometry_matches((195.50, 24.78, 14.77), combined, TOLERANCE),
            "combined geometry {combined:?}"
        );
        assert_eq!(actual.len(), EXPECTED.len());
        for (index, (expected, actual)) in EXPECTED.into_iter().zip(actual).enumerate() {
            assert!(
                geometry_matches(expected, actual, TOLERANCE),
                "expression {index}: expected {expected:?}, got {actual:?}"
            );
        }

        let manifest = include_str!("../../../scripts/officemath_oracle_manifest.json");
        for binding in [
            r#""word_identity": "Microsoft Word 16.104 build 16.104.25121423""#,
            r#""poppler_identity": "pdftoppm version 26.01.0""#,
            r#""dpi": 150"#,
            r#""geometry_tolerance_pt": 1.0"#,
            r#""source_sha256": "961f6c30011e922d334487e6f94be5f0021e1586be21df2a03c3f4783b7a795f""#,
            r#""page_size_pt": [612.0, 792.0]"#,
            r#""word_pdf_math_bbox_pt": [208.25, 93.28, 403.44, 132.80]"#,
            r#""word_expression_glyph_indices""#,
            r#""rust_expression_glyph_indices""#,
            r#""raster_expression_windows_pt""#,
        ] {
            assert!(
                manifest.contains(binding),
                "missing oracle binding {binding}"
            );
        }
        assert!(!manifest.contains("word_normalized_geometry_pt"));
    }

    #[test]
    fn officemath_golden_rejects_baseline_delimiter_and_operator_mutations() {
        const TOLERANCE: f64 = 1.0;
        const PERTURBATION: f64 = 1.01;
        let (combined, mut expressions) = golden_geometry();
        let baseline_mutation = (combined.0, combined.1 + PERTURBATION, combined.2);
        assert!(!geometry_matches(combined, baseline_mutation, TOLERANCE));

        let delimiter = expressions[11];
        expressions[11].0 += PERTURBATION;
        assert!(!geometry_matches(delimiter, expressions[11], TOLERANCE));

        let operator = expressions[10];
        expressions[10].1 += PERTURBATION;
        assert!(!geometry_matches(operator, expressions[10], TOLERANCE));
    }

    fn golden_geometry() -> (Geometry, Vec<Geometry>) {
        let mut fm = deterministic_font_manager();
        let expressions = supported_expressions();
        let measured = layout_expressions(
            &expressions,
            &mut fm,
            None,
            11.0,
            Color::BLACK,
            11.0 * EQUATION_SPACING_EM,
            "golden",
            &mut Vec::new(),
        )
        .expect("golden layout");
        let parts = expressions
            .iter()
            .map(|expression| {
                let value = layout_expression(
                    expression,
                    &mut fm,
                    None,
                    11.0,
                    Color::BLACK,
                    "golden",
                    &mut Vec::new(),
                )
                .expect("expression layout");
                (value.width, value.ascent, value.descent)
            })
            .collect();
        ((measured.width, measured.ascent, measured.descent), parts)
    }

    fn geometry_matches(expected: Geometry, actual: Geometry, tolerance: f64) -> bool {
        (expected.0 - actual.0).abs() <= tolerance
            && (expected.1 - actual.1).abs() <= tolerance
            && (expected.2 - actual.2).abs() <= tolerance
    }
}
