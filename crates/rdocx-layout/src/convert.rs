//! Conversion from WordprocessingML flow values to shared layout values.

use oxml_layout::{
    Align, LayoutLine, LineBreakParams, LineSpacing, TabAlign, TabLeader, TabStop, TextDirection,
    Underline,
};
use rdocx_oxml::borders::CT_TabStop;
use rdocx_oxml::properties::CT_PPr;
use rdocx_oxml::shared::{ST_Jc, ST_TabJc, ST_TabLeader, ST_Underline};

pub(crate) fn alignment(value: Option<ST_Jc>) -> Option<Align> {
    value.map(|value| match value {
        ST_Jc::Start | ST_Jc::Left => Align::Start,
        ST_Jc::Center => Align::Center,
        ST_Jc::End | ST_Jc::Right => Align::End,
        ST_Jc::Both => Align::Justify,
        ST_Jc::Distribute => Align::Distribute,
    })
}

pub(crate) fn alignment_for_direction(
    value: Option<ST_Jc>,
    direction: TextDirection,
) -> Option<Align> {
    value.map(|value| match (value, direction) {
        (ST_Jc::Start, TextDirection::RightToLeft) => Align::End,
        (ST_Jc::End, TextDirection::RightToLeft) => Align::Start,
        (ST_Jc::Start | ST_Jc::Left, _) => Align::Start,
        (ST_Jc::End | ST_Jc::Right, _) => Align::End,
        (ST_Jc::Center, _) => Align::Center,
        (ST_Jc::Both, _) => Align::Justify,
        (ST_Jc::Distribute, _) => Align::Distribute,
    })
}

pub(crate) fn tab_stops(values: &[CT_TabStop]) -> Vec<TabStop> {
    values
        .iter()
        .filter_map(|value| {
            let align = match value.val {
                ST_TabJc::Left => TabAlign::Left,
                ST_TabJc::Center => TabAlign::Center,
                ST_TabJc::Right => TabAlign::Right,
                ST_TabJc::Decimal | ST_TabJc::Num => TabAlign::Decimal,
                ST_TabJc::Bar => TabAlign::Bar,
                ST_TabJc::Clear => return None,
            };
            let leader = value.leader.map(|leader| match leader {
                ST_TabLeader::None => TabLeader::None,
                ST_TabLeader::Dot => TabLeader::Dot,
                ST_TabLeader::Hyphen => TabLeader::Hyphen,
                ST_TabLeader::Underscore => TabLeader::Underscore,
                ST_TabLeader::Heavy => TabLeader::Heavy,
                ST_TabLeader::MiddleDot => TabLeader::MiddleDot,
            });
            Some(TabStop {
                pos_pt: value.pos.to_pt(),
                align,
                leader,
            })
        })
        .collect()
}

pub(crate) fn underline(value: Option<ST_Underline>) -> Option<Underline> {
    value.and_then(|value| match value {
        ST_Underline::None => None,
        ST_Underline::Single => Some(Underline::Single),
        ST_Underline::Words => Some(Underline::Words),
        ST_Underline::Double => Some(Underline::Double),
        ST_Underline::Thick => Some(Underline::Thick),
        ST_Underline::Dotted => Some(Underline::Dotted),
        ST_Underline::Dash => Some(Underline::Dash),
        ST_Underline::DotDash => Some(Underline::DotDash),
        ST_Underline::DotDotDash => Some(Underline::DotDotDash),
        ST_Underline::Wave => Some(Underline::Wave),
    })
}

pub(crate) fn line_spacing(properties: &CT_PPr) -> LineSpacing {
    match (properties.line_spacing, properties.line_rule.as_deref()) {
        (Some(spacing), Some("exact")) => LineSpacing::Exact(spacing.to_pt()),
        (Some(spacing), Some("atLeast")) => LineSpacing::AtLeast(spacing.to_pt()),
        (Some(spacing), _) => LineSpacing::Multiple(spacing.0 as f64 / 240.0),
        (None, _) => LineSpacing::Single,
    }
}

pub(crate) fn line_break_params(properties: &CT_PPr, available_width: f64) -> LineBreakParams {
    LineBreakParams {
        line_prefix_widths: Vec::new(),
        line_suffix_widths: Vec::new(),
        available_width,
        ind_left: properties.ind_left.map_or(0.0, |value| value.to_pt()),
        ind_right: properties.ind_right.map_or(0.0, |value| value.to_pt()),
        ind_first_line: properties.ind_first_line.map_or(0.0, |value| value.to_pt()),
        ind_hanging: properties.ind_hanging.map_or(0.0, |value| value.to_pt()),
        tab_stops: properties
            .tabs
            .as_ref()
            .map_or_else(Vec::new, |tabs| tab_stops(&tabs.tabs)),
        line_spacing: line_spacing(properties),
        jc: alignment(properties.jc),
        wrap: true,
    }
}

pub(crate) fn restore_word_line_heights(lines: &mut [LayoutLine], properties: &CT_PPr) {
    for line in lines {
        let natural = line.ascent + line.descent;
        let natural = if natural < 1.0 { 12.0 } else { natural };
        line.line_gap = 0.0;
        line.height = match (properties.line_spacing, properties.line_rule.as_deref()) {
            (Some(spacing), Some("exact")) => spacing.to_pt(),
            (Some(spacing), Some("atLeast")) => natural.max(spacing.to_pt()),
            (Some(spacing), _) => natural * spacing.0 as f64 / 240.0,
            (None, _) => natural,
        };
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use oxml_layout::{Align, LineSpacing, TabAlign, TabLeader, Underline};
    use rdocx_oxml::units::Twips;

    #[test]
    fn every_word_alignment_maps_to_the_shared_alignment() {
        let cases = [
            (ST_Jc::Start, Align::Start),
            (ST_Jc::Left, Align::Start),
            (ST_Jc::Center, Align::Center),
            (ST_Jc::End, Align::End),
            (ST_Jc::Right, Align::End),
            (ST_Jc::Both, Align::Justify),
            (ST_Jc::Distribute, Align::Distribute),
        ];
        for (word, shared) in cases {
            assert_eq!(alignment(Some(word)), Some(shared));
        }
        assert_eq!(alignment(None), None);
    }

    #[test]
    fn logical_alignment_uses_the_paragraph_base_direction() {
        assert_eq!(
            alignment_for_direction(Some(ST_Jc::Start), TextDirection::RightToLeft),
            Some(Align::End)
        );
        assert_eq!(
            alignment_for_direction(Some(ST_Jc::End), TextDirection::RightToLeft),
            Some(Align::Start)
        );
        assert_eq!(
            alignment_for_direction(Some(ST_Jc::Left), TextDirection::RightToLeft),
            Some(Align::Start)
        );
        assert_eq!(
            alignment_for_direction(Some(ST_Jc::Right), TextDirection::LeftToRight),
            Some(Align::End)
        );
    }

    #[test]
    fn every_word_tab_and_leader_maps_without_rounding_its_position() {
        let cases = [
            (ST_TabJc::Left, TabAlign::Left),
            (ST_TabJc::Center, TabAlign::Center),
            (ST_TabJc::Right, TabAlign::Right),
            (ST_TabJc::Decimal, TabAlign::Decimal),
            (ST_TabJc::Bar, TabAlign::Bar),
            (ST_TabJc::Num, TabAlign::Decimal),
        ];
        let leaders = [
            (ST_TabLeader::None, TabLeader::None),
            (ST_TabLeader::Dot, TabLeader::Dot),
            (ST_TabLeader::Hyphen, TabLeader::Hyphen),
            (ST_TabLeader::Underscore, TabLeader::Underscore),
            (ST_TabLeader::Heavy, TabLeader::Heavy),
            (ST_TabLeader::MiddleDot, TabLeader::MiddleDot),
        ];

        for ((word_align, shared_align), (word_leader, shared_leader)) in
            cases.into_iter().zip(leaders)
        {
            let converted = tab_stops(&[CT_TabStop {
                val: word_align,
                pos: Twips(1),
                leader: Some(word_leader),
                source_occurrence: None,
            }]);
            assert_eq!(converted.len(), 1);
            assert_eq!(converted[0].align, shared_align);
            assert_eq!(converted[0].leader, Some(shared_leader));
            assert!((converted[0].pos_pt - 0.05).abs() < f64::EPSILON);
        }

        assert!(tab_stops(&[CT_TabStop::new(ST_TabJc::Clear, Twips(720))]).is_empty());
    }

    #[test]
    fn every_rendered_word_underline_maps_to_the_shared_style() {
        let cases = [
            (ST_Underline::Single, Underline::Single),
            (ST_Underline::Words, Underline::Words),
            (ST_Underline::Double, Underline::Double),
            (ST_Underline::Thick, Underline::Thick),
            (ST_Underline::Dotted, Underline::Dotted),
            (ST_Underline::Dash, Underline::Dash),
            (ST_Underline::DotDash, Underline::DotDash),
            (ST_Underline::DotDotDash, Underline::DotDotDash),
            (ST_Underline::Wave, Underline::Wave),
        ];
        for (word, shared) in cases {
            assert_eq!(underline(Some(word)), Some(shared));
        }
        assert_eq!(underline(Some(ST_Underline::None)), None);
        assert_eq!(underline(None), None);
    }

    #[test]
    fn every_word_line_spacing_rule_maps_exactly() {
        let cases = [
            (Some(Twips(1)), Some("exact"), LineSpacing::Exact(0.05)),
            (
                Some(Twips(480)),
                Some("atLeast"),
                LineSpacing::AtLeast(24.0),
            ),
            (Some(Twips(360)), Some("auto"), LineSpacing::Multiple(1.5)),
            (Some(Twips(240)), None, LineSpacing::Multiple(1.0)),
            (None, None, LineSpacing::Single),
        ];
        for (spacing, rule, expected) in cases {
            let properties = CT_PPr {
                line_spacing: spacing,
                line_rule: rule.map(str::to_owned),
                ..Default::default()
            };
            assert_eq!(line_spacing(&properties), expected);
        }
    }

    #[test]
    fn word_line_parameters_keep_wrap_enabled() {
        let params = line_break_params(&CT_PPr::default(), 321.0);
        assert_eq!(params.available_width, 321.0);
        assert!(params.wrap);
    }

    #[test]
    fn word_auto_spacing_uses_glyph_height_not_point_size() {
        let mut lines = vec![LayoutLine {
            items: Vec::new(),
            width: 0.0,
            ascent: 10.0,
            descent: 3.0,
            line_gap: 2.0,
            height: 12.0,
            indent_left: 0.0,
            available_width: 468.0,
            is_last: true,
            page_break_after: false,
        }];
        let properties = CT_PPr {
            line_spacing: Some(Twips(480)),
            line_rule: Some("auto".to_string()),
            ..Default::default()
        };
        restore_word_line_heights(&mut lines, &properties);
        assert_eq!(lines[0].height, 26.0);
        assert_eq!(lines[0].line_gap, 0.0);
    }
}
