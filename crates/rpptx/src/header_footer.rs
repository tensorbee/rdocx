//! Slide date, footer and slide-number placeholders, and text fields.
//!
//! PowerPoint draws a date, footer or slide number only through a placeholder
//! the slide owns. The layout and master copies are templates, and their
//! `p:hf` flags decide which of them a new slide receives. Google Slides reads
//! only the slide-owned placeholders too, so every call here writes them.

use oxml_drawing::text::{CT_TextBody, CT_TextField, CT_TextParagraph, TextRun};
use rpptx_oxml::placeholder::{CT_Placeholder, PhType};
use rpptx_oxml::shape_tree::{CT_Shape, ShapeIdAllocator, ShapeTreeChild};
use rpptx_oxml::slide_parts::{CT_HeaderFooter, CT_Slide, CT_SlideLayout, CT_SlideMaster};

use crate::{
    Error, Presentation, Result, invalid_slide_mutation, rel_types, related_internal_part,
    required_part,
};

/// The date, footer and slide number of slides, as PowerPoint's Header and
/// Footer dialog sets them.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct HeaderFooter {
    /// Whether the slide shows its number.
    pub slide_number: bool,
    /// The footer text, or `None` for no footer.
    pub footer: Option<String>,
    /// The date shown on the slide.
    pub date: HeaderFooterDate,
}

/// The date a slide shows in its date placeholder.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub enum HeaderFooterDate {
    /// No date placeholder.
    #[default]
    Off,
    /// Fixed text, written as a run.
    Fixed(String),
    /// A `datetime1` to `datetime13` field that PowerPoint refreshes when it
    /// opens the file. `text` is the cached value other readers show and
    /// rpptx renders, see [`date_field_text`].
    Automatic { field_type: String, text: String },
}

/// Which date, footer and slide-number placeholders a master's or layout's
/// `p:hf` gives new slides. An attribute the element omits is enabled.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct HeaderFooterFlags {
    pub date: bool,
    pub footer: bool,
    pub slide_number: bool,
}

impl From<&CT_HeaderFooter> for HeaderFooterFlags {
    fn from(header_footer: &CT_HeaderFooter) -> Self {
        Self {
            date: header_footer.date_time_enabled(),
            footer: header_footer.footer_enabled(),
            slide_number: header_footer.slide_number_enabled(),
        }
    }
}

/// The kinds of `a:fld` that [`crate::TextParagraphMut::add_field`] writes.
const FIELD_TYPES: [&str; 15] = [
    "slidenum",
    "datetime",
    "datetime1",
    "datetime2",
    "datetime3",
    "datetime4",
    "datetime5",
    "datetime6",
    "datetime7",
    "datetime8",
    "datetime9",
    "datetime10",
    "datetime11",
    "datetime12",
    "datetime13",
];

/// The text PowerPoint caches for a slide-number field before it numbers it.
pub const SLIDE_NUMBER_FIELD_TEXT: &str = "\u{2039}#\u{203a}";

const MONTHS: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];

const WEEKDAYS: [&str; 7] = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
];

/// Formats a calendar date as PowerPoint's English (United States) date
/// fields `datetime1` to `datetime7` show it, for use as cached field text.
///
/// `datetime` reads as `datetime1`. The time formats `datetime8` to
/// `datetime13` need a time of day and are refused, as is an invalid date.
pub fn date_field_text(field_type: &str, year: i32, month: u32, day: u32) -> Result<String> {
    let operation = "format a date field";
    let days_in_month = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) => 29,
        2 => 28,
        _ => 0,
    };
    if day == 0 || day > days_in_month {
        return Err(invalid_slide_mutation(
            operation,
            format!("{year:04}-{month:02}-{day:02} is not a calendar date"),
        ));
    }
    let month_name = MONTHS[month as usize - 1];
    let short_month = &month_name[..3];
    let short_year = year.rem_euclid(100);
    Ok(match field_type {
        "datetime" | "datetime1" => format!("{month}/{day}/{year}"),
        "datetime2" => format!(
            "{}, {month_name} {day}, {year}",
            WEEKDAYS[weekday_index(year, month, day)]
        ),
        "datetime3" => format!("{day} {month_name} {year}"),
        "datetime4" => format!("{month_name} {day}, {year}"),
        "datetime5" => format!("{day}-{short_month}-{short_year:02}"),
        "datetime6" => format!("{month_name} {short_year:02}"),
        "datetime7" => format!("{short_month}-{short_year:02}"),
        other => {
            return Err(invalid_slide_mutation(
                operation,
                format!(
                    "date field type {other:?} is not a date format, use datetime1 to datetime7"
                ),
            ));
        }
    })
}

/// Monday is 0, by the days-from-civil count of 1970-01-01, a Thursday.
fn weekday_index(year: i32, month: u32, day: u32) -> usize {
    let year = i64::from(year) - i64::from(month <= 2);
    let era = year.div_euclid(400);
    let year_of_era = year - era * 400;
    let month = i64::from(month);
    let day_of_year = (153 * (month + if month > 2 { -3 } else { 9 }) + 2) / 5 + i64::from(day) - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    let days = era * 146_097 + day_of_era - 719_468;
    (days + 3).rem_euclid(7) as usize
}

/// Builds one `a:fld` of a validated type with its cached text.
pub(crate) fn text_field(field_type: &str, text: &str) -> Result<CT_TextField> {
    if !FIELD_TYPES.contains(&field_type) {
        return Err(invalid_slide_mutation(
            "add a field",
            format!(
                "field type {field_type:?} is not supported, use slidenum or datetime1 to datetime13"
            ),
        ));
    }
    // PowerPoint reuses one field GUID across slides, so a fixed GUID per
    // kind is what its own decks carry.
    let id = if field_type == "slidenum" {
        "{8C1E7A2B-5F3D-4B6E-9A40-1D2C3B4A5F01}"
    } else {
        "{8C1E7A2B-5F3D-4B6E-9A40-1D2C3B4A5F02}"
    };
    let escaped = text
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;");
    CT_TextField::from_xml(
        format!(
            "<a:fld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" id=\"{id}\" type=\"{field_type}\"><a:t>{escaped}</a:t></a:fld>"
        )
        .as_bytes(),
    )
    .map_err(|error| invalid_slide_mutation("add a field", error.to_string()))
}

/// Appends a field after the paragraph's existing runs.
pub(crate) fn push_field(paragraph: &mut CT_TextParagraph, field: CT_TextField) {
    // `add_run` keeps preserved paragraph children in place, then the new
    // regular run becomes the field.
    paragraph.add_run("");
    if let Some(last) = paragraph.runs.last_mut() {
        *last = TextRun::Field(Box::new(field));
    }
}

/// The three latent placeholder kinds, in PowerPoint's slide order.
const LATENT_KINDS: [PhType; 3] = [PhType::DateTime, PhType::Footer, PhType::SlideNumber];

fn latent_name(kind: &PhType) -> &'static str {
    match kind {
        PhType::DateTime => "Date Placeholder",
        PhType::Footer => "Footer Placeholder",
        _ => "Slide Number Placeholder",
    }
}

fn latent_label(kind: &PhType) -> &'static str {
    match kind {
        PhType::DateTime => "date",
        PhType::Footer => "footer",
        _ => "slide number",
    }
}

fn is_kind(shape: &CT_Shape, kind: &PhType) -> bool {
    shape
        .placeholder
        .as_ref()
        .is_some_and(|placeholder| placeholder.ph_type.as_ref() == Some(kind))
}

/// Finds the first shape placeholder of one kind, groups included.
fn find_latent<'a>(children: &'a [ShapeTreeChild], kind: &PhType) -> Option<&'a CT_Shape> {
    children.iter().find_map(|child| match child {
        ShapeTreeChild::Shape(shape) if is_kind(shape, kind) => Some(shape),
        ShapeTreeChild::GroupShape(group) => find_latent(&group.children, kind),
        _ => None,
    })
}

fn find_latent_mut<'a>(
    children: &'a mut [ShapeTreeChild],
    kind: &PhType,
) -> Option<&'a mut CT_Shape> {
    for child in children {
        match child {
            ShapeTreeChild::Shape(shape) if is_kind(shape, kind) => return Some(shape),
            ShapeTreeChild::GroupShape(group) => {
                if let Some(found) = find_latent_mut(&mut group.children, kind) {
                    return Some(found);
                }
            }
            _ => {}
        }
    }
    None
}

fn remove_latent(children: &mut Vec<ShapeTreeChild>, kind: &PhType) {
    children.retain_mut(|child| match child {
        ShapeTreeChild::Shape(shape) => !is_kind(shape, kind),
        ShapeTreeChild::GroupShape(group) => {
            remove_latent(&mut group.children, kind);
            true
        }
        _ => true,
    });
}

fn has_field(body: &CT_TextBody, prefix: &str) -> bool {
    field_type(body).is_some_and(|field_type| field_type.starts_with(prefix))
}

fn field_type(body: &CT_TextBody) -> Option<&str> {
    body.paragraphs().iter().find_map(|paragraph| {
        paragraph.runs.iter().find_map(|run| match run {
            TextRun::Field(field) => field.field_type.as_deref(),
            _ => None,
        })
    })
}

/// One paragraph holding only `field`, formatted as `template`'s first one.
///
/// The field keeps the character properties, `lang` included, of the first
/// run or field it replaces, as PowerPoint does.
fn field_paragraph(template: Option<&CT_TextBody>, mut field: CT_TextField) -> CT_TextParagraph {
    let mut paragraph = template
        .and_then(|body| body.paragraphs().first())
        .cloned()
        .unwrap_or_default();
    let properties = paragraph.runs.iter().find_map(|run| match run {
        TextRun::Run(run) => run.properties.clone(),
        TextRun::Field(field) => field.run_properties.clone(),
        TextRun::Break(_) => None,
    });
    if field.run_properties.is_none() {
        field.run_properties = properties;
    }
    paragraph.set_text("");
    if let Some(first) = paragraph.runs.first_mut() {
        *first = TextRun::Field(Box::new(field));
    }
    paragraph
}

/// What one latent placeholder holds on a slide.
enum LatentContent<'a> {
    SlideNumber,
    Text(&'a str),
    DateField { field_type: &'a str, text: &'a str },
}

/// Rewrites a latent placeholder's paragraphs to hold `content`, keeping the
/// formatting of its first paragraph. A slide number that already holds a
/// `slidenum` field is kept as it is.
fn fill_latent_body(body: &mut CT_TextBody, content: &LatentContent<'_>) -> Result<()> {
    let paragraph = match content {
        LatentContent::SlideNumber if has_field(body, "slidenum") => return Ok(()),
        LatentContent::SlideNumber => {
            field_paragraph(Some(body), text_field("slidenum", SLIDE_NUMBER_FIELD_TEXT)?)
        }
        LatentContent::Text(text) => {
            body.set_text(text);
            return Ok(());
        }
        LatentContent::DateField { field_type, text } => {
            field_paragraph(Some(body), text_field(field_type, text)?)
        }
    };
    body.set_text("");
    if let Some(first) = body.paragraph_mut(0) {
        *first = paragraph;
    }
    Ok(())
}

/// The layout placeholder a new slide copies, else the master one placed by
/// an explicit transform, since a slide placeholder inherits only through
/// its layout.
fn latent_template<'a>(
    layout: &'a CT_SlideLayout,
    master: &'a CT_SlideMaster,
    kind: &PhType,
) -> Option<(&'a CT_Shape, bool)> {
    find_latent(&layout.common_slide_data.shape_tree.children, kind)
        .map(|shape| (shape, false))
        .or_else(|| {
            find_latent(&master.common_slide_data.shape_tree.children, kind)
                .map(|shape| (shape, true))
        })
}

/// A slide-owned copy of a layout or master latent placeholder.
fn latent_copy(
    template: &CT_Shape,
    from_master: bool,
    kind: &PhType,
    ids: &mut ShapeIdAllocator,
) -> Result<CT_Shape> {
    let placeholder = template
        .placeholder
        .clone()
        .unwrap_or_else(|| CT_Placeholder::new(Some(kind.clone()), None));
    let id = ids.allocate();
    let mut shape = CT_Shape::new_placeholder(id, placeholder)
        .and_then(|mut shape| {
            shape.set_name(&format!("{} {}", latent_name(kind), id))?;
            Ok(shape)
        })
        .map_err(|error| invalid_slide_mutation("add a header or footer", error.to_string()))?;
    if from_master {
        // A slide placeholder inherits only through its layout, so a copy of
        // a master placeholder carries the master's geometry, body
        // properties and list style itself, as its text keeps its size and
        // alignment.
        shape.shape_properties = template.shape_properties.clone();
        if template.text_body.is_some() {
            shape.text_body = template.text_body.clone();
        }
    } else if let (Some(body), Some(source)) =
        (shape.text_body.as_mut(), template.text_body.as_ref())
    {
        for (index, paragraph) in source.paragraphs().iter().enumerate() {
            if index == 0 {
                if let Some(first) = body.paragraph_mut(0) {
                    *first = paragraph.clone();
                }
            } else {
                *body.add_paragraph() = paragraph.clone();
            }
        }
    }
    Ok(shape)
}

/// Gives a slide exactly the date, footer and slide-number placeholders
/// `settings` asks for, copying missing ones from its layout or master.
fn apply_to_slide(
    slide: &mut CT_Slide,
    layout: &CT_SlideLayout,
    master: &CT_SlideMaster,
    layout_name: &str,
    settings: &HeaderFooter,
) -> Result<()> {
    let mut ids = ShapeIdAllocator::scan(&slide.common_slide_data.shape_tree);
    for kind in LATENT_KINDS {
        let content = match (&kind, settings) {
            (
                PhType::SlideNumber,
                HeaderFooter {
                    slide_number: true, ..
                },
            ) => Some(LatentContent::SlideNumber),
            (
                PhType::Footer,
                HeaderFooter {
                    footer: Some(text), ..
                },
            ) => Some(LatentContent::Text(text)),
            (PhType::DateTime, HeaderFooter { date, .. }) => match date {
                HeaderFooterDate::Off => None,
                HeaderFooterDate::Fixed(text) => Some(LatentContent::Text(text)),
                HeaderFooterDate::Automatic { field_type, text } => {
                    Some(LatentContent::DateField { field_type, text })
                }
            },
            _ => None,
        };
        let children = &mut slide.common_slide_data.shape_tree.children;
        let Some(content) = content else {
            remove_latent(children, &kind);
            continue;
        };
        if find_latent(children, &kind).is_none() {
            let (template, from_master) =
                latent_template(layout, master, &kind).ok_or_else(|| {
                    invalid_slide_mutation(
                        "set the header and footer",
                        format!(
                            "layout {layout_name:?} and its master have no {} placeholder to copy, add one to the master or layout first",
                            latent_label(&kind)
                        ),
                    )
                })?;
            let copy = latent_copy(template, from_master, &kind, &mut ids)?;
            children.push(ShapeTreeChild::Shape(copy));
        }
        let shape = find_latent_mut(children, &kind).expect("the placeholder exists");
        let body = shape.text_body.get_or_insert_with(CT_TextBody::new);
        fill_latent_body(body, &content)?;
    }
    Ok(())
}

/// Reads what a slide's own latent placeholders show.
fn read_slide(slide: &CT_Slide) -> HeaderFooter {
    let children = &slide.common_slide_data.shape_tree.children;
    let text = |kind: &PhType| {
        find_latent(children, kind).map(|shape| {
            (
                shape
                    .text_body
                    .as_ref()
                    .map(CT_TextBody::plain_text)
                    .unwrap_or_default(),
                shape.text_body.as_ref(),
            )
        })
    };
    let date = match text(&PhType::DateTime) {
        None => HeaderFooterDate::Off,
        Some((text, Some(body))) if has_field(body, "datetime") => HeaderFooterDate::Automatic {
            field_type: field_type(body).unwrap_or("datetime1").to_owned(),
            text,
        },
        Some((text, _)) => HeaderFooterDate::Fixed(text),
    };
    HeaderFooter {
        slide_number: find_latent(children, &PhType::SlideNumber).is_some(),
        footer: text(&PhType::Footer).map(|(text, _)| text),
        date,
    }
}

/// The `p:hf` flags that make new slides receive these placeholders.
fn header_footer_flags(settings: &HeaderFooter) -> CT_HeaderFooter {
    let flag = |enabled: bool| (!enabled).then_some(false);
    CT_HeaderFooter::new(
        flag(settings.slide_number),
        Some(false),
        flag(settings.footer.is_some()),
        flag(settings.date != HeaderFooterDate::Off),
    )
}

/// Updates the text of a template's footer and date placeholders, so a
/// slide added later copies what the deck shows.
fn update_templates(children: &mut [ShapeTreeChild], settings: &HeaderFooter) -> Result<()> {
    if let (Some(text), Some(shape)) = (
        settings.footer.as_deref(),
        find_latent_mut(children, &PhType::Footer),
    ) {
        fill_latent_body(
            shape.text_body.get_or_insert_with(CT_TextBody::new),
            &LatentContent::Text(text),
        )?;
    }
    if let Some(shape) = find_latent_mut(children, &PhType::DateTime) {
        let content = match &settings.date {
            HeaderFooterDate::Off => return Ok(()),
            HeaderFooterDate::Fixed(text) => LatentContent::Text(text),
            HeaderFooterDate::Automatic { field_type, text } => {
                LatentContent::DateField { field_type, text }
            }
        };
        fill_latent_body(
            shape.text_body.get_or_insert_with(CT_TextBody::new),
            &content,
        )?;
    }
    Ok(())
}

fn validate_settings(settings: &HeaderFooter) -> Result<()> {
    if let HeaderFooterDate::Automatic { field_type, .. } = &settings.date
        && !field_type.starts_with("datetime")
    {
        return Err(invalid_slide_mutation(
            "set the header and footer",
            format!(
                "date field type {field_type:?} is not a datetime field, use datetime1 to datetime13"
            ),
        ));
    }
    if let HeaderFooterDate::Automatic { field_type, text } = &settings.date {
        text_field(field_type, text)?;
    }
    Ok(())
}

fn malformed(part_name: &str, error: impl ToString) -> Error {
    Error::MalformedPart {
        part_name: part_name.to_owned(),
        message: error.to_string(),
    }
}

impl Presentation {
    /// Returns a layout's master part and parsed master.
    pub(crate) fn layout_master_part(
        &self,
        layout_index: usize,
    ) -> Result<(String, CT_SlideMaster)> {
        let record = &self.layouts[layout_index];
        let master_part =
            related_internal_part(&self.package, &record.part_name, rel_types::SLIDE_MASTER)?
                .ok_or_else(|| {
                    malformed(&record.part_name, "layout has no slide master relationship")
                })?;
        let master = CT_SlideMaster::from_xml(required_part(&self.package, &master_part)?)
            .map_err(|error| malformed(&master_part, error))?;
        Ok((master_part, master))
    }

    /// The latent placeholders a new slide on this layout receives.
    ///
    /// PowerPoint copies the layout's date, footer and slide-number
    /// placeholders that the layout's `p:hf`, else the master's, enables.
    /// Without either container a new slide receives none.
    pub(crate) fn new_slide_latent_shapes(
        &self,
        layout_index: usize,
        ids: &mut ShapeIdAllocator,
    ) -> Result<Vec<CT_Shape>> {
        let layout = &self.layouts[layout_index].layout;
        // A layout reached without a master relationship has no master flags
        // or placeholders to fall back on.
        let layout_part = &self.layouts[layout_index].part_name;
        let master =
            match related_internal_part(&self.package, layout_part, rel_types::SLIDE_MASTER)? {
                Some(master_part) => Some(
                    CT_SlideMaster::from_xml(required_part(&self.package, &master_part)?)
                        .map_err(|error| malformed(&master_part, error))?,
                ),
                None => None,
            };
        let Some(header_footer) = layout.header_footer.clone().or_else(|| {
            master
                .as_ref()
                .and_then(|master| master.header_footer.clone())
        }) else {
            return Ok(Vec::new());
        };
        let mut shapes = Vec::new();
        for kind in LATENT_KINDS {
            let enabled = match kind {
                PhType::DateTime => header_footer.date_time_enabled(),
                PhType::Footer => header_footer.footer_enabled(),
                _ => header_footer.slide_number_enabled(),
            };
            // As set_header_footer does, a layout without the placeholder
            // falls back on the master's.
            let template = match &master {
                Some(master) => latent_template(layout, master, &kind),
                None => find_latent(&layout.common_slide_data.shape_tree.children, &kind)
                    .map(|shape| (shape, false)),
            };
            if let Some((template, from_master)) = template
                && enabled
            {
                let mut copy = latent_copy(template, from_master, &kind, ids)?;
                if kind == PhType::SlideNumber {
                    let body = copy.text_body.get_or_insert_with(CT_TextBody::new);
                    fill_latent_body(body, &LatentContent::SlideNumber)?;
                }
                shapes.push(copy);
            }
        }
        Ok(shapes)
    }

    /// Returns the date, footer and slide number one slide shows through
    /// the placeholders it owns, or `None` for an unknown slide.
    pub fn slide_header_footer(&self, slide_index: usize) -> Option<HeaderFooter> {
        self.slides
            .get(slide_index)
            .map(|record| read_slide(&record.slide))
    }

    /// Sets one slide's date, footer and slide number.
    ///
    /// Missing placeholders are copied from the slide's layout, or from its
    /// master with an explicit position when the layout has none, and
    /// switched-off ones are removed. The masters and layouts are unchanged.
    /// A layout and master without the placeholder a setting needs is an
    /// error, and the slide is left as it was.
    pub fn set_slide_header_footer(
        &mut self,
        slide_index: usize,
        settings: &HeaderFooter,
    ) -> Result<()> {
        validate_settings(settings)?;
        let slide_count = self.slides.len();
        let layout_index =
            self.slide_layout_index(slide_index)
                .ok_or(Error::UnknownSlideIndex {
                    index: slide_index,
                    slide_count,
                })?;
        let (_, master) = self.layout_master_part(layout_index)?;
        let layout = &self.layouts[layout_index].layout;
        let layout_name = layout.common_slide_data.name.clone().unwrap_or_default();
        let mut slide = self.slides[slide_index].slide.clone();
        apply_to_slide(&mut slide, layout, &master, &layout_name, settings)?;
        self.slides[slide_index].slide = slide;
        Ok(())
    }

    /// Sets the date, footer and slide number of every slide, as
    /// PowerPoint's Header and Footer dialog does with Apply to All.
    ///
    /// Each slide owns the placeholders it shows, copied from its layout
    /// when missing. With `hide_on_title`, slides on a title layout show
    /// none, as "Don't show on title slide" does. Every master and layout
    /// receives the matching `p:hf` flags and the footer and date text, so a
    /// slide added later gets the same placeholders. A layout and master
    /// without a placeholder a slide needs is an error, and nothing changes.
    pub fn set_header_footer(
        &mut self,
        settings: &HeaderFooter,
        hide_on_title: bool,
    ) -> Result<()> {
        validate_settings(settings)?;
        let hidden = HeaderFooter::default();
        let title_layouts = self
            .layouts
            .iter()
            .map(|record| hide_on_title && record.layout.layout_type() == Some("title"))
            .collect::<Vec<_>>();
        let mut staged = self.clone();
        let mut masters = Vec::<(String, CT_SlideMaster)>::new();
        for layout_index in 0..staged.layouts.len() {
            let (master_part, master) = staged.layout_master_part(layout_index)?;
            if !masters.iter().any(|(part, _)| part == &master_part) {
                masters.push((master_part, master));
            }
        }
        for slide_index in 0..staged.slides.len() {
            let layout_index = staged.slide_layout_index(slide_index).ok_or_else(|| {
                malformed(
                    &staged.slides[slide_index].part_name,
                    "slide layout is not reached from a slide master",
                )
            })?;
            let (master_part, _) = staged.layout_master_part(layout_index)?;
            let master = &masters
                .iter()
                .find(|(part, _)| part == &master_part)
                .expect("every layout master was collected")
                .1;
            let layout = &staged.layouts[layout_index].layout;
            let layout_name = layout.common_slide_data.name.clone().unwrap_or_default();
            let effective = if title_layouts[layout_index] {
                &hidden
            } else {
                settings
            };
            let mut slide = staged.slides[slide_index].slide.clone();
            apply_to_slide(&mut slide, layout, master, &layout_name, effective)?;
            staged.slides[slide_index].slide = slide;
        }
        for (layout_index, record) in staged.layouts.iter_mut().enumerate() {
            let effective = if title_layouts[layout_index] {
                &hidden
            } else {
                settings
            };
            record.layout.header_footer = Some(header_footer_flags(effective));
            update_templates(
                &mut record.layout.common_slide_data.shape_tree.children,
                settings,
            )?;
            let xml = record
                .layout
                .to_xml()
                .map_err(|error| malformed(&record.part_name, error))?;
            staged.package.set_part(&record.part_name, xml);
        }
        staged
            .presentation
            .set_show_special_placeholders_on_title_slide(hide_on_title.then_some(false));
        for (master_part, mut master) in masters {
            master.header_footer = Some(header_footer_flags(settings));
            update_templates(&mut master.common_slide_data.shape_tree.children, settings)?;
            let xml = master
                .to_xml()
                .map_err(|error| malformed(&master_part, error))?;
            staged.package.set_part(&master_part, xml);
        }
        self.commit_candidate(staged)
    }
}
