//! Slide transitions: reading any `p:transition` and authoring the common set.
//!
//! The authorable kinds are the ones Google Slides offers and imports: fade,
//! push, wipe, split, cover, uncover, cut and zoom. A transition another
//! producer wrote with any other effect reads back as [`TransitionKind::Other`]
//! and stays untouched until it is replaced.

use rpptx_oxml::timing::{CT_SlideTransition, TransitionEffect};

use crate::{Presentation, Result, SlideMut, SlideRef, invalid_slide_mutation};

/// A transition effect.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TransitionKind {
    Fade,
    Push,
    Wipe,
    Split,
    Cover,
    /// PowerPoint's Uncover, the `p:pull` element.
    Uncover,
    Cut,
    Zoom,
    /// An effect rpptx reads but does not author, by its element name, such
    /// as `blinds` or `morph`.
    Other(String),
}

impl TransitionKind {
    /// The authorable kinds, by the names [`Self::parse`] accepts.
    pub const NAMES: [&'static str; 8] = [
        "fade", "push", "wipe", "split", "cover", "uncover", "cut", "zoom",
    ];

    /// Parses an authorable kind name, `None` for any other name.
    pub fn parse(name: &str) -> Option<Self> {
        Some(match name {
            "fade" => Self::Fade,
            "push" => Self::Push,
            "wipe" => Self::Wipe,
            "split" => Self::Split,
            "cover" => Self::Cover,
            "uncover" => Self::Uncover,
            "cut" => Self::Cut,
            "zoom" => Self::Zoom,
            _ => return None,
        })
    }

    /// The kind's name, the element name for [`Self::Other`].
    pub fn name(&self) -> &str {
        match self {
            Self::Fade => "fade",
            Self::Push => "push",
            Self::Wipe => "wipe",
            Self::Split => "split",
            Self::Cover => "cover",
            Self::Uncover => "uncover",
            Self::Cut => "cut",
            Self::Zoom => "zoom",
            Self::Other(name) => name,
        }
    }

    /// The directions this kind accepts, the first one being the default.
    pub fn directions(&self) -> &'static [TransitionDirection] {
        use TransitionDirection::*;
        match self {
            Self::Push | Self::Wipe => &[Left, Up, Right, Down],
            Self::Cover | Self::Uncover => {
                &[Left, Up, Right, Down, LeftUp, RightUp, LeftDown, RightDown]
            }
            Self::Split => &[HorizontalOut, HorizontalIn, VerticalOut, VerticalIn],
            Self::Zoom => &[Out, In],
            Self::Fade | Self::Cut | Self::Other(_) => &[],
        }
    }

    fn element(&self) -> &str {
        match self {
            Self::Uncover => "pull",
            other => other.name(),
        }
    }
}

/// The direction an effect moves the incoming slide, as the effect's `dir`
/// and, for split, `orient` attributes write it. `Left` is `dir="l"`, which
/// PowerPoint's user interface calls "From Right" for push and wipe.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TransitionDirection {
    Left,
    Up,
    Right,
    Down,
    LeftUp,
    RightUp,
    LeftDown,
    RightDown,
    In,
    Out,
    HorizontalIn,
    HorizontalOut,
    VerticalIn,
    VerticalOut,
}

impl TransitionDirection {
    /// Every direction, by the names [`Self::parse`] accepts.
    pub const NAMES: [&'static str; 14] = [
        "left",
        "up",
        "right",
        "down",
        "left-up",
        "right-up",
        "left-down",
        "right-down",
        "in",
        "out",
        "horizontal-in",
        "horizontal-out",
        "vertical-in",
        "vertical-out",
    ];

    const ALL: [Self; 14] = [
        Self::Left,
        Self::Up,
        Self::Right,
        Self::Down,
        Self::LeftUp,
        Self::RightUp,
        Self::LeftDown,
        Self::RightDown,
        Self::In,
        Self::Out,
        Self::HorizontalIn,
        Self::HorizontalOut,
        Self::VerticalIn,
        Self::VerticalOut,
    ];

    /// Parses a direction name such as `left` or `horizontal-in`.
    pub fn parse(name: &str) -> Option<Self> {
        Self::NAMES
            .iter()
            .position(|candidate| *candidate == name)
            .map(|index| Self::ALL[index])
    }

    /// The direction's name.
    pub fn name(self) -> &'static str {
        let index = Self::ALL
            .iter()
            .position(|candidate| *candidate == self)
            .expect("every direction is listed");
        Self::NAMES[index]
    }

    fn attributes(self) -> &'static [(&'static str, &'static str)] {
        match self {
            Self::Left => &[("dir", "l")],
            Self::Up => &[("dir", "u")],
            Self::Right => &[("dir", "r")],
            Self::Down => &[("dir", "d")],
            Self::LeftUp => &[("dir", "lu")],
            Self::RightUp => &[("dir", "ru")],
            Self::LeftDown => &[("dir", "ld")],
            Self::RightDown => &[("dir", "rd")],
            Self::In => &[("dir", "in")],
            Self::Out => &[("dir", "out")],
            Self::HorizontalIn => &[("orient", "horz"), ("dir", "in")],
            Self::HorizontalOut => &[("orient", "horz"), ("dir", "out")],
            Self::VerticalIn => &[("orient", "vert"), ("dir", "in")],
            Self::VerticalOut => &[("orient", "vert"), ("dir", "out")],
        }
    }
}

/// A slide transition and how the slide show advances past the slide.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct SlideTransition {
    /// The effect, `None` for a transition that only sets advance timing.
    pub kind: Option<TransitionKind>,
    /// The effect direction, `None` for the kind's default.
    pub direction: Option<TransitionDirection>,
    /// The effect duration in milliseconds, `None` for the default speed.
    pub duration_ms: Option<u64>,
    /// Whether a click advances, true by default.
    pub advance_on_click: bool,
    /// Advance automatically after this many milliseconds.
    pub advance_after_ms: Option<u64>,
}

impl Default for SlideTransition {
    fn default() -> Self {
        Self {
            kind: None,
            direction: None,
            duration_ms: None,
            advance_on_click: true,
            advance_after_ms: None,
        }
    }
}

impl SlideTransition {
    fn from_model(model: &CT_SlideTransition) -> Self {
        let kind = model.effect.as_ref().map(|effect| match effect {
            TransitionEffect::Cut => TransitionKind::Cut,
            TransitionEffect::Fade => TransitionKind::Fade,
            TransitionEffect::Wipe => TransitionKind::Wipe,
            TransitionEffect::Push => TransitionKind::Push,
            TransitionEffect::Zoom => TransitionKind::Zoom,
            TransitionEffect::Morph => TransitionKind::Other("morph".to_owned()),
            TransitionEffect::Other(name) => match name.as_str() {
                "split" => TransitionKind::Split,
                "cover" => TransitionKind::Cover,
                "pull" => TransitionKind::Uncover,
                _ => TransitionKind::Other(name.clone()),
            },
        });
        let parameter = |name: &str| {
            model
                .effect_parameters
                .iter()
                .find(|parameter| parameter.name == name)
                .map(|parameter| parameter.value.as_str())
        };
        // An effect that names any direction attribute reads with the
        // schema defaults for the others, so `<p:split orient="vert"/>` is
        // vertical-out. One that names none keeps the default, `None`.
        let direction = kind.as_ref().and_then(|kind| {
            let defaults = kind.directions().first()?.attributes();
            if !defaults.iter().any(|(name, _)| parameter(name).is_some()) {
                return None;
            }
            kind.directions().iter().copied().find(|direction| {
                direction.attributes().iter().all(|(name, value)| {
                    let default = defaults
                        .iter()
                        .find(|(default_name, _)| default_name == name)
                        .map(|(_, value)| *value);
                    parameter(name).or(default) == Some(*value)
                })
            })
        });
        Self {
            kind,
            direction,
            duration_ms: model.duration_ms,
            advance_on_click: model.advance_on_click.unwrap_or(true),
            advance_after_ms: model.advance_after_ms,
        }
    }

    /// Builds the `p:transition` this value writes, `None` when it has no
    /// effect and the default advance, which is no transition at all.
    fn to_model(&self) -> Result<Option<CT_SlideTransition>> {
        let operation = "set a slide transition";
        if let Some(TransitionKind::Other(name)) = &self.kind {
            return Err(invalid_slide_mutation(
                operation,
                format!(
                    "transition {name:?} cannot be authored, use one of {}",
                    TransitionKind::NAMES.join(", ")
                ),
            ));
        }
        if let Some(direction) = self.direction {
            let accepted = self
                .kind
                .as_ref()
                .map_or(&[][..], TransitionKind::directions);
            if !accepted.contains(&direction) {
                let kind = self.kind.as_ref().map_or("no effect", TransitionKind::name);
                let names = accepted
                    .iter()
                    .map(|direction| direction.name())
                    .collect::<Vec<_>>();
                return Err(invalid_slide_mutation(
                    operation,
                    if names.is_empty() {
                        format!("{kind} takes no direction")
                    } else {
                        format!(
                            "{kind} does not take direction {}, use one of {}",
                            direction.name(),
                            names.join(", ")
                        )
                    },
                ));
            }
        }
        if self.kind.is_none() && self.duration_ms.is_some() {
            return Err(invalid_slide_mutation(
                operation,
                "a duration needs an effect, set a transition type first",
            ));
        }
        if self.kind.is_none() && self.advance_on_click && self.advance_after_ms.is_none() {
            return Ok(None);
        }
        let attributes = self
            .direction
            .map_or(&[][..], TransitionDirection::attributes);
        let effect = self.kind.as_ref().map(|kind| (kind.element(), attributes));
        CT_SlideTransition::authored(
            effect,
            self.duration_ms,
            (!self.advance_on_click).then_some(false),
            self.advance_after_ms,
        )
        .map(Some)
        .map_err(|error| invalid_slide_mutation(operation, error.to_string()))
    }
}

impl SlideRef<'_> {
    /// Returns the slide's transition, `None` when it has none.
    pub fn transition(&self) -> Option<SlideTransition> {
        self.record
            .slide
            .transition
            .as_ref()
            .map(SlideTransition::from_model)
    }
}

impl SlideMut<'_> {
    /// Replaces the slide's transition, `None` removing it.
    ///
    /// The new element replaces the old one whole, sound included. An
    /// effect outside the authorable set, a direction the effect does not
    /// take or a duration without an effect is an error.
    pub fn set_transition(&mut self, transition: Option<&SlideTransition>) -> Result<()> {
        let model = match transition {
            Some(transition) => transition.to_model()?,
            None => None,
        };
        self.record.slide.transition = model;
        Ok(())
    }
}

impl Presentation {
    /// Gives every slide the same transition, as PowerPoint's Apply To All.
    pub fn set_all_transitions(&mut self, transition: Option<&SlideTransition>) -> Result<()> {
        let model = match transition {
            Some(transition) => transition.to_model()?,
            None => None,
        };
        for record in &mut self.slides {
            record.slide.transition = model.clone();
        }
        Ok(())
    }
}
