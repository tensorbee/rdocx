//! Length unit type with convenient constructors.

use rdocx_oxml::units::{Emu, HalfPoint, Twips};

/// A length measurement that can be expressed in various units.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Length {
    emu: i64,
}

impl Length {
    /// Create a length from inches.
    pub fn inches(val: f64) -> Self {
        Length {
            emu: (val * 914400.0) as i64,
        }
    }

    /// Create a length from centimeters.
    pub fn cm(val: f64) -> Self {
        Length {
            emu: (val * 360000.0) as i64,
        }
    }

    /// Create a length from points.
    pub fn pt(val: f64) -> Self {
        Length {
            emu: (val * 12700.0) as i64,
        }
    }

    /// Create a length from EMUs.
    pub fn emu(val: i64) -> Self {
        Length { emu: val }
    }

    /// Create a length from twips.
    pub fn twips(val: i32) -> Self {
        Length {
            emu: val as i64 * 635,
        }
    }

    /// Get the value in inches.
    pub fn to_inches(self) -> f64 {
        self.emu as f64 / 914400.0
    }

    /// Get the value in centimeters.
    pub fn to_cm(self) -> f64 {
        self.emu as f64 / 360000.0
    }

    /// Get the value in points.
    pub fn to_pt(self) -> f64 {
        self.emu as f64 / 12700.0
    }

    /// Get the value in EMUs.
    pub fn to_emu(self) -> i64 {
        self.emu
    }

    /// Get the value in twips.
    pub fn to_twips(self) -> i32 {
        (self.emu / 635) as i32
    }

    /// Convert to the OXML Twips type.
    pub fn as_twips(self) -> Twips {
        Twips(self.to_twips())
    }

    /// Convert to the OXML Emu type.
    pub fn as_emu(self) -> Emu {
        Emu(self.emu)
    }

    /// Convert to the OXML HalfPoint type (for font sizes).
    pub fn as_half_points(self) -> HalfPoint {
        HalfPoint((self.to_pt() * 2.0) as u32)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn length_conversions() {
        let one_inch = Length::inches(1.0);
        assert!((one_inch.to_inches() - 1.0).abs() < 0.001);
        assert!((one_inch.to_cm() - 2.54).abs() < 0.01);
        assert!((one_inch.to_pt() - 72.0).abs() < 0.01);
        assert_eq!(one_inch.to_twips(), 1440);
    }

    #[test]
    fn length_from_pt() {
        let twelve_pt = Length::pt(12.0);
        assert_eq!(twelve_pt.as_half_points().0, 24);
    }
}
