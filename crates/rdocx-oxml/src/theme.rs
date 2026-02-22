//! Theme parsing: extracts color scheme and font scheme from theme1.xml.

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Result;
use crate::namespace::matches_local_name;

/// Parsed theme information from `word/theme/theme1.xml`.
#[derive(Debug, Clone, Default)]
pub struct Theme {
    /// Theme color scheme (dk1, dk2, lt1, lt2, accent1-6, hlink, folHlink).
    pub colors: ThemeColors,
    /// Major font family (typically used for headings).
    pub major_font: Option<String>,
    /// Minor font family (typically used for body text).
    pub minor_font: Option<String>,
}

/// The 12 standard theme colors, stored as 6-character hex RGB strings.
#[derive(Debug, Clone, Default)]
pub struct ThemeColors {
    pub dk1: Option<String>,
    pub dk2: Option<String>,
    pub lt1: Option<String>,
    pub lt2: Option<String>,
    pub accent1: Option<String>,
    pub accent2: Option<String>,
    pub accent3: Option<String>,
    pub accent4: Option<String>,
    pub accent5: Option<String>,
    pub accent6: Option<String>,
    pub hlink: Option<String>,
    pub fol_hlink: Option<String>,
}

impl ThemeColors {
    /// Look up a theme color by its OOXML name (e.g., "accent1", "dark1").
    pub fn get(&self, name: &str) -> Option<&str> {
        match name {
            "dark1" | "dk1" => self.dk1.as_deref(),
            "dark2" | "dk2" => self.dk2.as_deref(),
            "light1" | "lt1" => self.lt1.as_deref(),
            "light2" | "lt2" => self.lt2.as_deref(),
            "accent1" => self.accent1.as_deref(),
            "accent2" => self.accent2.as_deref(),
            "accent3" => self.accent3.as_deref(),
            "accent4" => self.accent4.as_deref(),
            "accent5" => self.accent5.as_deref(),
            "accent6" => self.accent6.as_deref(),
            "hlink" | "hyperlink" => self.hlink.as_deref(),
            "folHlink" | "followedHyperlink" => self.fol_hlink.as_deref(),
            // Word also uses "text1" = dk1, "text2" = dk2, "background1" = lt1, "background2" = lt2
            "text1" => self.dk1.as_deref(),
            "text2" => self.dk2.as_deref(),
            "background1" | "bg1" => self.lt1.as_deref(),
            "background2" | "bg2" => self.lt2.as_deref(),
            _ => None,
        }
    }
}

impl Theme {
    /// Parse theme from XML bytes (the content of `word/theme/theme1.xml`).
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut theme = Theme::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"clrScheme") {
                        parse_color_scheme(&mut reader, &mut theme.colors)?;
                    } else if matches_local_name(name.as_ref(), b"majorFont") {
                        theme.major_font = parse_font_scheme(&mut reader, b"majorFont")?;
                    } else if matches_local_name(name.as_ref(), b"minorFont") {
                        theme.minor_font = parse_font_scheme(&mut reader, b"minorFont")?;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(theme)
    }
}

/// Parse the `<a:clrScheme>` element to extract theme colors.
fn parse_color_scheme(reader: &mut Reader<&[u8]>, colors: &mut ThemeColors) -> Result<()> {
    let mut buf = Vec::new();
    let mut current_slot: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                // Color slot elements: dk1, dk2, lt1, lt2, accent1-6, hlink, folHlink
                match local {
                    b"dk1" | b"dk2" | b"lt1" | b"lt2" | b"accent1" | b"accent2" | b"accent3"
                    | b"accent4" | b"accent5" | b"accent6" | b"hlink" | b"folHlink" => {
                        current_slot =
                            Some(std::str::from_utf8(local).unwrap_or("").to_string());
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                // Inside a color slot, look for <a:srgbClr val="RRGGBB"/> or <a:sysClr lastClr="RRGGBB"/>
                if let Some(ref slot) = current_slot {
                    let color = if matches_local_name(name.as_ref(), b"srgbClr") {
                        get_attr(e, b"val")
                    } else if matches_local_name(name.as_ref(), b"sysClr") {
                        // System color: use lastClr (the resolved value) if available
                        get_attr(e, b"lastClr").or_else(|| get_attr(e, b"val"))
                    } else {
                        None
                    };

                    if let Some(hex) = color {
                        match slot.as_str() {
                            "dk1" => colors.dk1 = Some(hex),
                            "dk2" => colors.dk2 = Some(hex),
                            "lt1" => colors.lt1 = Some(hex),
                            "lt2" => colors.lt2 = Some(hex),
                            "accent1" => colors.accent1 = Some(hex),
                            "accent2" => colors.accent2 = Some(hex),
                            "accent3" => colors.accent3 = Some(hex),
                            "accent4" => colors.accent4 = Some(hex),
                            "accent5" => colors.accent5 = Some(hex),
                            "accent6" => colors.accent6 = Some(hex),
                            "hlink" => colors.hlink = Some(hex),
                            "folHlink" => colors.fol_hlink = Some(hex),
                            _ => {}
                        }
                    }
                } else {
                    // Top-level color slot with inline color (dk1, etc.)
                    match local {
                        b"dk1" | b"dk2" | b"lt1" | b"lt2" | b"accent1" | b"accent2"
                        | b"accent3" | b"accent4" | b"accent5" | b"accent6" | b"hlink"
                        | b"folHlink" => {
                            // Shouldn't happen (they're Start not Empty), but handle anyway
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"dk1" | b"dk2" | b"lt1" | b"lt2" | b"accent1" | b"accent2" | b"accent3"
                    | b"accent4" | b"accent5" | b"accent6" | b"hlink" | b"folHlink" => {
                        current_slot = None;
                    }
                    b"clrScheme" => break,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.into()),
            _ => {}
        }
        buf.clear();
    }

    Ok(())
}

/// Parse a `<a:majorFont>` or `<a:minorFont>` element to extract the latin typeface.
fn parse_font_scheme(reader: &mut Reader<&[u8]>, end_tag: &[u8]) -> Result<Option<String>> {
    let mut buf = Vec::new();
    let mut latin_font = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) => {
                if matches_local_name(e.name().as_ref(), b"latin") {
                    latin_font = get_attr(e, b"typeface");
                }
            }
            Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), end_tag) => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.into()),
            _ => {}
        }
        buf.clear();
    }

    Ok(latin_font)
}

/// Get the local name (after namespace prefix) from a qualified name.
fn local_name(qname: &[u8]) -> &[u8] {
    match qname.iter().position(|&b| b == b':') {
        Some(pos) => &qname[pos + 1..],
        None => qname,
    }
}

/// Extract a named attribute value from an element.
fn get_attr(e: &quick_xml::events::BytesStart, attr_name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let local = local_name(key);
        if local == attr_name {
            return std::str::from_utf8(&attr.value).ok().map(|s| s.to_string());
        }
    }
    None
}

/// Apply theme tint/shade modifiers to a base color.
///
/// `tint_val` is 0-255 where 255 means full tint (lightest).
/// `shade_val` is 0-255 where 255 means full shade (darkest).
pub fn apply_tint_shade(hex: &str, tint_val: Option<u8>, shade_val: Option<u8>) -> String {
    let hex = hex.trim_start_matches('#');
    if hex.len() < 6 {
        return hex.to_string();
    }

    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f64 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f64 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f64 / 255.0;

    let (r, g, b) = if let Some(tint) = tint_val {
        // Tint: mix with white. tint=0 means no change, tint=255 means pure white
        let t = tint as f64 / 255.0;
        (r + (1.0 - r) * t, g + (1.0 - g) * t, b + (1.0 - b) * t)
    } else if let Some(shade) = shade_val {
        // Shade: mix with black. shade=0 means pure black, shade=255 means no change
        let s = shade as f64 / 255.0;
        (r * s, g * s, b * s)
    } else {
        (r, g, b)
    };

    format!(
        "{:02X}{:02X}{:02X}",
        (r.clamp(0.0, 1.0) * 255.0) as u8,
        (g.clamp(0.0, 1.0) * 255.0) as u8,
        (b.clamp(0.0, 1.0) * 255.0) as u8,
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_office_theme() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>"#;

        let theme = Theme::from_xml(xml).unwrap();

        assert_eq!(theme.colors.dk1.as_deref(), Some("000000"));
        assert_eq!(theme.colors.lt1.as_deref(), Some("FFFFFF"));
        assert_eq!(theme.colors.dk2.as_deref(), Some("44546A"));
        assert_eq!(theme.colors.lt2.as_deref(), Some("E7E6E6"));
        assert_eq!(theme.colors.accent1.as_deref(), Some("4472C4"));
        assert_eq!(theme.colors.accent2.as_deref(), Some("ED7D31"));
        assert_eq!(theme.colors.hlink.as_deref(), Some("0563C1"));
        assert_eq!(theme.colors.fol_hlink.as_deref(), Some("954F72"));

        assert_eq!(theme.major_font.as_deref(), Some("Calibri Light"));
        assert_eq!(theme.minor_font.as_deref(), Some("Calibri"));
    }

    #[test]
    fn theme_color_lookup() {
        let colors = ThemeColors {
            dk1: Some("000000".to_string()),
            lt1: Some("FFFFFF".to_string()),
            accent1: Some("4472C4".to_string()),
            ..Default::default()
        };
        assert_eq!(colors.get("dark1"), Some("000000"));
        assert_eq!(colors.get("text1"), Some("000000"));
        assert_eq!(colors.get("light1"), Some("FFFFFF"));
        assert_eq!(colors.get("background1"), Some("FFFFFF"));
        assert_eq!(colors.get("accent1"), Some("4472C4"));
        assert_eq!(colors.get("nonexistent"), None);
    }

    #[test]
    fn tint_shade_modifiers() {
        // Pure red with 50% tint → pinkish
        let result = apply_tint_shade("FF0000", Some(128), None);
        assert_eq!(result, "FF8080");

        // Pure red with no modification
        let result = apply_tint_shade("FF0000", None, None);
        assert_eq!(result, "FF0000");
    }
}
