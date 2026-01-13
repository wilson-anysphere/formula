use roxmltree::Document;

use crate::XlsxError;

pub(crate) mod convert;

/// Minimal palette extracted from `xl/theme/theme1.xml`.
///
/// Colors are stored as ARGB (`0xAARRGGBB`) with alpha always set to `0xFF`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ThemePalette {
    pub dk1: u32,
    pub lt1: u32,
    pub dk2: u32,
    pub lt2: u32,
    pub accent1: u32,
    pub accent2: u32,
    pub accent3: u32,
    pub accent4: u32,
    pub accent5: u32,
    pub accent6: u32,
    pub hlink: u32,
    pub followed_hlink: u32,
}

impl Default for ThemePalette {
    fn default() -> Self {
        // Default "Office" theme colors as used by the OpenXML spec.
        // Many XLSX writers omit the full theme definition and rely on these defaults.
        Self {
            dk1: 0xFF000000,
            lt1: 0xFFFFFFFF,
            dk2: 0xFF1F497D,
            lt2: 0xFFEEECE1,
            accent1: 0xFF4F81BD,
            accent2: 0xFFC0504D,
            accent3: 0xFF9BBB59,
            accent4: 0xFF8064A2,
            accent5: 0xFF4BACC6,
            accent6: 0xFFF79646,
            hlink: 0xFF0000FF,
            followed_hlink: 0xFF800080,
        }
    }
}

impl ThemePalette {
    pub fn accents(&self) -> [u32; 6] {
        [
            self.accent1,
            self.accent2,
            self.accent3,
            self.accent4,
            self.accent5,
            self.accent6,
        ]
    }
}

/// Parse an Excel theme (`xl/theme/theme1.xml`) and extract a minimal color palette.
///
/// If the theme is missing a color scheme, this returns the OpenXML default palette.
pub fn parse_theme_palette(theme_xml: &[u8]) -> Result<ThemePalette, XlsxError> {
    let xml = String::from_utf8(theme_xml.to_vec())?;
    let doc = Document::parse(&xml)?;

    let mut palette = ThemePalette::default();

    let Some(clr_scheme) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "clrScheme")
    else {
        return Ok(palette);
    };

    assign_color(&mut palette.dk1, parse_clr_scheme_entry(clr_scheme, "dk1"));
    assign_color(&mut palette.lt1, parse_clr_scheme_entry(clr_scheme, "lt1"));
    assign_color(&mut palette.dk2, parse_clr_scheme_entry(clr_scheme, "dk2"));
    assign_color(&mut palette.lt2, parse_clr_scheme_entry(clr_scheme, "lt2"));
    assign_color(&mut palette.accent1, parse_clr_scheme_entry(clr_scheme, "accent1"));
    assign_color(&mut palette.accent2, parse_clr_scheme_entry(clr_scheme, "accent2"));
    assign_color(&mut palette.accent3, parse_clr_scheme_entry(clr_scheme, "accent3"));
    assign_color(&mut palette.accent4, parse_clr_scheme_entry(clr_scheme, "accent4"));
    assign_color(&mut palette.accent5, parse_clr_scheme_entry(clr_scheme, "accent5"));
    assign_color(&mut palette.accent6, parse_clr_scheme_entry(clr_scheme, "accent6"));
    assign_color(&mut palette.hlink, parse_clr_scheme_entry(clr_scheme, "hlink"));
    // Excel stores this as <a:folHlink>.
    assign_color(
        &mut palette.followed_hlink,
        parse_clr_scheme_entry(clr_scheme, "folHlink"),
    );

    Ok(palette)
}

fn assign_color(slot: &mut u32, parsed: Option<u32>) {
    if let Some(color) = parsed {
        *slot = color;
    }
}

fn parse_clr_scheme_entry(clr_scheme: roxmltree::Node<'_, '_>, name: &str) -> Option<u32> {
    let entry = clr_scheme
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == name)?;
    let clr = entry.children().find(|n| n.is_element())?;

    match clr.tag_name().name() {
        "srgbClr" => clr.attribute("val").and_then(parse_rgb_hex),
        "sysClr" => clr
            .attribute("lastClr")
            .and_then(parse_rgb_hex)
            .or_else(|| clr.attribute("val").and_then(sys_clr_fallback)),
        _ => None,
    }
}

fn sys_clr_fallback(val: &str) -> Option<u32> {
    // Minimal mappings for common system colors. If `lastClr` is present, we always
    // prefer it since it reflects what Excel last resolved on the authoring system.
    match val {
        "windowText" | "WindowText" => Some(0xFF000000),
        "window" | "Window" => Some(0xFFFFFFFF),
        _ => None,
    }
}

fn parse_rgb_hex(value: &str) -> Option<u32> {
    let hex = value.trim().trim_start_matches('#');
    match hex.len() {
        6 => u32::from_str_radix(hex, 16).ok().map(|rgb| 0xFF00_0000 | rgb),
        8 => u32::from_str_radix(hex, 16).ok(),
        _ => None,
    }
}

/// Apply an Excel theme tint to an ARGB color.
///
/// `tint` is in thousandths (-1000..=1000), matching `formula_model::Color::Theme { tint }`.
///
/// Excel's tint algorithm:
/// - For `tint < 0` (darken): `c' = c * (1 + tint)`
/// - For `tint > 0` (lighten): `c' = c * (1 - tint) + 255 * tint`
pub fn apply_tint(argb: u32, tint: i16) -> u32 {
    if tint == 0 {
        return argb;
    }

    let tint = (tint as f64 / 1000.0).clamp(-1.0, 1.0);

    let a = (argb >> 24) & 0xFF;
    let r = (argb >> 16) & 0xFF;
    let g = (argb >> 8) & 0xFF;
    let b = argb & 0xFF;

    let r = apply_tint_channel(r as u8, tint);
    let g = apply_tint_channel(g as u8, tint);
    let b = apply_tint_channel(b as u8, tint);

    ((a as u32) << 24) | ((r as u32) << 16) | ((g as u32) << 8) | (b as u32)
}

fn apply_tint_channel(channel: u8, tint: f64) -> u8 {
    let c = channel as f64;
    let adjusted = if tint < 0.0 {
        c * (1.0 + tint)
    } else {
        c * (1.0 - tint) + 255.0 * tint
    };

    adjusted.round().clamp(0.0, 255.0) as u8
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_palette_from_theme_xml() {
        let theme = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="EEEEEE"/></a:lt1>
      <a:dk2><a:srgbClr val="222222"/></a:dk2>
      <a:lt2><a:srgbClr val="DDDDDD"/></a:lt2>
      <a:accent1><a:srgbClr val="010203"/></a:accent1>
      <a:accent2><a:srgbClr val="040506"/></a:accent2>
      <a:accent3><a:srgbClr val="070809"/></a:accent3>
      <a:accent4><a:srgbClr val="0A0B0C"/></a:accent4>
      <a:accent5><a:srgbClr val="0D0E0F"/></a:accent5>
      <a:accent6><a:srgbClr val="101112"/></a:accent6>
      <a:hlink><a:srgbClr val="131415"/></a:hlink>
      <a:folHlink><a:srgbClr val="161718"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"#;

        let palette = parse_theme_palette(theme.as_bytes()).expect("parse theme");
        assert_eq!(palette.dk1, 0xFF111111);
        assert_eq!(palette.lt1, 0xFFEEEEEE);
        assert_eq!(palette.dk2, 0xFF222222);
        assert_eq!(palette.lt2, 0xFFDDDDDD);
        assert_eq!(palette.accent1, 0xFF010203);
        assert_eq!(palette.accent6, 0xFF101112);
        assert_eq!(palette.hlink, 0xFF131415);
        assert_eq!(palette.followed_hlink, 0xFF161718);
    }

    #[test]
    fn apply_tint_handles_extremes_and_identity() {
        let blue = 0xFF0000FF;
        assert_eq!(apply_tint(blue, 0), blue);
        assert_eq!(apply_tint(blue, -1000), 0xFF000000);
        assert_eq!(apply_tint(blue, 1000), 0xFFFFFFFF);
    }

    #[test]
    fn apply_tint_lighten_and_darken() {
        // Darken pure blue by 50%.
        assert_eq!(apply_tint(0xFF0000FF, -500), 0xFF000080);
        // Lighten pure blue by 50%.
        assert_eq!(apply_tint(0xFF0000FF, 500), 0xFF8080FF);

        // 50% gray should tint symmetrically.
        assert_eq!(apply_tint(0xFF000000, 500), 0xFF808080);
        assert_eq!(apply_tint(0xFFFFFFFF, -500), 0xFF808080);
    }
}
