use core::fmt;

use serde::de::Error as _;
use serde::{Deserialize, Deserializer, Serialize, Serializer};

use crate::style::Color;

/// A concrete ARGB color (`0xAARRGGBB`).
///
/// Serialized as a `#AARRGGBB` hex string for JSON/IPC friendliness.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash)]
pub struct ArgbColor(pub u32);

impl ArgbColor {
    pub const fn new(argb: u32) -> Self {
        Self(argb)
    }

    pub const fn argb(self) -> u32 {
        self.0
    }
}

impl fmt::Display for ArgbColor {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "#{:08X}", self.0)
    }
}

impl Serialize for ArgbColor {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        serializer.serialize_str(&format!("#{:08X}", self.0))
    }
}

impl<'de> Deserialize<'de> for ArgbColor {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;
        parse_hex_argb(&s).map_err(D::Error::custom)
    }
}

fn parse_hex_argb(s: &str) -> Result<ArgbColor, String> {
    let hex = s.trim();
    let hex = hex.strip_prefix('#').unwrap_or(hex);
    let argb = if hex.len() == 8 {
        u32::from_str_radix(hex, 16).map_err(|_| "invalid hex".to_string())?
    } else if hex.len() == 6 {
        let rgb = u32::from_str_radix(hex, 16).map_err(|_| "invalid hex".to_string())?;
        0xFF00_0000 | rgb
    } else {
        return Err(
            "expected a #AARRGGBB (8 hex digits) or #RRGGBB (6 hex digits) string".to_string(),
        );
    };
    Ok(ArgbColor(argb))
}

/// Named theme color slots as defined by OOXML (`a:clrScheme`).
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash)]
pub enum ThemeColorSlot {
    /// Light 1 (theme index 0)
    Lt1,
    /// Dark 1 (theme index 1)
    Dk1,
    /// Light 2 (theme index 2)
    Lt2,
    /// Dark 2 (theme index 3)
    Dk2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    /// Hyperlink (theme index 10)
    Hlink,
    /// Followed hyperlink (theme index 11)
    FolHlink,
}

impl ThemeColorSlot {
    /// Convert the numeric `theme` index used in SpreadsheetML colors to a slot.
    pub const fn from_theme_index(theme: u16) -> Option<Self> {
        Some(match theme {
            0 => ThemeColorSlot::Lt1,
            1 => ThemeColorSlot::Dk1,
            2 => ThemeColorSlot::Lt2,
            3 => ThemeColorSlot::Dk2,
            4 => ThemeColorSlot::Accent1,
            5 => ThemeColorSlot::Accent2,
            6 => ThemeColorSlot::Accent3,
            7 => ThemeColorSlot::Accent4,
            8 => ThemeColorSlot::Accent5,
            9 => ThemeColorSlot::Accent6,
            10 => ThemeColorSlot::Hlink,
            11 => ThemeColorSlot::FolHlink,
            _ => return None,
        })
    }

    pub const fn theme_index(self) -> u16 {
        match self {
            ThemeColorSlot::Lt1 => 0,
            ThemeColorSlot::Dk1 => 1,
            ThemeColorSlot::Lt2 => 2,
            ThemeColorSlot::Dk2 => 3,
            ThemeColorSlot::Accent1 => 4,
            ThemeColorSlot::Accent2 => 5,
            ThemeColorSlot::Accent3 => 6,
            ThemeColorSlot::Accent4 => 7,
            ThemeColorSlot::Accent5 => 8,
            ThemeColorSlot::Accent6 => 9,
            ThemeColorSlot::Hlink => 10,
            ThemeColorSlot::FolHlink => 11,
        }
    }
}

/// Minimal theme palette needed to resolve SpreadsheetML `theme=` colors.
///
/// This matches the 12 standard theme colors in OOXML:
/// `dk1`, `lt1`, `dk2`, `lt2`, `accent1..accent6`, `hlink`, `folHlink`.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct ThemePalette {
    pub dk1: ArgbColor,
    pub lt1: ArgbColor,
    pub dk2: ArgbColor,
    pub lt2: ArgbColor,
    pub accent1: ArgbColor,
    pub accent2: ArgbColor,
    pub accent3: ArgbColor,
    pub accent4: ArgbColor,
    pub accent5: ArgbColor,
    pub accent6: ArgbColor,
    pub hlink: ArgbColor,
    #[serde(rename = "folHlink")]
    pub fol_hlink: ArgbColor,
}

impl ThemePalette {
    /// Excel/Office default theme used by modern `.xlsx` files (Office 2013+).
    pub const fn office_2013() -> Self {
        Self {
            dk1: ArgbColor(0xFF000000),
            lt1: ArgbColor(0xFFFFFFFF),
            dk2: ArgbColor(0xFF44546A),
            lt2: ArgbColor(0xFFE7E6E6),
            accent1: ArgbColor(0xFF5B9BD5),
            accent2: ArgbColor(0xFFED7D31),
            accent3: ArgbColor(0xFFA5A5A5),
            accent4: ArgbColor(0xFFFFC000),
            accent5: ArgbColor(0xFF4472C4),
            accent6: ArgbColor(0xFF70AD47),
            hlink: ArgbColor(0xFF0563C1),
            fol_hlink: ArgbColor(0xFF954F72),
        }
    }

    /// Legacy Excel/Office default theme (Office 2007/2010 era).
    pub const fn office_2007() -> Self {
        Self {
            dk1: ArgbColor(0xFF000000),
            lt1: ArgbColor(0xFFFFFFFF),
            dk2: ArgbColor(0xFF1F497D),
            lt2: ArgbColor(0xFFEEECE1),
            accent1: ArgbColor(0xFF4F81BD),
            accent2: ArgbColor(0xFFC0504D),
            accent3: ArgbColor(0xFF9BBB59),
            accent4: ArgbColor(0xFF8064A2),
            accent5: ArgbColor(0xFF4BACC6),
            accent6: ArgbColor(0xFFF79646),
            hlink: ArgbColor(0xFF0000FF),
            fol_hlink: ArgbColor(0xFF800080),
        }
    }

    pub fn is_default(&self) -> bool {
        self == &DEFAULT_THEME_PALETTE
    }

    pub fn slot(&self, slot: ThemeColorSlot) -> ArgbColor {
        match slot {
            ThemeColorSlot::Lt1 => self.lt1,
            ThemeColorSlot::Dk1 => self.dk1,
            ThemeColorSlot::Lt2 => self.lt2,
            ThemeColorSlot::Dk2 => self.dk2,
            ThemeColorSlot::Accent1 => self.accent1,
            ThemeColorSlot::Accent2 => self.accent2,
            ThemeColorSlot::Accent3 => self.accent3,
            ThemeColorSlot::Accent4 => self.accent4,
            ThemeColorSlot::Accent5 => self.accent5,
            ThemeColorSlot::Accent6 => self.accent6,
            ThemeColorSlot::Hlink => self.hlink,
            ThemeColorSlot::FolHlink => self.fol_hlink,
        }
    }

    pub fn color_for_theme_index(&self, theme: u16) -> Option<ArgbColor> {
        Some(self.slot(ThemeColorSlot::from_theme_index(theme)?))
    }
}

impl Default for ThemePalette {
    fn default() -> Self {
        ThemePalette::office_2013()
    }
}

/// Default theme palette used when a workbook/theme part is unavailable.
pub const DEFAULT_THEME_PALETTE: ThemePalette = ThemePalette::office_2013();

/// Context for resolving `Color::Auto`.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash)]
pub enum ColorContext {
    Font,
    Fill,
    Border,
    Other,
}

/// Resolve an Excel [`Color`] reference into a concrete ARGB value (`0xAARRGGBB`).
///
/// Returns `None` when the color is context-dependent (e.g. `Auto`) or cannot be resolved
/// with the available information.
pub fn resolve_color(color: Color, theme: Option<&ThemePalette>) -> Option<u32> {
    resolve_color_in_context(color, theme, ColorContext::Other)
}

/// Resolve an Excel [`Color`] reference into a concrete ARGB value (`0xAARRGGBB`) using
/// a context for `Color::Auto`.
pub fn resolve_color_in_context(
    color: Color,
    theme: Option<&ThemePalette>,
    context: ColorContext,
) -> Option<u32> {
    match color {
        Color::Argb(argb) => Some(argb),
        Color::Theme { theme: idx, tint } => {
            let theme = theme.unwrap_or(&DEFAULT_THEME_PALETTE);
            let mut argb = theme.color_for_theme_index(idx)?.argb();
            if let Some(tint) = tint {
                argb = apply_tint(argb, tint);
            }
            Some(argb)
        }
        Color::Indexed(index) => {
            // Excel (BIFF8) reserves index 64 as "automatic".
            if index == 64 {
                return resolve_auto(context);
            }
            indexed_color_argb(index)
        }
        Color::Auto => resolve_auto(context),
    }
}

fn resolve_auto(context: ColorContext) -> Option<u32> {
    match context {
        // Excel's automatic font and border colors render as black in most contexts.
        ColorContext::Font | ColorContext::Border => Some(0xFF000000),
        // Automatic fills are context-dependent (and often mean "no fill").
        ColorContext::Fill | ColorContext::Other => None,
    }
}

fn apply_tint(argb: u32, tint_thousandths: i16) -> u32 {
    let tint = (tint_thousandths as f64 / 1000.0).clamp(-1.0, 1.0);
    if tint == 0.0 {
        return argb;
    }

    let a = (argb >> 24) & 0xFF;
    let r = ((argb >> 16) & 0xFF) as u8;
    let g = ((argb >> 8) & 0xFF) as u8;
    let b = (argb & 0xFF) as u8;

    let r = tint_channel(r, tint) as u32;
    let g = tint_channel(g, tint) as u32;
    let b = tint_channel(b, tint) as u32;

    (a << 24) | (r << 16) | (g << 8) | b
}

fn tint_channel(value: u8, tint: f64) -> u8 {
    let v = value as f64;
    let out = if tint < 0.0 {
        // Shade toward black.
        v * (1.0 + tint)
    } else {
        // Tint toward white.
        v * (1.0 - tint) + 255.0 * tint
    };

    out.round().clamp(0.0, 255.0) as u8
}

/// Excel's standard indexed color table for indices `0..=63`.
///
/// Reference: ECMA-376 / SpreadsheetML default `indexedColors` palette.
///
/// Indices outside `0..=63` are not handled here (Excel uses index 64 for "automatic" and
/// may define custom palettes in `styles.xml`).
pub fn indexed_color_argb(index: u16) -> Option<u32> {
    EXCEL_INDEXED_COLORS.get(index as usize).copied()
}

const EXCEL_INDEXED_COLORS: [u32; 64] = [
    0xFF000000, // 0
    0xFFFFFFFF, // 1
    0xFFFF0000, // 2
    0xFF00FF00, // 3
    0xFF0000FF, // 4
    0xFFFFFF00, // 5
    0xFFFF00FF, // 6
    0xFF00FFFF, // 7
    0xFF000000, // 8
    0xFFFFFFFF, // 9
    0xFFFF0000, // 10
    0xFF00FF00, // 11
    0xFF0000FF, // 12
    0xFFFFFF00, // 13
    0xFFFF00FF, // 14
    0xFF00FFFF, // 15
    0xFF800000, // 16
    0xFF008000, // 17
    0xFF000080, // 18
    0xFF808000, // 19
    0xFF800080, // 20
    0xFF008080, // 21
    0xFFC0C0C0, // 22
    0xFF808080, // 23
    0xFF9999FF, // 24
    0xFF993366, // 25
    0xFFFFFFCC, // 26
    0xFFCCFFFF, // 27
    0xFF660066, // 28
    0xFFFF8080, // 29
    0xFF0066CC, // 30
    0xFFCCCCFF, // 31
    0xFF000080, // 32
    0xFFFF00FF, // 33
    0xFFFFFF00, // 34
    0xFF00FFFF, // 35
    0xFF800080, // 36
    0xFF800000, // 37
    0xFF008080, // 38
    0xFF0000FF, // 39
    0xFF00CCFF, // 40
    0xFFCCFFFF, // 41
    0xFFCCFFCC, // 42
    0xFFFFFF99, // 43
    0xFF99CCFF, // 44
    0xFFFF99CC, // 45
    0xFFCC99FF, // 46
    0xFFFFCC99, // 47
    0xFF3366FF, // 48
    0xFF33CCCC, // 49
    0xFF99CC00, // 50
    0xFFFFCC00, // 51
    0xFFFF9900, // 52
    0xFFFF6600, // 53
    0xFF666699, // 54
    0xFF969696, // 55
    0xFF003366, // 56
    0xFF339966, // 57
    0xFF003300, // 58
    0xFF333300, // 59
    0xFF993300, // 60
    0xFF993366, // 61
    0xFF333399, // 62
    0xFF333333, // 63
];

/// Parse an Excel number format bracket token into a [`Color`].
///
/// Supports:
/// - Named colors: `Black`, `Blue`, `Cyan`, `Green`, `Magenta`, `Red`, `White`, `Yellow`
/// - Indexed colors: `ColorN` (e.g. `Color10`).
pub fn parse_number_format_color_token(token: &str) -> Option<Color> {
    let t = token.trim();
    if t.is_empty() {
        return None;
    }

    if t.eq_ignore_ascii_case("black") {
        return Some(Color::Argb(0xFF000000));
    }
    if t.eq_ignore_ascii_case("blue") {
        return Some(Color::Argb(0xFF0000FF));
    }
    if t.eq_ignore_ascii_case("cyan") {
        return Some(Color::Argb(0xFF00FFFF));
    }
    if t.eq_ignore_ascii_case("green") {
        return Some(Color::Argb(0xFF00FF00));
    }
    if t.eq_ignore_ascii_case("magenta") {
        return Some(Color::Argb(0xFFFF00FF));
    }
    if t.eq_ignore_ascii_case("red") {
        return Some(Color::Argb(0xFFFF0000));
    }
    if t.eq_ignore_ascii_case("white") {
        return Some(Color::Argb(0xFFFFFFFF));
    }
    if t.eq_ignore_ascii_case("yellow") {
        return Some(Color::Argb(0xFFFFFF00));
    }

    let Some(prefix) = t.get(0..5) else {
        return None;
    };
    if !prefix.eq_ignore_ascii_case("color") {
        return None;
    }

    let num = &t[5..];
    if num.is_empty() || !num.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    let idx = num.parse::<u16>().ok()?;
    Some(Color::Indexed(idx))
}

/// Extract the number format section color for a numeric value.
///
/// Excel format codes can have up to 4 `;`-delimited sections, each optionally
/// prefixed with bracket tokens like `[Red]` or `[Color10]`. This helper picks the
/// appropriate section for `value` (including conditional sections like `[>=100]`)
/// and returns the color token if present.
pub fn number_format_color(format_code: &str, value: f64) -> Option<Color> {
    let sections = split_number_format_sections(format_code);
    let parsed: Vec<NumberFormatSection<'_>> = sections
        .iter()
        .map(|s| parse_number_format_section(s))
        .collect();

    if parsed.is_empty() {
        return None;
    }

    // Conditional sections (SpreadsheetML semantics): if any section has a condition,
    // Excel evaluates in-order and uses the first matching condition, otherwise the first
    // unconditional section as a fallback.
    if parsed.iter().any(|s| s.condition.is_some()) {
        let mut fallback: Option<&NumberFormatSection<'_>> = None;
        for section in &parsed {
            match section.condition {
                Some(cond) => {
                    if cond.matches(value) {
                        return section.color;
                    }
                }
                None => {
                    if fallback.is_none() {
                        fallback = Some(section);
                    }
                }
            }
        }
        return fallback.unwrap_or(&parsed[0]).color;
    }

    // Non-conditional sections: Excel semantics for 1-4 sections:
    // 1: positive; used for all numbers
    // 2: positive;negative
    // 3: positive;negative;zero
    // 4: positive;negative;zero;text
    let section_count = parsed.len();
    if value < 0.0 {
        if section_count >= 2 {
            parsed[1].color
        } else {
            parsed[0].color
        }
    } else if value == 0.0 {
        if section_count >= 3 {
            parsed[2].color
        } else {
            parsed[0].color
        }
    } else {
        parsed[0].color
    }
}

/// Convenience wrapper for `number_format_color` + `resolve_color_in_context` using `Font` context.
pub fn resolve_number_format_color(
    format_code: &str,
    value: f64,
    theme: Option<&ThemePalette>,
) -> Option<u32> {
    let color = number_format_color(format_code, value)?;
    resolve_color_in_context(color, theme, ColorContext::Font)
}

#[derive(Debug, Clone, Copy)]
struct NumberFormatSection<'a> {
    #[allow(dead_code)]
    raw: &'a str,
    color: Option<Color>,
    condition: Option<NumberFormatCondition>,
}

#[derive(Debug, Clone, Copy)]
struct NumberFormatCondition {
    op: NumberFormatCmpOp,
    rhs: f64,
}

impl NumberFormatCondition {
    fn matches(self, v: f64) -> bool {
        match self.op {
            NumberFormatCmpOp::Lt => v < self.rhs,
            NumberFormatCmpOp::Le => v <= self.rhs,
            NumberFormatCmpOp::Gt => v > self.rhs,
            NumberFormatCmpOp::Ge => v >= self.rhs,
            NumberFormatCmpOp::Eq => v == self.rhs,
            NumberFormatCmpOp::Ne => v != self.rhs,
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum NumberFormatCmpOp {
    Lt,
    Le,
    Gt,
    Ge,
    Eq,
    Ne,
}

fn split_number_format_sections(code: &str) -> Vec<&str> {
    let mut sections = Vec::new();
    let mut start = 0usize;
    let mut in_quotes = false;
    let mut chars = code.char_indices().peekable();
    while let Some((idx, ch)) = chars.next() {
        match ch {
            '"' => in_quotes = !in_quotes,
            '\\' => {
                // Skip escaped character.
                chars.next();
            }
            ';' if !in_quotes => {
                sections.push(&code[start..idx]);
                start = idx + 1;
            }
            _ => {}
        }
    }
    sections.push(&code[start..]);
    sections
}

fn parse_number_format_section<'a>(raw: &'a str) -> NumberFormatSection<'a> {
    let mut rest = raw.trim_start();
    let mut color: Option<Color> = None;
    let mut condition: Option<NumberFormatCondition> = None;

    // Parse leading `[ ... ]` tokens.
    loop {
        let Some(stripped) = rest.strip_prefix('[') else {
            break;
        };
        let Some(end) = stripped.find(']') else {
            break;
        };
        let content = &stripped[..end];

        // Elapsed time tokens are part of the format pattern.
        if content.eq_ignore_ascii_case("h")
            || content.eq_ignore_ascii_case("hh")
            || content.eq_ignore_ascii_case("m")
            || content.eq_ignore_ascii_case("mm")
            || content.eq_ignore_ascii_case("s")
            || content.eq_ignore_ascii_case("ss")
        {
            break;
        }

        if color.is_none() {
            color = parse_number_format_color_token(content);
        }
        if condition.is_none() {
            condition = parse_number_format_condition(content);
        }

        rest = stripped[end + 1..].trim_start();
    }

    NumberFormatSection {
        raw,
        color,
        condition,
    }
}

fn parse_number_format_condition(token: &str) -> Option<NumberFormatCondition> {
    let (op, rhs_str) = if let Some(rest) = token.strip_prefix(">=") {
        (NumberFormatCmpOp::Ge, rest)
    } else if let Some(rest) = token.strip_prefix("<=") {
        (NumberFormatCmpOp::Le, rest)
    } else if let Some(rest) = token.strip_prefix("<>") {
        (NumberFormatCmpOp::Ne, rest)
    } else if let Some(rest) = token.strip_prefix('>') {
        (NumberFormatCmpOp::Gt, rest)
    } else if let Some(rest) = token.strip_prefix('<') {
        (NumberFormatCmpOp::Lt, rest)
    } else if let Some(rest) = token.strip_prefix('=') {
        (NumberFormatCmpOp::Eq, rest)
    } else {
        return None;
    };

    let rhs = rhs_str.trim().parse::<f64>().ok()?;
    Some(NumberFormatCondition { op, rhs })
}
