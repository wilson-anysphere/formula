use core::fmt;
use std::collections::HashMap;

use serde::de::Error as _;
use serde::ser::SerializeMap;
use serde::{Deserialize, Deserializer, Serialize, Serializer};

use crate::theme::{ColorContext, ThemePalette};

/// An Excel color reference.
///
/// `Argb` serializes as a `#AARRGGBB` hex string for IPC friendliness.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash)]
pub enum Color {
    Argb(u32),
    Theme {
        theme: u16,
        /// Tint in thousandths (-1000..=1000).
        tint: Option<i16>,
    },
    Indexed(u16),
    Auto,
}

impl Color {
    pub const fn new_argb(argb: u32) -> Self {
        Self::Argb(argb)
    }

    pub const fn black() -> Self {
        Self::Argb(0xFF000000)
    }

    pub const fn white() -> Self {
        Self::Argb(0xFFFFFFFF)
    }

    pub fn argb(self) -> Option<u32> {
        match self {
            Color::Argb(v) => Some(v),
            _ => None,
        }
    }

    /// Resolve this color reference into a concrete ARGB value (`0xAARRGGBB`).
    ///
    /// For context-dependent colors like `Auto`, prefer [`Color::resolve_in_context`].
    pub fn resolve(self, theme: Option<&ThemePalette>) -> Option<u32> {
        crate::resolve_color(self, theme)
    }

    /// Resolve this color reference into a concrete ARGB value (`0xAARRGGBB`), using `context`
    /// to determine how `Auto` should be rendered.
    pub fn resolve_in_context(
        self,
        theme: Option<&ThemePalette>,
        context: ColorContext,
    ) -> Option<u32> {
        crate::resolve_color_in_context(self, theme, context)
    }
}

impl fmt::Display for Color {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Color::Argb(argb) => write!(f, "#{:08X}", argb),
            Color::Theme { theme, tint } => write!(
                f,
                "theme({}{})",
                theme,
                tint.map(|t| format!(", tint={t}")).unwrap_or_default()
            ),
            Color::Indexed(index) => write!(f, "indexed({index})"),
            Color::Auto => f.write_str("auto"),
        }
    }
}

impl Serialize for Color {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match self {
            Color::Argb(argb) => serializer.serialize_str(&format!("#{:08X}", argb)),
            Color::Theme { theme, tint } => {
                let mut map = serializer.serialize_map(None)?;
                map.serialize_entry("theme", theme)?;
                if let Some(tint) = tint {
                    map.serialize_entry("tint", tint)?;
                }
                map.end()
            }
            Color::Indexed(index) => {
                let mut map = serializer.serialize_map(None)?;
                map.serialize_entry("indexed", index)?;
                map.end()
            }
            Color::Auto => {
                let mut map = serializer.serialize_map(None)?;
                map.serialize_entry("auto", &true)?;
                map.end()
            }
        }
    }
}

impl<'de> Deserialize<'de> for Color {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        #[derive(Deserialize)]
        #[serde(untagged)]
        enum Helper {
            Hex(String),
            Map {
                #[serde(default)]
                argb: Option<String>,
                #[serde(default)]
                theme: Option<u16>,
                #[serde(default)]
                tint: Option<i16>,
                #[serde(default)]
                indexed: Option<u16>,
                #[serde(default)]
                auto: Option<bool>,
            },
        }

        match Helper::deserialize(deserializer)? {
            Helper::Hex(s) => parse_hex_color(&s).map_err(D::Error::custom),
            Helper::Map {
                argb,
                theme,
                tint,
                indexed,
                auto,
            } => {
                if let Some(argb) = argb {
                    return parse_hex_color(&argb).map_err(D::Error::custom);
                }
                if let Some(theme) = theme {
                    return Ok(Color::Theme { theme, tint });
                }
                if let Some(indexed) = indexed {
                    return Ok(Color::Indexed(indexed));
                }
                if auto.unwrap_or(false) {
                    return Ok(Color::Auto);
                }

                Err(D::Error::custom(
                    "color must be '#AARRGGBB' or {theme|indexed|auto}",
                ))
            }
        }
    }
}

fn parse_hex_color(s: &str) -> Result<Color, String> {
    let s = s.trim();
    let hex = s
        .strip_prefix('#')
        .ok_or_else(|| "color must be a #AARRGGBB hex string (missing '#')".to_string())?;
    if hex.len() != 8 {
        return Err("color must be a #AARRGGBB hex string (8 hex digits)".to_string());
    }
    let argb = u32::from_str_radix(hex, 16).map_err(|_| "invalid hex".to_string())?;
    Ok(Color::Argb(argb))
}

/// Font formatting.
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize, Default)]
pub struct Font {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Font size in 1/100 points (e.g. 1100 = 11pt).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_100pt: Option<u16>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub bold: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub italic: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub underline: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub strike: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Color>,
}

/// Fill pattern type.
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum FillPattern {
    None,
    Gray125,
    Solid,
    Other(String),
}

impl Default for FillPattern {
    fn default() -> Self {
        FillPattern::None
    }
}

/// Fill (background) formatting.
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize, Default)]
pub struct Fill {
    #[serde(default, skip_serializing_if = "is_default_fill_pattern")]
    pub pattern: FillPattern,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fg_color: Option<Color>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg_color: Option<Color>,
}

fn is_default_fill_pattern(p: &FillPattern) -> bool {
    matches!(p, FillPattern::None)
}

/// Border line style.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum BorderStyle {
    None,
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
}

impl Default for BorderStyle {
    fn default() -> Self {
        BorderStyle::None
    }
}

/// A single border edge.
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize, Default)]
pub struct BorderEdge {
    #[serde(default)]
    pub style: BorderStyle,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Color>,
}

/// Border formatting.
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize, Default)]
pub struct Border {
    #[serde(default)]
    pub left: BorderEdge,
    #[serde(default)]
    pub right: BorderEdge,
    #[serde(default)]
    pub top: BorderEdge,
    #[serde(default)]
    pub bottom: BorderEdge,
    #[serde(default)]
    pub diagonal: BorderEdge,
    #[serde(default, skip_serializing_if = "is_false")]
    pub diagonal_up: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub diagonal_down: bool,
}

/// Horizontal alignment options (subset).
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum HorizontalAlignment {
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
}

/// Vertical alignment options (subset).
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum VerticalAlignment {
    Top,
    Center,
    Bottom,
}

/// Alignment formatting.
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize, Default)]
pub struct Alignment {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<HorizontalAlignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical: Option<VerticalAlignment>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub wrap_text: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotation: Option<i16>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<u16>,
}

/// Protection flags.
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct Protection {
    #[serde(default = "default_locked", skip_serializing_if = "is_true")]
    pub locked: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub hidden: bool,
}

const fn default_locked() -> bool {
    true
}

fn is_true(b: &bool) -> bool {
    *b
}

impl Default for Protection {
    fn default() -> Self {
        Self {
            locked: true,
            hidden: false,
        }
    }
}

/// Complete cell style.
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize, Default)]
pub struct Style {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<Font>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<Fill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border: Option<Border>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<Alignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub protection: Option<Protection>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
}

fn is_false(b: &bool) -> bool {
    !*b
}

/// Deduplicated table of styles.
///
/// Cells store a `style_id` referencing this table. Style `0` is always the
/// default (empty) style.
#[derive(Clone, Debug, Serialize)]
pub struct StyleTable {
    pub styles: Vec<Style>,
    #[serde(skip)]
    index: HashMap<Style, u32>,
}

impl Default for StyleTable {
    fn default() -> Self {
        Self::new()
    }
}

impl StyleTable {
    pub fn new() -> Self {
        let mut table = Self {
            styles: vec![Style::default()],
            index: HashMap::new(),
        };
        table.rebuild_index();
        table
    }

    /// Insert (or reuse) a style, returning its ID.
    pub fn intern(&mut self, style: Style) -> u32 {
        if let Some(id) = self.index.get(&style) {
            return *id;
        }
        let id = self.styles.len() as u32;
        self.styles.push(style.clone());
        self.index.insert(style, id);
        id
    }

    /// Get a style by id.
    pub fn get(&self, style_id: u32) -> Option<&Style> {
        self.styles.get(style_id as usize)
    }

    pub fn len(&self) -> usize {
        self.styles.len()
    }

    fn rebuild_index(&mut self) {
        self.index.clear();
        for (i, style) in self.styles.iter().cloned().enumerate() {
            self.index.insert(style, i as u32);
        }
    }
}

impl<'de> Deserialize<'de> for StyleTable {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        #[derive(Deserialize)]
        struct Helper {
            #[serde(default)]
            styles: Vec<Style>,
        }

        let mut helper = Helper::deserialize(deserializer)?;
        if helper.styles.is_empty() {
            helper.styles.push(Style::default());
        }

        let mut table = StyleTable {
            styles: helper.styles,
            index: HashMap::new(),
        };
        table.rebuild_index();
        Ok(table)
    }
}
