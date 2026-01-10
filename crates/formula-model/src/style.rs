use core::fmt;
use std::collections::HashMap;

use serde::de::Error as _;
use serde::{Deserialize, Deserializer, Serialize, Serializer};

/// An ARGB color.
///
/// Serialized as a `#AARRGGBB` hex string for IPC friendliness.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash)]
pub struct Color {
    pub argb: u32,
}

impl Color {
    pub const fn new_argb(argb: u32) -> Self {
        Self { argb }
    }

    pub const fn black() -> Self {
        Self { argb: 0xFF000000 }
    }

    pub const fn white() -> Self {
        Self { argb: 0xFFFFFFFF }
    }

    fn to_hex(self) -> String {
        format!("#{:08X}", self.argb)
    }
}

impl fmt::Display for Color {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.to_hex())
    }
}

impl Serialize for Color {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        serializer.serialize_str(&self.to_hex())
    }
}

impl<'de> Deserialize<'de> for Color {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;
        let s = s.trim();
        let hex = s.strip_prefix('#').ok_or_else(|| {
            D::Error::custom("color must be a #AARRGGBB hex string (missing '#')")
        })?;
        if hex.len() != 8 {
            return Err(D::Error::custom(
                "color must be a #AARRGGBB hex string (8 hex digits)",
            ));
        }
        let argb = u32::from_str_radix(hex, 16).map_err(|_| D::Error::custom("invalid hex"))?;
        Ok(Color { argb })
    }
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
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Color>,
}

/// Fill (background) formatting.
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize, Default)]
pub struct Fill {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<Color>,
}

/// Border line style.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum BorderStyle {
    None,
    Thin,
    Medium,
    Thick,
}

impl Default for BorderStyle {
    fn default() -> Self {
        BorderStyle::None
    }
}

/// Border formatting (subset).
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize, Default)]
pub struct Border {
    #[serde(default)]
    pub top: BorderStyle,
    #[serde(default)]
    pub bottom: BorderStyle,
    #[serde(default)]
    pub left: BorderStyle,
    #[serde(default)]
    pub right: BorderStyle,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Color>,
}

/// Horizontal alignment options (subset).
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum HorizontalAlignment {
    General,
    Left,
    Center,
    Right,
}

/// Vertical alignment options (subset).
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum VerticalAlignment {
    Top,
    Center,
    Bottom,
}

/// Alignment formatting (subset).
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize, Default)]
pub struct Alignment {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<HorizontalAlignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical: Option<VerticalAlignment>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub wrap_text: bool,
    /// Excel text rotation in degrees.
    ///
    /// - `0` = horizontal
    /// - `-90..=90` = rotated text
    /// - Excel also uses `255` for vertical stacked text (supported as-is).
    #[serde(default, skip_serializing_if = "is_zero_i16")]
    pub text_rotation: i16,
}

/// Complete cell style (subset).
///
/// The style is designed to grow toward full Excel fidelity.
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
    pub number_format: Option<String>,
}

fn is_false(b: &bool) -> bool {
    !*b
}

fn is_zero_i16(v: &i16) -> bool {
    *v == 0
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
