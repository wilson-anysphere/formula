use serde::{Deserialize, Serialize};

use crate::Color;

/// A DrawingML color reference.
///
/// For now this is an alias of the workbook-wide [`Color`] type since the same
/// theme/index/tint machinery is used by charts and cell formatting.
pub type ColorRef = Color;

/// Solid fill formatting (`a:solidFill`).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SolidFill {
    pub color: ColorRef,
}

/// Pattern fill formatting (`a:pattFill`).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PatternFill {
    pub pattern: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fg_color: Option<ColorRef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg_color: Option<ColorRef>,
}

/// Gradient fill formatting (`a:gradFill`).
///
/// Full gradient modeling is not implemented yet; we preserve the raw XML to
/// allow renderers to make a best-effort attempt or round-trip the data.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GradientFill {
    pub raw_xml: String,
}

/// Catch-all fill formatting for unsupported DrawingML fill types (`a:*Fill`).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct UnknownFill {
    pub name: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub raw_xml: Option<String>,
}

/// Shape fill formatting extracted from DrawingML.
///
/// This enum is `serde(untagged)` to remain backwards compatible with the
/// historical `ShapeStyle.fill: Option<SolidFill>` representation where the fill
/// was always a `{ color: ... }` object.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum FillStyle {
    Solid(SolidFill),
    None { none: bool },
    Pattern(PatternFill),
    Gradient(GradientFill),
    Unknown(UnknownFill),
}

/// Line dash style (`a:prstDash`).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum LineDash {
    Solid,
    Dash,
    Dot,
    DashDot,
    LongDash,
    LongDashDot,
    LongDashDotDot,
    SysDash,
    SysDot,
    SysDashDot,
    SysDashDotDot,
    #[serde(untagged)]
    Unknown(String),
}

/// Line formatting (`a:ln`).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct LineStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorRef>,
    /// Width in 1/100 points.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width_100pt: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dash: Option<LineDash>,
}

/// Shape properties (`c:spPr`) as a simplified fill+stroke model.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct ShapeStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<FillStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<LineStyle>,
}

impl ShapeStyle {
    pub fn is_empty(&self) -> bool {
        self.fill.is_none() && self.line.is_none()
    }
}

/// Marker symbol (`c:marker/c:symbol`).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum MarkerShape {
    Auto,
    Circle,
    Dash,
    Diamond,
    Dot,
    None,
    Plus,
    Square,
    Star,
    Triangle,
    X,
    #[serde(untagged)]
    Unknown(String),
}

/// Marker formatting (`c:marker`).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct MarkerStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape: Option<MarkerShape>,
    /// Marker size (Excel units, stored as an integer in the XML).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<SolidFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stroke: Option<LineStyle>,
}

/// A simplified text run style model extracted from `c:txPr`.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TextRunStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    /// Font size in 1/100 points.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_100pt: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,
    /// Baseline shift in 1/1000 of the font size (DrawingML `baseline` attribute).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub baseline: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorRef>,
}

impl TextRunStyle {
    pub fn is_empty(&self) -> bool {
        self.font_family.is_none()
            && self.size_100pt.is_none()
            && self.bold.is_none()
            && self.italic.is_none()
            && self.underline.is_none()
            && self.strike.is_none()
            && self.baseline.is_none()
            && self.color.is_none()
    }
}
