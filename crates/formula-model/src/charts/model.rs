use serde::{Deserialize, Serialize};

use crate::RichText;

use super::{MarkerStyle, ShapeStyle, TextRunStyle};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartModel {
    pub chart_kind: ChartKind,
    pub title: Option<TextModel>,
    pub legend: Option<LegendModel>,
    pub plot_area: PlotAreaModel,
    pub axes: Vec<AxisModel>,
    pub series: Vec<SeriesModel>,
    /// Chart area shape properties (`c:chartSpace/c:spPr`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_area_style: Option<ShapeStyle>,
    /// Plot area shape properties (`c:plotArea/c:spPr`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_style: Option<ShapeStyle>,
    pub diagnostics: Vec<ChartDiagnostic>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", tag = "kind")]
pub enum ChartKind {
    Bar,
    Line,
    Pie,
    Scatter,
    Unknown { name: String },
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextModel {
    pub rich_text: RichText,
    pub formula: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<TextRunStyle>,
}

impl TextModel {
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            rich_text: RichText::new(text),
            formula: None,
            style: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LegendModel {
    pub position: LegendPosition,
    pub overlay: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_style: Option<TextRunStyle>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum LegendPosition {
    Left,
    Right,
    Top,
    Bottom,
    TopRight,
    Unknown,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", tag = "kind")]
pub enum PlotAreaModel {
    Bar(BarChartModel),
    Line(LineChartModel),
    Pie(PieChartModel),
    Scatter(ScatterChartModel),
    Combo(ComboPlotAreaModel),
    Unknown { name: String },
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ComboPlotAreaModel {
    pub charts: Vec<ComboChartEntry>,
}

/// A stable index range into [`ChartModel::series`] for a given subplot.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SeriesIndexRange {
    pub start: usize,
    pub end: usize,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", tag = "kind")]
pub enum ComboChartEntry {
    Bar {
        #[serde(flatten)]
        model: BarChartModel,
        series: SeriesIndexRange,
    },
    Line {
        #[serde(flatten)]
        model: LineChartModel,
        series: SeriesIndexRange,
    },
    Pie {
        #[serde(flatten)]
        model: PieChartModel,
        series: SeriesIndexRange,
    },
    Scatter {
        #[serde(flatten)]
        model: ScatterChartModel,
        series: SeriesIndexRange,
    },
    Unknown {
        name: String,
        series: SeriesIndexRange,
    },
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct BarChartModel {
    pub bar_direction: Option<String>,
    pub grouping: Option<String>,
    pub ax_ids: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct LineChartModel {
    pub grouping: Option<String>,
    pub ax_ids: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct PieChartModel {
    pub vary_colors: Option<bool>,
    pub first_slice_angle: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct ScatterChartModel {
    pub scatter_style: Option<String>,
    pub ax_ids: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AxisModel {
    pub id: u32,
    pub kind: AxisKind,
    pub position: Option<AxisPosition>,
    pub scaling: AxisScalingModel,
    pub num_fmt: Option<NumberFormatModel>,
    pub tick_label_position: Option<String>,
    pub major_gridlines: bool,
    /// Axis line shape properties (`c:*Ax/c:spPr`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis_line_style: Option<ShapeStyle>,
    /// Major gridline formatting (`c:*Ax/c:majorGridlines/c:spPr`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridlines_style: Option<ShapeStyle>,
    /// Minor gridline formatting (`c:*Ax/c:minorGridlines/c:spPr`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridlines_style: Option<ShapeStyle>,
    /// Tick label text formatting (`c:*Ax/c:txPr`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_label_text_style: Option<TextRunStyle>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum AxisKind {
    Category,
    Value,
    Unknown,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum AxisPosition {
    Left,
    Right,
    Top,
    Bottom,
    Unknown,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct AxisScalingModel {
    pub min: Option<f64>,
    pub max: Option<f64>,
    pub log_base: Option<f64>,
    pub reverse: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NumberFormatModel {
    pub format_code: String,
    pub source_linked: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct SeriesModel {
    pub name: Option<TextModel>,
    pub categories: Option<SeriesTextData>,
    pub values: Option<SeriesNumberData>,
    pub x_values: Option<SeriesData>,
    pub y_values: Option<SeriesData>,
    /// Series shape properties (`c:ser/c:spPr`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ShapeStyle>,
    /// Series marker properties (`c:ser/c:marker`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<MarkerStyle>,
    /// Per-point overrides (`c:ser/c:dPt`).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub points: Vec<SeriesPointStyle>,
    /// If this chart has multiple subplots (combo chart), the index of the
    /// subplot within the combo plot area.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_index: Option<usize>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SeriesPointStyle {
    pub idx: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ShapeStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<MarkerStyle>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct SeriesTextData {
    pub formula: Option<String>,
    pub cache: Option<Vec<String>>,
    /// Multi-level category label cache (`c:multiLvlStrRef` / `c:multiLvlStrLit`).
    ///
    /// Outer index = level, inner index = point.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub multi_cache: Option<Vec<Vec<String>>>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct SeriesNumberData {
    pub formula: Option<String>,
    pub cache: Option<Vec<f64>>,
    pub format_code: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", tag = "kind")]
pub enum SeriesData {
    Text(SeriesTextData),
    Number(SeriesNumberData),
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartDiagnostic {
    pub level: ChartDiagnosticLevel,
    pub message: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ChartDiagnosticLevel {
    Warning,
}
