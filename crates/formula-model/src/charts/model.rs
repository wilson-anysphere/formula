use serde::{Deserialize, Serialize};

use crate::RichText;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartModel {
    pub chart_kind: ChartKind,
    pub title: Option<TextModel>,
    pub legend: Option<LegendModel>,
    pub plot_area: PlotAreaModel,
    pub axes: Vec<AxisModel>,
    pub series: Vec<SeriesModel>,
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
}

impl TextModel {
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            rich_text: RichText::new(text),
            formula: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LegendModel {
    pub position: LegendPosition,
    pub overlay: bool,
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
    Unknown { name: String },
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
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct SeriesTextData {
    pub formula: Option<String>,
    pub cache: Option<Vec<String>>,
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
