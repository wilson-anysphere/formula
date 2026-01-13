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
    /// Built-in chart style index (`c:chartSpace/c:style/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_id: Option<u32>,
    /// Whether rounded corners are enabled (`c:chartSpace/c:roundedCorners/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rounded_corners: Option<bool>,
    /// How to display blanks (`c:chart/c:dispBlanksAs/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_blanks_as: Option<String>,
    /// Whether only visible cells are plotted (`c:chart/c:plotVisOnly/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_vis_only: Option<bool>,
    /// Chart area shape properties (`c:chartSpace/c:spPr`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_area_style: Option<ShapeStyle>,
    /// Plot area shape properties (`c:plotArea/c:spPr`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_style: Option<ShapeStyle>,
    /// External workbook link relationship id (`c:chartSpace/c:externalData/@r:id`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub external_data_rel_id: Option<String>,
    /// Whether the chart should auto-update when the linked workbook changes
    /// (`c:chartSpace/c:externalData/c:autoUpdate/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub external_data_auto_update: Option<bool>,
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
    /// Cross-axis id (`c:*Ax/c:crossAx/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cross_axis_id: Option<u32>,
    /// Cross behavior (`c:*Ax/c:crosses/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses: Option<String>,
    /// Explicit crossing position (`c:*Ax/c:crossesAt/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses_at: Option<f64>,
    /// Tick mark style (`c:*Ax/c:majorTickMark/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_tick_mark: Option<String>,
    /// Tick mark style (`c:*Ax/c:minorTickMark/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<String>,
    /// Major unit (`c:*Ax/c:majorUnit/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<f64>,
    /// Minor unit (`c:*Ax/c:minorUnit/@val`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_unit: Option<f64>,
    /// Axis title (`c:*Ax/c:title`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<TextModel>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum AxisKind {
    Category,
    Value,
    Date,
    Series,
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

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct DataLabelsModel {
    pub show_val: Option<bool>,
    pub show_cat_name: Option<bool>,
    pub show_ser_name: Option<bool>,
    pub position: Option<String>,
    pub num_fmt: Option<NumberFormatModel>,
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
    /// Series data label properties (`c:ser/c:dLbls`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_labels: Option<DataLabelsModel>,
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
    /// Literal values embedded in the chart XML (`c:strLit`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub literal: Option<Vec<String>>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct SeriesNumberData {
    pub formula: Option<String>,
    pub cache: Option<Vec<f64>>,
    pub format_code: Option<String>,
    /// Literal values embedded in the chart XML (`c:numLit`).
    ///
    /// See [`SeriesTextData::literal`] for rationale.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub literal: Option<Vec<f64>>,
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
