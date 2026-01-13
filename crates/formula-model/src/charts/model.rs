use serde::{Deserialize, Serialize};

use crate::RichText;

use super::{
    ChartColorStylePartModel, ChartStylePartModel, MarkerStyle, ShapeStyle, TextRunStyle,
};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct ManualLayoutModel {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout_target: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartModel {
    pub chart_kind: ChartKind,
    pub title: Option<TextModel>,
    pub legend: Option<LegendModel>,
    pub plot_area: PlotAreaModel,
    /// Plot area position and size (`c:plotArea/c:layout/c:manualLayout`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_layout: Option<ManualLayoutModel>,
    pub axes: Vec<AxisModel>,
    #[serde(deserialize_with = "deserialize_series_vec")]
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
    /// Optional external chart style part (`xl/charts/style*.xml`).
    ///
    /// This is considered an implementation detail for now and is skipped during
    /// serde (de)serialization to avoid changing the public JSON schema.
    #[serde(default, skip)]
    pub style_part: Option<ChartStylePartModel>,
    /// Optional external chart color style part (`xl/charts/colors*.xml`).
    ///
    /// This is considered an implementation detail for now and is skipped during
    /// serde (de)serialization to avoid changing the public JSON schema.
    #[serde(default, skip)]
    pub colors_part: Option<ChartColorStylePartModel>,
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
    /// Raw XML for the chartSpace extension list (`c:chartSpace/c:extLst`).
    ///
    /// Chart extensions are currently treated as opaque blobs for forward
    /// compatibility and debugging.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_space_ext_lst_xml: Option<String>,
    /// Raw XML for the chart extension list (`c:chartSpace/c:chart/c:extLst`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_ext_lst_xml: Option<String>,
    /// Raw XML for the plotArea extension list (`c:chartSpace/c:chart/c:plotArea/c:extLst`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_ext_lst_xml: Option<String>,
    pub diagnostics: Vec<ChartDiagnostic>,
}

fn deserialize_series_vec<'de, D>(deserializer: D) -> Result<Vec<SeriesModel>, D::Error>
where
    D: serde::Deserializer<'de>,
{
    let mut series = Vec::<SeriesModel>::deserialize(deserializer)?;
    // Older serialized models (including fixtures) predate `SeriesModel::{idx, order}`.
    // When absent, Excel implies a default `idx`/`order` matching the series' position.
    for (pos, ser) in series.iter_mut().enumerate() {
        let pos_u32 = pos as u32;
        if ser.idx.is_none() {
            ser.idx = Some(pos_u32);
        }
        if ser.order.is_none() {
            ser.order = Some(pos_u32);
        }
    }
    Ok(series)
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", tag = "kind")]
pub enum ChartKind {
    Area,
    Bar,
    Bubble,
    Doughnut,
    Line,
    Pie,
    Radar,
    Scatter,
    Stock,
    Surface,
    Unknown { name: String },
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextModel {
    pub rich_text: RichText,
    pub formula: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<TextRunStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub box_style: Option<ShapeStyle>,
    /// Manual layout (used by chart titles and other chart text elements).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<ManualLayoutModel>,
}

impl TextModel {
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            rich_text: RichText::new(text),
            formula: None,
            style: None,
            box_style: None,
            layout: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LegendModel {
    pub position: LegendPosition,
    pub overlay: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_style: Option<TextRunStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ShapeStyle>,
    /// Manual layout (`c:legend/c:layout/c:manualLayout`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<ManualLayoutModel>,
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
    Area(AreaChartModel),
    Bar(BarChartModel),
    Bubble(BubbleChartModel),
    Doughnut(DoughnutChartModel),
    Line(LineChartModel),
    Pie(PieChartModel),
    Radar(RadarChartModel),
    Scatter(ScatterChartModel),
    Stock(StockChartModel),
    Surface(SurfaceChartModel),
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
    Area {
        #[serde(flatten)]
        model: AreaChartModel,
        series: SeriesIndexRange,
    },
    Bar {
        #[serde(flatten)]
        model: BarChartModel,
        series: SeriesIndexRange,
    },
    Bubble {
        #[serde(flatten)]
        model: BubbleChartModel,
        series: SeriesIndexRange,
    },
    Doughnut {
        #[serde(flatten)]
        model: DoughnutChartModel,
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
    Radar {
        #[serde(flatten)]
        model: RadarChartModel,
        series: SeriesIndexRange,
    },
    Scatter {
        #[serde(flatten)]
        model: ScatterChartModel,
        series: SeriesIndexRange,
    },
    Stock {
        #[serde(flatten)]
        model: StockChartModel,
        series: SeriesIndexRange,
    },
    Surface {
        #[serde(flatten)]
        model: SurfaceChartModel,
        series: SeriesIndexRange,
    },
    Unknown {
        name: String,
        series: SeriesIndexRange,
    },
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct AreaChartModel {
    pub grouping: Option<String>,
    pub ax_ids: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct BarChartModel {
    pub vary_colors: Option<bool>,
    pub bar_direction: Option<String>,
    pub grouping: Option<String>,
    pub gap_width: Option<u16>,
    pub overlap: Option<i16>,
    pub ax_ids: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct BubbleChartModel {
    pub bubble_scale: Option<u32>,
    pub show_neg_bubbles: Option<bool>,
    pub size_represents: Option<String>,
    pub ax_ids: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct DoughnutChartModel {
    pub vary_colors: Option<bool>,
    pub first_slice_angle: Option<u32>,
    pub hole_size: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct LineChartModel {
    pub vary_colors: Option<bool>,
    pub grouping: Option<String>,
    pub ax_ids: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct PieChartModel {
    pub vary_colors: Option<bool>,
    pub first_slice_angle: Option<u32>,
    pub hole_size: Option<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct RadarChartModel {
    pub radar_style: Option<String>,
    pub ax_ids: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct ScatterChartModel {
    pub vary_colors: Option<bool>,
    pub scatter_style: Option<String>,
    pub ax_ids: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct StockChartModel {
    pub ax_ids: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct SurfaceChartModel {
    pub wireframe: Option<bool>,
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
    /// Raw XML for the axis extension list (`c:*Ax/c:extLst`).
    ///
    /// Axis extensions are currently treated as opaque blobs for forward
    /// compatibility and debugging.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst_xml: Option<String>,
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
    /// Series index (`c:ser/c:idx/@val`, `cx:ser/cx:idx/@val`).
    ///
    /// Excel uses this value as a stable identity for a series, independent of display order.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub idx: Option<u32>,
    /// Series order (`c:ser/c:order/@val`, `cx:ser/cx:order/@val`).
    ///
    /// This generally matches the series' visual ordering.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub order: Option<u32>,
    pub name: Option<TextModel>,
    pub categories: Option<SeriesTextData>,
    /// Numeric categories (`c:cat/c:numRef`, `c:cat/c:numLit`), commonly used by
    /// date/time axes. Stored separately to preserve numeric/date semantics
    /// without stringifying.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub categories_num: Option<SeriesNumberData>,
    pub values: Option<SeriesNumberData>,
    pub x_values: Option<SeriesData>,
    pub y_values: Option<SeriesData>,
    pub smooth: Option<bool>,
    pub invert_if_negative: Option<bool>,
    /// Bubble size data (`c:ser/c:bubbleSize`), for bubble charts.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_size: Option<SeriesNumberData>,
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
    /// Raw XML for the series extension list (`c:ser/c:extLst`).
    ///
    /// Series extensions are currently treated as opaque blobs for forward
    /// compatibility and debugging.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst_xml: Option<String>,
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
    /// Optional part path this diagnostic refers to (e.g. `xl/charts/chart1.xml`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub part: Option<String>,
    /// Optional XPath/location hint within `part`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub xpath: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ChartDiagnosticLevel {
    Info,
    Warning,
    Error,
}
