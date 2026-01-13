use serde::{Deserialize, Serialize};

mod formatting;
mod model;
mod style_parts;
pub use formatting::*;
pub use model::*;
pub use style_parts::*;

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", tag = "kind")]
pub enum ChartType {
    Area,
    Bar,
    Doughnut,
    Line,
    Pie,
    Scatter,
    Unknown { name: String },
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    pub name: Option<String>,
    pub categories: Option<String>,
    pub values: Option<String>,
    pub x_values: Option<String>,
    pub y_values: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", tag = "kind")]
pub enum ChartAnchor {
    TwoCell {
        from_col: u32,
        from_row: u32,
        from_col_off_emu: i64,
        from_row_off_emu: i64,
        to_col: u32,
        to_row: u32,
        to_col_off_emu: i64,
        to_row_off_emu: i64,
    },
    OneCell {
        from_col: u32,
        from_row: u32,
        from_col_off_emu: i64,
        from_row_off_emu: i64,
        cx_emu: i64,
        cy_emu: i64,
    },
    Absolute {
        x_emu: i64,
        y_emu: i64,
        cx_emu: i64,
        cy_emu: i64,
    },
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Chart {
    pub sheet_name: Option<String>,
    pub sheet_part: Option<String>,
    pub drawing_part: String,
    pub chart_part: Option<String>,
    pub rel_id: String,
    pub chart_type: ChartType,
    pub title: Option<String>,
    pub series: Vec<ChartSeries>,
    pub anchor: ChartAnchor,
}
