use formula_model::charts::{ChartKind, ChartModel, ChartSeries, ChartType, SeriesData, TextModel};

use crate::drawingml::charts::{parse_chart_space, ChartSpaceParseError};
use crate::workbook::ChartExtractionError;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParsedChart {
    pub chart_type: ChartType,
    pub title: Option<String>,
    pub series: Vec<ChartSeries>,
}

/// Convert a richer [`ChartModel`] into the legacy [`ParsedChart`] shape.
///
/// The legacy pipeline only needs the chart kind, title, and formula strings.
pub(crate) fn legacy_parsed_chart_from_model(model: &ChartModel) -> ParsedChart {
    let chart_type = chart_type_from_kind(&model.chart_kind);
    let title = model.title.as_ref().and_then(legacy_text_string);
    let series = model.series.iter().map(legacy_series).collect();

    ParsedChart {
        chart_type,
        title,
        series,
    }
}

/// Legacy parser for classic `c:chartSpace` parts.
///
/// This is retained for downstream users that call `formula_xlsx::charts::parse_chart` directly,
/// but is no longer used by the `XlsxPackage::extract_charts()` pipeline.
#[deprecated(
    note = "use `drawingml::charts::parse_chart_space` or `XlsxPackage::extract_chart_objects` instead"
)]
pub fn parse_chart(
    chart_xml: &[u8],
    part_name: &str,
) -> Result<Option<ParsedChart>, ChartExtractionError> {
    match parse_chart_space(chart_xml, part_name) {
        Ok(model) => Ok(Some(legacy_parsed_chart_from_model(&model))),
        Err(ChartSpaceParseError::XmlStructure(msg)) if msg.contains("missing <c:chart>") => {
            Ok(None)
        }
        Err(err) => Err(chart_space_error_to_extraction(err)),
    }
}

fn chart_space_error_to_extraction(err: ChartSpaceParseError) -> ChartExtractionError {
    match err {
        ChartSpaceParseError::XmlNonUtf8 { part_name, source } => {
            ChartExtractionError::XmlNonUtf8(part_name, source)
        }
        ChartSpaceParseError::XmlParse { part_name, source } => {
            ChartExtractionError::XmlParse(part_name, source)
        }
        ChartSpaceParseError::XmlStructure(message) => ChartExtractionError::XmlStructure(message),
    }
}

fn chart_type_from_kind(kind: &ChartKind) -> ChartType {
    match kind {
        ChartKind::Area => ChartType::Area,
        ChartKind::Bar => ChartType::Bar,
        ChartKind::Bubble => ChartType::Unknown {
            name: "bubble".to_string(),
        },
        ChartKind::Doughnut => ChartType::Doughnut,
        ChartKind::Line => ChartType::Line,
        ChartKind::Pie => ChartType::Pie,
        ChartKind::Radar => ChartType::Unknown {
            name: "radar".to_string(),
        },
        ChartKind::Scatter => ChartType::Scatter,
        ChartKind::Stock => ChartType::Unknown {
            name: "stock".to_string(),
        },
        ChartKind::Surface => ChartType::Unknown {
            name: "surface".to_string(),
        },
        ChartKind::Unknown { name } => ChartType::Unknown { name: name.clone() },
    }
}

fn legacy_text_string(text: &TextModel) -> Option<String> {
    if let Some(formula) = &text.formula {
        Some(formula.clone())
    } else if text.rich_text.text.is_empty() {
        None
    } else {
        Some(text.rich_text.text.clone())
    }
}

fn legacy_series(series: &formula_model::charts::SeriesModel) -> ChartSeries {
    ChartSeries {
        name: series.name.as_ref().and_then(legacy_text_string),
        categories: series
            .categories
            .as_ref()
            .and_then(|data| data.formula.clone()),
        values: series.values.as_ref().and_then(|data| data.formula.clone()),
        x_values: series
            .x_values
            .as_ref()
            .and_then(|data| legacy_series_data_formula(data)),
        y_values: series
            .y_values
            .as_ref()
            .and_then(|data| legacy_series_data_formula(data)),
    }
}

fn legacy_series_data_formula(data: &SeriesData) -> Option<String> {
    match data {
        SeriesData::Text(text) => text.formula.clone(),
        SeriesData::Number(num) => num.formula.clone(),
    }
}
