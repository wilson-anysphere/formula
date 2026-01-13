use formula_model::charts::{ChartSeries, ChartType};
use roxmltree::Document;

use crate::workbook::ChartExtractionError;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParsedChart {
    pub chart_type: ChartType,
    pub title: Option<String>,
    pub series: Vec<ChartSeries>,
}

pub fn parse_chart(chart_xml: &[u8], part_name: &str) -> Result<Option<ParsedChart>, ChartExtractionError> {
    let xml = std::str::from_utf8(chart_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;
    let doc =
        Document::parse(xml).map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;

    let Some(chart_node) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "chart")
    else {
        return Ok(None);
    };

    let title = chart_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "title")
        .and_then(extract_title);

    let plot_area = chart_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "plotArea");

    let Some(plot_area) = plot_area else {
        return Ok(Some(ParsedChart {
            chart_type: ChartType::Unknown {
                name: "missingPlotArea".to_string(),
            },
            title,
            series: Vec::new(),
        }));
    };

    let chart_elems: Vec<_> = plot_area
        .children()
        .filter(|n| n.is_element() && n.tag_name().name().ends_with("Chart"))
        .collect();

    let Some(primary_chart) = chart_elems.first().copied() else {
        return Ok(Some(ParsedChart {
            chart_type: ChartType::Unknown {
                name: "missingChartType".to_string(),
            },
            title,
            series: Vec::new(),
        }));
    };

    let chart_type = map_chart_type(primary_chart.tag_name().name());

    // Some Excel charts (combo charts) contain multiple chart type nodes (e.g.
    // `<c:barChart>` + `<c:lineChart>`) within the same plotArea. When present,
    // preserve series from all chart nodes, in document order.
    let series = chart_elems
        .iter()
        .flat_map(|chart| {
            chart
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "ser")
                .map(parse_series)
        })
        .collect();

    Ok(Some(ParsedChart {
        chart_type,
        title,
        series,
    }))
}

fn extract_title(title_node: roxmltree::Node<'_, '_>) -> Option<String> {
    let rich_text: String = title_node
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "t")
        .filter_map(|n| n.text())
        .collect();

    if !rich_text.is_empty() {
        return Some(rich_text);
    }

    title_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "f")
        .and_then(|n| n.text())
        .map(|s| s.to_string())
}

fn parse_series(series_node: roxmltree::Node<'_, '_>) -> ChartSeries {
    ChartSeries {
        name: series_node
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "tx")
            .and_then(extract_tx),
        categories: extract_formula(series_node, "cat"),
        values: extract_formula(series_node, "val"),
        x_values: extract_formula(series_node, "xVal"),
        y_values: extract_formula(series_node, "yVal"),
    }
}

fn extract_tx(tx_node: roxmltree::Node<'_, '_>) -> Option<String> {
    tx_node
        .descendants()
        .find(|n| n.is_element() && (n.tag_name().name() == "f" || n.tag_name().name() == "v"))
        .and_then(|n| n.text())
        .map(|s| s.to_string())
}

fn extract_formula(series_node: roxmltree::Node<'_, '_>, series_child: &str) -> Option<String> {
    series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == series_child)
        .and_then(|n| {
            n.descendants()
                .find(|d| d.is_element() && d.tag_name().name() == "f")
        })
        .and_then(|n| n.text())
        .map(|s| s.to_string())
}

fn map_chart_type(name: &str) -> ChartType {
    match name {
        "areaChart" | "area3DChart" => ChartType::Area,
        "barChart" | "bar3DChart" => ChartType::Bar,
        "doughnutChart" => ChartType::Doughnut,
        "lineChart" | "line3DChart" => ChartType::Line,
        "pieChart" | "pie3DChart" => ChartType::Pie,
        "scatterChart" => ChartType::Scatter,
        other => ChartType::Unknown {
            name: other.to_string(),
        },
    }
}
