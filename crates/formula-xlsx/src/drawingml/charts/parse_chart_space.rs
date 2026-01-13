use formula_model::charts::{
    AxisKind, AxisModel, AxisPosition, AxisScalingModel, BarChartModel, ChartDiagnostic,
    ChartDiagnosticLevel, ChartKind, ChartModel, ComboChartEntry, ComboPlotAreaModel, DataLabelsModel,
    LegendModel, LegendPosition, LineChartModel, NumberFormatModel, PieChartModel, PlotAreaModel,
    ScatterChartModel, SeriesData, SeriesIndexRange, SeriesModel, SeriesNumberData, SeriesPointStyle,
    SeriesTextData, TextModel,
};
use formula_model::RichText;
use roxmltree::{Document, Node};

use super::REL_NS;
use crate::drawingml::style::{parse_marker, parse_sppr, parse_txpr};

#[derive(Debug, thiserror::Error)]
pub enum ChartSpaceParseError {
    #[error("part is not valid UTF-8: {part_name}: {source}")]
    XmlNonUtf8 {
        part_name: String,
        #[source]
        source: std::str::Utf8Error,
    },
    #[error("failed to parse XML: {part_name}: {source}")]
    XmlParse {
        part_name: String,
        #[source]
        source: roxmltree::Error,
    },
    #[error("invalid XML structure: {0}")]
    XmlStructure(String),
}

pub fn parse_chart_space(
    chart_xml: &[u8],
    part_name: &str,
) -> Result<ChartModel, ChartSpaceParseError> {
    let xml = std::str::from_utf8(chart_xml).map_err(|e| ChartSpaceParseError::XmlNonUtf8 {
        part_name: part_name.to_string(),
        source: e,
    })?;
    let doc = Document::parse(xml).map_err(|e| ChartSpaceParseError::XmlParse {
        part_name: part_name.to_string(),
        source: e,
    })?;

    let mut diagnostics = Vec::new();

    let chart_space = doc.root_element();
    let style_id = child_attr(chart_space, "style", "val").and_then(|v| v.parse::<u32>().ok());
    let rounded_corners = child_attr(chart_space, "roundedCorners", "val").map(parse_ooxml_bool);
    let chart_area_style = chart_space
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(parse_sppr);

    let external_data_node = chart_space
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "externalData");
    let external_data_rel_id = external_data_node
        .and_then(|n| {
            n.attribute((REL_NS, "id"))
                .or_else(|| n.attribute("r:id"))
                .or_else(|| n.attribute("id"))
        })
        .map(str::to_string);
    let external_data_auto_update = external_data_node
        .and_then(|n| {
            n.children()
                .find(|c| c.is_element() && c.tag_name().name() == "autoUpdate")
        })
        .and_then(|n| n.attribute("val"))
        .map(parse_ooxml_bool);

    if doc
        .descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "AlternateContent")
    {
        warn(
            &mut diagnostics,
            "mc:AlternateContent encountered; content choice is not yet modeled",
        );
    }
    if doc
        .descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "extLst")
    {
        warn(
            &mut diagnostics,
            "c:extLst encountered; chart extensions are not yet modeled",
        );
    }

    let chart_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "chart")
        .ok_or_else(|| {
            ChartSpaceParseError::XmlStructure(format!("{part_name}: missing <c:chart>"))
        })?;

    let disp_blanks_as = child_attr(chart_node, "dispBlanksAs", "val").map(str::to_string);
    let plot_vis_only = child_attr(chart_node, "plotVisOnly", "val").map(parse_ooxml_bool);

    let title = parse_title(chart_node, &mut diagnostics);
    let legend = parse_legend(chart_node, &mut diagnostics);

    let Some(plot_area_node) = chart_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "plotArea")
    else {
        warn(&mut diagnostics, "missing c:plotArea");
        return Ok(ChartModel {
            chart_kind: ChartKind::Unknown {
                name: "missingPlotArea".to_string(),
            },
            title,
            legend,
            plot_area: PlotAreaModel::Unknown {
                name: "missingPlotArea".to_string(),
            },
            axes: Vec::new(),
            series: Vec::new(),
            style_id,
            rounded_corners,
            disp_blanks_as,
            plot_vis_only,
            chart_area_style,
            plot_area_style: None,
            external_data_rel_id,
            external_data_auto_update,
            diagnostics,
        });
    };

    let plot_area_style = plot_area_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(parse_sppr);

    let (chart_kind, plot_area, series) = parse_plot_area_chart(plot_area_node, &mut diagnostics);
    let axes = parse_axes(plot_area_node, &mut diagnostics);

    Ok(ChartModel {
        chart_kind,
        title,
        legend,
        plot_area,
        axes,
        series,
        style_id,
        rounded_corners,
        disp_blanks_as,
        plot_vis_only,
        chart_area_style,
        plot_area_style,
        external_data_rel_id,
        external_data_auto_update,
        diagnostics,
    })
}

fn parse_plot_area_chart(
    plot_area_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> (ChartKind, PlotAreaModel, Vec<SeriesModel>) {
    let chart_elems: Vec<_> = plot_area_node
        .children()
        .filter(|n| n.is_element())
        .filter(|n| n.tag_name().name().ends_with("Chart"))
        .collect();

    let Some(primary_chart) = chart_elems.first().copied() else {
        warn(
            diagnostics,
            "plotArea is missing a supported *Chart element",
        );
        return (
            ChartKind::Unknown {
                name: "missingChartType".to_string(),
            },
            PlotAreaModel::Unknown {
                name: "missingChartType".to_string(),
            },
            Vec::new(),
        );
    };

    let chart_kind = map_chart_kind(primary_chart.tag_name().name());

    if chart_elems.len() == 1 {
        let plot_area = parse_plot_area_model(primary_chart, &chart_kind, diagnostics);
        let series = primary_chart
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "ser")
            .map(|ser| parse_series(ser, diagnostics, None))
            .collect();
        return (chart_kind, plot_area, series);
    }

    // Combo chart: multiple chart types are present in the same plotArea.
    let mut charts = Vec::new();
    let mut series = Vec::new();

    for (plot_index, chart_node) in chart_elems.iter().copied().enumerate() {
        let subplot_kind = map_chart_kind(chart_node.tag_name().name());
        let subplot_plot_area = parse_plot_area_model(chart_node, &subplot_kind, diagnostics);

        let start = series.len();
        for ser in chart_node
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "ser")
        {
            series.push(parse_series(ser, diagnostics, Some(plot_index)));
        }
        let end = series.len();

        let series_range = SeriesIndexRange { start, end };
        let entry = match subplot_plot_area {
            PlotAreaModel::Bar(model) => ComboChartEntry::Bar {
                model,
                series: series_range,
            },
            PlotAreaModel::Line(model) => ComboChartEntry::Line {
                model,
                series: series_range,
            },
            PlotAreaModel::Pie(model) => ComboChartEntry::Pie {
                model,
                series: series_range,
            },
            PlotAreaModel::Scatter(model) => ComboChartEntry::Scatter {
                model,
                series: series_range,
            },
            PlotAreaModel::Combo(_) => unreachable!("nested combo plot area is not supported"),
            PlotAreaModel::Unknown { name } => ComboChartEntry::Unknown {
                name,
                series: series_range,
            },
        };

        charts.push(entry);
    }

    (
        chart_kind,
        PlotAreaModel::Combo(ComboPlotAreaModel { charts }),
        series,
    )
}

fn parse_plot_area_model(
    chart_node: Node<'_, '_>,
    chart_kind: &ChartKind,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> PlotAreaModel {
    match chart_kind {
        ChartKind::Bar => PlotAreaModel::Bar(BarChartModel {
            bar_direction: child_attr(chart_node, "barDir", "val").map(str::to_string),
            grouping: child_attr(chart_node, "grouping", "val").map(str::to_string),
            ax_ids: parse_ax_ids(chart_node),
        }),
        ChartKind::Line => PlotAreaModel::Line(LineChartModel {
            grouping: child_attr(chart_node, "grouping", "val").map(str::to_string),
            ax_ids: parse_ax_ids(chart_node),
        }),
        ChartKind::Pie => PlotAreaModel::Pie(PieChartModel {
            vary_colors: child_attr(chart_node, "varyColors", "val").map(parse_ooxml_bool),
            first_slice_angle: child_attr(chart_node, "firstSliceAng", "val")
                .and_then(|v| v.parse::<u32>().ok()),
        }),
        ChartKind::Scatter => PlotAreaModel::Scatter(ScatterChartModel {
            scatter_style: child_attr(chart_node, "scatterStyle", "val").map(str::to_string),
            ax_ids: parse_ax_ids(chart_node),
        }),
        ChartKind::Unknown { name } => {
            warn(
                diagnostics,
                format!("unsupported chart type {name}; rendering may be incomplete"),
            );
            PlotAreaModel::Unknown { name: name.clone() }
        }
    }
}

fn parse_series(
    series_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    plot_index: Option<usize>,
) -> SeriesModel {
    let name = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tx")
        .and_then(|tx| parse_text_from_tx(tx, diagnostics, "series.tx"));

    let categories = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "cat")
        .and_then(|cat| parse_series_text_data(cat, diagnostics, "series.cat"));

    let values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "val")
        .and_then(|val| parse_series_number_data(val, diagnostics, "series.val"));

    let x_values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "xVal")
        .and_then(|x| parse_series_data(x, diagnostics, "series.xVal"));

    let y_values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "yVal")
        .and_then(|y| parse_series_data(y, diagnostics, "series.yVal"));

    let style = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(parse_sppr);

    let marker = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "marker")
        .and_then(parse_marker);

    let data_labels = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .map(|n| parse_data_labels(n, diagnostics));

    let points = series_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "dPt")
        .filter_map(parse_series_point_style)
        .collect();

    SeriesModel {
        name,
        categories,
        values,
        x_values,
        y_values,
        style,
        marker,
        data_labels,
        points,
        plot_index,
    }
}

fn parse_data_labels(
    data_labels_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> DataLabelsModel {
    let show_val = child_attr(data_labels_node, "showVal", "val").map(parse_ooxml_bool);
    let show_cat_name = child_attr(data_labels_node, "showCatName", "val").map(parse_ooxml_bool);
    let show_ser_name = child_attr(data_labels_node, "showSerName", "val").map(parse_ooxml_bool);
    let position = child_attr(data_labels_node, "dLblPos", "val").map(str::to_string);

    let num_fmt = data_labels_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numFmt")
        .and_then(|n| parse_number_format(n, diagnostics));

    DataLabelsModel {
        show_val,
        show_cat_name,
        show_ser_name,
        position,
        num_fmt,
    }
}

fn parse_series_point_style(dpt_node: Node<'_, '_>) -> Option<SeriesPointStyle> {
    let idx = dpt_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "idx")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok())?;

    let style = dpt_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(parse_sppr);

    let marker = dpt_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "marker")
        .and_then(parse_marker);

    if style.is_none() && marker.is_none() {
        return None;
    }

    Some(SeriesPointStyle { idx, style, marker })
}

fn parse_axes(
    plot_area_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Vec<AxisModel> {
    let mut axes = Vec::new();

    for axis in plot_area_node.children().filter(|n| n.is_element()) {
        let tag = axis.tag_name().name();
        if tag == "catAx" || tag == "valAx" || tag == "dateAx" || tag == "serAx" {
            if let Some(axis_model) = parse_axis(axis, diagnostics) {
                axes.push(axis_model);
            }
            continue;
        }

        // Surface and other specialized axes are currently ignored but can affect rendering.
        if tag.ends_with("Ax") {
            warn(
                diagnostics,
                format!("unsupported axis type <c:{tag}> encountered; axis will be ignored"),
            );
        }
    }

    axes
}

fn parse_axis(
    axis_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Option<AxisModel> {
    let id = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "axId")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok());

    let Some(id) = id else {
        warn(
            diagnostics,
            format!(
                "axis <c:{}> is missing c:axId/@val",
                axis_node.tag_name().name()
            ),
        );
        return None;
    };

    let kind = match axis_node.tag_name().name() {
        "catAx" => AxisKind::Category,
        "valAx" => AxisKind::Value,
        "dateAx" => AxisKind::Date,
        "serAx" => AxisKind::Series,
        _ => AxisKind::Unknown,
    };

    let position = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "axPos")
        .and_then(|n| n.attribute("val"))
        .map(|v| parse_axis_position(v, diagnostics));

    let scaling = parse_axis_scaling(axis_node, diagnostics);

    let num_fmt = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numFmt")
        .and_then(|n| parse_number_format(n, diagnostics));

    let tick_label_position = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tickLblPos")
        .and_then(|n| n.attribute("val"))
        .map(str::to_string);

    let major_gridlines = axis_node
        .children()
        .any(|n| n.is_element() && n.tag_name().name() == "majorGridlines");

    let axis_line_style = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(parse_sppr);

    let major_gridlines_style = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "majorGridlines")
        .and_then(|n| {
            n.children()
                .find(|c| c.is_element() && c.tag_name().name() == "spPr")
        })
        .and_then(parse_sppr);

    let minor_gridlines_style = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "minorGridlines")
        .and_then(|n| {
            n.children()
                .find(|c| c.is_element() && c.tag_name().name() == "spPr")
        })
        .and_then(parse_sppr);

    let tick_label_text_style = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "txPr")
        .and_then(parse_txpr);

    let cross_axis_id = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "crossAx")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok());
    let crosses = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "crosses")
        .and_then(|n| n.attribute("val"))
        .map(str::to_string);
    let crosses_at = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "crossesAt")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());

    let major_tick_mark = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "majorTickMark")
        .and_then(|n| n.attribute("val"))
        .map(str::to_string);
    let minor_tick_mark = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "minorTickMark")
        .and_then(|n| n.attribute("val"))
        .map(str::to_string);

    let major_unit = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "majorUnit")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    let minor_unit = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "minorUnit")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());

    let title = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "title")
        .and_then(|title| parse_axis_title(title, diagnostics, id));

    Some(AxisModel {
        id,
        kind,
        position,
        scaling,
        num_fmt,
        tick_label_position,
        major_gridlines,
        axis_line_style,
        major_gridlines_style,
        minor_gridlines_style,
        tick_label_text_style,
        cross_axis_id,
        crosses,
        crosses_at,
        major_tick_mark,
        minor_tick_mark,
        major_unit,
        minor_unit,
        title,
    })
}

fn parse_axis_title(
    title_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    axis_id: u32,
) -> Option<TextModel> {
    let tx_node = title_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tx");
    let context = format!("axis[{axis_id}].title.tx");
    let mut parsed = tx_node.and_then(|tx| parse_text_from_tx(tx, diagnostics, &context));

    let style = title_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "txPr")
        .and_then(parse_txpr);
    if let (Some(style), Some(text)) = (style, parsed.as_mut()) {
        text.style = Some(style);
    }

    parsed
}

fn parse_axis_position(value: &str, diagnostics: &mut Vec<ChartDiagnostic>) -> AxisPosition {
    match value {
        "l" => AxisPosition::Left,
        "r" => AxisPosition::Right,
        "t" => AxisPosition::Top,
        "b" => AxisPosition::Bottom,
        other => {
            warn(
                diagnostics,
                format!("unsupported axis position axPos={other:?}"),
            );
            AxisPosition::Unknown
        }
    }
}

fn parse_axis_scaling(
    axis_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> AxisScalingModel {
    let Some(scaling_node) = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "scaling")
    else {
        return AxisScalingModel::default();
    };

    let min = child_attr(scaling_node, "min", "val").and_then(|v| v.parse::<f64>().ok());
    let max = child_attr(scaling_node, "max", "val").and_then(|v| v.parse::<f64>().ok());
    let log_base = child_attr(scaling_node, "logBase", "val").and_then(|v| v.parse::<f64>().ok());

    let reverse = child_attr(scaling_node, "orientation", "val").map_or(false, |v| {
        if v == "maxMin" {
            true
        } else if v == "minMax" {
            false
        } else {
            warn(
                diagnostics,
                format!("unsupported axis orientation scaling/orientation={v:?}"),
            );
            false
        }
    });

    AxisScalingModel {
        min,
        max,
        log_base,
        reverse,
    }
}

fn parse_number_format(
    num_fmt_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Option<NumberFormatModel> {
    let Some(format_code) = num_fmt_node.attribute("formatCode") else {
        warn(diagnostics, "c:numFmt missing formatCode attribute");
        return None;
    };

    let source_linked = num_fmt_node.attribute("sourceLinked").map(parse_ooxml_bool);

    Some(NumberFormatModel {
        format_code: format_code.to_string(),
        source_linked,
    })
}

fn parse_title(
    chart_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Option<TextModel> {
    let auto_deleted = chart_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "autoTitleDeleted")
        .and_then(|n| n.attribute("val"))
        .map(parse_ooxml_bool)
        .unwrap_or(false);

    let Some(title_node) = chart_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "title")
    else {
        return None;
    };

    let tx_node = title_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tx");

    let mut parsed = tx_node.and_then(|tx| parse_text_from_tx(tx, diagnostics, "title.tx"));
    let style = title_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "txPr")
        .and_then(parse_txpr);
    if let (Some(style), Some(text)) = (style, parsed.as_mut()) {
        text.style = Some(style);
    }

    if auto_deleted && parsed.is_none() {
        return None;
    }

    parsed
}

fn parse_legend(
    chart_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Option<LegendModel> {
    let legend_node = chart_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "legend")?;

    let position = legend_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "legendPos")
        .and_then(|n| n.attribute("val"))
        .map(|v| parse_legend_position(v, diagnostics))
        .unwrap_or(LegendPosition::Unknown);

    let overlay = legend_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "overlay")
        .and_then(|n| n.attribute("val"))
        .map(parse_ooxml_bool)
        .unwrap_or(false);

    let text_style = legend_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "txPr")
        .and_then(parse_txpr);

    Some(LegendModel {
        position,
        overlay,
        text_style,
    })
}

fn parse_legend_position(value: &str, diagnostics: &mut Vec<ChartDiagnostic>) -> LegendPosition {
    match value {
        "l" => LegendPosition::Left,
        "r" => LegendPosition::Right,
        "t" => LegendPosition::Top,
        "b" => LegendPosition::Bottom,
        "tr" => LegendPosition::TopRight,
        other => {
            warn(
                diagnostics,
                format!("unsupported legend position legendPos={other:?}"),
            );
            LegendPosition::Unknown
        }
    }
}

fn parse_text_from_tx(
    tx_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> Option<TextModel> {
    if let Some(rich_node) = tx_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "rich")
    {
        let text = collect_rich_text(rich_node);
        return Some(TextModel {
            rich_text: RichText::new(text),
            formula: None,
            style: None,
        });
    }

    if let Some(str_ref) = tx_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "strRef")
    {
        let formula = descendant_text(str_ref, "f").map(str::to_string);
        let cache = str_ref
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "strCache")
            .and_then(|cache| parse_str_cache(cache, diagnostics, context));
        let cached_value = cache.as_ref().and_then(|v| v.first()).cloned();

        return Some(TextModel {
            rich_text: RichText::new(cached_value.unwrap_or_default()),
            formula,
            style: None,
        });
    }

    if let Some(v) = tx_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
    {
        return Some(TextModel {
            rich_text: RichText::new(v),
            formula: None,
            style: None,
        });
    }

    None
}

fn collect_rich_text(rich_node: Node<'_, '_>) -> String {
    rich_node
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "t")
        .filter_map(|n| n.text())
        .collect::<String>()
}

fn parse_series_text_data(
    data_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> Option<SeriesTextData> {
    if let Some(multi_lvl_str_ref) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "multiLvlStrRef")
    {
        return Some(parse_multi_lvl_str_ref(
            multi_lvl_str_ref,
            diagnostics,
            context,
        ));
    }

    if let Some(multi_lvl_str_lit) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "multiLvlStrLit")
    {
        return Some(parse_multi_lvl_str_lit(
            multi_lvl_str_lit,
            diagnostics,
            context,
        ));
    }

    if let Some(str_ref) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "strRef")
    {
        return Some(parse_str_ref(str_ref, diagnostics, context));
    }

    if let Some(num_ref) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numRef")
    {
        warn(
            diagnostics,
            format!("{context}: numeric category axis detected; values will be stringified"),
        );
        let num = parse_num_ref(num_ref, diagnostics, context);
        let cache = num
            .cache
            .map(|vals| vals.into_iter().map(|v| v.to_string()).collect());
        return Some(SeriesTextData {
            formula: num.formula,
            cache,
            multi_cache: None,
            literal: None,
        });
    }

    if let Some(str_lit) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "strLit")
    {
        let values = parse_str_cache(str_lit, diagnostics, context);
        return Some(SeriesTextData {
            formula: None,
            cache: values.clone(),
            multi_cache: None,
            literal: values,
        });
    }

    if let Some(num_lit) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numLit")
    {
        warn(
            diagnostics,
            format!("{context}: numeric category axis detected; values will be stringified"),
        );
        let (values, _format_code) = parse_num_cache(num_lit, diagnostics, context);
        let cache = values.map(|vals| vals.into_iter().map(|v| v.to_string()).collect());
        return Some(SeriesTextData {
            formula: None,
            cache: cache.clone(),
            multi_cache: None,
            literal: cache,
        });
    }

    None
}

fn parse_multi_lvl_str_ref(
    multi_lvl_str_ref_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> SeriesTextData {
    let formula = descendant_text(multi_lvl_str_ref_node, "f").map(str::to_string);
    let multi_cache = multi_lvl_str_ref_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "multiLvlStrCache")
        .and_then(|cache| parse_multi_lvl_str_cache(cache, diagnostics, context));

    SeriesTextData {
        formula,
        cache: None,
        multi_cache,
        literal: None,
    }
}

fn parse_multi_lvl_str_lit(
    multi_lvl_str_lit_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> SeriesTextData {
    let multi_cache = parse_multi_lvl_str_cache(multi_lvl_str_lit_node, diagnostics, context);

    SeriesTextData {
        formula: None,
        cache: None,
        multi_cache,
        literal: None,
    }
}

fn parse_multi_lvl_str_cache(
    cache_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> Option<Vec<Vec<String>>> {
    let mut levels = Vec::new();

    for (lvl_idx, lvl) in cache_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "lvl")
        .enumerate()
    {
        let lvl_context = format!("{context}: multi-level label lvl[{lvl_idx}]");
        let values = lvl
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "strCache")
            .and_then(|cache| parse_str_cache(cache, diagnostics, &lvl_context))
            .or_else(|| parse_str_cache(lvl, diagnostics, &lvl_context));
        if let Some(values) = values {
            levels.push(values);
        }
    }

    if levels.is_empty() {
        None
    } else {
        Some(levels)
    }
}

fn parse_series_number_data(
    data_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> Option<SeriesNumberData> {
    if let Some(num_ref) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numRef")
    {
        return Some(parse_num_ref(num_ref, diagnostics, context));
    }

    if let Some(num_lit) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numLit")
    {
        let (cache, format_code) = parse_num_cache(num_lit, diagnostics, context);
        return Some(SeriesNumberData {
            formula: None,
            cache: cache.clone(),
            format_code,
            literal: cache,
        });
    }

    None
}

fn parse_series_data(
    data_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> Option<SeriesData> {
    if let Some(str_ref) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "strRef")
    {
        return Some(SeriesData::Text(parse_str_ref(
            str_ref,
            diagnostics,
            context,
        )));
    }

    if let Some(num_ref) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numRef")
    {
        return Some(SeriesData::Number(parse_num_ref(
            num_ref,
            diagnostics,
            context,
        )));
    }

    if let Some(str_lit) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "strLit")
    {
        let values = parse_str_cache(str_lit, diagnostics, context);
        return Some(SeriesData::Text(SeriesTextData {
            formula: None,
            cache: values.clone(),
            multi_cache: None,
            literal: values,
        }));
    }

    if let Some(num_lit) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numLit")
    {
        let (cache, format_code) = parse_num_cache(num_lit, diagnostics, context);
        return Some(SeriesData::Number(SeriesNumberData {
            formula: None,
            cache: cache.clone(),
            format_code,
            literal: cache,
        }));
    }

    None
}

fn parse_str_ref(
    str_ref_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> SeriesTextData {
    let formula = descendant_text(str_ref_node, "f").map(str::to_string);
    let cache = str_ref_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "strCache")
        .and_then(|cache| parse_str_cache(cache, diagnostics, context));

    SeriesTextData {
        formula,
        cache,
        multi_cache: None,
        literal: None,
    }
}

fn parse_num_ref(
    num_ref_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> SeriesNumberData {
    let formula = descendant_text(num_ref_node, "f").map(str::to_string);
    let (cache, format_code) = num_ref_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numCache")
        .map(|cache| parse_num_cache(cache, diagnostics, context))
        .unwrap_or((None, None));

    SeriesNumberData {
        formula,
        cache,
        format_code,
        literal: None,
    }
}

fn parse_str_cache(
    cache_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> Option<Vec<String>> {
    let pt_count = cache_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "ptCount")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok());

    let mut points = Vec::new();
    let mut max_idx = None::<usize>;

    for pt in cache_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "pt")
    {
        let idx = pt.attribute("idx").and_then(|v| v.parse::<usize>().ok());
        let Some(idx) = idx else {
            warn(
                diagnostics,
                format!("{context}: <c:pt> missing idx attribute"),
            );
            continue;
        };
        let value = pt
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "v")
            .and_then(|n| n.text())
            .unwrap_or("")
            .to_string();
        max_idx = Some(max_idx.map_or(idx, |m| m.max(idx)));
        points.push((idx, value));
    }

    if points.is_empty() {
        return None;
    }

    let inferred_len = max_idx.map(|v| v + 1).unwrap_or(0);
    let len = pt_count.unwrap_or(inferred_len);

    let mut values = vec![String::new(); len];
    let mut seen = vec![false; len];
    for (idx, value) in points {
        if idx >= len {
            warn(
                diagnostics,
                format!("{context}: cache point idx={idx} exceeds ptCount={len}"),
            );
            continue;
        }
        if seen[idx] {
            warn(
                diagnostics,
                format!("{context}: duplicate cache point idx={idx}"),
            );
        }
        values[idx] = value;
        seen[idx] = true;
    }

    if let Some(expected) = pt_count {
        let missing = seen.iter().filter(|&&v| !v).count();
        if missing > 0 {
            warn(
                diagnostics,
                format!("{context}: strCache missing {missing} of {expected} points"),
            );
        }
    }

    Some(values)
}

fn parse_num_cache(
    cache_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> (Option<Vec<f64>>, Option<String>) {
    let format_code = cache_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "formatCode")
        .and_then(|n| n.text())
        .map(str::to_string);

    let pt_count = cache_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "ptCount")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok());

    let mut points = Vec::new();
    let mut max_idx = None::<usize>;

    for pt in cache_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "pt")
    {
        let idx = pt.attribute("idx").and_then(|v| v.parse::<usize>().ok());
        let Some(idx) = idx else {
            warn(
                diagnostics,
                format!("{context}: <c:pt> missing idx attribute"),
            );
            continue;
        };
        let raw = pt
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "v")
            .and_then(|n| n.text())
            .unwrap_or("")
            .trim();
        let value = match raw.parse::<f64>() {
            Ok(v) => v,
            Err(_) => {
                warn(
                    diagnostics,
                    format!("{context}: invalid numeric cache value {raw:?}"),
                );
                f64::NAN
            }
        };
        max_idx = Some(max_idx.map_or(idx, |m| m.max(idx)));
        points.push((idx, value));
    }

    if points.is_empty() {
        return (None, format_code);
    }

    let inferred_len = max_idx.map(|v| v + 1).unwrap_or(0);
    let len = pt_count.unwrap_or(inferred_len);

    let mut values = vec![f64::NAN; len];
    let mut seen = vec![false; len];
    for (idx, value) in points {
        if idx >= len {
            warn(
                diagnostics,
                format!("{context}: cache point idx={idx} exceeds ptCount={len}"),
            );
            continue;
        }
        if seen[idx] {
            warn(
                diagnostics,
                format!("{context}: duplicate cache point idx={idx}"),
            );
        }
        values[idx] = value;
        seen[idx] = true;
    }

    if let Some(expected) = pt_count {
        let missing = seen.iter().filter(|&&v| !v).count();
        if missing > 0 {
            warn(
                diagnostics,
                format!("{context}: numCache missing {missing} of {expected} points"),
            );
        }
    }

    (Some(values), format_code)
}

fn parse_ax_ids(chart_node: Node<'_, '_>) -> Vec<u32> {
    chart_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "axId")
        .filter_map(|n| n.attribute("val").and_then(|v| v.parse::<u32>().ok()))
        .collect()
}

fn map_chart_kind(name: &str) -> ChartKind {
    match name {
        "barChart" | "bar3DChart" => ChartKind::Bar,
        "lineChart" | "line3DChart" => ChartKind::Line,
        "pieChart" | "pie3DChart" | "doughnutChart" => ChartKind::Pie,
        "scatterChart" => ChartKind::Scatter,
        other => ChartKind::Unknown {
            name: other.to_string(),
        },
    }
}

fn parse_ooxml_bool(value: &str) -> bool {
    matches!(value, "1" | "true" | "True")
}

fn warn(diagnostics: &mut Vec<ChartDiagnostic>, message: impl Into<String>) {
    diagnostics.push(ChartDiagnostic {
        level: ChartDiagnosticLevel::Warning,
        message: message.into(),
    });
}

fn child_attr<'a>(node: Node<'a, 'a>, child: &str, attr: &str) -> Option<&'a str> {
    node.children()
        .find(|n| n.is_element() && n.tag_name().name() == child)
        .and_then(|n| n.attribute(attr))
}

fn descendant_text<'a>(node: Node<'a, 'a>, name: &str) -> Option<&'a str> {
    node.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == name)
        .and_then(|n| n.text())
}
