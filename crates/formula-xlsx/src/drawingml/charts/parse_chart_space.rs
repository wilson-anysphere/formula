use formula_model::charts::{
    AreaChartModel, AxisKind, AxisModel, AxisPosition, AxisScalingModel, BarChartModel,
    BubbleChartModel, ChartDiagnostic, ChartDiagnosticLevel, ChartKind, ChartModel,
    ComboChartEntry, ComboPlotAreaModel, DataLabelsModel, DoughnutChartModel, LegendModel,
    LegendPosition, LineChartModel, ManualLayoutModel, NumberFormatModel, PieChartModel,
    PlotAreaModel, RadarChartModel, ScatterChartModel, SeriesData, SeriesIndexRange, SeriesModel,
    SeriesNumberData, SeriesPointStyle, SeriesTextData, StockChartModel, SurfaceChartModel,
    TextModel,
};
use formula_model::rich_text::RichTextRunStyle;
use formula_model::RichText;
use roxmltree::{Document, Node};

use super::cache::{parse_num_cache, parse_num_ref, parse_str_cache, parse_str_ref};
use super::REL_NS;
use crate::drawingml::anchor::flatten_alternate_content;
use crate::drawingml::style::{parse_marker, parse_solid_fill, parse_sppr, parse_txpr};

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
    let style_id = chart_space
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_style_node))
        .find(|n| n.tag_name().name() == "style")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok());
    let rounded_corners = chart_space
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_rounded_corners_node))
        .find(|n| n.tag_name().name() == "roundedCorners")
        .and_then(|n| n.attribute("val"))
        .map(parse_ooxml_bool);
    let chart_space_ext_lst_xml = chart_space
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_ext_lst_node))
        .find(|n| n.tag_name().name() == "extLst")
        .and_then(|n| super::slice_node_xml(&n, xml))
        .filter(|s| !s.is_empty());
    let chart_area_style = chart_space
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_sppr_node))
        .find(|n| n.tag_name().name() == "spPr")
        .and_then(parse_sppr);

    let external_data_node = chart_space
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_external_data_node))
        .find(|n| is_external_data_node(*n));
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
                .filter(|c| c.is_element())
                .flat_map(|c| flatten_alternate_content(c, is_auto_update_node))
                .find(|c| c.tag_name().name() == "autoUpdate")
        })
        // `<c:autoUpdate>` is a CT_Boolean where the `@val` attribute is optional (default=true).
        .map(|n| n.attribute("val").map_or(true, parse_ooxml_bool));

    if doc
        .descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "AlternateContent")
    {
        warn(
            &mut diagnostics,
            "mc:AlternateContent encountered; Choice/Fallback selection is heuristic (mc:Choice/@Requires is ignored)",
        );
    }
    if doc
        .descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "extLst")
    {
        warn(
            &mut diagnostics,
            "c:extLst encountered; extension blocks were captured as raw XML but are not yet modeled",
        );
    }

    let chart_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "chart")
        .ok_or_else(|| {
            ChartSpaceParseError::XmlStructure(format!("{part_name}: missing <c:chart>"))
        })?;
    let chart_ext_lst_xml = chart_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_ext_lst_node))
        .find(|n| n.tag_name().name() == "extLst")
        .and_then(|n| super::slice_node_xml(&n, xml))
        .filter(|s| !s.is_empty());

    let disp_blanks_as = child_attr(chart_node, "dispBlanksAs", "val").map(str::to_string);
    let plot_vis_only = child_attr(chart_node, "plotVisOnly", "val").map(parse_ooxml_bool);

    let title = parse_title(chart_node, &mut diagnostics);
    let legend = parse_legend(chart_node, &mut diagnostics);

    let Some(plot_area_node) = chart_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_plot_area_node))
        .find(|n| n.tag_name().name() == "plotArea")
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
            plot_area_layout: None,
            axes: Vec::new(),
            series: Vec::new(),
            style_id,
            rounded_corners,
            disp_blanks_as,
            plot_vis_only,
            style_part: None,
            colors_part: None,
            chart_area_style,
            plot_area_style: None,
            external_data_rel_id,
            external_data_auto_update,
            chart_space_ext_lst_xml,
            chart_ext_lst_xml,
            plot_area_ext_lst_xml: None,
            diagnostics,
        });
    };

    let plot_area_ext_lst_xml = plot_area_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_ext_lst_node))
        .find(|n| n.tag_name().name() == "extLst")
        .and_then(|n| super::slice_node_xml(&n, xml))
        .filter(|s| !s.is_empty());
    let plot_area_style = plot_area_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_sppr_node))
        .find(|n| n.tag_name().name() == "spPr")
        .and_then(parse_sppr);

    let plot_area_layout = parse_layout_manual(plot_area_node, &mut diagnostics, "plotArea.layout");

    let (chart_kind, plot_area, series) = parse_plot_area_chart(plot_area_node, xml, &mut diagnostics);
    let axes = parse_axes(plot_area_node, xml, &mut diagnostics);
    warn_on_numeric_categories_with_non_numeric_axis(plot_area_node, &series, &mut diagnostics);

    Ok(ChartModel {
        chart_kind,
        title,
        legend,
        plot_area,
        plot_area_layout,
        axes,
        series,
        style_id,
        rounded_corners,
        disp_blanks_as,
        plot_vis_only,
        style_part: None,
        colors_part: None,
        chart_area_style,
        plot_area_style,
        external_data_rel_id,
        external_data_auto_update,
        chart_space_ext_lst_xml,
        chart_ext_lst_xml,
        plot_area_ext_lst_xml,
        diagnostics,
    })
}

fn parse_plot_area_chart(
    plot_area_node: Node<'_, '_>,
    xml: &str,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> (ChartKind, PlotAreaModel, Vec<SeriesModel>) {
    let chart_elems: Vec<_> = plot_area_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_plot_area_chart_node))
        .filter(|n| n.tag_name().name().ends_with("Chart"))
        .collect();

    let primary_chart = chart_elems
        .iter()
        .copied()
        .find(|n| {
            !matches!(
                map_chart_kind(n.tag_name().name()),
                ChartKind::Unknown { .. }
            )
        })
        .or_else(|| chart_elems.first().copied());

    let Some(primary_chart) = primary_chart else {
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
            .filter(|n| n.is_element())
            .flat_map(|n| flatten_alternate_content(n, is_ser_node))
            .filter(|n| n.tag_name().name() == "ser")
            .map(|ser| parse_series(ser, xml, diagnostics, None))
            .collect();
        return (chart_kind, plot_area, series);
    }

    // Combo chart: multiple chart types are present in the same plotArea.
    let mut charts = Vec::new();
    let mut series = Vec::new();

    for (plot_index, chart_node) in chart_elems.iter().copied().enumerate() {
        let raw_chart_type = chart_node.tag_name().name();
        let subplot_kind = map_chart_kind(raw_chart_type);
        let subplot_plot_area = parse_plot_area_model(chart_node, &subplot_kind, diagnostics);

        let start = series.len();
        for ser in chart_node
            .children()
            .filter(|n| n.is_element())
            .flat_map(|n| flatten_alternate_content(n, is_ser_node))
            .filter(|n| n.tag_name().name() == "ser")
        {
            series.push(parse_series(ser, xml, diagnostics, Some(plot_index)));
        }
        let end = series.len();

        let series_range = SeriesIndexRange { start, end };
        let entry = match subplot_plot_area {
            PlotAreaModel::Area(model) => ComboChartEntry::Area {
                model,
                series: series_range,
            },
            PlotAreaModel::Bar(model) => ComboChartEntry::Bar {
                model,
                series: series_range,
            },
            PlotAreaModel::Bubble(model) => ComboChartEntry::Bubble {
                model,
                series: series_range,
            },
            PlotAreaModel::Doughnut(model) => ComboChartEntry::Doughnut {
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
            PlotAreaModel::Radar(model) => ComboChartEntry::Radar {
                model,
                series: series_range,
            },
            PlotAreaModel::Scatter(model) => ComboChartEntry::Scatter {
                model,
                series: series_range,
            },
            PlotAreaModel::Stock(model) => ComboChartEntry::Stock {
                model,
                series: series_range,
            },
            PlotAreaModel::Surface(model) => ComboChartEntry::Surface {
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
        ChartKind::Area => PlotAreaModel::Area(AreaChartModel {
            grouping: child_attr(chart_node, "grouping", "val").map(str::to_string),
            ax_ids: parse_ax_ids(chart_node),
        }),
        ChartKind::Bar => PlotAreaModel::Bar(BarChartModel {
            vary_colors: child_attr(chart_node, "varyColors", "val").map(parse_ooxml_bool),
            bar_direction: child_attr(chart_node, "barDir", "val").map(str::to_string),
            grouping: child_attr(chart_node, "grouping", "val").map(str::to_string),
            gap_width: child_attr(chart_node, "gapWidth", "val")
                .and_then(|v| v.parse::<u16>().ok()),
            overlap: child_attr(chart_node, "overlap", "val").and_then(|v| v.parse::<i16>().ok()),
            ax_ids: parse_ax_ids(chart_node),
        }),
        ChartKind::Bubble => PlotAreaModel::Bubble(BubbleChartModel {
            bubble_scale: child_attr(chart_node, "bubbleScale", "val")
                .and_then(|v| v.parse::<u32>().ok()),
            show_neg_bubbles: child_attr(chart_node, "showNegBubbles", "val").map(parse_ooxml_bool),
            size_represents: child_attr(chart_node, "sizeRepresents", "val").map(str::to_string),
            ax_ids: parse_ax_ids(chart_node),
        }),
        ChartKind::Doughnut => PlotAreaModel::Doughnut(DoughnutChartModel {
            vary_colors: child_attr(chart_node, "varyColors", "val").map(parse_ooxml_bool),
            first_slice_angle: child_attr(chart_node, "firstSliceAng", "val")
                .and_then(|v| v.parse::<u32>().ok()),
            hole_size: child_attr(chart_node, "holeSize", "val")
                .and_then(|v| v.parse::<u32>().ok()),
        }),
        ChartKind::Line => PlotAreaModel::Line(LineChartModel {
            vary_colors: child_attr(chart_node, "varyColors", "val").map(parse_ooxml_bool),
            grouping: child_attr(chart_node, "grouping", "val").map(str::to_string),
            ax_ids: parse_ax_ids(chart_node),
        }),
        ChartKind::Pie => PlotAreaModel::Pie(PieChartModel {
            vary_colors: child_attr(chart_node, "varyColors", "val").map(parse_ooxml_bool),
            first_slice_angle: child_attr(chart_node, "firstSliceAng", "val")
                .and_then(|v| v.parse::<u32>().ok()),
            hole_size: child_attr(chart_node, "holeSize", "val").and_then(|v| v.parse::<u8>().ok()),
        }),
        ChartKind::Radar => PlotAreaModel::Radar(RadarChartModel {
            radar_style: child_attr(chart_node, "radarStyle", "val").map(str::to_string),
            ax_ids: parse_ax_ids(chart_node),
        }),
        ChartKind::Scatter => PlotAreaModel::Scatter(ScatterChartModel {
            vary_colors: child_attr(chart_node, "varyColors", "val").map(parse_ooxml_bool),
            scatter_style: child_attr(chart_node, "scatterStyle", "val").map(str::to_string),
            ax_ids: parse_ax_ids(chart_node),
        }),
        ChartKind::Stock => PlotAreaModel::Stock(StockChartModel {
            ax_ids: parse_ax_ids(chart_node),
        }),
        ChartKind::Surface => PlotAreaModel::Surface(SurfaceChartModel {
            wireframe: child_attr(chart_node, "wireframe", "val").map(parse_ooxml_bool),
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
    xml: &str,
    diagnostics: &mut Vec<ChartDiagnostic>,
    plot_index: Option<usize>,
) -> SeriesModel {
    let idx = parse_series_u32_child(series_node, "idx", diagnostics);
    let order = parse_series_u32_child(series_node, "order", diagnostics);
    let name = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tx")
        .and_then(|tx| parse_text_from_tx(tx, diagnostics, "series.tx"));

    let (categories, categories_num) = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "cat")
        .map(|cat| parse_series_categories(cat, diagnostics, "series.cat"))
        .unwrap_or((None, None));

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

    let bubble_size = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "bubbleSize")
        .and_then(|b| parse_series_number_data(b, diagnostics, "series.bubbleSize"));

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

    let smooth = child_attr(series_node, "smooth", "val").map(parse_ooxml_bool);
    let invert_if_negative =
        child_attr(series_node, "invertIfNegative", "val").map(parse_ooxml_bool);
    let ext_lst_xml = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "extLst")
        .and_then(|n| super::slice_node_xml(&n, xml))
        .filter(|s| !s.is_empty());

    SeriesModel {
        idx,
        order,
        name,
        categories,
        categories_num,
        values,
        x_values,
        y_values,
        smooth,
        invert_if_negative,
        bubble_size,
        style,
        marker,
        data_labels,
        points,
        plot_index,
        ext_lst_xml,
    }
}

fn parse_series_u32_child(
    series_node: Node<'_, '_>,
    child_name: &str,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Option<u32> {
    let Some(raw) = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == child_name)
        .and_then(|n| n.attribute("val"))
    else {
        return None;
    };

    match raw.parse::<u32>() {
        Ok(v) => Some(v),
        Err(_) => {
            warn(
                diagnostics,
                format!("series {child_name} is not a valid u32: {raw:?}"),
            );
            None
        }
    }
}

fn parse_data_labels(
    data_labels_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> DataLabelsModel {
    let show_val = child_bool_attr(data_labels_node, "showVal");
    let show_cat_name = child_bool_attr(data_labels_node, "showCatName");
    let show_ser_name = child_bool_attr(data_labels_node, "showSerName");
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

fn parse_series_categories(
    cat_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> (Option<SeriesTextData>, Option<SeriesNumberData>) {
    // `c:cat` can contain either string or numeric categories. Preserve numeric
    // categories separately rather than stringifying them.
    if cat_node
        .children()
        .any(|n| n.is_element() && matches!(n.tag_name().name(), "numRef" | "numLit"))
    {
        return (
            None,
            parse_series_number_data(cat_node, diagnostics, context),
        );
    }

    (parse_series_text_data(cat_node, diagnostics, context), None)
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
    xml: &str,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Vec<AxisModel> {
    let mut axes = Vec::new();
    let mut seen_ids = std::collections::HashSet::<u32>::new();

    for axis in plot_area_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_plot_area_axis_node))
    {
        let tag = axis.tag_name().name();
        if tag == "catAx" || tag == "valAx" || tag == "dateAx" || tag == "serAx" {
            if let Some(axis_model) = parse_axis(axis, xml, diagnostics) {
                if seen_ids.insert(axis_model.id) {
                    axes.push(axis_model);
                }
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

fn is_plot_area_chart_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    if !node.is_element() {
        return false;
    }
    // Prefer chart types we know how to model; if a Choice branch contains an
    // unsupported chart type but the Fallback contains a supported one, we want to
    // follow the Fallback so we can still parse something useful.
    !matches!(
        map_chart_kind(node.tag_name().name()),
        ChartKind::Unknown { .. }
    )
}

fn is_plot_area_axis_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    if !node.is_element() {
        return false;
    }
    matches!(
        node.tag_name().name(),
        "catAx" | "valAx" | "dateAx" | "serAx"
    )
}

fn is_sppr_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "spPr"
}

fn is_layout_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "layout"
}

fn is_style_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "style"
}

fn is_rounded_corners_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "roundedCorners"
}

fn is_external_data_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "externalData"
}

fn is_auto_update_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "autoUpdate"
}

fn is_ser_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "ser"
}

fn is_title_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "title"
}

fn is_legend_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "legend"
}

fn is_legend_pos_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "legendPos"
}

fn is_overlay_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "overlay"
}

fn is_tx_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "tx"
}

fn is_txpr_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "txPr"
}

fn is_plot_area_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "plotArea"
}

fn is_ext_lst_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "extLst"
}

/// Flattens `mc:AlternateContent` wrappers for the chartSpace parser.
///
/// Excel often uses markup-compatibility wrappers to conditionally include newer
/// OOXML content. For chart parsing, we treat these wrappers as transparent.
///
/// Branch selection is search-dependent: for an `mc:AlternateContent` node we
/// first try the first `mc:Choice` branch that contains a node matching `desired`
/// (searching within that branch), falling back to `mc:Fallback` branches when no
/// choice matches. This lets us successfully locate e.g. chart-type nodes in
/// `Fallback` even when `Choice` contains other (unrelated) elements.
///
/// The caller is responsible for any further filtering (e.g., `*Chart`, `*Ax`).
fn warn_on_numeric_categories_with_non_numeric_axis(
    plot_area_node: Node<'_, '_>,
    series: &[SeriesModel],
    diagnostics: &mut Vec<ChartDiagnostic>,
) {
    if !series.iter().any(|s| s.categories_num.is_some()) {
        return;
    }

    // For classic charts, numeric categories are typically paired with `<c:dateAx>`.
    // Scatter charts can have a `<c:valAx>` in the category position.
    let plot_area_children: Vec<_> = plot_area_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_plot_area_axis_node))
        .collect();
    let has_date_ax = plot_area_children
        .iter()
        .any(|n| n.is_element() && n.tag_name().name() == "dateAx");
    let has_cat_ax = plot_area_children
        .iter()
        .any(|n| n.is_element() && n.tag_name().name() == "catAx");
    let has_val_ax = plot_area_children
        .iter()
        .any(|n| n.is_element() && n.tag_name().name() == "valAx");

    let category_axis_is_numeric = has_date_ax || (!has_cat_ax && has_val_ax);
    if category_axis_is_numeric {
        return;
    }

    warn(
        diagnostics,
        "numeric series categories detected, but the category axis is not a date/value axis; rendering may interpret categories as text",
    );
}

fn parse_axis(
    axis_node: Node<'_, '_>,
    xml: &str,
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
    let ext_lst_xml = axis_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "extLst")
        .and_then(|n| super::slice_node_xml(&n, xml))
        .filter(|s| !s.is_empty());

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
        ext_lst_xml,
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

    let box_style = title_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(parse_sppr);
    if let (Some(box_style), Some(text)) = (box_style, parsed.as_mut()) {
        text.box_style = Some(box_style);
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
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_title_node))
        .find(|n| n.tag_name().name() == "title")
    else {
        return None;
    };

    let tx_node = title_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_tx_node))
        .find(|n| n.tag_name().name() == "tx");

    let mut parsed = tx_node.and_then(|tx| parse_text_from_tx(tx, diagnostics, "title.tx"));
    let style = title_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_txpr_node))
        .find(|n| n.tag_name().name() == "txPr")
        .and_then(parse_txpr);
    if let (Some(style), Some(text)) = (style, parsed.as_mut()) {
        text.style = Some(style);
    }

    let box_style = title_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_sppr_node))
        .find(|n| n.tag_name().name() == "spPr")
        .and_then(parse_sppr);
    if let (Some(box_style), Some(text)) = (box_style, parsed.as_mut()) {
        text.box_style = Some(box_style);
    }

    let layout = parse_layout_manual(title_node, diagnostics, "title.layout");
    if let (Some(layout), Some(text)) = (layout, parsed.as_mut()) {
        text.layout = Some(layout);
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
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_legend_node))
        .find(|n| n.tag_name().name() == "legend")?;

    let position = legend_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_legend_pos_node))
        .find(|n| n.tag_name().name() == "legendPos")
        .and_then(|n| n.attribute("val"))
        .map(|v| parse_legend_position(v, diagnostics))
        .unwrap_or(LegendPosition::Unknown);

    let overlay = legend_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_overlay_node))
        .find(|n| n.tag_name().name() == "overlay")
        .and_then(|n| n.attribute("val"))
        .map(parse_ooxml_bool)
        .unwrap_or(false);

    let text_style = legend_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_txpr_node))
        .find(|n| n.tag_name().name() == "txPr")
        .and_then(parse_txpr);

    let style = legend_node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_sppr_node))
        .find(|n| n.tag_name().name() == "spPr")
        .and_then(parse_sppr);

    let layout = parse_layout_manual(legend_node, diagnostics, "legend.layout");

    Some(LegendModel {
        position,
        overlay,
        text_style,
        style,
        layout,
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
        let rich_text = parse_rich_text(rich_node);
        return Some(TextModel {
            rich_text,
            formula: None,
            style: None,
            box_style: None,
            layout: None,
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
            box_style: None,
            layout: None,
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
            box_style: None,
            layout: None,
        });
    }

    None
}

pub(crate) fn parse_rich_text(rich_node: Node<'_, '_>) -> RichText {
    // Chart text is DrawingML (`a:p` + `a:r` runs). Preserve run boundaries and
    // capture basic formatting from `a:rPr` when present.
    let mut segments: Vec<(String, RichTextRunStyle)> = Vec::new();

    for p in rich_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "p")
    {
        for child in p.children().filter(|n| n.is_element()) {
            match child.tag_name().name() {
                "r" | "fld" => {
                    let (text, style) = parse_rich_text_run(child);
                    if !text.is_empty() {
                        segments.push((text, style));
                    }
                }
                _ => {}
            }
        }
    }

    if segments.is_empty() {
        // Fallback: preserve the historical behavior of concatenating all `<a:t>`
        // descendants (best-effort for weird producer output).
        let text = rich_node
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "t")
            .filter_map(|n| n.text())
            .collect::<String>();
        return RichText::new(text);
    }

    if segments.iter().all(|(_, style)| style.is_empty()) {
        RichText::new(segments.into_iter().map(|(t, _)| t).collect::<String>())
    } else {
        RichText::from_segments(segments)
    }
}

fn parse_rich_text_run(run_node: Node<'_, '_>) -> (String, RichTextRunStyle) {
    let style = run_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "rPr")
        .map(parse_a_rpr)
        .unwrap_or_default();

    let mut text = String::new();
    for t in run_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "t")
    {
        if let Some(s) = t.text() {
            text.push_str(s);
        }
    }
    (text, style)
}

fn parse_a_rpr(rpr: Node<'_, '_>) -> RichTextRunStyle {
    let mut style = RichTextRunStyle::default();

    style.bold = rpr.attribute("b").and_then(parse_drawingml_bool_attr);
    style.italic = rpr.attribute("i").and_then(parse_drawingml_bool_attr);

    style.size_100pt = rpr
        .attribute("sz")
        .and_then(|v| v.parse::<u32>().ok())
        .and_then(|v| u16::try_from(v).ok());

    style.font = rpr
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "latin")
        .and_then(|n| n.attribute("typeface"))
        .map(str::to_string);

    style.color = rpr
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        .and_then(parse_solid_fill)
        .map(|f| f.color);

    style
}

fn parse_drawingml_bool_attr(v: &str) -> Option<bool> {
    match v {
        "1" | "true" | "True" | "TRUE" => Some(true),
        "0" | "false" | "False" | "FALSE" => Some(false),
        _ => None,
    }
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

fn parse_ax_ids(chart_node: Node<'_, '_>) -> Vec<u32> {
    chart_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "axId")
        .filter_map(|n| n.attribute("val").and_then(|v| v.parse::<u32>().ok()))
        .collect()
}

fn map_chart_kind(name: &str) -> ChartKind {
    match name {
        "areaChart" | "area3DChart" => ChartKind::Area,
        "barChart" | "bar3DChart" => ChartKind::Bar,
        "bubbleChart" => ChartKind::Bubble,
        "doughnutChart" => ChartKind::Doughnut,
        "lineChart" | "line3DChart" => ChartKind::Line,
        "pieChart" | "pie3DChart" => ChartKind::Pie,
        "radarChart" => ChartKind::Radar,
        "scatterChart" => ChartKind::Scatter,
        "stockChart" => ChartKind::Stock,
        "surfaceChart" | "surface3DChart" => ChartKind::Surface,
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

fn child_bool_attr(node: Node<'_, '_>, child: &str) -> Option<bool> {
    let child_node = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == child)?;

    Some(child_node.attribute("val").map_or(true, parse_ooxml_bool))
}

fn descendant_text<'a>(node: Node<'a, 'a>, name: &str) -> Option<&'a str> {
    node.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == name)
        .and_then(|n| n.text())
}

fn parse_layout_manual(
    node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> Option<ManualLayoutModel> {
    let layout_node = node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_layout_node))
        .find(|n| n.tag_name().name() == "layout")?;
    let manual_node = layout_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "manualLayout")?;

    let model = ManualLayoutModel {
        x: parse_manual_layout_f64(manual_node, "x", diagnostics, context),
        y: parse_manual_layout_f64(manual_node, "y", diagnostics, context),
        w: parse_manual_layout_f64(manual_node, "w", diagnostics, context),
        h: parse_manual_layout_f64(manual_node, "h", diagnostics, context),
        x_mode: child_attr(manual_node, "xMode", "val").map(|v| v.trim().to_string()),
        y_mode: child_attr(manual_node, "yMode", "val").map(|v| v.trim().to_string()),
        w_mode: child_attr(manual_node, "wMode", "val").map(|v| v.trim().to_string()),
        h_mode: child_attr(manual_node, "hMode", "val").map(|v| v.trim().to_string()),
        layout_target: child_attr(manual_node, "layoutTarget", "val").map(|v| v.trim().to_string()),
    };

    if model == ManualLayoutModel::default() {
        None
    } else {
        Some(model)
    }
}

fn parse_manual_layout_f64(
    manual_node: Node<'_, '_>,
    child: &str,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> Option<f64> {
    let Some(raw) = child_attr(manual_node, child, "val") else {
        return None;
    };

    match raw.trim().parse::<f64>() {
        Ok(v) => Some(v),
        Err(_) => {
            warn(
                diagnostics,
                format!("{context}: invalid manualLayout {child}/@val {raw:?}"),
            );
            None
        }
    }
}
