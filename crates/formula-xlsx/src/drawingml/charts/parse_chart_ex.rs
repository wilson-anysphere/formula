use formula_model::charts::{
    ChartDiagnostic, ChartDiagnosticLevel, ChartKind, ChartModel, LegendModel, LegendPosition,
    ManualLayoutModel, PlotAreaModel, SeriesData, SeriesModel, SeriesNumberData, SeriesTextData,
    TextModel,
};
use formula_model::RichText;
use roxmltree::{Document, Node};
use std::collections::{HashMap, HashSet};

use crate::drawingml::anchor::flatten_alternate_content;

use super::cache::{parse_num_cache, parse_num_ref, parse_str_cache, parse_str_ref};
use super::parse_chart_space::parse_rich_text;
use super::REL_NS;

#[derive(Debug, thiserror::Error)]
pub enum ChartExParseError {
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
}

/// Parse a ChartEx (`cx:*`) chart part into a best-effort [`ChartModel`].
///
/// ChartEx is used by Excel for several "modern" chart types (e.g. histogram,
/// waterfall). This parser intentionally produces a placeholder model that:
/// - preserves series formulas and cached values when present,
/// - surfaces a stable chart kind string via `ChartKind::Unknown { name:
///   \"ChartEx:<kind>\" }`,
/// - records diagnostics indicating that ChartEx is not yet fully modeled.
pub fn parse_chart_ex(
    chart_ex_xml: &[u8],
    part_name: &str,
) -> Result<ChartModel, ChartExParseError> {
    let xml = std::str::from_utf8(chart_ex_xml).map_err(|e| ChartExParseError::XmlNonUtf8 {
        part_name: part_name.to_string(),
        source: e,
    })?;
    let doc = Document::parse(xml).map_err(|e| ChartExParseError::XmlParse {
        part_name: part_name.to_string(),
        source: e,
    })?;

    let root = doc.root_element();
    let root_name = root.tag_name().name();
    let root_ns = root.tag_name().namespace().unwrap_or("<none>");

    let external_data_node = root
        .descendants()
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
        // `<c:autoUpdate>` / `<cx:autoUpdate>` is a CT_Boolean where the `@val` attribute is optional
        // (default=true).
        .map(|n| n.attribute("val").map_or(true, parse_ooxml_bool));

    let mut diagnostics = vec![ChartDiagnostic {
        level: ChartDiagnosticLevel::Warning,
        message: format!(
            "ChartEx root <{root_name}> (ns={root_ns}) parsed as placeholder model"
        ),
        part: None,
        xpath: None,
    }];

    let kind = detect_chart_kind(&doc, &mut diagnostics);
    let chart_name = format!("ChartEx:{kind}");

    let chart_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "chart");
    let title = chart_node.and_then(|chart| parse_title(chart, &mut diagnostics));
    let legend = chart_node.and_then(|chart| parse_legend(chart, &mut diagnostics));
    let plot_area_layout = chart_node
        .and_then(|chart| {
            chart
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "plotArea")
        })
        .and_then(parse_layout_manual);
    let chart_data = parse_chart_data(&doc, &mut diagnostics);

    // ChartEx series can appear either under `<cx:*Chart>` wrappers or
    // directly under `<cx:plotArea>`. Both variants are descendants of
    // `<cx:chart>`, so scan the chart subtree for series nodes and exclude
    // `<cx:chartData>` definitions.
    let mut series = Vec::new();
    for series_node in doc.descendants().filter(|n| {
        is_series_node(n) && has_ancestor_named(*n, "chart") && !has_ancestor_named(*n, "chartData")
    }) {
        series.push(parse_series(series_node, &chart_data, &mut diagnostics));
    }

    if series.is_empty() {
        // Defensive fallback: if the part omits `<cx:chart>`, parse any series
        // nodes outside of `<cx:chartData>`.
        for series_node in doc
            .descendants()
            .filter(|n| is_series_node(n) && !has_ancestor_named(*n, "chartData"))
        {
            series.push(parse_series(series_node, &chart_data, &mut diagnostics));
        }
    }

    // Keep the in-memory model aligned with the serde round-trip behavior in
    // `formula_model::charts::ChartModel`: if a series lacks explicit `idx`/`order`,
    // Excel treats the series' position in the document as the implied values.
    for (pos, ser) in series.iter_mut().enumerate() {
        let pos_u32 = pos as u32;
        if ser.idx.is_none() {
            ser.idx = Some(pos_u32);
        }
        if ser.order.is_none() {
            ser.order = Some(pos_u32);
        }
    }

    attach_part(&mut diagnostics, part_name);
    Ok(ChartModel {
        chart_kind: ChartKind::Unknown {
            name: chart_name.clone(),
        },
        title,
        legend,
        plot_area: PlotAreaModel::Unknown { name: chart_name },
        plot_area_layout,
        axes: Vec::new(),
        series,
        style_id: None,
        rounded_corners: None,
        disp_blanks_as: None,
        plot_vis_only: None,
        style_part: None,
        colors_part: None,
        chart_area_style: None,
        plot_area_style: None,
        external_data_rel_id,
        external_data_auto_update,
        chart_space_ext_lst_xml: None,
        chart_ext_lst_xml: None,
        plot_area_ext_lst_xml: None,
        diagnostics,
    })
}

fn parse_title(
    chart_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Option<TextModel> {
    // ChartEx sometimes stores a plain-text title directly as `<cx:title>My title</cx:title>`.
    // It can also store a structured title as `<cx:title><cx:tx>...</cx:tx></cx:title>`.
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

    let mut parsed = title_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tx")
        .and_then(|tx| parse_text_from_tx(tx, diagnostics, "title.tx"));

    if parsed.is_none() {
        parsed = title_node
            .text()
            .map(str::trim)
            .filter(|t| !t.is_empty())
            .map(TextModel::plain);
    }

    if parsed.is_none() {
        parsed = descendant_text(title_node, "v")
            .map(str::trim)
            .filter(|t| !t.is_empty())
            .map(TextModel::plain);
    }

    let layout = parse_layout_manual(title_node);
    if let (Some(layout), Some(text)) = (layout, parsed.as_mut()) {
        text.layout = Some(layout);
    }

    if auto_deleted && parsed.is_none() {
        None
    } else {
        parsed
    }
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

    let layout = parse_layout_manual(legend_node);

    Some(LegendModel {
        position,
        overlay,
        text_style: None,
        style: None,
        layout,
    })
}
fn detect_chart_kind(
    doc: &Document<'_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> String {
    // 1) Prefer explicit chart-type nodes like `<cx:waterfallChart>`.
    if let Some(node) = find_chart_type_node(doc) {
        let raw = node.tag_name().name();
        let base = raw.strip_suffix("Chart").unwrap_or(raw);
        return lowercase_first(base);
    }

    // 2) Some real-world ChartEx parts omit the `<*Chart>` node; in those cases the
    // chart kind is surfaced via attributes like `layoutId="treemap"` on
    // `<cx:series>`.
    if let Some(layout_id) = doc
        .descendants()
        .filter(|n| n.is_element())
        .filter(|n| {
            let name = n.tag_name().name();
            name.eq_ignore_ascii_case("series") || name.eq_ignore_ascii_case("ser")
        })
        .find_map(|series| attribute_case_insensitive(series, "layoutId"))
        .and_then(normalize_chart_ex_kind_hint)
    {
        return layout_id;
    }

    // 3) Another common hint is an explicit `chartType` attribute (seen in some
    // producer variations).
    if let Some(chart_type) = doc
        .descendants()
        .filter(|n| n.is_element())
        .find_map(|node| attribute_case_insensitive(node, "chartType"))
        .and_then(normalize_chart_ex_kind_hint)
    {
        return chart_type;
    }

    let hints = collect_chart_ex_kind_hints(doc);
    let hint_list = if hints.is_empty() {
        "<none>".to_string()
    } else {
        hints.join(", ")
    };
    let root_ns_display = doc
        .root_element()
        .tag_name()
        .namespace()
        .filter(|ns| !ns.is_empty())
        .unwrap_or("<none>");
    // 4) Unknown: capture a richer diagnostic to make it easier to debug/extend
    // detection for new ChartEx variants.
    diagnostics.push(ChartDiagnostic {
        level: ChartDiagnosticLevel::Warning,
        message: format!(
            "ChartEx chart kind could not be inferred (root ns={root_ns_display}); hints: {hint_list}"
        ),
        part: None,
        xpath: None,
    });

    "unknown".to_string()
}
fn collect_chart_ex_kind_hints(doc: &Document<'_>) -> Vec<String> {
    // This helper only feeds diagnostics when ChartEx kind detection fails.
    // Keep output small + stable (tests depend on ordering), but include enough context to extend
    // detection later.
    const MAX_HINTS: usize = 12;

    fn hint_priority(hint: &str) -> u8 {
        // Put the most specific / actionable hints first.
        if hint.starts_with("typeNode=") || hint.starts_with("typeHint=") {
            0
        } else if hint.starts_with("chartType=") {
            1
        } else if hint.starts_with("layoutId=") {
            2
        } else {
            3
        }
    }

    fn finalize(mut hints: Vec<String>) -> Vec<String> {
        // `sort_by_key` is stable, so within each priority bucket we preserve document order.
        hints.sort_by_key(|h| hint_priority(h));
        hints
    }

    // Best-effort: capture any attribute-based hints that might identify the ChartEx chart kind.
    //
    // This is diagnostic-only; keep output small + stable by:
    // - de-duplicating
    // - capping the total output size (so diagnostics remain readable)
    let mut out: Vec<String> = Vec::new();
    let mut seen: HashSet<String> = HashSet::new();

    let maybe_push = |hint: String, out: &mut Vec<String>, seen: &mut HashSet<String>| -> bool {
        if seen.insert(hint.clone()) {
            out.push(hint);
            return out.len() >= MAX_HINTS;
        }
        false
    };

    if let Some(node) = find_chart_type_node(doc) {
        let name = node.tag_name().name();
        let hint = format!("typeNode={name}");
        if maybe_push(hint, &mut out, &mut seen) {
            return finalize(out);
        }
        if let Some(kind) = normalize_chart_ex_kind_hint(name) {
            let hint = format!("typeHint={kind}");
            if maybe_push(hint, &mut out, &mut seen) {
                return finalize(out);
            }
        }
    }

    for node in doc.descendants().filter(|n| n.is_element()) {
        for attr in ["layoutId", "chartType"] {
            let Some(raw) = attribute_case_insensitive(node, attr) else {
                continue;
            };
            let raw = raw.trim();
            if raw.is_empty() {
                continue;
            }
            let value = normalize_chart_ex_kind_hint(raw).unwrap_or_else(|| raw.to_string());
            let hint = format!("{attr}={value}");
            if maybe_push(hint, &mut out, &mut seen) {
                return finalize(out);
            }
        }

        // Collect element names that look like chart type containers. This helps debug cases where
        // producers omit explicit attributes but still include type-like nodes.
        let name = node.tag_name().name();
        let lower = name.to_ascii_lowercase();
        if lower.ends_with("chart") && lower != "chart" && lower != "chartspace" {
            let hint = format!("node={name}");
            if maybe_push(hint, &mut out, &mut seen) {
                return finalize(out);
            }
        }
    }
    finalize(out)
}

fn find_chart_type_node<'a>(doc: &'a Document<'a>) -> Option<Node<'a, 'a>> {
    // Prefer explicit known chart-type nodes. Some ChartEx parts contain other
    // `*Chart`-suffixed elements (e.g. style/theme) that can appear before the
    // actual chart type.
    const KNOWN_CHART_TYPES: &[&str] = &[
        "waterfallChart",
        "histogramChart",
        "paretoChart",
        "boxWhiskerChart",
        "funnelChart",
        "regionMapChart",
        "treemapChart",
        "sunburstChart",
    ];

    if let Some(node) = doc.descendants().find(|n| {
        if !n.is_element() {
            return false;
        }
        let name = n.tag_name().name();
        KNOWN_CHART_TYPES
            .iter()
            .any(|known| known.eq_ignore_ascii_case(name))
    }) {
        return Some(node);
    }

    // Fallback heuristic: the first element whose local name ends with "Chart"
    // (case-insensitive) but isn't the generic `<chart>` container.
    doc.descendants().find(|n| {
        if !n.is_element() {
            return false;
        }
        let name = n.tag_name().name();
        let lower = name.to_ascii_lowercase();
        lower.ends_with("chart") && lower != "chart" && lower != "chartspace"
    })
}

fn is_series_node(node: &Node<'_, '_>) -> bool {
    node.is_element() && (node.tag_name().name() == "ser" || node.tag_name().name() == "series")
}

fn has_ancestor_named(node: Node<'_, '_>, name: &str) -> bool {
    node.ancestors()
        .any(|a| a.is_element() && a.tag_name().name() == name)
}

fn attribute_case_insensitive<'a>(node: Node<'a, 'a>, name: &str) -> Option<&'a str> {
    node.attribute(name).or_else(|| {
        node.attributes()
            .find(|attr| attr.name().eq_ignore_ascii_case(name))
            .map(|attr| attr.value())
    })
}

fn normalize_chart_ex_kind_hint(raw: &str) -> Option<String> {
    let raw = raw.trim();
    if raw.is_empty() {
        return None;
    }

    // Attributes may include a prefix (e.g. `cx:treemap`); keep only the local
    // portion.
    let raw = raw.split(':').last().unwrap_or(raw).trim();
    if raw.is_empty() {
        return None;
    }

    // Excel often uses camelCase identifiers, with some values sometimes ending
    // in `Chart` (e.g. `treemapChart`).
    let base = if raw.len() >= 5 && raw.to_ascii_lowercase().ends_with("chart") {
        &raw[..raw.len() - 5]
    } else {
        raw
    };

    if base.is_empty() {
        return None;
    }

    Some(lowercase_first(base))
}

#[derive(Debug, Clone, Default)]
struct ChartExDataDefinition {
    categories: Option<SeriesTextData>,
    values: Option<SeriesNumberData>,
    size: Option<SeriesNumberData>,
    x_values: Option<SeriesData>,
    y_values: Option<SeriesData>,
}

fn parse_chart_data<'a>(
    doc: &'a Document<'a>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> HashMap<String, ChartExDataDefinition> {
    let Some(chart_data) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "chartData")
    else {
        return HashMap::new();
    };

    let mut out: HashMap<String, ChartExDataDefinition> = HashMap::new();
    for data in chart_data
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "data")
    {
        let Some(id) = data.attribute("id") else {
            diagnostics.push(ChartDiagnostic {
                level: ChartDiagnosticLevel::Warning,
                message: "ChartEx <chartData> contains <data> without an id attribute".to_string(),
                part: None,
                xpath: None,
            });
            continue;
        };

        let mut def = ChartExDataDefinition::default();
        for dim in data.descendants().filter(|n| n.is_element()) {
            match dim.tag_name().name() {
                "strDim" => {
                    if dim.attribute("type") != Some("cat") {
                        continue;
                    }
                    let incoming = SeriesTextData {
                        formula: descendant_text(dim, "f").map(str::to_string),
                        cache: dim
                            .descendants()
                            .find(|n| n.is_element() && n.tag_name().name() == "strCache")
                            .and_then(|cache| {
                                parse_str_cache(cache, diagnostics, "chartData.strDim.strCache")
                            }),
                        multi_cache: None,
                        literal: None,
                    };
                    merge_series_text_data_slot(&mut def.categories, incoming);
                }
                "numDim" => {
                    let Some(typ) = dim.attribute("type") else {
                        continue;
                    };
                    let (cache, format_code) = dim
                        .descendants()
                        .find(|n| n.is_element() && n.tag_name().name() == "numCache")
                        .map(|cache| {
                            parse_num_cache(cache, diagnostics, "chartData.numDim.numCache")
                        })
                        .unwrap_or((None, None));

                    let num = SeriesNumberData {
                        formula: descendant_text(dim, "f").map(str::to_string),
                        cache,
                        format_code,
                        literal: None,
                    };
                    match typ {
                        "val" => {
                            merge_series_number_data_slot(&mut def.values, num);
                        }
                        "size" => {
                            merge_series_number_data_slot(&mut def.size, num);
                        }
                        "x" => {
                            fill_series_data_from_chart_data(
                                &mut def.x_values,
                                &Some(SeriesData::Number(num)),
                            );
                        }
                        "y" => {
                            fill_series_data_from_chart_data(
                                &mut def.y_values,
                                &Some(SeriesData::Number(num)),
                            );
                        }
                        _ => {}
                    }
                }
                _ => {}
            }
        }

        out.insert(id.to_string(), def);
    }

    out
}

fn parse_series_data_id(series_node: Node<'_, '_>) -> Option<String> {
    if let Some(id) = series_node.attribute("dataId") {
        return Some(id.to_string());
    }

    series_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dataId")
        .and_then(|n| {
            n.attribute("val")
                .or_else(|| n.text())
                .map(|v| v.to_string())
        })
}

fn parse_series_u32_child(
    series_node: Node<'_, '_>,
    child_name: &str,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Option<u32> {
    let raw = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == child_name)
        .and_then(|n| n.attribute("val").or_else(|| n.text()))
        // Fallback: some producers may encode idx/order as attributes on the series element.
        .or_else(|| attribute_case_insensitive(series_node, child_name));

    let Some(raw) = raw else {
        return None;
    };

    match raw.parse::<u32>() {
        Ok(v) => Some(v),
        Err(_) => {
            diagnostics.push(ChartDiagnostic {
                level: ChartDiagnosticLevel::Warning,
                message: format!("ChartEx series {child_name} is not a valid u32: {raw:?}"),
                part: None,
                xpath: None,
            });
            None
        }
    }
}

fn parse_series(
    series_node: Node<'_, '_>,
    chart_data: &HashMap<String, ChartExDataDefinition>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> SeriesModel {
    let idx = parse_series_u32_child(series_node, "idx", diagnostics);
    let order = parse_series_u32_child(series_node, "order", diagnostics);
    let name = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tx")
        .and_then(|tx| parse_text_from_tx(tx, diagnostics, "series.tx"));

    let (mut categories, categories_num) = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "cat")
        .map(|cat| parse_series_categories(cat, diagnostics, "series.cat"))
        .unwrap_or((None, None));

    let mut values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "val")
        .and_then(|val| parse_series_number_data(val, diagnostics, "series.val"));

    let mut x_values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "xVal")
        .and_then(|x| parse_series_data(x, diagnostics, "series.xVal"));

    let mut y_values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "yVal")
        .and_then(|y| parse_series_data(y, diagnostics, "series.yVal"));

    if !chart_data.is_empty() {
        if let Some(data_id) = parse_series_data_id(series_node) {
            if let Some(def) = chart_data.get(&data_id) {
                if categories.is_none() && categories_num.is_none() {
                    categories = def.categories.clone();
                } else if let (Some(dst), Some(src)) =
                    (categories.as_mut(), def.categories.as_ref())
                {
                    if dst.formula.is_none() {
                        dst.formula = src.formula.clone();
                    }
                    if dst.cache.is_none() {
                        dst.cache = src.cache.clone();
                    }
                    if dst.multi_cache.is_none() {
                        dst.multi_cache = src.multi_cache.clone();
                    }
                    if dst.literal.is_none() {
                        dst.literal = src.literal.clone();
                    }
                }

                let src_values = def.values.as_ref().or(def.size.as_ref());
                if values.is_none() {
                    values = src_values.cloned();
                } else if let (Some(dst), Some(src)) = (values.as_mut(), src_values) {
                    if dst.formula.is_none() {
                        dst.formula = src.formula.clone();
                    }
                    if dst.cache.is_none() {
                        dst.cache = src.cache.clone();
                    }
                    if dst.format_code.is_none() {
                        dst.format_code = src.format_code.clone();
                    }
                    if dst.literal.is_none() {
                        dst.literal = src.literal.clone();
                    }
                }

                fill_series_data_from_chart_data(&mut x_values, &def.x_values);
                fill_series_data_from_chart_data(&mut y_values, &def.y_values);
            } else {
                diagnostics.push(ChartDiagnostic {
                    level: ChartDiagnosticLevel::Warning,
                    message: format!(
                        "ChartEx series references dataId={data_id}, but no matching <chartData>/<data> was found"
                    ),
                    part: None,
                    xpath: None,
                });
            }
        }
    }

    SeriesModel {
        idx,
        order,
        name,
        categories,
        categories_num,
        values,
        x_values,
        y_values,
        smooth: None,
        invert_if_negative: None,
        bubble_size: None,
        style: None,
        marker: None,
        data_labels: None,
        points: Vec::new(),
        plot_index: None,
        ext_lst_xml: None,
    }
}

fn merge_series_text_data_slot(slot: &mut Option<SeriesTextData>, incoming: SeriesTextData) {
    let has_data = incoming.formula.is_some()
        || incoming.cache.is_some()
        || incoming.multi_cache.is_some()
        || incoming.literal.is_some();
    match slot {
        None => {
            if has_data {
                *slot = Some(incoming);
            }
        }
        Some(dst) => {
            if dst.formula.is_none() {
                dst.formula = incoming.formula;
            }
            if dst.cache.is_none() {
                dst.cache = incoming.cache;
            }
            if dst.multi_cache.is_none() {
                dst.multi_cache = incoming.multi_cache;
            }
            if dst.literal.is_none() {
                dst.literal = incoming.literal;
            }
        }
    }
}

fn merge_series_number_data_slot(slot: &mut Option<SeriesNumberData>, incoming: SeriesNumberData) {
    let has_data = incoming.formula.is_some()
        || incoming.cache.is_some()
        || incoming.format_code.is_some()
        || incoming.literal.is_some();
    match slot {
        None => {
            if has_data {
                *slot = Some(incoming);
            }
        }
        Some(dst) => {
            if dst.formula.is_none() {
                dst.formula = incoming.formula;
            }
            if dst.cache.is_none() {
                dst.cache = incoming.cache;
            }
            if dst.format_code.is_none() {
                dst.format_code = incoming.format_code;
            }
            if dst.literal.is_none() {
                dst.literal = incoming.literal;
            }
        }
    }
}

fn fill_series_data_from_chart_data(dst: &mut Option<SeriesData>, src: &Option<SeriesData>) {
    let Some(src_data) = src.as_ref() else {
        return;
    };

    let has_data = match src_data {
        SeriesData::Text(text) => {
            text.formula.is_some()
                || text.cache.is_some()
                || text.multi_cache.is_some()
                || text.literal.is_some()
        }
        SeriesData::Number(num) => {
            num.formula.is_some()
                || num.cache.is_some()
                || num.format_code.is_some()
                || num.literal.is_some()
        }
    };

    if dst.is_none() {
        if has_data {
            *dst = Some(src_data.clone());
        }
        return;
    }

    let Some(dst_data) = dst.as_mut() else {
        return;
    };

    match (dst_data, src_data) {
        (SeriesData::Text(dst_text), SeriesData::Text(src_text)) => {
            if dst_text.formula.is_none() {
                dst_text.formula = src_text.formula.clone();
            }
            if dst_text.cache.is_none() {
                dst_text.cache = src_text.cache.clone();
            }
            if dst_text.multi_cache.is_none() {
                dst_text.multi_cache = src_text.multi_cache.clone();
            }
            if dst_text.literal.is_none() {
                dst_text.literal = src_text.literal.clone();
            }
        }
        (SeriesData::Number(dst_num), SeriesData::Number(src_num)) => {
            if dst_num.formula.is_none() {
                dst_num.formula = src_num.formula.clone();
            }
            if dst_num.cache.is_none() {
                dst_num.cache = src_num.cache.clone();
            }
            if dst_num.format_code.is_none() {
                dst_num.format_code = src_num.format_code.clone();
            }
            if dst_num.literal.is_none() {
                dst_num.literal = src_num.literal.clone();
            }
        }
        _ => {}
    }
}

fn parse_series_categories(
    cat_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> (Option<SeriesTextData>, Option<SeriesNumberData>) {
    // Preserve numeric categories separately rather than stringifying them.
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

fn parse_text_from_tx(
    tx_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> Option<TextModel> {
    // Similar to classic charts: `tx/strRef/f` + `tx/strRef/strCache`, `tx/rich`, or a direct `v`.
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

    tx_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .map(|v| TextModel {
            rich_text: RichText::new(v),
            formula: None,
            style: None,
            box_style: None,
            layout: None,
        })
}

fn parse_series_text_data(
    data_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
    context: &str,
) -> Option<SeriesTextData> {
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

fn descendant_text<'a>(node: Node<'a, 'a>, name: &str) -> Option<&'a str> {
    node.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == name)
        .and_then(|n| n.text())
}

fn child_attr<'a>(node: Node<'a, 'a>, child: &str, attr: &str) -> Option<&'a str> {
    node.children()
        .find(|n| n.is_element() && n.tag_name().name() == child)
        .and_then(|n| n.attribute(attr))
}

fn is_layout_node<'a, 'input>(node: Node<'a, 'input>) -> bool {
    node.is_element() && node.tag_name().name() == "layout"
}

fn parse_layout_manual(node: Node<'_, '_>) -> Option<ManualLayoutModel> {
    let layout_node = node
        .children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, is_layout_node))
        .find(|n| n.tag_name().name() == "layout")?;
    let manual_node = layout_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "manualLayout")?;

    fn parse_f64(value: &str) -> Option<f64> {
        value.trim().parse::<f64>().ok()
    }

    let model = ManualLayoutModel {
        x: child_attr(manual_node, "x", "val").and_then(parse_f64),
        y: child_attr(manual_node, "y", "val").and_then(parse_f64),
        w: child_attr(manual_node, "w", "val").and_then(parse_f64),
        h: child_attr(manual_node, "h", "val").and_then(parse_f64),
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
fn parse_legend_position(value: &str, diagnostics: &mut Vec<ChartDiagnostic>) -> LegendPosition {
    match value {
        "l" => LegendPosition::Left,
        "r" => LegendPosition::Right,
        "t" => LegendPosition::Top,
        "b" => LegendPosition::Bottom,
        "tr" => LegendPosition::TopRight,
        other => {
            diagnostics.push(ChartDiagnostic {
                level: ChartDiagnosticLevel::Warning,
                message: format!("unsupported legend position legendPos={other:?}"),
                part: None,
                xpath: None,
            });
            LegendPosition::Unknown
        }
    }
}

fn parse_ooxml_bool(value: &str) -> bool {
    matches!(value, "1" | "true" | "True")
}

fn lowercase_first(s: &str) -> String {
    let mut chars = s.chars();
    let Some(first) = chars.next() else {
        return String::new();
    };
    first.to_lowercase().collect::<String>() + chars.as_str()
}

fn attach_part(diagnostics: &mut [ChartDiagnostic], part_name: &str) {
    let part = part_name.to_string();
    for diag in diagnostics {
        if diag.part.is_none() {
            diag.part = Some(part.clone());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_legend_position_and_overlay() {
        let xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chart>
    <cx:title><cx:tx><cx:v>My chart</cx:v></cx:tx></cx:title>
    <cx:legend>
      <cx:legendPos val="r"/>
      <cx:overlay val="1"/>
    </cx:legend>
    <cx:plotArea>
      <cx:histogramChart/>
    </cx:plotArea>
  </cx:chart>
 </cx:chartSpace>
"#;

        let model = parse_chart_ex(xml.as_bytes(), "unit-test").expect("parse");
        assert_eq!(model.title, Some(TextModel::plain("My chart")));

        let legend = model.legend.expect("legend should be parsed");
        assert_eq!(legend.position, LegendPosition::Right);
        assert!(legend.overlay);
    }

    #[test]
    fn parses_manual_layout_under_alternate_content() {
        let xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="cx">
  <cx:chart>
    <cx:legend>
      <cx:legendPos val="r"/>
      <mc:AlternateContent>
        <mc:Choice Requires="cx">
          <cx:layout>
            <cx:manualLayout>
              <cx:x val="0.1"/>
            </cx:manualLayout>
          </cx:layout>
        </mc:Choice>
        <mc:Fallback>
          <cx:layout>
            <cx:manualLayout>
              <cx:x val="0.2"/>
            </cx:manualLayout>
          </cx:layout>
        </mc:Fallback>
      </mc:AlternateContent>
    </cx:legend>
    <cx:plotArea>
      <cx:histogramChart/>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>
"#;

        let model = parse_chart_ex(xml.as_bytes(), "unit-test").expect("parse");
        let legend = model.legend.expect("legend should be parsed");
        let layout = legend.layout.expect("legend should contain layout");
        assert_eq!(layout.x, Some(0.1));
    }

    #[test]
    fn parses_manual_layout_for_title() {
        let xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chart>
    <cx:title>
      <cx:layout>
        <cx:manualLayout>
          <cx:x val="0.3"/>
        </cx:manualLayout>
      </cx:layout>
      <cx:tx><cx:v>My title</cx:v></cx:tx>
    </cx:title>
    <cx:plotArea>
      <cx:histogramChart/>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>
"#;

        let model = parse_chart_ex(xml.as_bytes(), "unit-test").expect("parse");
        let title = model.title.expect("title should be parsed");
        assert_eq!(title.rich_text.plain_text(), "My title");
        assert_eq!(title.layout.as_ref().and_then(|l| l.x), Some(0.3));
    }

    #[test]
    fn parses_manual_layout_for_plot_area() {
        let xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chart>
    <cx:plotArea>
      <cx:layout>
        <cx:manualLayout>
          <cx:y val="0.4"/>
        </cx:manualLayout>
      </cx:layout>
      <cx:histogramChart/>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>
"#;

        let model = parse_chart_ex(xml.as_bytes(), "unit-test").expect("parse");
        let layout = model.plot_area_layout.expect("plot area layout present");
        assert_eq!(layout.y, Some(0.4));
    }

    #[test]
    fn chart_ex_unknown_kind_diagnostic_includes_root_namespace() {
        let xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart/></cx:chartSpace>"#;
        let model = parse_chart_ex(xml.as_bytes(), "unit-test").expect("parse");
        assert!(
            model.diagnostics.iter().any(|d| d.message
                == "ChartEx chart kind could not be inferred (root ns=http://schemas.microsoft.com/office/drawing/2014/chartex); hints: <none>"),
            "expected kind inference diagnostic with root namespace, got: {:#?}",
            model.diagnostics
        );
    }

    #[test]
    fn chart_ex_unknown_kind_diagnostic_handles_missing_namespace() {
        let xml = r#"<chartSpace><chart/></chartSpace>"#;
        let model = parse_chart_ex(xml.as_bytes(), "unit-test").expect("parse");
        assert!(
            model.diagnostics.iter().any(|d| d.message
                == "ChartEx chart kind could not be inferred (root ns=<none>); hints: <none>"),
            "expected kind inference diagnostic with <none> namespace, got: {:#?}",
            model.diagnostics
        );
    }

    #[test]
    fn collect_chart_ex_kind_hints_normalizes_and_deduplicates() {
        let xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chart>
    <cx:plotArea>
      <cx:chartData>
        <cx:series layoutId="cx:treemapChart" chartType="WaterfallChart"/>
        <cx:series layoutId="treemap" chartType="waterfall"/>
        <cx:series layoutId="treemapChart" chartType="waterfallChart"/>
      </cx:chartData>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>"#;

        let doc = Document::parse(xml).expect("parse xml");
        let hints = collect_chart_ex_kind_hints(&doc);
        assert_eq!(
            hints,
            vec![
                // Output is sorted for stability (diagnostic-only helper).
                "chartType=waterfall".to_string(),
                "layoutId=treemap".to_string(),
            ]
        );
    }

    #[test]
    fn collect_chart_ex_kind_hints_caps_output() {
        let mut xml = String::from(
            r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart><cx:plotArea><cx:chartData>"#,
        );
        for i in 0..20 {
            xml.push_str(&format!(
                r#"<cx:series layoutId="kind{i}" chartType="type{i}"/>"#
            ));
        }
        xml.push_str("</cx:chartData></cx:plotArea></cx:chart></cx:chartSpace>");

        let doc = Document::parse(&xml).expect("parse xml");
        let hints = collect_chart_ex_kind_hints(&doc);
        assert_eq!(hints.len(), 12);
        assert_eq!(
            hints,
            vec![
                "chartType=type0",
                "chartType=type1",
                "chartType=type2",
                "chartType=type3",
                "chartType=type4",
                "chartType=type5",
                // Followed by the sorted layoutId hints.
                "layoutId=kind0",
                "layoutId=kind1",
                "layoutId=kind2",
                "layoutId=kind3",
                "layoutId=kind4",
                "layoutId=kind5",
            ]
            .into_iter()
            .map(str::to_string)
            .collect::<Vec<_>>()
        );
    }
}
