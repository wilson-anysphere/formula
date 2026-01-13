use formula_model::charts::{
    ChartDiagnostic, ChartDiagnosticLevel, ChartKind, ChartModel, PlotAreaModel, SeriesData,
    SeriesModel, SeriesNumberData, SeriesTextData, TextModel,
};
use formula_model::RichText;
use roxmltree::{Document, Node};
use std::collections::{BTreeSet, HashMap};

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
    let root_ns = root.tag_name().namespace().unwrap_or("");

    let mut diagnostics = vec![ChartDiagnostic {
        level: ChartDiagnosticLevel::Warning,
        message: format!("ChartEx root <{root_name}> (ns={root_ns}) parsed as placeholder model"),
    }];

    let kind = detect_chart_kind(&doc, &mut diagnostics);
    let chart_name = format!("ChartEx:{kind}");

    let chart_data = parse_chart_data(&doc, &mut diagnostics);

    let series = find_chart_type_node(&doc)
        .map(|chart_type_node| {
            chart_type_node
                .descendants()
                .filter(|n| {
                    n.is_element()
                        && (n.tag_name().name() == "ser" || n.tag_name().name() == "series")
                })
                .map(|n| parse_series(n, &chart_data, &mut diagnostics))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();

    Ok(ChartModel {
        chart_kind: ChartKind::Unknown {
            name: chart_name.clone(),
        },
        title: None,
        legend: None,
        plot_area: PlotAreaModel::Unknown { name: chart_name },
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
        external_data_rel_id: None,
        external_data_auto_update: None,
        diagnostics,
    })
}

fn detect_chart_kind(doc: &Document<'_>, diagnostics: &mut Vec<ChartDiagnostic>) -> String {
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

    // 4) Unknown: capture a richer diagnostic to make it easier to debug/extend
    // detection for new ChartEx variants.
    let root_ns = doc.root_element().tag_name().namespace().unwrap_or("");
    let hints = collect_chart_ex_kind_hints(doc);
    let hint_list = if hints.is_empty() {
        "<none>".to_string()
    } else {
        hints.join(", ")
    };
    diagnostics.push(ChartDiagnostic {
        level: ChartDiagnosticLevel::Warning,
        message: format!(
            "ChartEx chart kind could not be inferred (root ns={root_ns}); hints: {hint_list}"
        ),
    });

    "unknown".to_string()
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

fn collect_chart_ex_kind_hints(doc: &Document<'_>) -> Vec<String> {
    let mut hints = BTreeSet::new();

    for node in doc.descendants().filter(|n| n.is_element()) {
        let name = node.tag_name().name();
        let lower = name.to_ascii_lowercase();
        if lower.ends_with("chart") && lower != "chart" && lower != "chartspace" {
            hints.insert(format!("node:{name}"));
        }

        for attr in node.attributes() {
            let attr_name = attr.name();
            if attr_name.eq_ignore_ascii_case("layoutId") {
                hints.insert(format!("layoutId={}", attr.value()));
            } else if attr_name.eq_ignore_ascii_case("chartType") {
                hints.insert(format!("chartType={}", attr.value()));
            }
        }
    }

    hints.into_iter().collect()
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
                    let Some(f) = descendant_text(dim, "f") else {
                        continue;
                    };
                    if def.categories.is_none() {
                        def.categories = Some(SeriesTextData {
                            formula: Some(f.to_string()),
                            cache: None,
                            multi_cache: None,
                            literal: None,
                        });
                    }
                }
                "numDim" => {
                    let Some(typ) = dim.attribute("type") else {
                        continue;
                    };
                    let Some(f) = descendant_text(dim, "f") else {
                        continue;
                    };
                    let num = SeriesNumberData {
                        formula: Some(f.to_string()),
                        cache: None,
                        format_code: None,
                        literal: None,
                    };
                    match typ {
                        "val" => {
                            if def.values.is_none() {
                                def.values = Some(num);
                            }
                        }
                        "size" => {
                            if def.size.is_none() {
                                def.size = Some(num);
                            }
                        }
                        "x" => {
                            if def.x_values.is_none() {
                                def.x_values = Some(SeriesData::Number(num));
                            }
                        }
                        "y" => {
                            if def.y_values.is_none() {
                                def.y_values = Some(SeriesData::Number(num));
                            }
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

fn parse_series(
    series_node: Node<'_, '_>,
    chart_data: &HashMap<String, ChartExDataDefinition>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> SeriesModel {
    let name = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tx")
        .and_then(|tx| parse_text_from_tx(tx));

    let mut categories = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "cat")
        .and_then(|cat| parse_series_text_data(cat, diagnostics));

    let mut values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "val")
        .and_then(|val| parse_series_number_data(val, diagnostics));

    let mut x_values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "xVal")
        .and_then(|x| parse_series_data(x, diagnostics));

    let mut y_values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "yVal")
        .and_then(|y| parse_series_data(y, diagnostics));

    if !chart_data.is_empty() {
        if let Some(data_id) = parse_series_data_id(series_node) {
            if let Some(def) = chart_data.get(&data_id) {
                if categories.is_none() {
                    categories = def.categories.clone();
                } else if let (Some(dst), Some(src)) =
                    (categories.as_mut(), def.categories.as_ref())
                {
                    if dst.formula.is_none() {
                        dst.formula = src.formula.clone();
                    }
                }

                let src_values = def.values.as_ref().or(def.size.as_ref());
                if values.is_none() {
                    values = src_values.cloned();
                } else if let (Some(dst), Some(src)) = (values.as_mut(), src_values) {
                    if dst.formula.is_none() {
                        dst.formula = src.formula.clone();
                    }
                }

                fill_series_data_formula(&mut x_values, &def.x_values);
                fill_series_data_formula(&mut y_values, &def.y_values);
            } else {
                diagnostics.push(ChartDiagnostic {
                    level: ChartDiagnosticLevel::Warning,
                    message: format!(
                        "ChartEx series references dataId={data_id}, but no matching <chartData>/<data> was found"
                    ),
                });
            }
        }
    }

    SeriesModel {
        name,
        categories,
        values,
        x_values,
        y_values,
        smooth: None,
        invert_if_negative: None,
        style: None,
        marker: None,
        data_labels: None,
        points: Vec::new(),
        plot_index: None,
    }
}

fn fill_series_data_formula(dst: &mut Option<SeriesData>, src: &Option<SeriesData>) {
    if dst.is_none() {
        *dst = src.clone();
        return;
    }

    let (Some(dst_data), Some(src_data)) = (dst.as_mut(), src.as_ref()) else {
        return;
    };

    match (dst_data, src_data) {
        (SeriesData::Text(dst_text), SeriesData::Text(src_text)) => {
            if dst_text.formula.is_none() {
                dst_text.formula = src_text.formula.clone();
            }
        }
        (SeriesData::Number(dst_num), SeriesData::Number(src_num)) => {
            if dst_num.formula.is_none() {
                dst_num.formula = src_num.formula.clone();
            }
        }
        _ => {}
    }
}

fn parse_text_from_tx(tx_node: Node<'_, '_>) -> Option<TextModel> {
    // Similar to classic charts: `tx/strRef/f` + `tx/strRef/strCache` or a direct `v`.
    if let Some(str_ref) = tx_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "strRef")
    {
        let formula = descendant_text(str_ref, "f").map(str::to_string);
        let cache = str_ref
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "strCache")
            .and_then(parse_str_cache);
        let cached_value = cache.as_ref().and_then(|v| v.first()).cloned();
        return Some(TextModel {
            rich_text: RichText::new(cached_value.unwrap_or_default()),
            formula,
            style: None,
            box_style: None,
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
        })
}

fn parse_series_text_data(
    data_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Option<SeriesTextData> {
    if let Some(str_ref) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "strRef")
    {
        return Some(parse_str_ref(str_ref));
    }

    if let Some(num_ref) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numRef")
    {
        let num = parse_num_ref(num_ref, diagnostics);
        let cache = num
            .cache
            .as_ref()
            .map(|vals| vals.iter().map(|v| v.to_string()).collect());
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
        let values = parse_str_cache(str_lit);
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
        let (values, _format_code) = parse_num_cache(num_lit, diagnostics);
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

fn parse_series_number_data(
    data_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Option<SeriesNumberData> {
    if let Some(num_ref) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numRef")
    {
        return Some(parse_num_ref(num_ref, diagnostics));
    }

    if let Some(num_lit) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numLit")
    {
        let (cache, format_code) = parse_num_cache(num_lit, diagnostics);
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
) -> Option<SeriesData> {
    if let Some(str_ref) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "strRef")
    {
        return Some(SeriesData::Text(parse_str_ref(str_ref)));
    }

    if let Some(num_ref) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numRef")
    {
        return Some(SeriesData::Number(parse_num_ref(num_ref, diagnostics)));
    }

    if let Some(str_lit) = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "strLit")
    {
        let values = parse_str_cache(str_lit);
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
        let (cache, format_code) = parse_num_cache(num_lit, diagnostics);
        return Some(SeriesData::Number(SeriesNumberData {
            formula: None,
            cache: cache.clone(),
            format_code,
            literal: cache,
        }));
    }

    None
}

fn parse_str_ref(str_ref_node: Node<'_, '_>) -> SeriesTextData {
    let formula = descendant_text(str_ref_node, "f").map(str::to_string);
    let cache = str_ref_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "strCache")
        .and_then(parse_str_cache);

    SeriesTextData {
        formula,
        cache,
        literal: None,
        multi_cache: None,
    }
}

fn parse_num_ref(
    num_ref_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> SeriesNumberData {
    let formula = descendant_text(num_ref_node, "f").map(str::to_string);
    let (cache, format_code) = num_ref_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numCache")
        .map(|cache| parse_num_cache(cache, diagnostics))
        .unwrap_or((None, None));

    SeriesNumberData {
        formula,
        cache,
        format_code,
        literal: None,
    }
}

fn parse_str_cache(cache_node: Node<'_, '_>) -> Option<Vec<String>> {
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
    for (idx, value) in points {
        if idx < len {
            values[idx] = value;
        }
    }
    Some(values)
}

fn parse_num_cache(
    cache_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
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
        let Some(idx) = idx else { continue };
        let raw = pt
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "v")
            .and_then(|n| n.text())
            .unwrap_or("")
            .trim();
        let value = match raw.parse::<f64>() {
            Ok(v) => v,
            Err(_) => {
                diagnostics.push(ChartDiagnostic {
                    level: ChartDiagnosticLevel::Warning,
                    message: format!("invalid numeric cache value {raw:?}"),
                });
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
    for (idx, value) in points {
        if idx < len {
            values[idx] = value;
        }
    }
    (Some(values), format_code)
}

fn descendant_text<'a>(node: Node<'a, 'a>, name: &str) -> Option<&'a str> {
    node.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == name)
        .and_then(|n| n.text())
}

fn lowercase_first(s: &str) -> String {
    let mut chars = s.chars();
    let Some(first) = chars.next() else {
        return String::new();
    };
    first.to_lowercase().collect::<String>() + chars.as_str()
}
