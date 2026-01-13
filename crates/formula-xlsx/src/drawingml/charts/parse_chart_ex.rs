use formula_model::charts::{
    ChartDiagnostic, ChartDiagnosticLevel, ChartKind, ChartModel, PlotAreaModel, SeriesData,
    SeriesModel, SeriesNumberData, SeriesTextData, TextModel,
};
use formula_model::RichText;
use roxmltree::{Document, Node};

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

    let kind = detect_chart_kind(&doc).unwrap_or_else(|| "unknown".to_string());
    let chart_name = format!("ChartEx:{kind}");

    let mut diagnostics = vec![ChartDiagnostic {
        level: ChartDiagnosticLevel::Warning,
        message: format!("ChartEx root <{root_name}> (ns={root_ns}) parsed as placeholder model"),
    }];

    if kind == "unknown" {
        diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Warning,
            message: "ChartEx chart kind could not be inferred".to_string(),
        });
    }

    let series = find_chart_type_node(&doc)
        .map(|chart_type_node| {
            chart_type_node
                .descendants()
                .filter(|n| {
                    n.is_element()
                        && (n.tag_name().name() == "ser" || n.tag_name().name() == "series")
                })
                .map(|n| parse_series(n, &mut diagnostics))
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
        chart_area_style: None,
        plot_area_style: None,
        diagnostics,
    })
}

fn detect_chart_kind(doc: &Document<'_>) -> Option<String> {
    find_chart_type_node(doc).map(|node| {
        let raw = node.tag_name().name();
        let base = raw.strip_suffix("Chart").unwrap_or(raw);
        lowercase_first(base)
    })
}

fn find_chart_type_node<'a>(doc: &'a Document<'a>) -> Option<Node<'a, 'a>> {
    // Heuristic: the first element whose local name ends with "Chart" (case-insensitive)
    // but isn't the generic `<chart>` container.
    doc.descendants().find(|n| {
        if !n.is_element() {
            return false;
        }
        let name = n.tag_name().name();
        let lower = name.to_ascii_lowercase();
        lower.ends_with("chart") && lower != "chart" && lower != "chartspace"
    })
}

fn parse_series(series_node: Node<'_, '_>, diagnostics: &mut Vec<ChartDiagnostic>) -> SeriesModel {
    let name = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tx")
        .and_then(|tx| parse_text_from_tx(tx));

    let categories = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "cat")
        .and_then(|cat| parse_series_text_data(cat, diagnostics));

    let values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "val")
        .and_then(|val| parse_series_number_data(val, diagnostics));

    let x_values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "xVal")
        .and_then(|x| parse_series_data(x, diagnostics));

    let y_values = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "yVal")
        .and_then(|y| parse_series_data(y, diagnostics));

    SeriesModel {
        name,
        categories,
        values,
        x_values,
        y_values,
        style: None,
        marker: None,
        points: Vec::new(),
        plot_index: None,
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
        });
    }

    None
}

fn parse_series_number_data(
    data_node: Node<'_, '_>,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> Option<SeriesNumberData> {
    let num_ref = data_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "numRef")?;
    Some(parse_num_ref(num_ref, diagnostics))
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
