use formula_model::charts::{
    ChartDiagnostic, ChartDiagnosticLevel, SeriesNumberData, SeriesTextData,
};
use roxmltree::Node;

pub fn parse_str_cache(
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

pub fn parse_num_cache(
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

pub fn parse_str_ref(
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

pub fn parse_num_ref(
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

fn warn(diagnostics: &mut Vec<ChartDiagnostic>, message: impl Into<String>) {
    diagnostics.push(ChartDiagnostic {
        level: ChartDiagnosticLevel::Warning,
        message: message.into(),
    });
}

fn descendant_text<'a>(node: Node<'a, 'a>, name: &str) -> Option<&'a str> {
    node.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == name)
        .and_then(|n| n.text())
}
