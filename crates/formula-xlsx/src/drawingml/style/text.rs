use formula_model::charts::TextRunStyle;
use roxmltree::Node;

use crate::drawingml::style::parse_solid_fill;

pub fn parse_txpr(node: Node<'_, '_>) -> Option<TextRunStyle> {
    if node.tag_name().name() != "txPr" {
        return None;
    }

    // Prefer default run properties from the list style if present, but allow fallback to a run
    // (`a:rPr`) or paragraph end properties (`a:endParaRPr`) for any properties that are not
    // specified on `a:defRPr`.
    let def_rpr = node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "defRPr");
    let run_rpr = node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "rPr");
    let end_para_rpr = node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "endParaRPr");

    let mut style = TextRunStyle::default();
    if let Some(def_rpr) = def_rpr {
        apply_rpr(&mut style, def_rpr);
    }
    if let Some(run_rpr) = run_rpr {
        apply_rpr_fallback(&mut style, run_rpr);
    }
    if let Some(end_para_rpr) = end_para_rpr {
        apply_rpr_fallback(&mut style, end_para_rpr);
    }

    // Bail early if we didn't find any `a:*RPr` nodes at all.
    if def_rpr.is_none() && run_rpr.is_none() && end_para_rpr.is_none() {
        return None;
    }

    if style.is_empty() {
        None
    } else {
        Some(style)
    }
}

fn apply_rpr(style: &mut TextRunStyle, rpr: Node<'_, '_>) {
    style.font_family = parse_font_family(rpr);
    style.size_100pt = rpr.attribute("sz").and_then(|v| v.parse::<u32>().ok());
    style.bold = rpr.attribute("b").and_then(parse_bool_attr);
    style.italic = rpr.attribute("i").and_then(parse_bool_attr);
    style.underline = rpr.attribute("u").and_then(parse_underline_attr);
    style.strike = rpr.attribute("strike").and_then(parse_strike_attr);
    style.baseline = rpr.attribute("baseline").and_then(|v| v.parse::<i32>().ok());
    style.color = parse_color(rpr);
}

fn apply_rpr_fallback(style: &mut TextRunStyle, rpr: Node<'_, '_>) {
    if style.font_family.is_none() {
        style.font_family = parse_font_family(rpr);
    }
    if style.size_100pt.is_none() {
        style.size_100pt = rpr.attribute("sz").and_then(|v| v.parse::<u32>().ok());
    }
    if style.bold.is_none() {
        style.bold = rpr.attribute("b").and_then(parse_bool_attr);
    }
    if style.italic.is_none() {
        style.italic = rpr.attribute("i").and_then(parse_bool_attr);
    }
    if style.underline.is_none() {
        style.underline = rpr.attribute("u").and_then(parse_underline_attr);
    }
    if style.strike.is_none() {
        style.strike = rpr.attribute("strike").and_then(parse_strike_attr);
    }
    if style.baseline.is_none() {
        style.baseline = rpr.attribute("baseline").and_then(|v| v.parse::<i32>().ok());
    }
    if style.color.is_none() {
        style.color = parse_color(rpr);
    }
}

fn parse_font_family(rpr: Node<'_, '_>) -> Option<String> {
    // DrawingML uses `typeface` on `<a:latin>` for both concrete font names and theme placeholders
    // like `+mn-lt` / `+mj-lt`. Preserve the raw string so callers can later resolve theme fonts.
    for tag in ["latin", "ea", "cs"] {
        let family = rpr
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
            .and_then(|n| n.attribute("typeface"))
            .map(str::trim)
            .filter(|s| !s.is_empty());
        if let Some(family) = family {
            return Some(family.to_string());
        }
    }
    None
}

fn parse_color(rpr: Node<'_, '_>) -> Option<formula_model::charts::ColorRef> {
    rpr.children()
        .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        .and_then(parse_solid_fill)
        .map(|f| f.color)
}

fn parse_bool_attr(v: &str) -> Option<bool> {
    match v {
        "1" | "true" | "True" | "TRUE" => Some(true),
        "0" | "false" | "False" | "FALSE" => Some(false),
        _ => None,
    }
}

fn parse_underline_attr(v: &str) -> Option<bool> {
    // ST_TextUnderlineType includes values like `none`, `sng`, `dbl`, ...
    if v.is_empty() {
        None
    } else if v == "none" {
        Some(false)
    } else {
        Some(true)
    }
}

fn parse_strike_attr(v: &str) -> Option<bool> {
    // ST_TextStrikeType includes values like `noStrike`, `sngStrike`, `dblStrike`.
    if v.is_empty() {
        None
    } else if v == "noStrike" {
        Some(false)
    } else {
        Some(true)
    }
}
