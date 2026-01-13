use formula_model::charts::TextRunStyle;
use roxmltree::Node;

use crate::drawingml::style::parse_solid_fill;

pub fn parse_txpr(node: Node<'_, '_>) -> Option<TextRunStyle> {
    if node.tag_name().name() != "txPr" {
        return None;
    }

    // DrawingML text styles can come from multiple run property sources. For chart `txPr` we
    // model a single "effective" run style by merging a best-effort cascade:
    //
    //   1) list-style defaults (`a:lstStyle/*/a:defRPr`)
    //   2) direct defaults (`defRPr` directly under `txPr`)
    //   3) first paragraph defaults (`a:p/a:pPr/a:defRPr`)
    //   4) first run overrides (`a:p/a:r/a:rPr`)
    //   5) paragraph-end run props (`a:p/a:endParaRPr`)
    //
    // This matches how Excel often structures chart text, while remaining resilient to
    // missing/empty elements.
    let first_paragraph = node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "p");

    let list_style_def_rpr = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "lstStyle")
        .and_then(|lst| {
            lst.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
        });

    // Some producers (and some chart sub-elements) use a simplified `<c:txPr>` encoding that places
    // `<defRPr>` directly under `<txPr>` instead of nesting it under `<a:p>/<a:pPr>`.
    //
    // This form still carries the same run properties (`sz`, `b`, `<latin typeface=...>`, etc), so
    // treat it as an additional default run-property source.
    let direct_def_rpr = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "defRPr");

    let paragraph_def_rpr = first_paragraph.and_then(|p| {
        p.children()
            .find(|n| n.is_element() && n.tag_name().name() == "pPr")
            .and_then(|ppr| {
                ppr.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
            })
    });

    let run_rpr = first_paragraph.and_then(|p| {
        p.descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "rPr")
    });

    let end_para_rpr = first_paragraph
        .and_then(|p| p.children().find(|n| n.is_element() && n.tag_name().name() == "endParaRPr"));

    let mut style = TextRunStyle::default();
    let mut saw_rpr = false;

    for rpr in [
        list_style_def_rpr,
        direct_def_rpr,
        paragraph_def_rpr,
        run_rpr,
        end_para_rpr,
    ]
    .into_iter()
    .flatten()
    {
        saw_rpr = true;
        apply_rpr_override(&mut style, rpr);
    }

    if !saw_rpr {
        return None;
    }

    if style.is_empty() {
        None
    } else {
        Some(style)
    }
}

fn apply_rpr_override(style: &mut TextRunStyle, rpr: Node<'_, '_>) {
    if let Some(font_family) = parse_font_family(rpr) {
        style.font_family = Some(font_family);
    }
    if let Some(sz) = rpr.attribute("sz").and_then(|v| v.parse::<u32>().ok()) {
        style.size_100pt = Some(sz);
    }
    if let Some(bold) = rpr.attribute("b").and_then(parse_bool_attr) {
        style.bold = Some(bold);
    }
    if let Some(italic) = rpr.attribute("i").and_then(parse_bool_attr) {
        style.italic = Some(italic);
    }
    if let Some(underline) = rpr.attribute("u").and_then(parse_underline_attr) {
        style.underline = Some(underline);
    }
    if let Some(strike) = rpr.attribute("strike").and_then(parse_strike_attr) {
        style.strike = Some(strike);
    }
    if let Some(baseline) = rpr.attribute("baseline").and_then(|v| v.parse::<i32>().ok()) {
        style.baseline = Some(baseline);
    }
    if let Some(color) = parse_color(rpr) {
        style.color = Some(color);
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
