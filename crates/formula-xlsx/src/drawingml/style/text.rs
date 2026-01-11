use formula_model::charts::TextRunStyle;
use roxmltree::Node;

use crate::drawingml::style::parse_solid_fill;

pub fn parse_txpr(node: Node<'_, '_>) -> Option<TextRunStyle> {
    if node.tag_name().name() != "txPr" {
        return None;
    }

    // Prefer default run properties from the list style if present.
    let rpr = node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
        .or_else(|| {
            node.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "rPr")
        })?;

    let font_family = rpr
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "latin")
        .and_then(|n| n.attribute("typeface"))
        .map(|s| s.to_string());

    let size_100pt = rpr
        .attribute("sz")
        .and_then(|v| v.parse::<u32>().ok());

    let bold = rpr.attribute("b").and_then(parse_bool_attr);
    let italic = rpr.attribute("i").and_then(parse_bool_attr);

    let color = rpr
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        .and_then(parse_solid_fill)
        .map(|f| f.color);

    let style = TextRunStyle {
        font_family,
        size_100pt,
        bold,
        italic,
        color,
    };

    if style.is_empty() {
        None
    } else {
        Some(style)
    }
}

fn parse_bool_attr(v: &str) -> Option<bool> {
    match v {
        "1" | "true" | "True" | "TRUE" => Some(true),
        "0" | "false" | "False" | "FALSE" => Some(false),
        _ => None,
    }
}
