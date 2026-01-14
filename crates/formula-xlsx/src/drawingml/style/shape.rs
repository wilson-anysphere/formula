use formula_model::charts::{FillStyle, GradientFill, PatternFill, ShapeStyle, SolidFill, UnknownFill};
use roxmltree::Node;

use crate::drawingml::style::{parse_color, parse_ln};

pub fn parse_sppr(node: Node<'_, '_>) -> Option<ShapeStyle> {
    if node.tag_name().name() != "spPr" {
        return None;
    }

    let fill = node
        .children()
        .filter(|n| n.is_element())
        .find_map(parse_fill_style);

    let line = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "ln")
        .and_then(parse_ln);

    let style = ShapeStyle { fill, line };
    if style.is_empty() {
        None
    } else {
        Some(style)
    }
}

fn parse_fill_style(node: Node<'_, '_>) -> Option<FillStyle> {
    match node.tag_name().name() {
        "noFill" => Some(FillStyle::None { none: true }),
        "solidFill" => parse_solid_fill(node).map(FillStyle::Solid),
        "pattFill" => parse_patt_fill(node).map(FillStyle::Pattern),
        "gradFill" => parse_grad_fill(node).map(FillStyle::Gradient),
        other if other.ends_with("Fill") => Some(FillStyle::Unknown(UnknownFill {
            name: other.to_string(),
            raw_xml: outer_xml(node),
        })),
        _ => None,
    }
}

pub fn parse_solid_fill(node: Node<'_, '_>) -> Option<SolidFill> {
    if node.tag_name().name() != "solidFill" {
        return None;
    }

    // `a:solidFill` contains exactly one color element (e.g. `a:srgbClr`, `a:schemeClr`,
    // `a:sysClr`, `a:prstClr`, `a:scrgbClr`) and may also include non-color children like
    // `a:extLst`. Iterate until we find a supported color node.
    let color = node
        .children()
        .filter(|n| n.is_element())
        .find_map(parse_color)?;
    Some(SolidFill { color })
}

fn parse_patt_fill(node: Node<'_, '_>) -> Option<PatternFill> {
    if node.tag_name().name() != "pattFill" {
        return None;
    }

    let pattern = node.attribute("prst").unwrap_or("unknown").to_string();
    let fg_color = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "fgClr")
        .and_then(parse_color_container);
    let bg_color = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "bgClr")
        .and_then(parse_color_container);

    Some(PatternFill {
        pattern,
        fg_color,
        bg_color,
    })
}

fn parse_grad_fill(node: Node<'_, '_>) -> Option<GradientFill> {
    if node.tag_name().name() != "gradFill" {
        return None;
    }

    outer_xml(node).map(|raw_xml| GradientFill { raw_xml })
}

fn parse_color_container(node: Node<'_, '_>) -> Option<formula_model::charts::ColorRef> {
    node.children()
        .filter(|n| n.is_element())
        .find_map(parse_color)
}

fn outer_xml(node: Node<'_, '_>) -> Option<String> {
    let doc = node.document();
    let xml = doc.input_text();
    xml.get(node.range()).map(str::to_string)
}
