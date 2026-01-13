use formula_model::charts::{ShapeStyle, SolidFill};
use roxmltree::Node;

use crate::drawingml::style::{parse_color, parse_ln};

pub fn parse_sppr(node: Node<'_, '_>) -> Option<ShapeStyle> {
    if node.tag_name().name() != "spPr" {
        return None;
    }

    let fill = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        .and_then(parse_solid_fill);

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
