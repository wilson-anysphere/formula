use formula_model::charts::{MarkerShape, MarkerStyle};
use roxmltree::Node;

use crate::drawingml::style::parse_sppr;

pub fn parse_marker(node: Node<'_, '_>) -> Option<MarkerStyle> {
    if node.tag_name().name() != "marker" {
        return None;
    }

    let shape = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "symbol")
        .and_then(|n| n.attribute("val"))
        .map(parse_marker_shape);

    let size = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "size")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u8>().ok());

    let sppr = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(parse_sppr);

    let fill = sppr.as_ref().and_then(|s| {
        s.fill.as_ref().and_then(|fill| match fill {
            formula_model::charts::FillStyle::Solid(solid) => Some(solid.clone()),
            _ => None,
        })
    });
    let stroke = sppr.and_then(|s| s.line);

    let style = MarkerStyle {
        shape,
        size,
        fill,
        stroke,
    };

    if style.shape.is_none()
        && style.size.is_none()
        && style.fill.is_none()
        && style.stroke.is_none()
    {
        None
    } else {
        Some(style)
    }
}

fn parse_marker_shape(val: &str) -> MarkerShape {
    match val {
        "auto" => MarkerShape::Auto,
        "circle" => MarkerShape::Circle,
        "dash" => MarkerShape::Dash,
        "diamond" => MarkerShape::Diamond,
        "dot" => MarkerShape::Dot,
        "none" => MarkerShape::None,
        "plus" => MarkerShape::Plus,
        "square" => MarkerShape::Square,
        "star" => MarkerShape::Star,
        "triangle" => MarkerShape::Triangle,
        "x" => MarkerShape::X,
        other => MarkerShape::Unknown(other.to_string()),
    }
}
