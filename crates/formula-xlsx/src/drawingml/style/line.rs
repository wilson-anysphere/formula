use formula_model::charts::{LineDash, LineStyle};
use roxmltree::Node;

use crate::drawingml::style::parse_solid_fill;

pub fn parse_ln(node: Node<'_, '_>) -> Option<LineStyle> {
    if node.tag_name().name() != "ln" {
        return None;
    }

    let width_100pt = node
        .attribute("w")
        .and_then(|v| v.parse::<u64>().ok())
        .map(|w_emu| ((w_emu * 100) + 6_350) / 12_700)
        .map(|v| v as u32);

    let color = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        .and_then(parse_solid_fill)
        .map(|fill| fill.color);

    let dash = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "prstDash")
        .and_then(|d| d.attribute("val"))
        .map(parse_dash);

    let style = LineStyle {
        color,
        width_100pt,
        dash,
    };

    if style.color.is_none() && style.width_100pt.is_none() && style.dash.is_none() {
        None
    } else {
        Some(style)
    }
}

fn parse_dash(val: &str) -> LineDash {
    match val {
        "solid" => LineDash::Solid,
        "dash" => LineDash::Dash,
        "dot" => LineDash::Dot,
        "dashDot" => LineDash::DashDot,
        "lgDash" => LineDash::LongDash,
        "lgDashDot" => LineDash::LongDashDot,
        "lgDashDotDot" => LineDash::LongDashDotDot,
        "sysDash" => LineDash::SysDash,
        "sysDot" => LineDash::SysDot,
        "sysDashDot" => LineDash::SysDashDot,
        "sysDashDotDot" => LineDash::SysDashDotDot,
        other => LineDash::Unknown(other.to_string()),
    }
}
