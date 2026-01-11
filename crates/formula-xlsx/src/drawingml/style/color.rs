use formula_model::{charts::ColorRef, Color};
use roxmltree::Node;

pub fn parse_color(node: Node<'_, '_>) -> Option<ColorRef> {
    match node.tag_name().name() {
        "srgbClr" => node.attribute("val").and_then(parse_srgb).map(Color::Argb),
        "schemeClr" => parse_scheme(node),
        _ => None,
    }
}

fn parse_scheme(node: Node<'_, '_>) -> Option<ColorRef> {
    let scheme = node.attribute("val")?;
    let theme = scheme_to_theme_index(scheme)?;
    let tint = parse_tint_thousandths(node);
    Some(Color::Theme { theme, tint })
}

fn parse_srgb(val: &str) -> Option<u32> {
    let hex = val.trim().strip_prefix('#').unwrap_or(val.trim());
    match hex.len() {
        6 => u32::from_str_radix(hex, 16)
            .ok()
            .map(|rgb| 0xFF00_0000 | rgb),
        8 => u32::from_str_radix(hex, 16).ok(),
        _ => None,
    }
}

fn parse_tint_thousandths(node: Node<'_, '_>) -> Option<i16> {
    // DrawingML uses color transforms (tint/shade) as fixed percentages in the
    // range 0..=100000 (100% = 100000). We map them into the same thousandths
    // representation used by `formula-model::Color::Theme` (-1000..=1000).
    if let Some(tint) = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tint")
        .and_then(|t| t.attribute("val"))
        .and_then(|v| v.parse::<i32>().ok())
    {
        return Some(pct_to_thousandths(tint, false));
    }

    if let Some(shade) = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "shade")
        .and_then(|t| t.attribute("val"))
        .and_then(|v| v.parse::<i32>().ok())
    {
        // `shade` is the inverse operation: 100% (100000) means no change, and
        // smaller values darken the color.
        return Some(pct_to_thousandths(shade, true));
    }

    None
}

fn pct_to_thousandths(value: i32, invert_for_shade: bool) -> i16 {
    let value = value.clamp(0, 100_000) as f64 / 100_000.0;
    let frac = if invert_for_shade {
        // shade=1.0 => 0 change, shade=0.0 => 100% darken
        -(1.0 - value)
    } else {
        value
    };
    (frac.clamp(-1.0, 1.0) * 1000.0).round() as i16
}

fn scheme_to_theme_index(scheme: &str) -> Option<u16> {
    // These indices are the same as the `theme` attribute in `styles.xml`.
    // 0=lt1/bg1, 1=dk1/tx1, 2=lt2/bg2, 3=dk2/tx2, 4..9=accent1..6, 10=hlink, 11=folHlink.
    match scheme {
        "lt1" | "bg1" => Some(0),
        "dk1" | "tx1" => Some(1),
        "lt2" | "bg2" => Some(2),
        "dk2" | "tx2" => Some(3),
        "accent1" => Some(4),
        "accent2" => Some(5),
        "accent3" => Some(6),
        "accent4" => Some(7),
        "accent5" => Some(8),
        "accent6" => Some(9),
        "hlink" => Some(10),
        "folHlink" => Some(11),
        // `phClr` is resolved via placeholder color inheritance; we don't model it yet.
        _ => None,
    }
}
