use formula_model::{charts::ColorRef, Color};
use roxmltree::Node;

pub fn parse_color(node: Node<'_, '_>) -> Option<ColorRef> {
    let color = match node.tag_name().name() {
        "srgbClr" => node.attribute("val").and_then(parse_srgb).map(Color::Argb),
        "schemeClr" => parse_scheme(node),
        // System colors are dynamic; Excel often includes a `lastClr` fallback with an sRGB value.
        "sysClr" => node
            .attribute("lastClr")
            .and_then(parse_srgb)
            .map(Color::Argb),
        "prstClr" => node
            .attribute("val")
            .and_then(preset_to_argb)
            .map(Color::Argb),
        "scrgbClr" => parse_scrgb(node).map(Color::Argb),
        _ => None,
    }?;

    // DrawingML represents color adjustments as child transform elements on the color node.
    //
    // For theme colors (`schemeClr`) we preserve the existing Theme+tint/shade representation so
    // that colors can be resolved later against the workbook theme palette.
    //
    // For concrete ARGB colors, apply basic transforms directly so the renderer sees a closer
    // match to Excel's output.
    if let Color::Argb(mut argb) = color {
        // Absolute alpha transform (`<a:alpha val="..."/>`).
        if let Some(alpha) = parse_alpha(node) {
            argb = (argb & 0x00FF_FFFF) | ((alpha as u32) << 24);
        }

        // Tint/shade transforms (`<a:tint>` / `<a:shade>`).
        if let Some(tint) = parse_tint_thousandths(node) {
            argb = apply_tint(argb, tint);
        }

        return Some(Color::Argb(argb));
    }

    Some(color)
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

fn preset_to_argb(preset: &str) -> Option<u32> {
    // DrawingML preset colors (a:prstClr) are a set of named sRGB colors.
    // We only map the most common names used by Excel charts/shapes.
    //
    // Spec reference: ECMA-376 5th ed. Part 1, §20.1.2.3.30 (PresetColorVal).
    let key = preset.trim();
    Some(if key.eq_ignore_ascii_case("black") {
        0xFF000000
    } else if key.eq_ignore_ascii_case("white") {
        0xFFFFFFFF
    } else if key.eq_ignore_ascii_case("red") {
        0xFFFF0000
    } else if key.eq_ignore_ascii_case("green") {
        0xFF00FF00
    } else if key.eq_ignore_ascii_case("blue") {
        0xFF0000FF
    } else if key.eq_ignore_ascii_case("yellow") {
        0xFFFFFF00
    } else if key.eq_ignore_ascii_case("cyan") || key.eq_ignore_ascii_case("aqua") {
        0xFF00FFFF
    } else if key.eq_ignore_ascii_case("magenta") || key.eq_ignore_ascii_case("fuchsia") {
        0xFFFF00FF
    } else if key.eq_ignore_ascii_case("gray") || key.eq_ignore_ascii_case("grey") {
        0xFF808080
    } else if key.eq_ignore_ascii_case("ltgray") || key.eq_ignore_ascii_case("ltgrey") {
        0xFFC0C0C0
    } else if key.eq_ignore_ascii_case("dkgray") || key.eq_ignore_ascii_case("dkgrey") {
        0xFF404040
    } else if key.eq_ignore_ascii_case("orange") {
        0xFFFFA500
    } else if key.eq_ignore_ascii_case("brown") {
        0xFFA52A2A
    } else if key.eq_ignore_ascii_case("purple") {
        0xFF800080
    } else {
        return None;
    })
}

fn parse_scrgb(node: Node<'_, '_>) -> Option<u32> {
    // scRGB values are expressed as fixed-point percentages in the range 0..=100000.
    // The values are linear; convert to gamma-encoded sRGB for rendering.
    let r = node.attribute("r")?.parse::<i32>().ok()?;
    let g = node.attribute("g")?.parse::<i32>().ok()?;
    let b = node.attribute("b")?.parse::<i32>().ok()?;

    let r = linear_to_srgb8(scrgb_pct_to_linear(r));
    let g = linear_to_srgb8(scrgb_pct_to_linear(g));
    let b = linear_to_srgb8(scrgb_pct_to_linear(b));

    Some(0xFF00_0000 | ((r as u32) << 16) | ((g as u32) << 8) | (b as u32))
}

fn scrgb_pct_to_linear(v: i32) -> f64 {
    (v.clamp(0, 100_000) as f64) / 100_000.0
}

fn linear_to_srgb8(v: f64) -> u8 {
    // https://en.wikipedia.org/wiki/SRGB#From_CIE_XYZ_to_sRGB
    // (Standard piecewise linear->sRGB transfer function.)
    let v = v.clamp(0.0, 1.0);
    let srgb = if v <= 0.003_130_8 {
        12.92 * v
    } else {
        1.055 * v.powf(1.0 / 2.4) - 0.055
    };
    (srgb.clamp(0.0, 1.0) * 255.0).round() as u8
}

fn parse_alpha(node: Node<'_, '_>) -> Option<u8> {
    let alpha = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "alpha")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok())?
        .clamp(0, 100_000);

    // Convert percentage-in-100000 to 8-bit alpha.
    // Use integer math with rounding half-up.
    Some(((alpha * 255 + 50_000) / 100_000) as u8)
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

fn apply_tint(argb: u32, tint_thousandths: i16) -> u32 {
    // Keep this in sync with `formula-model` tinting so theme-based and concrete colors behave
    // consistently.
    let tint = (tint_thousandths as f64 / 1000.0).clamp(-1.0, 1.0);
    if tint == 0.0 {
        return argb;
    }

    let a = (argb >> 24) & 0xFF;
    let r = ((argb >> 16) & 0xFF) as u8;
    let g = ((argb >> 8) & 0xFF) as u8;
    let b = (argb & 0xFF) as u8;

    let r = tint_channel(r, tint) as u32;
    let g = tint_channel(g, tint) as u32;
    let b = tint_channel(b, tint) as u32;

    (a << 24) | (r << 16) | (g << 8) | b
}

fn tint_channel(value: u8, tint: f64) -> u8 {
    let v = value as f64;
    let out = if tint < 0.0 {
        // Shade toward black.
        v * (1.0 + tint)
    } else {
        // Tint toward white.
        v * (1.0 - tint) + 255.0 * tint
    };

    out.round().clamp(0.0, 255.0) as u8
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
