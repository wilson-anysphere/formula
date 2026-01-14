use formula_model::charts::FillStyle;
use formula_model::charts::LineDash;
use formula_model::Color;
use roxmltree::Document;

use super::{parse_ln, parse_solid_fill, parse_sppr, parse_txpr};

#[test]
fn solid_fill_srgb() {
    let xml = r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:srgbClr val="FF0000"/>
    </a:solidFill>"#;
    let doc = Document::parse(xml).unwrap();
    let fill = parse_solid_fill(doc.root_element()).unwrap();
    assert_eq!(fill.color, Color::Argb(0xFFFF0000));
}

#[test]
fn solid_fill_scheme_with_tint() {
    let xml = r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:schemeClr val="accent1">
            <a:tint val="50000"/>
        </a:schemeClr>
    </a:solidFill>"#;
    let doc = Document::parse(xml).unwrap();
    let fill = parse_solid_fill(doc.root_element()).unwrap();
    assert_eq!(
        fill.color,
        Color::Theme {
            theme: 4,
            tint: Some(500)
        }
    );
}

#[test]
fn solid_fill_sys_clr_uses_last_clr() {
    let xml = r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:sysClr val="windowText" lastClr="112233"/>
    </a:solidFill>"#;
    let doc = Document::parse(xml).unwrap();
    let fill = parse_solid_fill(doc.root_element()).unwrap();
    assert_eq!(fill.color, Color::Argb(0xFF112233));
}

#[test]
fn solid_fill_prst_clr_mapping() {
    let xml = r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:prstClr val="red"/>
    </a:solidFill>"#;
    let doc = Document::parse(xml).unwrap();
    let fill = parse_solid_fill(doc.root_element()).unwrap();
    assert_eq!(fill.color, Color::Argb(0xFFFF0000));
}

#[test]
fn solid_fill_prst_clr_mapping_is_case_insensitive() {
    let xml = r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:prstClr val="LtGray"/>
    </a:solidFill>"#;
    let doc = Document::parse(xml).unwrap();
    let fill = parse_solid_fill(doc.root_element()).unwrap();
    assert_eq!(fill.color, Color::Argb(0xFFC0C0C0));
}

#[test]
fn solid_fill_alpha_transform() {
    let xml = r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:srgbClr val="FF0000">
            <a:alpha val="50000"/>
        </a:srgbClr>
    </a:solidFill>"#;
    let doc = Document::parse(xml).unwrap();
    let fill = parse_solid_fill(doc.root_element()).unwrap();
    assert_eq!(fill.color, Color::Argb(0x80FF0000));
}

#[test]
fn solid_fill_tint_transform_on_srgb() {
    let xml = r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:srgbClr val="000000">
            <a:tint val="50000"/>
        </a:srgbClr>
    </a:solidFill>"#;
    let doc = Document::parse(xml).unwrap();
    let fill = parse_solid_fill(doc.root_element()).unwrap();
    assert_eq!(fill.color, Color::Argb(0xFF808080));
}

#[test]
fn solid_fill_scrgb_clr_converts_to_srgb() {
    let xml = r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:scrgbClr r="100000" g="0" b="0"/>
    </a:solidFill>"#;
    let doc = Document::parse(xml).unwrap();
    let fill = parse_solid_fill(doc.root_element()).unwrap();
    assert_eq!(fill.color, Color::Argb(0xFFFF0000));
}

#[test]
fn solid_fill_skips_extlst_and_finds_color() {
    let xml = r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:extLst><a:ext uri="{00000000-0000-0000-0000-000000000000}"/></a:extLst>
        <a:srgbClr val="00FF00"/>
    </a:solidFill>"#;
    let doc = Document::parse(xml).unwrap();
    let fill = parse_solid_fill(doc.root_element()).unwrap();
    assert_eq!(fill.color, Color::Argb(0xFF00FF00));
}

#[test]
fn line_width_and_dash() {
    let xml = r#"<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="12700">
        <a:prstDash val="dash"/>
    </a:ln>"#;
    let doc = Document::parse(xml).unwrap();
    let line = parse_ln(doc.root_element()).unwrap();
    assert_eq!(line.width_100pt, Some(100));
    assert_eq!(line.dash, Some(LineDash::Dash));
}

#[test]
fn txpr_parses_underline_and_strike_true() {
    let xml = r#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:p>
            <a:pPr>
                <a:defRPr u="sng" strike="sngStrike">
                    <a:latin typeface="Calibri"/>
                </a:defRPr>
            </a:pPr>
        </a:p>
    </c:txPr>"#;
    let doc = Document::parse(xml).unwrap();
    let style = parse_txpr(doc.root_element()).unwrap();
    assert_eq!(style.underline, Some(true));
    assert_eq!(style.strike, Some(true));
}

#[test]
fn txpr_parses_underline_and_strike_false() {
    let xml = r#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:p>
            <a:pPr>
                <a:defRPr u="none" strike="noStrike"/>
            </a:pPr>
        </a:p>
    </c:txPr>"#;
    let doc = Document::parse(xml).unwrap();
    let style = parse_txpr(doc.root_element()).unwrap();
    assert_eq!(style.underline, Some(false));
    assert_eq!(style.strike, Some(false));
}

#[test]
fn txpr_preserves_theme_font_placeholders() {
    let xml = r#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:p>
            <a:pPr>
                <a:defRPr>
                    <a:latin typeface="+mn-lt"/>
                </a:defRPr>
            </a:pPr>
        </a:p>
    </c:txPr>"#;
    let doc = Document::parse(xml).unwrap();
    let style = parse_txpr(doc.root_element()).unwrap();
    assert_eq!(style.font_family.as_deref(), Some("+mn-lt"));
}

#[test]
fn txpr_falls_back_to_rpr_when_defrpr_missing_attrs() {
    let xml = r#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:p>
            <a:pPr>
                <a:defRPr/>
            </a:pPr>
            <a:r>
                <a:rPr u="sng" strike="sngStrike">
                    <a:latin typeface="Calibri"/>
                </a:rPr>
                <a:t>Text</a:t>
            </a:r>
        </a:p>
    </c:txPr>"#;
    let doc = Document::parse(xml).unwrap();
    let style = parse_txpr(doc.root_element()).unwrap();
    assert_eq!(style.underline, Some(true));
    assert_eq!(style.strike, Some(true));
    assert_eq!(style.font_family.as_deref(), Some("Calibri"));
}

#[test]
fn txpr_falls_back_to_end_para_rpr_when_no_runs() {
    let xml = r#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:p>
            <a:pPr/>
            <a:endParaRPr u="sng" strike="sngStrike">
                <a:latin typeface="Calibri"/>
            </a:endParaRPr>
        </a:p>
    </c:txPr>"#;
    let doc = Document::parse(xml).unwrap();
    let style = parse_txpr(doc.root_element()).unwrap();
    assert_eq!(style.underline, Some(true));
    assert_eq!(style.strike, Some(true));
    assert_eq!(style.font_family.as_deref(), Some("Calibri"));
}

#[test]
fn txpr_parses_baseline_shift() {
    let xml = r#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:p>
            <a:pPr>
                <a:defRPr baseline="30000"/>
            </a:pPr>
        </a:p>
    </c:txPr>"#;
    let doc = Document::parse(xml).unwrap();
    let style = parse_txpr(doc.root_element()).unwrap();
    assert_eq!(style.baseline, Some(30000));
}

#[test]
fn txpr_falls_back_to_ea_font_when_latin_missing() {
    let xml = r#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:p>
            <a:pPr>
                <a:defRPr>
                    <a:ea typeface="MS Gothic"/>
                </a:defRPr>
            </a:pPr>
        </a:p>
    </c:txPr>"#;
    let doc = Document::parse(xml).unwrap();
    let style = parse_txpr(doc.root_element()).unwrap();
    assert_eq!(style.font_family.as_deref(), Some("MS Gothic"));
}

#[test]
fn txpr_paragraph_defrpr_overrides_lststyle_defrpr() {
    let xml = r#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:bodyPr/>
        <a:lstStyle>
            <a:lvl1pPr>
                <a:defRPr sz="1000">
                    <a:latin typeface="Calibri"/>
                </a:defRPr>
            </a:lvl1pPr>
        </a:lstStyle>
        <a:p>
            <a:pPr>
                <a:defRPr sz="1200">
                    <a:latin typeface="Arial"/>
                </a:defRPr>
            </a:pPr>
        </a:p>
    </c:txPr>"#;
    let doc = Document::parse(xml).unwrap();
    let style = parse_txpr(doc.root_element()).unwrap();
    assert_eq!(style.font_family.as_deref(), Some("Arial"));
    assert_eq!(style.size_100pt, Some(1200));
}

#[test]
fn sppr_no_fill() {
    let xml = r#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:noFill/>
    </c:spPr>"#;
    let doc = Document::parse(xml).unwrap();
    let sppr = parse_sppr(doc.root_element()).unwrap();
    assert_eq!(sppr.fill, Some(FillStyle::None { none: true }));
}

#[test]
fn sppr_pattern_fill_with_fg_bg() {
    let xml = r#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:pattFill prst="pct50">
            <a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>
            <a:bgClr><a:srgbClr val="00FF00"/></a:bgClr>
        </a:pattFill>
    </c:spPr>"#;
    let doc = Document::parse(xml).unwrap();
    let sppr = parse_sppr(doc.root_element()).unwrap();
    let FillStyle::Pattern(fill) = sppr.fill.unwrap() else {
        panic!("expected pattFill");
    };
    assert_eq!(fill.pattern, "pct50");
    assert_eq!(fill.fg_color, Some(Color::Argb(0xFFFF0000)));
    assert_eq!(fill.bg_color, Some(Color::Argb(0xFF00FF00)));
}

#[test]
fn sppr_pattern_fill_supports_prst_clr_and_sys_clr() {
    let xml = r#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:pattFill prst="pct20">
            <a:fgClr><a:prstClr val="LtGray"/></a:fgClr>
            <a:bgClr><a:sysClr val="windowText" lastClr="112233"/></a:bgClr>
        </a:pattFill>
    </c:spPr>"#;
    let doc = Document::parse(xml).unwrap();
    let sppr = parse_sppr(doc.root_element()).unwrap();
    let FillStyle::Pattern(fill) = sppr.fill.unwrap() else {
        panic!("expected pattFill");
    };
    assert_eq!(fill.pattern, "pct20");
    assert_eq!(fill.fg_color, Some(Color::Argb(0xFFC0C0C0)));
    assert_eq!(fill.bg_color, Some(Color::Argb(0xFF112233)));
}
