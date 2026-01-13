use formula_model::charts::LineDash;
use formula_model::Color;
use roxmltree::Document;

use super::{parse_ln, parse_solid_fill, parse_txpr};

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
