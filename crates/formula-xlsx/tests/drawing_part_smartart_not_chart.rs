use formula_model::drawings::DrawingObjectKind;
use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::XlsxPackage;

const FIXTURE: &[u8] = include_bytes!("../../../fixtures/xlsx/basic/smartart.xlsx");

#[test]
fn drawing_part_does_not_treat_smartart_graphic_frame_as_chart_placeholder() {
    let pkg = XlsxPackage::from_bytes(FIXTURE).expect("load smartart.xlsx fixture");
    let mut workbook = formula_model::Workbook::new();

    let part = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        pkg.parts_map(),
        &mut workbook,
    )
    .expect("parse drawing part");

    assert!(
        !part.objects.iter().any(|obj| matches!(
            obj.kind,
            DrawingObjectKind::ChartPlaceholder { .. }
        )),
        "SmartArt drawings are represented as graphicFrames but should not be parsed as charts"
    );

    let raw = part
        .objects
        .iter()
        .filter_map(|obj| match &obj.kind {
            DrawingObjectKind::Unknown { raw_xml } => Some(raw_xml.as_str()),
            _ => None,
        })
        .next()
        .expect("expected at least one unknown drawing object for the SmartArt graphicFrame");

    assert!(raw.contains("SmartArt 1"));
    assert!(raw.contains("drawingml/2006/diagram"));
}

