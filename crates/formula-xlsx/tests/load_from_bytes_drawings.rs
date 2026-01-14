use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::drawings::{DrawingObjectKind, ImageId};
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn load_from_bytes_populates_worksheet_drawings_with_images() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/image.xlsx");
    let bytes = std::fs::read(&path).expect("read image fixture");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load_from_bytes should succeed");

    assert_eq!(doc.workbook.sheets.len(), 1);
    let ws = &doc.workbook.sheets[0];
    assert!(
        ws.drawings
            .iter()
            .any(|o| matches!(o.kind, DrawingObjectKind::Image { .. })),
        "expected sheet to contain at least one Image drawing object, got {:#?}",
        ws.drawings
    );

    let image_id = ImageId::new("image1.png");
    let stored = doc
        .workbook
        .images
        .get(&image_id)
        .expect("expected workbook.images to contain image1.png");
    assert_eq!(
        stored.bytes,
        zip_part(&bytes, "xl/media/image1.png"),
        "expected workbook.images bytes for image1.png to match the xl/media/image1.png part"
    );
}

#[test]
fn load_from_bytes_populates_worksheet_drawings_with_chart_placeholders() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/charts/xlsx/bar.xlsx");
    let bytes = std::fs::read(&path).expect("read chart fixture");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load_from_bytes should succeed");

    assert_eq!(doc.workbook.sheets.len(), 1);
    let ws = &doc.workbook.sheets[0];
    assert!(
        ws.drawings.iter().any(|o| matches!(
            &o.kind,
            DrawingObjectKind::ChartPlaceholder { rel_id, .. } if rel_id == "rId1"
        )),
        "expected sheet to contain a ChartPlaceholder with rel_id=rId1, got {:#?}",
        ws.drawings
    );
}

#[test]
fn load_from_bytes_populates_chartsheet_drawings_with_chart_placeholders() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/charts/chart-sheet.xlsx");
    let bytes = std::fs::read(&path).expect("read chartsheet fixture");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load_from_bytes should succeed");

    assert_eq!(doc.workbook.sheets.len(), 2, "expected workbook to contain worksheet + chartsheet");
    let chart_sheet = doc
        .workbook
        .sheets
        .iter()
        .find(|s| s.name == "Chart1")
        .expect("expected workbook to contain a sheet named Chart1");

    assert!(
        chart_sheet.drawings.iter().any(|o| matches!(
            &o.kind,
            DrawingObjectKind::ChartPlaceholder { rel_id, .. } if rel_id == "rId1"
        )),
        "expected chartsheet to contain a ChartPlaceholder with rel_id=rId1, got {:#?}",
        chart_sheet.drawings
    );
}
