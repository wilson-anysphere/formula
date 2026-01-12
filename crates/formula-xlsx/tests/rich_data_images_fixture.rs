use std::path::Path;

use formula_xlsx::{resolve_rich_value_image_targets, XlsxPackage};

#[test]
fn resolves_rich_value_image_target_from_fixture_if_present() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/rich-data/richdata-minimal.xlsx"
    ));

    if !fixture_path.exists() {
        // Fixture is optional (may be added by another task/agent); skip if not present.
        return;
    }

    let bytes = std::fs::read(fixture_path).expect("read fixture");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("parse xlsx package");
    let resolved =
        resolve_rich_value_image_targets(pkg.parts_map()).expect("resolve rich value targets");

    assert_eq!(
        resolved.get(0).and_then(|v| v.as_deref()),
        Some("xl/media/image1.png")
    );
}

#[test]
fn resolves_rdrichvalue_image_targets_from_excel_fixture() {
    // This real Excel fixture uses the `rdrichvalue.xml` rich value schema (no `richValue.xml`).
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/image-in-cell.xlsx"
    ));

    let bytes = std::fs::read(fixture_path).expect("read excel image-in-cell fixture");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("parse xlsx package");
    let resolved =
        resolve_rich_value_image_targets(pkg.parts_map()).expect("resolve rich value targets");

    assert_eq!(
        resolved,
        vec![
            Some("xl/media/image1.png".to_string()),
            Some("xl/media/image2.png".to_string())
        ]
    );
}
