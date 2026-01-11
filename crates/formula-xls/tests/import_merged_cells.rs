use std::path::PathBuf;

use formula_model::{CellRef, CellValue, Range};

#[test]
fn imports_merged_cells_and_row_col_properties() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("merged_hidden.xls");

    let result = formula_xls::import_xls_path(&fixture_path).expect("import xls");

    let sheet = result
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing");

    assert!(sheet.merged_regions.region_count() > 0);

    let merge_range = Range::from_a1("A1:C1").unwrap();
    assert!(
        sheet
            .merged_regions
            .iter()
            .any(|region| region.range == merge_range),
        "missing expected merged range A1:C1"
    );

    // Cell addresses inside the merged region should resolve to the top-left anchor (A1).
    assert_eq!(
        sheet
            .merged_regions
            .resolve_cell(CellRef::from_a1("B1").unwrap()),
        CellRef::from_a1("A1").unwrap()
    );
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("Merged".to_string())
    );

    // Back-compat: merged ranges are still returned on the import result.
    assert!(
        result
            .merged_ranges
            .iter()
            .any(|r| r.sheet_name == "Sheet1" && r.range == merge_range),
        "expected merged range in import result metadata"
    );

    // Row and column properties.
    assert_eq!(sheet.row_properties(0).unwrap().height, Some(20.0));
    assert!(sheet.row_properties(2).unwrap().hidden);

    assert_eq!(sheet.col_properties(0).unwrap().width, Some(20.0));
    assert!(sheet.col_properties(3).unwrap().hidden);
}

