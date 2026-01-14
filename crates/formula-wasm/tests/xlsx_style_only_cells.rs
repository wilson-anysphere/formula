#![cfg(not(target_arch = "wasm32"))]

use formula_engine::Value as EngineValue;
use formula_format::cell_format_code;
use formula_model::{Alignment, HorizontalAlignment, Protection, Style, Workbook};
use formula_wasm::{WasmWorkbook, DEFAULT_SHEET};

#[test]
fn from_xlsx_bytes_imports_style_only_cells_for_cell_metadata_functions() {
    let mut workbook = Workbook::new();

    let style_id = workbook.styles.intern(Style {
        number_format: Some("0.00".to_string()),
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::Left),
            ..Alignment::default()
        }),
        protection: Some(Protection {
            locked: false,
            hidden: false,
        }),
        ..Style::default()
    });

    let sheet_id = workbook.add_sheet(DEFAULT_SHEET).unwrap();
    let sheet = workbook.sheet_mut(sheet_id).unwrap();

    // Style-only formatted-but-empty cell.
    sheet.set_style_id_a1("A1", style_id).unwrap();
    sheet
        .set_formula_a1("B1", Some(r#"CELL("format",A1)"#.to_string()))
        .unwrap();
    sheet
        .set_formula_a1("B2", Some(r#"CELL("protect",A1)"#.to_string()))
        .unwrap();
    sheet
        .set_formula_a1("B3", Some(r#"CELL("prefix",A1)"#.to_string()))
        .unwrap();

    let bytes = formula_xlsx::XlsxDocument::new(workbook)
        .save_to_vec()
        .unwrap();

    let mut wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();
    wb.debug_recalculate();

    assert_eq!(
        wb.debug_get_engine_value(DEFAULT_SHEET, "B1"),
        EngineValue::Text(cell_format_code(Some("0.00")))
    );
    assert_eq!(
        wb.debug_get_engine_value(DEFAULT_SHEET, "B2"),
        EngineValue::Number(0.0)
    );
    assert_eq!(
        wb.debug_get_engine_value(DEFAULT_SHEET, "B3"),
        EngineValue::Text("'".to_string())
    );

    // Keep `toJson` sparse: style-only cells should not appear in the input map.
    let json = wb.to_json().unwrap();
    let parsed: serde_json::Value = serde_json::from_str(&json).unwrap();
    let cells = parsed
        .get("sheets")
        .and_then(|sheets| sheets.get(DEFAULT_SHEET))
        .and_then(|sheet| sheet.get("cells"))
        .and_then(|cells| cells.as_object())
        .expect("toJson output should include a sheet cell map");
    assert!(
        !cells.contains_key("A1"),
        "toJson should not include an explicit A1 entry: {json}"
    );
}
