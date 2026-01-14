use formula_model::{Cell, CellRef, Style};
use formula_storage::{CellChange, CellData, CellValue, ImportModelWorkbookOptions, Storage};

#[test]
fn export_model_workbook_includes_styles_added_after_import() {
    let mut model = formula_model::Workbook::new();
    let sheet_id = model.add_sheet("Sheet1").expect("add sheet");

    let existing_style_id = model.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Default::default()
    });

    {
        let sheet = model.sheet_mut(sheet_id).expect("sheet");
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Number(1.0),
                formula: None,
                phonetic: None,
                style_id: existing_style_id,
                phonetic: None,
            },
        );
    }

    let storage = Storage::open_in_memory().expect("open storage");
    let meta = storage
        .import_model_workbook(&model, ImportModelWorkbookOptions::new("Book"))
        .expect("import");
    let sheet_uuid = storage
        .list_sheets(meta.id)
        .expect("list sheets")
        .into_iter()
        .find(|s| s.name == "Sheet1")
        .expect("sheet meta")
        .id;

    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet_uuid,
            row: 1,
            col: 0,
            data: CellData {
                value: CellValue::Number(2.0),
                formula: None,
                style: Some(formula_storage::Style {
                    font_id: None,
                    fill_id: None,
                    border_id: None,
                    number_format: Some("0.000".to_string()),
                    alignment: None,
                    protection: None,
                }),
            },
            user_id: None,
        }])
        .expect("apply style after import");

    let exported = storage.export_model_workbook(meta.id).expect("export");
    let sheet = exported.sheet_by_name("Sheet1").expect("sheet exists");
    let cell = sheet.cell(CellRef::new(1, 0)).expect("cell exists");
    assert_ne!(
        cell.style_id, 0,
        "expected non-default style id for cell updated via legacy APIs"
    );
    let style = exported
        .styles
        .get(cell.style_id)
        .expect("style exists in exported table");
    assert_eq!(style.number_format.as_deref(), Some("0.000"));
}
