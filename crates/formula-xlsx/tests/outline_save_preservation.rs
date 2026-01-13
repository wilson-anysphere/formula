use formula_model::{CellRef, CellValue};
use formula_xlsx::outline::read_outline_from_xlsx_bytes;
use formula_xlsx::load_from_bytes;

const FIXTURE: &[u8] = include_bytes!("fixtures/grouped_rows.xlsx");
const SHEET_PATH: &str = "xl/worksheets/sheet1.xml";

#[test]
fn outline_save_to_vec_preserves_outline_after_sheetdata_patch() {
    let mut doc = load_from_bytes(FIXTURE).expect("load fixture");
    let sheet_id = doc
        .workbook
        .sheets
        .first()
        .map(|s| s.id)
        .expect("sheet exists");

    // Add a new cell in an outlined row (row 2) so the sheetdata patcher rewrites the row.
    let b2 = CellRef::from_a1("B2").unwrap();
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .set_value(b2, CellValue::Number(42.0));

    let saved = doc.save_to_vec().expect("save");
    let outline = read_outline_from_xlsx_bytes(&saved, SHEET_PATH).expect("read outline");

    for row in 2..=4 {
        assert_eq!(outline.rows.entry(row).level, 1, "row {row} outlineLevel");
    }
    for col in 2..=4 {
        assert_eq!(outline.cols.entry(col).level, 1, "col {col} outlineLevel");
    }
}

#[test]
fn outline_save_to_vec_emits_outline_for_inserted_rows() {
    let mut doc = load_from_bytes(FIXTURE).expect("load fixture");
    let sheet_id = doc
        .workbook
        .sheets
        .first()
        .map(|s| s.id)
        .expect("sheet exists");

    // Populate the worksheet model with the fixture's outline metadata.
    let outline = read_outline_from_xlsx_bytes(FIXTURE, SHEET_PATH).expect("read outline");
    let sheet = doc.workbook.sheet_mut(sheet_id).expect("sheet exists");
    sheet.outline = outline;

    // Synthesize an outlined row that does not exist in the original XML so the sheetdata patcher
    // has to insert a new <row> element. The inserted row must carry the outline attrs.
    sheet.outline.rows.entry_mut(6).level = 1;

    let a6 = CellRef::from_a1("A6").unwrap();
    sheet.set_value(a6, CellValue::String("Inserted".to_string()));

    let saved = doc.save_to_vec().expect("save");
    let outline2 = read_outline_from_xlsx_bytes(&saved, SHEET_PATH).expect("read outline");
    assert_eq!(outline2.rows.entry(6).level, 1);
}

#[test]
fn outline_save_to_vec_rewrites_cols_when_outline_changes() {
    let mut doc = load_from_bytes(FIXTURE).expect("load fixture");
    let sheet_id = doc
        .workbook
        .sheets
        .first()
        .map(|s| s.id)
        .expect("sheet exists");

    // Populate outline metadata in the worksheet model so we can mutate it.
    let outline = read_outline_from_xlsx_bytes(FIXTURE, SHEET_PATH).expect("read outline");
    let sheet = doc.workbook.sheet_mut(sheet_id).expect("sheet exists");
    sheet.outline = outline;

    // Collapse the column group controlled by summary column 5 (hiding detail cols 2-4).
    sheet.outline.toggle_col_group(5);

    let saved = doc.save_to_vec().expect("save");
    let outline2 = read_outline_from_xlsx_bytes(&saved, SHEET_PATH).expect("read outline");

    assert!(outline2.cols.entry(5).collapsed, "summary col should be collapsed");
    for col in 2..=4 {
        assert!(
            outline2.cols.entry(col).hidden.is_hidden(),
            "detail col {col} should be hidden after collapse"
        );
    }
}

