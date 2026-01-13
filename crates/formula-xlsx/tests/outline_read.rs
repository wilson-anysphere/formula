use formula_xlsx::outline::{read_outline_from_xlsx_bytes, write_outline_to_xlsx_bytes};
use formula_xlsx::{load_from_bytes, read_workbook_model_from_bytes};

const FIXTURE: &[u8] = include_bytes!("fixtures/grouped_rows.xlsx");
const SHEET_PATH: &str = "xl/worksheets/sheet1.xml";

#[test]
fn outline_load_from_bytes_populates_worksheet_outline() {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");
    let sheet = &doc.workbook.sheets[0];

    assert_eq!(sheet.name, "Sheet1");
    assert!(sheet.outline.pr.summary_below);
    assert!(sheet.outline.pr.summary_right);

    for row in 2..=4 {
        assert_eq!(sheet.outline.rows.entry(row).level, 1);
    }
    for col in 2..=4 {
        assert_eq!(sheet.outline.cols.entry(col).level, 1);
    }
}

#[test]
fn outline_load_from_bytes_does_not_mark_outline_hidden_rows_user_hidden() {
    // Construct a variant of the fixture with the row group collapsed. Excel stores collapsed
    // outline detail rows as `hidden="1"` in the worksheet XML, but those rows should not end up
    // marked as *user hidden* in `row_properties`.
    let mut outline = read_outline_from_xlsx_bytes(FIXTURE, SHEET_PATH).expect("fixture outline");
    outline.toggle_row_group(5);
    outline.toggle_col_group(5);
    let collapsed_bytes =
        write_outline_to_xlsx_bytes(FIXTURE, SHEET_PATH, &outline).expect("write collapsed xlsx");

    let doc = load_from_bytes(&collapsed_bytes).expect("load collapsed fixture");
    let sheet = &doc.workbook.sheets[0];

    for row in 2..=4 {
        let entry = sheet.outline.rows.entry(row);
        assert!(
            entry.hidden.outline,
            "expected row {row} to be hidden due to collapsed outline"
        );
        assert!(
            !entry.hidden.user,
            "expected row {row} to not be marked user-hidden"
        );

        let row0 = row - 1;
        let persisted_user_hidden = sheet
            .row_properties
            .get(&row0)
            .map(|p| p.hidden)
            .unwrap_or(false);
        assert!(
            !persisted_user_hidden,
            "expected row_properties.hidden to be false for outline-hidden row {row}"
        );
    }

    for col in 2..=4 {
        let entry = sheet.outline.cols.entry(col);
        assert!(
            entry.hidden.outline,
            "expected col {col} to be hidden due to collapsed outline"
        );
        assert!(
            !entry.hidden.user,
            "expected col {col} to not be marked user-hidden"
        );

        let col0 = col - 1;
        let persisted_user_hidden = sheet
            .col_properties
            .get(&col0)
            .map(|p| p.hidden)
            .unwrap_or(false);
        assert!(
            !persisted_user_hidden,
            "expected col_properties.hidden to be false for outline-hidden col {col}"
        );
    }
}

#[test]
fn outline_read_workbook_model_from_bytes_parity() {
    let workbook = read_workbook_model_from_bytes(FIXTURE).expect("load workbook model");
    let sheet = &workbook.sheets[0];

    assert_eq!(sheet.name, "Sheet1");
    assert!(sheet.outline.pr.summary_below);
    assert!(sheet.outline.pr.summary_right);

    for row in 2..=4 {
        assert_eq!(sheet.outline.rows.entry(row).level, 1);
    }
    for col in 2..=4 {
        assert_eq!(sheet.outline.cols.entry(col).level, 1);
    }
}

#[test]
fn outline_read_workbook_model_from_bytes_does_not_mark_outline_hidden_rows_user_hidden() {
    let mut outline = read_outline_from_xlsx_bytes(FIXTURE, SHEET_PATH).expect("fixture outline");
    outline.toggle_row_group(5);
    outline.toggle_col_group(5);
    let collapsed_bytes =
        write_outline_to_xlsx_bytes(FIXTURE, SHEET_PATH, &outline).expect("write collapsed xlsx");

    let workbook =
        read_workbook_model_from_bytes(&collapsed_bytes).expect("load collapsed workbook model");
    let sheet = &workbook.sheets[0];

    for row in 2..=4 {
        let entry = sheet.outline.rows.entry(row);
        assert!(entry.hidden.outline, "expected row {row} outline-hidden");
        assert!(!entry.hidden.user, "expected row {row} not user-hidden");

        let row0 = row - 1;
        let persisted_user_hidden = sheet
            .row_properties
            .get(&row0)
            .map(|p| p.hidden)
            .unwrap_or(false);
        assert!(
            !persisted_user_hidden,
            "expected row_properties.hidden to be false for outline-hidden row {row}"
        );
    }

    for col in 2..=4 {
        let entry = sheet.outline.cols.entry(col);
        assert!(entry.hidden.outline, "expected col {col} outline-hidden");
        assert!(!entry.hidden.user, "expected col {col} not user-hidden");

        let col0 = col - 1;
        let persisted_user_hidden = sheet
            .col_properties
            .get(&col0)
            .map(|p| p.hidden)
            .unwrap_or(false);
        assert!(
            !persisted_user_hidden,
            "expected col_properties.hidden to be false for outline-hidden col {col}"
        );
    }
}

