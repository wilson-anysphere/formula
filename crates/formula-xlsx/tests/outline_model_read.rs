use formula_xlsx::load_from_bytes;
use formula_xlsx::outline::{read_outline_from_xlsx_bytes, write_outline_to_xlsx_bytes};

const FIXTURE: &[u8] = include_bytes!("fixtures/grouped_rows.xlsx");
const SHEET_PATH: &str = "xl/worksheets/sheet1.xml";

#[test]
fn load_from_bytes_populates_worksheet_outline() {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");
    let sheet = &doc.workbook.sheets[0];

    for row in 2..=4 {
        let entry = sheet.outline.rows.entry(row);
        assert_eq!(entry.level, 1, "expected row {row} to be outline level 1");
    }

    for col in 2..=4 {
        let entry = sheet.outline.cols.entry(col);
        assert_eq!(entry.level, 1, "expected col {col} to be outline level 1");
    }
}

#[test]
fn outline_hidden_rows_are_not_marked_user_hidden_in_row_properties() {
    // The base fixture has grouping levels but is not collapsed. Generate a collapsed variant so
    // we can verify that `row/@hidden="1"` caused by outline collapse does not set the model's
    // *user hidden* bit.
    let mut outline = read_outline_from_xlsx_bytes(FIXTURE, SHEET_PATH).expect("read outline");
    outline.toggle_row_group(5);
    outline.toggle_col_group(5);

    let bytes = write_outline_to_xlsx_bytes(FIXTURE, SHEET_PATH, &outline).expect("write outline");
    let doc = load_from_bytes(&bytes).expect("load collapsed fixture");
    let sheet = &doc.workbook.sheets[0];

    for row in 2..=4 {
        let entry = sheet.outline.rows.entry(row);
        assert_eq!(
            entry.level, 1,
            "expected row {row} to remain outline level 1"
        );
        assert!(
            entry.hidden.outline,
            "expected row {row} to be hidden by outline in collapsed group"
        );
        assert!(
            !entry.hidden.user,
            "expected row {row} to not be marked user-hidden when hidden only by outline"
        );

        let row_0based = row - 1;
        let props_hidden = sheet
            .row_properties
            .get(&row_0based)
            .map(|p| p.hidden)
            .unwrap_or(false);
        assert!(
            !props_hidden,
            "expected sheet.row_properties[{row_0based}].hidden to be false for outline-hidden row"
        );
    }

    for col in 2..=4 {
        let entry = sheet.outline.cols.entry(col);
        assert_eq!(
            entry.level, 1,
            "expected col {col} to remain outline level 1"
        );
        assert!(
            entry.hidden.outline,
            "expected col {col} to be hidden by outline in collapsed group"
        );
        assert!(
            !entry.hidden.user,
            "expected col {col} to not be marked user-hidden when hidden only by outline"
        );

        let col_0based = col - 1;
        let props_hidden = sheet
            .col_properties
            .get(&col_0based)
            .map(|p| p.hidden)
            .unwrap_or(false);
        assert!(
            !props_hidden,
            "expected sheet.col_properties[{col_0based}].hidden to be false for outline-hidden col"
        );
    }
}
