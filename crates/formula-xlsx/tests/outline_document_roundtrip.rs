use formula_xlsx::outline::read_outline_from_xlsx_bytes;
use formula_xlsx::load_from_bytes;

const FIXTURE: &[u8] = include_bytes!("fixtures/grouped_rows.xlsx");
const SHEET_PATH: &str = "xl/worksheets/sheet1.xml";

#[test]
fn document_round_trip_preserves_outline_grouping() {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    let outline = read_outline_from_xlsx_bytes(&saved, SHEET_PATH).expect("read outline");
    assert!(outline.pr.summary_below);
    assert!(outline.pr.summary_right);

    for row in 2..=4 {
        assert_eq!(outline.rows.entry(row).level, 1);
    }
    for col in 2..=4 {
        assert_eq!(outline.cols.entry(col).level, 1);
    }
}

