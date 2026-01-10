use formula_xlsx::outline::{read_outline_from_xlsx_bytes, write_outline_to_xlsx_bytes};

const FIXTURE: &[u8] = include_bytes!("fixtures/grouped_rows.xlsx");
const SHEET_PATH: &str = "xl/worksheets/sheet1.xml";

#[test]
fn reads_outline_levels_from_fixture() {
    let outline = read_outline_from_xlsx_bytes(FIXTURE, SHEET_PATH).expect("fixture should parse");
    assert!(outline.pr.summary_below);
    assert!(outline.pr.summary_right);

    for row in 2..=4 {
        let entry = outline.rows.entry(row);
        assert_eq!(entry.level, 1);
        assert!(!entry.hidden.is_hidden());
    }

    let summary = outline.rows.entry(5);
    assert_eq!(summary.level, 0);
    assert!(!summary.collapsed);

    for col in 2..=4 {
        let entry = outline.cols.entry(col);
        assert_eq!(entry.level, 1);
        assert!(!entry.hidden.is_hidden());
    }
}

#[test]
fn outline_round_trip_preserves_grouping() {
    let outline = read_outline_from_xlsx_bytes(FIXTURE, SHEET_PATH).expect("fixture should parse");
    let bytes = write_outline_to_xlsx_bytes(FIXTURE, SHEET_PATH, &outline)
        .expect("should be able to write xlsx");
    let outline2 = read_outline_from_xlsx_bytes(&bytes, SHEET_PATH).expect("round-tripped xlsx");

    assert_eq!(outline2.pr.summary_below, outline.pr.summary_below);
    assert_eq!(outline2.pr.summary_right, outline.pr.summary_right);

    for row in 2..=4 {
        let entry = outline2.rows.entry(row);
        assert_eq!(entry.level, 1);
        assert!(!entry.hidden.is_hidden());
    }

    for col in 2..=4 {
        let entry = outline2.cols.entry(col);
        assert_eq!(entry.level, 1);
        assert!(!entry.hidden.is_hidden());
    }
}

#[test]
fn collapse_and_expand_group_round_trip() {
    let mut outline = read_outline_from_xlsx_bytes(FIXTURE, SHEET_PATH).expect("fixture should parse");

    // Collapse the group with summary row 5 (rows 2-4 at outline level 1).
    outline.toggle_row_group(5);
    assert!(outline.rows.entry(5).collapsed);
    for row in 2..=4 {
        assert!(outline.rows.entry(row).hidden.is_hidden());
    }

    let bytes = write_outline_to_xlsx_bytes(FIXTURE, SHEET_PATH, &outline)
        .expect("should be able to write collapsed xlsx");
    let mut collapsed = read_outline_from_xlsx_bytes(&bytes, SHEET_PATH).expect("collapsed xlsx");
    assert!(collapsed.rows.entry(5).collapsed);
    for row in 2..=4 {
        assert!(collapsed.rows.entry(row).hidden.is_hidden());
    }

    // Expand again.
    collapsed.toggle_row_group(5);
    assert!(!collapsed.rows.entry(5).collapsed);
    for row in 2..=4 {
        assert!(!collapsed.rows.entry(row).hidden.is_hidden());
    }

    let bytes2 = write_outline_to_xlsx_bytes(&bytes, SHEET_PATH, &collapsed)
        .expect("should be able to write expanded xlsx");
    let expanded = read_outline_from_xlsx_bytes(&bytes2, SHEET_PATH).expect("expanded xlsx");
    assert!(!expanded.rows.entry(5).collapsed);
    for row in 2..=4 {
        assert!(!expanded.rows.entry(row).hidden.is_hidden());
    }
}

#[test]
fn collapse_and_expand_column_group_round_trip() {
    let mut outline = read_outline_from_xlsx_bytes(FIXTURE, SHEET_PATH).expect("fixture should parse");

    outline.toggle_col_group(5);
    assert!(outline.cols.entry(5).collapsed);
    for col in 2..=4 {
        assert!(outline.cols.entry(col).hidden.is_hidden());
    }

    let bytes = write_outline_to_xlsx_bytes(FIXTURE, SHEET_PATH, &outline)
        .expect("should be able to write collapsed xlsx");
    let mut collapsed = read_outline_from_xlsx_bytes(&bytes, SHEET_PATH).expect("collapsed xlsx");
    assert!(collapsed.cols.entry(5).collapsed);
    for col in 2..=4 {
        assert!(collapsed.cols.entry(col).hidden.is_hidden());
    }

    collapsed.toggle_col_group(5);
    assert!(!collapsed.cols.entry(5).collapsed);
    for col in 2..=4 {
        assert!(!collapsed.cols.entry(col).hidden.is_hidden());
    }
}
