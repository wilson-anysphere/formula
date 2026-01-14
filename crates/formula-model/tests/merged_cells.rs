use formula_model::{CellRef, CellValue, Range, Worksheet};

#[test]
fn merge_edit_unmerge_behaves_like_excel_anchor_cell() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    sheet.set_value(CellRef::new(0, 0), CellValue::String("keep".into()));
    sheet.set_value(CellRef::new(0, 1), CellValue::String("drop".into()));

    sheet
        .merge_range(Range::new(CellRef::new(0, 0), CellRef::new(0, 1)))
        .expect("merge");

    // Only the top-left cell is stored.
    assert_eq!(sheet.iter_cells().count(), 1);
    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::String("keep".into())
    );
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("keep".into())
    );

    // Editing any cell inside a merge writes the anchor cell.
    sheet.set_value(CellRef::new(0, 1), CellValue::String("hello".into()));
    assert_eq!(sheet.iter_cells().count(), 1);
    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::String("hello".into())
    );

    // Unmerge enables independent cells again.
    sheet.unmerge_range(Range::new(CellRef::new(0, 0), CellRef::new(0, 0)));
    sheet.set_value(CellRef::new(0, 1), CellValue::String("b".into()));
    assert_eq!(sheet.iter_cells().count(), 2);
    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::String("hello".into())
    );
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("b".into())
    );
}
