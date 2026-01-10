use formula_engine::sort_filter::{sort_worksheet_range, HeaderOption, SortKey, SortOrder, SortSpec, SortValueType};
use formula_model::{CellRef, CellValue, Range, Worksheet};

#[test]
fn worksheet_sort_moves_cells_and_rewrites_relative_formulas() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    // Header row.
    sheet.set_value(CellRef::new(0, 0), CellValue::String("Val".into()));
    sheet.set_value(CellRef::new(0, 1), CellValue::String("Calc".into()));

    // Data rows (A2/B2, A3/B3).
    sheet.set_value(CellRef::new(1, 0), CellValue::Number(2.0));
    sheet.set_formula(CellRef::new(1, 1), Some("=A2*10".into()));

    sheet.set_value(CellRef::new(2, 0), CellValue::Number(1.0));
    sheet.set_formula(CellRef::new(2, 1), Some("=A3*10".into()));

    // Add a row-level property to ensure it permutes with the row data.
    sheet.set_row_height(2, Some(50.0));

    let range = Range::from_a1("A1:B3").unwrap();
    let spec = SortSpec {
        header: HeaderOption::HasHeader,
        keys: vec![SortKey {
            column: 0,
            order: SortOrder::Ascending,
            value_type: SortValueType::Auto,
            case_sensitive: false,
        }],
    };

    sort_worksheet_range(&mut sheet, range, &spec);

    // Values swapped.
    assert_eq!(sheet.value(CellRef::new(1, 0)), CellValue::Number(1.0));
    assert_eq!(sheet.value(CellRef::new(2, 0)), CellValue::Number(2.0));

    // Formulas moved and updated to keep relative references consistent.
    assert_eq!(sheet.formula(CellRef::new(1, 1)), Some("=A2*10"));
    assert_eq!(sheet.formula(CellRef::new(2, 1)), Some("=A3*10"));

    // Row height moved with the row that contained the `1.0` value.
    assert_eq!(sheet.row_properties(1).and_then(|p| p.height), Some(50.0));
}

