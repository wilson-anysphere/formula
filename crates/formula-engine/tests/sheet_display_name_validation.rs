use formula_engine::pivot::{PivotConfig, PivotDestination, PivotSource, PivotTableDefinition};
use formula_engine::{Engine, Value};
use formula_model::{CellRef, EXCEL_MAX_SHEET_NAME_LEN};

#[test]
fn set_sheet_display_name_rejects_invalid_excel_sheet_names() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");

    // Start with a valid display name that differs from the stable sheet key.
    engine.set_sheet_display_name("Sheet1", "Budget");

    engine
        .set_cell_formula("Sheet2", "A1", r#"=CELL("address",Sheet1!A1)"#)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet2", "A1"),
        Value::Text("Budget!$A$1".to_string())
    );

    // Seed a pivot definition that references the sheet by its display name.
    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Test Pivot".to_string(),
        source: PivotSource::Range {
            sheet: "Budget".to_string(),
            range: None,
        },
        destination: PivotDestination {
            sheet: "Budget".to_string(),
            cell: CellRef::new(0, 0),
        },
        config: PivotConfig::default(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: false,
    });

    let too_long = "A".repeat(EXCEL_MAX_SHEET_NAME_LEN + 1);
    let invalid_names = vec!["Bad:Name", "[Bad]", "'Bad", "Bad'", too_long.as_str()];

    for invalid in invalid_names {
        engine.set_sheet_display_name("Sheet1", invalid);
        engine.recalculate_single_threaded();

        // Invalid names should be ignored (display name unchanged).
        let sheet_id = engine.sheet_id("Sheet1").unwrap();
        assert_eq!(engine.sheet_name(sheet_id), Some("Budget"));
        assert_eq!(
            engine.get_cell_value("Sheet2", "A1"),
            Value::Text("Budget!$A$1".to_string())
        );

        // Pivot metadata should not be rewritten when the display-name update is rejected.
        let pivot = engine.pivot_table(pivot_id).unwrap();
        assert_eq!(pivot.destination.sheet, "Budget");
        match &pivot.source {
            PivotSource::Range { sheet, .. } => assert_eq!(sheet, "Budget"),
            other => panic!("unexpected pivot source: {other:?}"),
        }
    }
}
