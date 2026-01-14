use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotDestination, PivotField, PivotSource,
    PivotTableDefinition, SubtotalPosition, ValueField,
};
use formula_engine::{EditOp, Engine, Value};
use formula_model::{CellRef, Range};
use pretty_assertions::assert_eq;

fn cell(a1: &str) -> CellRef {
    CellRef::from_a1(a1).unwrap()
}

fn range(a1: &str) -> Range {
    Range::from_a1(a1).unwrap()
}

fn seed_sales_data_on(engine: &mut Engine, sheet: &str) {
    engine.set_cell_value(sheet, "A1", "Region").unwrap();
    engine.set_cell_value(sheet, "B1", "Sales").unwrap();
    engine.set_cell_value(sheet, "A2", "East").unwrap();
    engine.set_cell_value(sheet, "B2", 100.0).unwrap();
    engine.set_cell_value(sheet, "A3", "East").unwrap();
    engine.set_cell_value(sheet, "B3", 150.0).unwrap();
    engine.set_cell_value(sheet, "A4", "West").unwrap();
    engine.set_cell_value(sheet, "B4", 200.0).unwrap();
    engine.set_cell_value(sheet, "A5", "West").unwrap();
    engine.set_cell_value(sheet, "B5", 250.0).unwrap();
}

fn seed_sales_data(engine: &mut Engine) {
    seed_sales_data_on(engine, "Sheet1");
}

fn sum_sales_by_region_config() -> PivotConfig {
    PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".into(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: false,
        },
    }
}

#[test]
fn rename_sheet_updates_pivot_definition_and_refresh_uses_new_name() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B5")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Text("Region".to_string())
    );

    let sheet1_id = engine.sheet_id("Sheet1").expect("Sheet1 id");
    engine.rename_sheet_by_id(sheet1_id, "Data").unwrap();

    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(pivot.destination.sheet, "Data");
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Data".to_string(),
            range: Some(range("A1:B5")),
        }
    );

    // Refresh should use the updated sheet names and must not recreate the old sheet name via
    // `set_cell_value` / `ensure_sheet`.
    engine.refresh_pivot_table(pivot_id).unwrap();
    assert_eq!(engine.sheet_id("Sheet1"), None);
    assert_eq!(engine.get_cell_value("Data", "E4"), Value::Number(700.0));
}

#[test]
fn insert_rows_shifts_pivot_definition_and_refresh_uses_updated_addresses() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B5")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 2,
        })
        .unwrap();

    // Both the source and destination should shift down by 2 rows.
    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(pivot.destination.cell, cell("D3"));
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A3:B7")),
        }
    );

    engine.refresh_pivot_table(pivot_id).unwrap();

    // Output written at shifted destination (D3).
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Text("Region".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "E6"), Value::Number(700.0));
    // Original destination remains empty.
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Blank);
}

#[test]
fn insert_rows_matches_sheet_names_case_insensitively_for_unicode_text() {
    let mut engine = Engine::new();

    // Use a Unicode sheet name that requires Unicode-aware case folding (ß -> SS).
    seed_sales_data_on(&mut engine, "Straße");

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Straße".to_string(),
            range: Some(range("A1:B5")),
        },
        destination: PivotDestination {
            sheet: "Straße".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    // Apply the edit using a different (Unicode-folded) spelling of the sheet name.
    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "STRASSE".to_string(),
            row: 0,
            count: 2,
        })
        .unwrap();

    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(pivot.destination.cell, cell("D3"));
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Straße".to_string(),
            range: Some(range("A3:B7")),
        }
    );

    engine.refresh_pivot_table(pivot_id).unwrap();
    assert_eq!(
        engine.get_cell_value("Straße", "D3"),
        Value::Text("Region".to_string())
    );
    assert_eq!(engine.get_cell_value("Straße", "E6"), Value::Number(700.0));
}

#[test]
fn insert_rows_matches_sheet_names_nfkc_case_insensitively() {
    let mut engine = Engine::new();

    // Use a Unicode sheet name that requires NFKC normalization to match.
    // U+212A KELVIN SIGN (K) is NFKC-equivalent to ASCII 'K'.
    seed_sales_data_on(&mut engine, "Kelvin");

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Kelvin".to_string(),
            range: Some(range("A1:B5")),
        },
        destination: PivotDestination {
            sheet: "Kelvin".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    // Apply the edit using a different spelling that should match under NFKC + Unicode case.
    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "KELVIN".to_string(),
            row: 0,
            count: 2,
        })
        .unwrap();

    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(pivot.destination.cell, cell("D3"));
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Kelvin".to_string(),
            range: Some(range("A3:B7")),
        }
    );

    engine.refresh_pivot_table(pivot_id).unwrap();
    assert_eq!(
        engine.get_cell_value("Kelvin", "D3"),
        Value::Text("Region".to_string())
    );
    assert_eq!(engine.get_cell_value("Kelvin", "E6"), Value::Number(700.0));
}

#[test]
fn insert_cols_shifts_pivot_definition_and_refresh_uses_updated_addresses() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B5")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine
        .apply_operation(EditOp::InsertCols {
            sheet: "Sheet1".to_string(),
            col: 0,
            count: 1,
        })
        .unwrap();

    // Both the source and destination should shift right by 1 column.
    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(pivot.destination.cell, cell("E1"));
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("B1:C5")),
        }
    );

    engine.refresh_pivot_table(pivot_id).unwrap();

    // Output written at shifted destination (E1).
    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Text("Region".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "F4"), Value::Number(700.0));
}

#[test]
fn move_range_updates_pivot_source_and_refresh_reads_from_moved_location() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B5")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine
        .apply_operation(EditOp::MoveRange {
            sheet: "Sheet1".to_string(),
            src: range("A1:B5"),
            dst_top_left: cell("A10"),
        })
        .unwrap();

    // Source range should track the move.
    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A10:B14")),
        }
    );

    engine.refresh_pivot_table(pivot_id).unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "E4"), Value::Number(700.0));
}

#[test]
fn pivot_definitions_match_sheet_names_nfkc_case_insensitively() {
    let mut engine = Engine::new();

    // U+212A KELVIN SIGN (K) is compatibility-equivalent (NFKC) to ASCII 'K'.
    // The engine treats these as the same sheet name for lookups, so pivot definitions should too
    // when applying edits.
    let sheet = "Kelvin";
    seed_sales_data_on(&mut engine, sheet);

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: sheet.to_string(),
            range: Some(range("A1:B5")),
        },
        destination: PivotDestination {
            sheet: sheet.to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();
    assert!(!engine.pivot_table(pivot_id).unwrap().needs_refresh);

    // Apply a fill edit using the compatibility-equivalent sheet name (`Kelvin`). This should
    // still invalidate the pivot output on `Kelvin`.
    engine
        .apply_operation(EditOp::Fill {
            sheet: "Kelvin".to_string(),
            src: range("A1:B1"),
            dst: range("D1:D1"),
        })
        .unwrap();

    assert!(engine.pivot_table(pivot_id).unwrap().needs_refresh);
}

#[test]
fn rename_sheet_updates_pivot_definition_and_refresh_does_not_recreate_old_sheet() {
    let mut engine = Engine::new();
    // Use a stable sheet key that differs from the user-visible display name so we can assert that
    // renaming the display name makes the old name unresolvable (and so a stale pivot definition
    // would incorrectly resurrect it via `ensure_sheet`).
    seed_sales_data_on(&mut engine, "sheet1_key");
    engine.set_sheet_display_name("sheet1_key", "Sheet1");

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B5")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    assert!(engine.rename_sheet("Sheet1", "Data"));

    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(pivot.destination.sheet, "Data");
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Data".to_string(),
            range: Some(range("A1:B5")),
        }
    );

    // The old *display name* should not be present anymore (sheet keys remain stable across renames
    // for persistence/back-compat, so `sheet_id("Sheet1")` can still resolve).
    assert!(
        !engine
            .sheet_names_in_order()
            .iter()
            .any(|name| name == "Sheet1"),
        "expected old display name to be absent after rename"
    );

    engine.refresh_pivot_table(pivot_id).unwrap();

    // Output written to renamed sheet.
    assert_eq!(
        engine.get_cell_value("Data", "D1"),
        Value::Text("Region".to_string())
    );

    // Refresh should *not* recreate the old sheet.
    assert!(
        !engine
            .sheet_names_in_order()
            .iter()
            .any(|name| name == "Sheet1"),
        "expected refresh not to recreate a new sheet with the old display name"
    );
}

#[test]
fn delete_sheet_drops_pivot_definitions_that_referenced_it() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);
    // The engine matches Excel semantics and disallows deleting the last remaining sheet.
    // Ensure an extra sheet exists so the deletion succeeds.
    engine.set_cell_value("Sheet2", "A1", "ok").unwrap();

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B5")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    // The engine disallows deleting the last remaining sheet.
    engine.set_cell_value("Sheet2", "A1", "").unwrap();

    engine.delete_sheet("Sheet1").unwrap();

    assert!(engine.pivot_table(pivot_id).is_none());
    assert!(engine.refresh_pivot_table(pivot_id).is_err());
}
