use formula_engine::eval::{CellAddr, Expr, Parser};
use formula_engine::structured_refs::{
    resolve_structured_ref, StructuredColumn, StructuredColumns,
    StructuredRefItem,
};
use formula_engine::{Engine, Value};
use formula_model::table::TableColumn;
use formula_model::{Range, Table};

fn table_fixture_single_col() -> Table {
    Table {
        id: 1,
        name: "Table1".into(),
        display_name: "Table1".into(),
        range: Range::from_a1("A1:A3").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![TableColumn {
            id: 1,
            name: "Col".into(),
            formula: None,
            totals_formula: None,
        }],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    }
}

fn table_fixture_multi_col() -> Table {
    Table {
        id: 1,
        name: "Table1".into(),
        display_name: "Table1".into(),
        range: Range::from_a1("A1:D4").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Col1".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Col2".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 3,
                name: "Col3".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 4,
                name: "Col4".into(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    }
}

fn table_fixture_multi_col_with_totals() -> Table {
    Table {
        id: 1,
        name: "Table1".into(),
        display_name: "Table1".into(),
        range: Range::from_a1("A1:D5").unwrap(),
        header_row_count: 1,
        totals_row_count: 1,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Col1".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Col2".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 3,
                name: "Col3".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 4,
                name: "Col4".into(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    }
}

fn table_fixture_escaped_bracket_column() -> Table {
    Table {
        id: 1,
        name: "Table1".into(),
        display_name: "Table1".into(),
        range: Range::from_a1("A1:A3").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![TableColumn {
            id: 1,
            name: "A]B".into(),
            formula: None,
            totals_formula: None,
        }],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    }
}

fn setup_engine_with_table() -> Engine {
    let mut engine = Engine::new();
    engine.set_sheet_tables("Sheet1", vec![table_fixture_multi_col()]);

    // Header row.
    engine.set_cell_value("Sheet1", "A1", "Col1").expect("A1");
    engine.set_cell_value("Sheet1", "B1", "Col2").expect("B1");
    engine.set_cell_value("Sheet1", "C1", "Col3").expect("C1");
    engine.set_cell_value("Sheet1", "D1", "Col4").expect("D1");

    // Data rows.
    engine.set_cell_value("Sheet1", "A2", 1.0_f64).expect("A2");
    engine.set_cell_value("Sheet1", "A3", 2.0_f64).expect("A3");
    engine.set_cell_value("Sheet1", "A4", 3.0_f64).expect("A4");

    engine.set_cell_value("Sheet1", "B2", 10.0_f64).expect("B2");
    engine.set_cell_value("Sheet1", "B3", 20.0_f64).expect("B3");
    engine.set_cell_value("Sheet1", "B4", 30.0_f64).expect("B4");

    engine
        .set_cell_value("Sheet1", "C2", 100.0_f64)
        .expect("C2");
    engine
        .set_cell_value("Sheet1", "C3", 200.0_f64)
        .expect("C3");
    engine
        .set_cell_value("Sheet1", "C4", 300.0_f64)
        .expect("C4");

    engine
}

fn setup_engine_with_table_and_totals() -> Engine {
    let mut engine = Engine::new();
    engine.set_sheet_tables("Sheet1", vec![table_fixture_multi_col_with_totals()]);

    // Header row.
    engine.set_cell_value("Sheet1", "A1", "Col1").expect("A1");
    engine.set_cell_value("Sheet1", "B1", "Col2").expect("B1");
    engine.set_cell_value("Sheet1", "C1", "Col3").expect("C1");
    engine.set_cell_value("Sheet1", "D1", "Col4").expect("D1");

    // Data rows.
    engine.set_cell_value("Sheet1", "A2", 1.0_f64).expect("A2");
    engine.set_cell_value("Sheet1", "A3", 2.0_f64).expect("A3");
    engine.set_cell_value("Sheet1", "A4", 3.0_f64).expect("A4");
    engine.set_cell_value("Sheet1", "A5", 6.0_f64).expect("A5 totals");

    engine.set_cell_value("Sheet1", "B2", 10.0_f64).expect("B2");
    engine.set_cell_value("Sheet1", "B3", 20.0_f64).expect("B3");
    engine.set_cell_value("Sheet1", "B4", 30.0_f64).expect("B4");
    engine.set_cell_value("Sheet1", "B5", 60.0_f64).expect("B5 totals");

    engine.set_cell_value("Sheet1", "C2", 100.0_f64).expect("C2");
    engine.set_cell_value("Sheet1", "C3", 200.0_f64).expect("C3");
    engine.set_cell_value("Sheet1", "C4", 300.0_f64).expect("C4");
    engine.set_cell_value("Sheet1", "C5", 600.0_f64).expect("C5 totals");

    engine.set_cell_value("Sheet1", "D2", 1000.0_f64).expect("D2");
    engine.set_cell_value("Sheet1", "D3", 2000.0_f64).expect("D3");
    engine.set_cell_value("Sheet1", "D4", 3000.0_f64).expect("D4");
    engine.set_cell_value("Sheet1", "D5", 6000.0_f64).expect("D5 totals");

    engine
}

#[test]
fn resolves_table_name_case_insensitively() {
    let tables_by_sheet = vec![vec![table_fixture_single_col()]];

    let parsed = Parser::parse("=SUM(table1[Col])").unwrap();
    let Expr::FunctionCall { args, .. } = parsed else {
        panic!("expected function call expression");
    };
    assert_eq!(args.len(), 1);
    let Expr::StructuredRef(sref) = &args[0] else {
        panic!("expected structured ref argument");
    };

    let ranges =
        resolve_structured_ref(&tables_by_sheet, 0, CellAddr { row: 0, col: 0 }, sref).unwrap();
    let [(_sheet, start, end)] = ranges.as_slice() else {
        panic!("expected a single resolved range");
    };
    assert_eq!(*start, CellAddr { row: 1, col: 0 });
    assert_eq!(*end, CellAddr { row: 2, col: 0 });
}

#[test]
fn parses_multi_column_selection() {
    let parsed = Parser::parse("=SUM(Table1[[Col1],[Col3]])").unwrap();
    let Expr::FunctionCall { args, .. } = parsed else {
        panic!("expected function call expression");
    };
    assert_eq!(args.len(), 1);
    let Expr::StructuredRef(sref) = &args[0] else {
        panic!("expected structured ref argument");
    };

    assert_eq!(sref.table_name.as_deref(), Some("Table1"));
    assert_eq!(sref.items, Vec::<StructuredRefItem>::new());
    assert_eq!(
        sref.columns,
        StructuredColumns::Multi(vec![
            StructuredColumn::Single("Col1".into()),
            StructuredColumn::Single("Col3".into()),
        ])
    );
}

#[test]
fn parses_multi_item_structured_ref() {
    let parsed = Parser::parse("=SUM(Table1[[#Headers],[#Data],[Col1]])").unwrap();
    let Expr::FunctionCall { args, .. } = parsed else {
        panic!("expected function call expression");
    };
    let [Expr::StructuredRef(sref)] = args.as_slice() else {
        panic!("expected structured ref argument");
    };

    assert_eq!(sref.table_name.as_deref(), Some("Table1"));
    assert_eq!(
        sref.items,
        vec![StructuredRefItem::Headers, StructuredRefItem::Data]
    );
    assert_eq!(sref.columns, StructuredColumns::Single("Col1".into()));
}

#[test]
fn parses_multi_item_structured_ref_without_columns() {
    let parsed = Parser::parse("=COUNTA(Table1[[#All],[#Totals]])").unwrap();
    let Expr::FunctionCall { args, .. } = parsed else {
        panic!("expected function call expression");
    };
    let [Expr::StructuredRef(sref)] = args.as_slice() else {
        panic!("expected structured ref argument");
    };

    assert_eq!(sref.table_name.as_deref(), Some("Table1"));
    assert_eq!(sref.items, vec![StructuredRefItem::All, StructuredRefItem::Totals]);
    assert_eq!(sref.columns, StructuredColumns::All);
}

#[test]
fn parses_multi_item_structured_ref_with_column_range() {
    let parsed = Parser::parse("=COUNTA(Table1[[#Headers],[#Data],[Col1]:[Col3]])").unwrap();
    let Expr::FunctionCall { args, .. } = parsed else {
        panic!("expected function call expression");
    };
    let [Expr::StructuredRef(sref)] = args.as_slice() else {
        panic!("expected structured ref argument");
    };

    assert_eq!(sref.table_name.as_deref(), Some("Table1"));
    assert_eq!(
        sref.items,
        vec![StructuredRefItem::Headers, StructuredRefItem::Data]
    );
    assert_eq!(
        sref.columns,
        StructuredColumns::Range {
            start: "Col1".into(),
            end: "Col3".into()
        }
    );
}

#[test]
fn parses_escaped_bracket_nested_group_even_with_bracket_in_string_literal() {
    let parsed = Parser::parse("=COUNTA(Table1[[#Headers],[A]]B]])&\"]\"").unwrap();
    let Expr::Binary { left, .. } = parsed else {
        panic!("expected binary expression");
    };

    let Expr::FunctionCall { args, .. } = &*left else {
        panic!("expected function call on left side");
    };
    let [Expr::StructuredRef(sref)] = args.as_slice() else {
        panic!("expected structured ref argument");
    };

    assert_eq!(sref.table_name.as_deref(), Some("Table1"));
    assert_eq!(sref.items, vec![StructuredRefItem::Headers]);
    assert_eq!(sref.columns, StructuredColumns::Single("A]B".into()));
}

#[test]
fn evaluates_multi_column_structured_ref_sum() {
    let mut engine = setup_engine_with_table();
    engine
        .set_cell_formula("Sheet1", "E1", "=SUM(Table1[[Col1],[Col3]])")
        .expect("formula");
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(606.0));
}

#[test]
fn evaluates_header_area_multi_column_selection() {
    let mut engine = setup_engine_with_table();
    engine
        .set_cell_formula("Sheet1", "E1", "=COUNTA(Table1[[#Headers],[Col1],[Col2]])")
        .expect("formula");
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(2.0));
}

#[test]
fn does_not_double_count_overlapping_column_ranges() {
    let mut engine = setup_engine_with_table();
    engine
        .set_cell_formula("Sheet1", "E1", "=SUM(Table1[[Col1]:[Col3],[Col2]])")
        .expect("formula");
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(666.0));
}

#[test]
fn this_row_structured_refs_still_work() {
    let mut engine = setup_engine_with_table();

    // `[@Col]` (implicit table) works when the formula is in the table.
    engine
        .set_cell_formula("Sheet1", "D2", "=[@Col2]")
        .expect("formula");
    // `[#This Row]` works with an explicit table name.
    engine
        .set_cell_formula("Sheet1", "D3", "=SUM(Table1[[#This Row],[Col1],[Col3]])")
        .expect("formula");

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(202.0));
}

#[test]
fn this_row_bracketed_column_range_syntax_works() {
    let mut engine = setup_engine_with_table();
    engine
        .set_cell_formula("Sheet1", "D2", "=SUM([@[Col1]:[Col3]])")
        .expect("formula");
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(111.0));
}

#[test]
fn evaluates_structured_ref_with_escaped_bracket_in_nested_group() {
    let mut engine = Engine::new();
    engine.set_sheet_tables("Sheet1", vec![table_fixture_escaped_bracket_column()]);
    engine.set_cell_value("Sheet1", "A1", "A]B").expect("A1");
    engine
        .set_cell_formula("Sheet1", "B1", "=COUNTA(Table1[[#Headers],[A]]B]])")
        .expect("formula");
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
}

#[test]
fn this_row_structured_refs_do_not_resolve_outside_the_table_sheet() {
    let mut engine = setup_engine_with_table();
    // Use a cell address that would otherwise fall within the table's data-range coordinates.
    engine
        .set_cell_formula("Sheet2", "D2", "=SUM(Table1[[#This Row],[Col1],[Col3]])")
        .expect("formula");
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet2", "D2"),
        Value::Error(formula_engine::ErrorKind::Name)
    );
}

#[test]
fn dependency_graph_tracks_multi_area_structured_refs() {
    let mut engine = setup_engine_with_table();
    engine
        .set_cell_formula("Sheet1", "E1", "=SUM(Table1[[Col1],[Col3]])")
        .expect("formula");
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(606.0));
    assert!(!engine.is_dirty("Sheet1", "E1"));

    // Edit a cell in Col3 (referenced by the structured-ref union) and ensure the dependent
    // formula is marked dirty.
    engine
        .set_cell_value("Sheet1", "C3", 999.0_f64)
        .expect("edit C3");
    assert!(engine.is_dirty("Sheet1", "E1"));

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1405.0));
}

#[test]
fn evaluates_multi_item_structured_ref_union_header_and_data() {
    let mut engine = setup_engine_with_table_and_totals();
    engine
        .set_cell_formula("Sheet1", "E1", "=COUNTA(Table1[[#Headers],[#Data],[Col1]])")
        .expect("formula");
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(4.0));
}

#[test]
fn evaluates_multi_item_structured_ref_union_dedups_overlaps() {
    let mut engine = setup_engine_with_table_and_totals();
    engine
        .set_cell_formula("Sheet1", "E1", "=COUNTA(Table1[[#All],[#Totals]])")
        .expect("formula");
    engine.recalculate_single_threaded();

    // Should not double-count the totals row (it's already included in #All).
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(20.0));
}

#[test]
fn evaluates_multi_item_structured_ref_union_column_range() {
    let mut engine = setup_engine_with_table_and_totals();
    engine
        .set_cell_formula("Sheet1", "E1", "=COUNTA(Table1[[#Headers],[#Data],[Col1]:[Col3]])")
        .expect("formula");
    engine.recalculate_single_threaded();

    // Header + 3 data rows = 4 rows; 3 columns => 12 cells.
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(12.0));
}

#[test]
fn evaluates_multi_item_structured_ref_union_discontiguous_rows() {
    let mut engine = setup_engine_with_table_and_totals();
    engine
        .set_cell_formula("Sheet1", "E1", "=COUNTA(Table1[[#Headers],[#Totals],[Col1]])")
        .expect("formula");
    engine.recalculate_single_threaded();

    // Header cell + totals cell for Col1.
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(2.0));
}

fn assert_table1_spill_works(mut engine: Engine) {
    engine
        .set_cell_formula("Sheet1", "F1", "=Table1")
        .expect("formula");
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(100.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(1000.0));

    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H2"), Value::Number(200.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I2"), Value::Number(2000.0));

    assert_eq!(engine.get_cell_value("Sheet1", "F3"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G3"), Value::Number(30.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H3"), Value::Number(300.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I3"), Value::Number(3000.0));
}

#[test]
fn bare_table_name_spills_data_body_into_grid() {
    assert_table1_spill_works(setup_engine_with_table_and_totals());
}

#[test]
fn bare_table_name_works_with_bytecode_disabled() {
    let mut engine = setup_engine_with_table_and_totals();
    engine.set_bytecode_enabled(false);
    assert_table1_spill_works(engine);
}

#[test]
fn bare_table_name_resolves_case_insensitive_and_unicode_names() {
    let mut engine = Engine::new();
    engine.set_sheet_tables(
        "Sheet1",
        vec![Table {
            id: 1,
            name: "Täßle".into(),
            display_name: "Täßle".into(),
            range: Range::from_a1("A1:A3").unwrap(),
            header_row_count: 1,
            totals_row_count: 0,
            columns: vec![TableColumn {
                id: 1,
                name: "Col".into(),
                formula: None,
                totals_formula: None,
            }],
            style: None,
            auto_filter: None,
            relationship_id: None,
            part_path: None,
        }],
    );
    engine.set_cell_value("Sheet1", "A1", "Col").expect("A1");
    engine.set_cell_value("Sheet1", "A2", 1.0_f64).expect("A2");
    engine.set_cell_value("Sheet1", "A3", 2.0_f64).expect("A3");

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(TÄSSLE)")
        .expect("formula");
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
}

#[test]
fn bare_table_name_dependency_tracking_marks_cells_dirty() {
    let mut engine = setup_engine_with_table_and_totals();
    engine
        .set_cell_formula("Sheet1", "E1", "=SUM(Table1)")
        .expect("formula");
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(6666.0));
    assert!(!engine.is_dirty("Sheet1", "E1"));

    engine
        .set_cell_value("Sheet1", "B3", 999.0_f64)
        .expect("edit B3");
    assert!(engine.is_dirty("Sheet1", "E1"));

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(7645.0));
}

#[test]
fn structured_ref_errors_have_excel_like_error_kinds() {
    let mut engine = setup_engine_with_table();

    engine
        .set_cell_formula("Sheet1", "E1", "=SUM(UnknownTable[Col1])")
        .expect("formula");
    engine
        .set_cell_formula("Sheet1", "E2", "=SUM(Table1[NoSuchColumn])")
        .expect("formula");
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Error(formula_engine::ErrorKind::Name)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "E2"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}

#[test]
fn missing_headers_or_totals_yield_ref_error() {
    // Missing totals.
    let mut engine = setup_engine_with_table();
    engine
        .set_cell_formula("Sheet1", "E1", "=Table1[#Totals]")
        .expect("formula");
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );

    // Missing headers.
    let mut engine = Engine::new();
    engine.set_sheet_tables(
        "Sheet1",
        vec![Table {
            id: 1,
            name: "Table1".into(),
            display_name: "Table1".into(),
            range: Range::from_a1("A1:A2").unwrap(),
            header_row_count: 0,
            totals_row_count: 0,
            columns: vec![TableColumn {
                id: 1,
                name: "Col".into(),
                formula: None,
                totals_formula: None,
            }],
            style: None,
            auto_filter: None,
            relationship_id: None,
            part_path: None,
        }],
    );
    engine.set_cell_value("Sheet1", "A1", 1.0_f64).expect("A1");
    engine.set_cell_value("Sheet1", "A2", 2.0_f64).expect("A2");
    engine
        .set_cell_formula("Sheet1", "B1", "=Table1[#Headers]")
        .expect("formula");
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}
