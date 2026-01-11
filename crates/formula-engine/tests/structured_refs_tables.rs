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
    assert_eq!(sref.item, None);
    assert_eq!(
        sref.columns,
        StructuredColumns::Multi(vec![
            StructuredColumn::Single("Col1".into()),
            StructuredColumn::Single("Col3".into()),
        ])
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
    assert_eq!(sref.item, Some(StructuredRefItem::Headers));
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
