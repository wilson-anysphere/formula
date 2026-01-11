use formula_engine::eval::{CellAddr, Expr, Parser};
use formula_engine::structured_refs::resolve_structured_ref;
use formula_model::table::TableColumn;
use formula_model::{Range, Table};

fn table_fixture() -> Table {
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

#[test]
fn resolves_table_name_case_insensitively() {
    let tables_by_sheet = vec![vec![table_fixture()]];

    let parsed = Parser::parse("=SUM(table1[Col])").unwrap();
    let Expr::FunctionCall { args, .. } = parsed else {
        panic!("expected function call expression");
    };
    assert_eq!(args.len(), 1);
    let Expr::StructuredRef(sref) = &args[0] else {
        panic!("expected structured ref argument");
    };

    let (_sheet, start, end) =
        resolve_structured_ref(&tables_by_sheet, 0, CellAddr { row: 0, col: 0 }, sref).unwrap();
    assert_eq!(start, CellAddr { row: 1, col: 0 });
    assert_eq!(end, CellAddr { row: 2, col: 0 });
}

