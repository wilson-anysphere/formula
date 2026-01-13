mod common;

use common::build_model;
use formula_dax::{DaxEngine, FilterContext, RowContext, Value};
use pretty_assertions::assert_eq;

#[test]
fn dax_function_golden_suite() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let engine = DaxEngine::new();
    let cases: Vec<(&str, Value)> = vec![
        ("COUNTROWS(Orders)", 4.into()),
        ("COUNTROWS(FILTER(Orders, Orders[Amount] > 10))", 1.into()),
        ("COUNTROWS(ALL(Customers))", 3.into()),
        ("COUNT(Orders[Amount])", 4.into()),
        ("COUNTA(Orders[Amount])", 4.into()),
        ("COUNTBLANK(Orders[Amount])", 0.into()),
        // COUNT ignores text values (treat them as blanks).
        ("COUNT(Customers[Name])", 0.into()),
        ("COUNTA(Customers[Name])", 3.into()),
        ("COUNTROWS(VALUES(Customers[Region]))", 2.into()),
        ("COUNTROWS(DISTINCT(Customers[Region]))", 2.into()),
        ("COUNTROWS(SUMMARIZE(Orders, Orders[CustomerId]))", 3.into()),
        ("COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))", 2.into()),
        (
            "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region], \"X\", 1))",
            2.into(),
        ),
        (
            "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region], FILTER(Customers, Customers[Region] = \"East\")))",
            1.into(),
        ),
        (
            "COUNTROWS(CALCULATETABLE(Orders, Customers[Region] = \"East\"))",
            3.into(),
        ),
        ("DISTINCTCOUNT(Customers[Region])", 2.into()),
        ("AVERAGEX(Orders, Orders[Amount])", 10.75.into()),
        ("MINX(Orders, Orders[Amount])", 5.0.into()),
        ("MAXX(Orders, Orders[Amount])", 20.0.into()),
        // BLANK and type coercion (DAX treats BLANK as 0/FALSE in comparisons).
        ("BLANK() = 0", true.into()),
        ("BLANK() = FALSE()", true.into()),
        ("TRUE() + 1", 2.into()),
        ("BLANK() + 1", 1.into()),
        // ISBLANK
        ("ISBLANK(BLANK())", true.into()),
        ("ISBLANK(0)", false.into()),
        ("ISBLANK(\"\")", false.into()),
        ("ISBLANK([Total Sales])", false.into()),
        (
            "ISBLANK(CALCULATE([Total Sales], Customers[Region] = \"Nowhere\"))",
            true.into(),
        ),
    ];

    for (expr, expected) in cases {
        let value = engine
            .evaluate(
                &model,
                expr,
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap();
        assert_eq!(value, expected, "expression: {expr}");
    }
}
