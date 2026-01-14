use formula_model::pivots::PivotFieldRef;

#[test]
fn pivot_field_ref_display_quotes_table_names_exactly_once() {
    let simple = PivotFieldRef::DataModelColumn {
        table: "Sales".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(simple.to_string(), "Sales[Amount]");

    let spaced = PivotFieldRef::DataModelColumn {
        table: "Dim Product".to_string(),
        column: "Category".to_string(),
    };
    assert_eq!(spaced.to_string(), "'Dim Product'[Category]");

    let escaped = PivotFieldRef::DataModelColumn {
        table: "O'Reilly".to_string(),
        column: "Name".to_string(),
    };
    assert_eq!(escaped.to_string(), "'O''Reilly'[Name]");
}
