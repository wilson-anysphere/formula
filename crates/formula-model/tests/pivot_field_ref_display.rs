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

#[test]
fn pivot_field_ref_display_quotes_dax_keywords() {
    let var_kw = PivotFieldRef::DataModelColumn {
        table: "VAR".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(var_kw.to_string(), "'VAR'[Amount]");

    let return_kw = PivotFieldRef::DataModelColumn {
        table: "Return".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(return_kw.to_string(), "'Return'[Amount]");

    let in_kw = PivotFieldRef::DataModelColumn {
        table: "in".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(in_kw.to_string(), "'in'[Amount]");

    // Ensure we don't over-quote identifiers that merely start with a keyword.
    let var_suffix = PivotFieldRef::DataModelColumn {
        table: "VAR_1".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(var_suffix.to_string(), "VAR_1[Amount]");
}

#[test]
fn pivot_field_ref_display_quotes_non_identifier_table_names() {
    let leading_digit = PivotFieldRef::DataModelColumn {
        table: "123Sales".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(leading_digit.to_string(), "'123Sales'[Amount]");

    let punct = PivotFieldRef::DataModelColumn {
        table: "Sales-2024".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(punct.to_string(), "'Sales-2024'[Amount]");
}

#[test]
fn pivot_field_ref_display_allows_unicode_identifier_table_names() {
    let unicode = PivotFieldRef::DataModelColumn {
        table: "Straße".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(unicode.to_string(), "Straße[Amount]");
}
