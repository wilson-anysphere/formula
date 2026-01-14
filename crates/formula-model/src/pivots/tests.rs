use super::*;

use serde_json::json;
use std::collections::HashSet;

#[test]
fn pivot_config_serde_roundtrips_with_calculated_fields_and_items() {
    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![PivotField::new("Product")],
        value_fields: vec![ValueField {
            source_field: PivotFieldRef::CacheFieldName("Sales".to_string()),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![FilterField {
            source_field: PivotFieldRef::CacheFieldName("Region".to_string()),
            allowed: Some(std::collections::HashSet::from([PivotKeyPart::Text(
                "East".to_string(),
            )])),
        }],
        calculated_fields: vec![CalculatedField {
            name: "Profit".to_string(),
            formula: "Sales - Cost".to_string(),
        }],
        calculated_items: vec![CalculatedItem {
            field: "Region".to_string(),
            name: "East+West".to_string(),
            formula: "\"East\" + \"West\"".to_string(),
        }],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals::default(),
    };

    let json_value = serde_json::to_value(&cfg).unwrap();
    assert!(json_value.get("calculatedFields").is_some());
    assert!(json_value.get("calculatedItems").is_some());

    let decoded: PivotConfig = serde_json::from_value(json_value.clone()).unwrap();
    assert_eq!(decoded, cfg);

    // Backward-compat: missing keys should default to empty vectors.
    let mut json_without = json_value;
    if let Some(obj) = json_without.as_object_mut() {
        obj.remove("calculatedFields");
        obj.remove("calculatedItems");
    }
    let decoded: PivotConfig = serde_json::from_value(json_without).unwrap();
    assert!(decoded.calculated_fields.is_empty());
    assert!(decoded.calculated_items.is_empty());
    assert_eq!(decoded.row_fields, cfg.row_fields);
    assert_eq!(decoded.column_fields, cfg.column_fields);
    assert_eq!(decoded.value_fields, cfg.value_fields);
    assert_eq!(decoded.filter_fields, cfg.filter_fields);
    assert_eq!(decoded.layout, cfg.layout);
    assert_eq!(decoded.subtotals, cfg.subtotals);
    assert_eq!(decoded.grand_totals, cfg.grand_totals);
}

#[test]
fn pivot_grand_totals_defaults_true_when_fields_missing() {
    let decoded: GrandTotals = serde_json::from_value(json!({})).unwrap();
    assert_eq!(decoded, GrandTotals::default());

    let decoded: PivotConfig = serde_json::from_value(json!({})).unwrap();
    assert_eq!(decoded.grand_totals, GrandTotals::default());

    // Back-compat: allow partial `grandTotals` payloads to deserialize by defaulting
    // missing keys to Excel-like defaults (true).
    let decoded: PivotConfig = serde_json::from_value(json!({"grandTotals": {}})).unwrap();
    assert_eq!(decoded.grand_totals, GrandTotals::default());

    let decoded: PivotConfig =
        serde_json::from_value(json!({"grandTotals": {"rows": false}})).unwrap();
    assert!(!decoded.grand_totals.rows);
    assert!(decoded.grand_totals.columns);
}

#[test]
fn pivot_grand_totals_partial_payload_defaults_missing_keys_to_true() {
    let decoded: GrandTotals = serde_json::from_value(json!({ "rows": false })).unwrap();
    assert_eq!(
        decoded,
        GrandTotals {
            rows: false,
            columns: true,
        }
    );

    let decoded: GrandTotals = serde_json::from_value(json!({ "columns": false })).unwrap();
    assert_eq!(
        decoded,
        GrandTotals {
            rows: true,
            columns: false,
        }
    );
}

#[test]
fn pivot_field_ref_from_unstructured_parses_dax_like_refs() {
    assert_eq!(
        PivotFieldRef::from_unstructured("Sales"),
        PivotFieldRef::CacheFieldName("Sales".to_string())
    );
    assert_eq!(
        PivotFieldRef::from_unstructured("[Total Sales]"),
        PivotFieldRef::DataModelMeasure("Total Sales".to_string())
    );
    assert_eq!(
        PivotFieldRef::from_unstructured("Orders[Order ID]"),
        PivotFieldRef::DataModelColumn {
            table: "Orders".to_string(),
            column: "Order ID".to_string(),
        }
    );
    assert_eq!(
        PivotFieldRef::from_unstructured("'Order Details'[Unit Price]"),
        PivotFieldRef::DataModelColumn {
            table: "Order Details".to_string(),
            column: "Unit Price".to_string(),
        }
    );
    assert_eq!(
        PivotFieldRef::from_unstructured("'O''Reilly'[Book]"),
        PivotFieldRef::DataModelColumn {
            table: "O'Reilly".to_string(),
            column: "Book".to_string(),
        }
    );
    assert_eq!(
        PivotFieldRef::from_unstructured("  Orders  [  Order ID  ]  "),
        PivotFieldRef::DataModelColumn {
            table: "Orders".to_string(),
            column: "Order ID".to_string(),
        }
    );
}

#[test]
fn pivot_field_ref_parses_dax_refs_and_serializes_structured_data_model_fields() {
    let cache: PivotFieldRef = serde_json::from_value(json!("Region")).unwrap();
    assert_eq!(cache, PivotFieldRef::CacheFieldName("Region".to_string()));
    assert_eq!(cache.as_cache_field_name(), Some("Region"));
    assert_eq!(serde_json::to_value(&cache).unwrap(), json!("Region"));
    assert_eq!(cache.to_string(), "Region");
    assert!(cache == "Region");

    let col: PivotFieldRef = serde_json::from_value(json!("'Sales Table'[Amount]")).unwrap();
    assert_eq!(
        col,
        PivotFieldRef::DataModelColumn {
            table: "Sales Table".to_string(),
            column: "Amount".to_string()
        }
    );
    assert_eq!(
        serde_json::to_value(&col).unwrap(),
        json!({"table": "Sales Table", "column": "Amount"})
    );
    assert_eq!(col.to_string(), "'Sales Table'[Amount]");
    assert!(col != "Amount");

    let escaped_col: PivotFieldRef = serde_json::from_value(json!("'O''Brien'[Amount]")).unwrap();
    assert_eq!(
        escaped_col,
        PivotFieldRef::DataModelColumn {
            table: "O'Brien".to_string(),
            column: "Amount".to_string()
        }
    );
    assert_eq!(escaped_col.to_string(), "'O''Brien'[Amount]");

    let col_with_brackets: PivotFieldRef =
        serde_json::from_value(json!("Orders[Amount]]USD]")).unwrap();
    assert_eq!(
        col_with_brackets,
        PivotFieldRef::DataModelColumn {
            table: "Orders".to_string(),
            column: "Amount]USD".to_string()
        }
    );
    assert_eq!(col_with_brackets.to_string(), "Orders[Amount]]USD]");

    let measure: PivotFieldRef = serde_json::from_value(json!("[Total Sales]")).unwrap();
    assert_eq!(
        measure,
        PivotFieldRef::DataModelMeasure("Total Sales".to_string())
    );
    assert_eq!(measure.as_cache_field_name(), None);
    assert_eq!(
        serde_json::to_value(&measure).unwrap(),
        json!({"measure": "Total Sales"})
    );
    assert_eq!(measure.to_string(), "[Total Sales]");

    let measure_with_brackets: PivotFieldRef =
        serde_json::from_value(json!("[Total]]USD]")).unwrap();
    assert_eq!(
        measure_with_brackets,
        PivotFieldRef::DataModelMeasure("Total]USD".to_string())
    );
    assert_eq!(measure_with_brackets.to_string(), "[Total]]USD]");

    // Back-compat: allow `{ name: \"...\" }` as an alternate measure payload shape.
    let measure_alt: PivotFieldRef =
        serde_json::from_value(json!({"name": "Total Sales"})).unwrap();
    assert_eq!(measure_alt, measure);
}

#[test]
fn pivot_field_ref_display_and_canonical_name_handle_dax_quoting_and_escaping() {
    let simple_table = PivotFieldRef::DataModelColumn {
        table: "Orders".to_string(),
        column: "Order ID".to_string(),
    };
    assert_eq!(simple_table.to_string(), "Orders[Order ID]");

    // `Display` uses DAX quoting rules for table names, while `canonical_name` and
    // `display_string` intentionally keep table names unquoted for friendlier UI labels and
    // cache-field matching.
    let spaced_table = PivotFieldRef::DataModelColumn {
        table: "Sales Table".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(spaced_table.to_string(), "'Sales Table'[Amount]");
    assert_eq!(
        spaced_table.canonical_name().as_ref(),
        "Sales Table[Amount]"
    );
    assert_eq!(spaced_table.display_string(), "Sales Table[Amount]");

    // Table names that are not valid "C identifiers" need quoting.
    let leading_digit_table = PivotFieldRef::DataModelColumn {
        table: "2024Orders".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(leading_digit_table.to_string(), "'2024Orders'[Amount]");

    // Unicode table names should remain unquoted as long as they match identifier-like rules.
    let unicode_table = PivotFieldRef::DataModelColumn {
        table: "Straße".to_string(),
        column: "Category".to_string(),
    };
    assert_eq!(unicode_table.to_string(), "Straße[Category]");

    // DAX keywords must be quoted to avoid ambiguity in DAX expressions.
    let keyword_table = PivotFieldRef::DataModelColumn {
        table: "VAR".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(keyword_table.to_string(), "'VAR'[Amount]");
    assert_eq!(keyword_table.canonical_name().as_ref(), "VAR[Amount]");
    assert_eq!(keyword_table.display_string(), "VAR[Amount]");

    // Column and measure names escape `]` as `]]` inside DAX brackets.
    let bracketed_column = PivotFieldRef::DataModelColumn {
        table: "Orders".to_string(),
        column: "Gross]Margin".to_string(),
    };
    assert_eq!(bracketed_column.to_string(), "Orders[Gross]]Margin]");

    let bracketed_measure = PivotFieldRef::DataModelMeasure("My]Measure".to_string());
    assert_eq!(bracketed_measure.to_string(), "[My]]Measure]");

    // Parsing is best-effort but should round-trip DAX escaping for `]`.
    let parsed_col = PivotFieldRef::from_unstructured("Orders[Gross]]Margin]");
    assert_eq!(
        parsed_col,
        PivotFieldRef::DataModelColumn {
            table: "Orders".to_string(),
            column: "Gross]Margin".to_string(),
        }
    );
    assert_eq!(parsed_col.to_string(), "Orders[Gross]]Margin]");

    let parsed_measure = PivotFieldRef::from_unstructured("[My]]Measure]");
    assert_eq!(
        parsed_measure,
        PivotFieldRef::DataModelMeasure("My]Measure".to_string())
    );
    assert_eq!(parsed_measure.to_string(), "[My]]Measure]");

    // If the table name contains `[`, `canonical_name`/`display_string` must quote it to avoid
    // producing an ambiguous `Table[Column]` shape.
    let bracket_table = PivotFieldRef::DataModelColumn {
        table: "My[Table]".to_string(),
        column: "Col".to_string(),
    };
    assert_eq!(bracket_table.to_string(), "'My[Table]'[Col]");
    assert_eq!(bracket_table.canonical_name().as_ref(), "'My[Table]'[Col]");
    assert_eq!(bracket_table.display_string(), "'My[Table]'[Col]");
    assert_eq!(
        PivotFieldRef::from_unstructured(bracket_table.canonical_name().as_ref()),
        bracket_table
    );
}

#[test]
fn pivot_value_to_key_part_canonicalizes_numbers() {
    assert_eq!(
        PivotValue::Number(0.0).to_key_part(),
        PivotValue::Number(-0.0).to_key_part()
    );

    let alt_nan = f64::from_bits(0x7ff8_0000_0000_0001);
    assert!(alt_nan.is_nan());
    assert_eq!(
        PivotValue::Number(f64::NAN).to_key_part(),
        PivotValue::Number(alt_nan).to_key_part()
    );
    assert_eq!(
        PivotValue::Number(alt_nan).to_key_part(),
        PivotKeyPart::Number(f64::NAN.to_bits())
    );

    let n = 12.5;
    assert_eq!(
        PivotValue::Number(n).to_key_part(),
        PivotKeyPart::Number(n.to_bits())
    );
}

#[test]
fn pivot_value_uses_tagged_ipc_serde_layout() {
    assert_eq!(
        serde_json::to_value(&PivotValue::Blank).unwrap(),
        json!({"type": "blank"})
    );
    assert_eq!(
        serde_json::to_value(&PivotValue::Text("hi".to_string())).unwrap(),
        json!({"type": "text", "value": "hi"})
    );
    assert_eq!(
        serde_json::to_value(&PivotValue::Number(1.5)).unwrap(),
        json!({"type": "number", "value": 1.5})
    );
}

#[test]
fn pivot_key_part_display_string_matches_excel_like_expectations() {
    assert_eq!(PivotKeyPart::Blank.display_string(), "(blank)");
    assert_eq!(
        PivotKeyPart::Number(PivotValue::canonical_number_bits(-0.0)).display_string(),
        "0"
    );
}

#[test]
fn pivot_reference_rewrite_helpers_are_case_insensitive() {
    let mut source = PivotSource::Table {
        table: crate::table::TableIdentifier::Name("TABLE1".to_string()),
    };
    assert!(source.rewrite_table_name("table1", "Renamed"));
    assert_eq!(
        source,
        PivotSource::Table {
            table: crate::table::TableIdentifier::Name("Renamed".to_string())
        }
    );

    let mut source = PivotSource::NamedRange {
        name: DefinedNameIdentifier::Name("MyRange".to_string()),
    };
    assert!(source.rewrite_defined_name("MYRANGE", "RenamedRange"));
    assert_eq!(
        source,
        PivotSource::NamedRange {
            name: DefinedNameIdentifier::Name("RenamedRange".to_string())
        }
    );

    let mut source = PivotSource::RangeName {
        sheet_name: "Data".to_string(),
        range: crate::Range::from_a1("A1:B2").unwrap(),
    };
    assert!(source.rewrite_sheet_name("DATA", "RenamedSheet"));
    assert_eq!(
        source,
        PivotSource::RangeName {
            sheet_name: "RenamedSheet".to_string(),
            range: crate::Range::from_a1("A1:B2").unwrap(),
        }
    );

    let mut dest = PivotDestination::CellName {
        sheet_name: "Data".to_string(),
        cell: crate::CellRef::new(0, 0),
    };
    assert!(dest.rewrite_sheet_name("data", "RenamedSheet"));
    assert_eq!(
        dest,
        PivotDestination::CellName {
            sheet_name: "RenamedSheet".to_string(),
            cell: crate::CellRef::new(0, 0),
        }
    );
}

#[test]
fn dax_column_ref_parser_handles_basic_and_quoted_tables() {
    assert_eq!(
        parse_dax_column_ref("Table[Column]"),
        Some(("Table".to_string(), "Column".to_string()))
    );
    assert_eq!(
        parse_dax_column_ref("Table[Column]]Name]"),
        Some(("Table".to_string(), "Column]Name".to_string()))
    );
    assert_eq!(
        parse_dax_column_ref("  Table [ Column ]  "),
        Some(("Table".to_string(), "Column".to_string()))
    );
    assert_eq!(
        parse_dax_column_ref("'My Table'[Column]"),
        Some(("My Table".to_string(), "Column".to_string()))
    );
    assert_eq!(
        parse_dax_column_ref("'O''Reilly'[Col]"),
        Some(("O'Reilly".to_string(), "Col".to_string()))
    );
    assert_eq!(
        parse_dax_column_ref("'My[Table]'[Col]"),
        Some(("My[Table]".to_string(), "Col".to_string()))
    );

    // Invalid shapes.
    assert_eq!(parse_dax_column_ref(""), None);
    assert_eq!(parse_dax_column_ref("[Column]"), None);
    assert_eq!(parse_dax_column_ref("Table[]"), None);
    assert_eq!(parse_dax_column_ref("Table[Column"), None);
    assert_eq!(parse_dax_column_ref("Table[Column] trailing"), None);
    assert_eq!(parse_dax_column_ref("'Table'X[Column]"), None);
}

#[test]
fn dax_measure_ref_parser_handles_basic_and_rejects_nested_brackets() {
    assert_eq!(
        parse_dax_measure_ref("[Measure]"),
        Some("Measure".to_string())
    );
    assert_eq!(
        parse_dax_measure_ref("[Measure]]Name]"),
        Some("Measure]Name".to_string())
    );
    assert_eq!(
        parse_dax_measure_ref(" [ Measure ] "),
        Some("Measure".to_string())
    );

    assert_eq!(parse_dax_measure_ref("[]"), None);
    assert_eq!(parse_dax_measure_ref("[Table[Column]]"), None);
    assert_eq!(parse_dax_measure_ref("[Measure] trailing"), None);
}

#[test]
fn dax_table_identifier_display_quotes_keywords_and_invalid_identifiers() {
    let cases = [
        // DAX keywords used by VAR expressions should be quoted to avoid ambiguity.
        ("IN", "'IN'[Col]"),
        ("var", "'var'[Col]"),
        ("Return", "'Return'[Col]"),
        // Invalid "C identifier" forms must be quoted.
        ("123", "'123'[Col]"),
        ("Sales-2024", "'Sales-2024'[Col]"),
    ];

    for (table, expected) in cases {
        let field = PivotFieldRef::DataModelColumn {
            table: table.to_string(),
            column: "Col".to_string(),
        };
        assert_eq!(field.to_string(), expected);
    }
}

#[test]
fn pivot_field_ref_deserializes_dax_refs_from_strings() {
    assert_eq!(
        serde_json::from_value::<PivotFieldRef>(json!("Table[Column]")).unwrap(),
        PivotFieldRef::DataModelColumn {
            table: "Table".to_string(),
            column: "Column".to_string(),
        }
    );
    assert_eq!(
        serde_json::from_value::<PivotFieldRef>(json!("'My Table'[Column]")).unwrap(),
        PivotFieldRef::DataModelColumn {
            table: "My Table".to_string(),
            column: "Column".to_string(),
        }
    );
    assert_eq!(
        serde_json::from_value::<PivotFieldRef>(json!("[Measure]")).unwrap(),
        PivotFieldRef::DataModelMeasure("Measure".to_string())
    );
    assert_eq!(
        serde_json::from_value::<PivotFieldRef>(json!("Regular Field")).unwrap(),
        PivotFieldRef::CacheFieldName("Regular Field".to_string())
    );
}

#[test]
fn pivot_field_ref_serializes_data_model_refs_as_structured_objects() {
    let column = PivotFieldRef::DataModelColumn {
        table: "Sales".to_string(),
        column: "Amount".to_string(),
    };
    assert_eq!(
        serde_json::to_value(&column).unwrap(),
        json!({ "table": "Sales", "column": "Amount" })
    );

    let measure = PivotFieldRef::DataModelMeasure("Total Sales".to_string());
    assert_eq!(
        serde_json::to_value(&measure).unwrap(),
        json!({ "measure": "Total Sales" })
    );
}

#[test]
fn pivot_field_ref_cache_field_name_helpers() {
    let cache = PivotFieldRef::CacheFieldName("X".to_string());
    assert_eq!(cache.as_cache_field_name(), Some("X"));
    assert_eq!(cache.cache_field_name(), Some("X"));

    let col = PivotFieldRef::DataModelColumn {
        table: "T".to_string(),
        column: "C".to_string(),
    };
    assert_eq!(col.as_cache_field_name(), None);
    assert_eq!(col.cache_field_name(), None);

    let measure = PivotFieldRef::DataModelMeasure("M".to_string());
    assert_eq!(measure.as_cache_field_name(), None);
    assert_eq!(measure.cache_field_name(), None);
}
