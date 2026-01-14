#![cfg(feature = "pivot-model")]

use formula_dax::{
    build_data_model_pivot_plan, pivot, Cardinality, CrossFilterDirection, DataModel, FilterContext,
    GroupByColumn, Relationship, Table, Value,
};
use formula_model::pivots::{
    AggregationType, PivotConfig, PivotField, PivotFieldRef, PivotSource, SortOrder, ValueField,
};
use pretty_assertions::assert_eq;

#[test]
fn data_model_pivot_plan_maps_group_by_and_measures() {
    let mut model = DataModel::new();

    model
        .add_table(Table::new(
            "FactSales",
            vec!["ProductId", "Region", "Amount"],
        ))
        .unwrap();
    model
        .add_table(Table::new("Dim Product", vec!["ProductId", "Category"]))
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "FactSales_Product".to_string(),
            from_table: "FactSales".to_string(),
            from_column: "ProductId".to_string(),
            to_table: "Dim Product".to_string(),
            to_column: "ProductId".to_string(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
        .add_measure("Total Sales", "SUM(FactSales[Amount])")
        .unwrap();

    let source = PivotSource::DataModel {
        table: "FactSales".to_string(),
    };

    let cfg = PivotConfig {
        row_fields: vec![PivotField {
            source_field: PivotFieldRef::DataModelColumn {
                table: "Dim Product".to_string(),
                column: "Category".to_string(),
            },
            sort_order: SortOrder::Ascending,
            manual_sort: None,
        }],
        column_fields: vec![PivotField {
            source_field: PivotFieldRef::DataModelColumn {
                table: "FactSales".to_string(),
                column: "Region".to_string(),
            },
            sort_order: SortOrder::Ascending,
            manual_sort: None,
        }],
        value_fields: vec![
            ValueField {
                source_field: PivotFieldRef::DataModelMeasure("Total Sales".to_string()),
                name: "Total Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            },
            ValueField {
                source_field: PivotFieldRef::DataModelColumn {
                    table: "FactSales".to_string(),
                    column: "Amount".to_string(),
                },
                name: "Sum of Amount".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            },
        ],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Default::default(),
        subtotals: Default::default(),
        grand_totals: Default::default(),
    };

    let plan = build_data_model_pivot_plan(&model, &source, &cfg).unwrap();
    assert_eq!(plan.base_table, "FactSales");
    assert_eq!(
        plan.group_by,
        vec![
            GroupByColumn::new("Dim Product", "Category"),
            GroupByColumn::new("FactSales", "Region")
        ]
    );

    assert_eq!(plan.measures.len(), 2);
    assert_eq!(plan.measures[0].name, "Total Sales");
    assert_eq!(plan.measures[0].expression, "[Total Sales]");
    assert_eq!(plan.measures[1].name, "Sum of Amount");
    assert_eq!(plan.measures[1].expression, "SUM('FactSales'[Amount])");
}

#[test]
fn data_model_pivot_plan_rejects_cache_field_names() {
    let mut model = DataModel::new();
    model.add_table(Table::new("FactSales", vec!["Region", "Amount"]))
        .unwrap();
    model.add_measure("Total", "SUM(FactSales[Amount])").unwrap();

    let source = PivotSource::DataModel {
        table: "FactSales".to_string(),
    };

    let cfg = PivotConfig {
        row_fields: vec![PivotField {
            source_field: PivotFieldRef::CacheFieldName("Region".to_string()),
            sort_order: SortOrder::Ascending,
            manual_sort: None,
        }],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: PivotFieldRef::DataModelMeasure("Total".to_string()),
            name: "Total".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Default::default(),
        subtotals: Default::default(),
        grand_totals: Default::default(),
    };

    let err = build_data_model_pivot_plan(&model, &source, &cfg).unwrap_err();
    assert!(
        err.to_string().contains("invalid pivot config"),
        "unexpected error: {err}"
    );
}

#[test]
fn data_model_pivot_plan_resolves_unicode_identifiers_case_insensitively() {
    // Ensure the pivot-model adapter uses Unicode-aware case folding (ß -> SS) when resolving
    // table/column references, including across relationships.
    let mut model = DataModel::new();

    let mut streets = Table::new("Straße", vec!["StraßenId", "Region"]);
    streets.push_row(vec![1.into(), "East".into()]).unwrap();
    streets.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(streets).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "StraßenId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 1.into(), 20.0.into()])
        .unwrap();
    orders
        .push_row(vec![102.into(), 2.into(), 5.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders->Straße".into(),
            from_table: "orders".into(),
            from_column: "straßenid".into(),
            to_table: "STRASSE".into(),
            to_column: "STRASSENID".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Total", "SUM(orders[amount])").unwrap();

    let source = PivotSource::DataModel {
        table: "orders".to_string(),
    };

    let cfg = PivotConfig {
        row_fields: vec![PivotField {
            source_field: PivotFieldRef::DataModelColumn {
                table: "STRASSE".to_string(),
                column: "region".to_string(),
            },
            sort_order: SortOrder::Ascending,
            manual_sort: None,
        }],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: PivotFieldRef::DataModelMeasure("TOTAL".to_string()),
            name: "Total".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Default::default(),
        subtotals: Default::default(),
        grand_totals: Default::default(),
    };

    let plan = build_data_model_pivot_plan(&model, &source, &cfg).unwrap();
    assert_eq!(plan.base_table, "Orders");
    assert_eq!(plan.group_by, vec![GroupByColumn::new("Straße", "Region")]);
    assert_eq!(plan.measures[0].expression, "[Total]");

    let result = pivot(
        &model,
        &plan.base_table,
        &plan.group_by,
        &plan.measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(result.columns, vec!["Straße[Region]", "Total"]);
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("East"), Value::from(30.0)],
            vec![Value::from("West"), Value::from(5.0)],
        ]
    );
}

#[test]
fn data_model_pivot_plan_resolves_unicode_measure_names_case_insensitively() {
    // Measure names participate in the same Unicode-aware identifier folding as table/column
    // identifiers.
    let mut model = DataModel::new();

    let mut fact = Table::new("Fact", vec!["Category", "Amount"]);
    fact.push_row(vec!["A".into(), 10.0.into()]).unwrap();
    fact.push_row(vec!["B".into(), 5.0.into()]).unwrap();
    model.add_table(fact).unwrap();

    model.add_measure("Maß Total", "SUM(Fact[Amount])").unwrap();

    let source = PivotSource::DataModel {
        table: "fact".to_string(),
    };

    let cfg = PivotConfig {
        row_fields: vec![PivotField {
            source_field: PivotFieldRef::DataModelColumn {
                table: "FACT".to_string(),
                column: "category".to_string(),
            },
            sort_order: SortOrder::Ascending,
            manual_sort: None,
        }],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: PivotFieldRef::DataModelMeasure("MASS TOTAL".to_string()),
            name: "Maß Total".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Default::default(),
        subtotals: Default::default(),
        grand_totals: Default::default(),
    };

    let plan = build_data_model_pivot_plan(&model, &source, &cfg).unwrap();
    assert_eq!(plan.base_table, "Fact");
    assert_eq!(plan.measures[0].expression, "[Maß Total]");

    let result = pivot(
        &model,
        &plan.base_table,
        &plan.group_by,
        &plan.measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(result.columns, vec!["Fact[Category]", "Maß Total"]);
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), Value::from(10.0)],
            vec![Value::from("B"), Value::from(5.0)],
        ]
    );
}
