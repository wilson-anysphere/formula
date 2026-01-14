#![cfg(feature = "pivot-model")]

use formula_dax::{
    build_data_model_pivot_plan, Cardinality, CrossFilterDirection, DataModel, GroupByColumn,
    Relationship, Table,
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
