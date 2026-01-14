#![cfg(all(target_arch = "wasm32", feature = "dax"))]

use serde_json::json;
use wasm_bindgen::JsValue;
use wasm_bindgen_test::wasm_bindgen_test;

use formula_wasm::{DaxFilterContext, DaxModel};

#[derive(Debug, serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct RelationshipDto {
    name: String,
    from_table: String,
    from_column: String,
    to_table: String,
    to_column: String,
    cardinality: String,
    cross_filter_direction: String,
    is_active: bool,
    enforce_referential_integrity: bool,
}

#[derive(Debug, serde::Serialize)]
struct GroupByDto {
    table: String,
    column: String,
}

#[derive(Debug, serde::Serialize)]
struct PivotMeasureDto {
    name: String,
    expression: String,
}

#[derive(Debug, serde::Deserialize)]
struct PivotResultDto {
    columns: Vec<String>,
    rows: Vec<Vec<serde_json::Value>>,
}

#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct DaxModelSchemaDto {
    tables: Vec<DaxTableSchemaDto>,
    measures: Vec<DaxMeasureSchemaDto>,
    relationships: Vec<DaxRelationshipSchemaDto>,
}

#[derive(Debug, serde::Deserialize)]
struct DaxTableSchemaDto {
    name: String,
    columns: Vec<String>,
}

#[derive(Debug, serde::Deserialize)]
struct DaxMeasureSchemaDto {
    name: String,
    expression: String,
}

#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct DaxRelationshipSchemaDto {
    name: String,
    from_table: String,
    from_column: String,
    to_table: String,
    to_column: String,
    cardinality: String,
    cross_filter_direction: String,
    is_active: bool,
    enforce_referential_integrity: bool,
}

fn build_basic_model(enforce_referential_integrity: bool) -> DaxModel {
    let mut model = DaxModel::new();

    // Build the same sample model used by `crates/formula-dax/tests/common.rs`.
    let customers_rows = serde_wasm_bindgen::to_value(&vec![
        vec![json!(1), json!("Alice"), json!("East")],
        vec![json!(2), json!("Bob"), json!("West")],
        vec![json!(3), json!("Carol"), json!("East")],
    ])
    .unwrap();
    model
        .add_table(
            "Customers",
            vec!["CustomerId".into(), "Name".into(), "Region".into()],
            customers_rows,
        )
        .unwrap();

    let orders_rows = serde_wasm_bindgen::to_value(&vec![
        vec![json!(100), json!(1), json!(10.0)],
        vec![json!(101), json!(1), json!(20.0)],
        vec![json!(102), json!(2), json!(5.0)],
        vec![json!(103), json!(3), json!(8.0)],
    ])
    .unwrap();
    model
        .add_table(
            "Orders",
            vec!["OrderId".into(), "CustomerId".into(), "Amount".into()],
            orders_rows,
        )
        .unwrap();

    let relationship = RelationshipDto {
        name: "Orders_Customers".into(),
        from_table: "Orders".into(),
        from_column: "CustomerId".into(),
        to_table: "Customers".into(),
        to_column: "CustomerId".into(),
        cardinality: "OneToMany".into(),
        cross_filter_direction: "Single".into(),
        is_active: true,
        enforce_referential_integrity,
    };
    let relationship_js = serde_wasm_bindgen::to_value(&relationship).unwrap();
    model.add_relationship(relationship_js).unwrap();

    model.add_measure("Total", "SUM(Orders[Amount])").unwrap();

    model
}

#[wasm_bindgen_test]
fn dax_model_evaluate_and_pivot() {
    let model = build_basic_model(true);

    let schema_js = model.get_schema().unwrap();
    let schema: DaxModelSchemaDto = serde_wasm_bindgen::from_value(schema_js).unwrap();
    assert_eq!(schema.tables.len(), 2);
    assert_eq!(schema.tables[0].name, "Customers");
    assert_eq!(
        schema.tables[0].columns,
        vec!["CustomerId", "Name", "Region"]
    );
    assert_eq!(schema.tables[1].name, "Orders");
    assert_eq!(
        schema.tables[1].columns,
        vec!["OrderId", "CustomerId", "Amount"]
    );
    assert_eq!(schema.measures.len(), 1);
    assert_eq!(schema.measures[0].name, "Total");
    assert_eq!(schema.measures[0].expression, "SUM(Orders[Amount])");
    assert_eq!(schema.relationships.len(), 1);
    assert_eq!(schema.relationships[0].name, "Orders_Customers");
    assert_eq!(schema.relationships[0].from_table, "Orders");
    assert_eq!(schema.relationships[0].from_column, "CustomerId");
    assert_eq!(schema.relationships[0].to_table, "Customers");
    assert_eq!(schema.relationships[0].to_column, "CustomerId");
    assert_eq!(schema.relationships[0].cardinality, "OneToMany");
    assert_eq!(schema.relationships[0].cross_filter_direction, "Single");
    assert!(schema.relationships[0].is_active);
    assert!(schema.relationships[0].enforce_referential_integrity);

    // Measure evaluation (no filter context).
    let total = model.evaluate("Total", None).unwrap();
    assert_eq!(total.as_f64().unwrap(), 43.0);

    // Measure evaluation with a filter context (Customers[Region] == "East" should filter Orders).
    let mut filter = DaxFilterContext::new();
    filter
        .set_column_equals("Customers", "Region", JsValue::from_str("East"))
        .unwrap();
    let total_east = model.evaluate_with_filter("Total", &filter).unwrap();
    assert_eq!(total_east.as_f64().unwrap(), 38.0);

    // Multi-value filter (Customers[CustomerId] IN {1,2}).
    let mut filter_multi = DaxFilterContext::new();
    filter_multi
        .set_column_in(
            "Customers",
            "CustomerId",
            vec![JsValue::from_f64(1.0), JsValue::from_f64(2.0)],
        )
        .unwrap();
    let total_1_2 = model.evaluate_with_filter("Total", &filter_multi).unwrap();
    assert_eq!(total_1_2.as_f64().unwrap(), 35.0);

    // Clearing the filter should return to the unfiltered total.
    filter_multi.clear_column_filter("Customers", "CustomerId");
    let total_after_clear = model.evaluate_with_filter("Total", &filter_multi).unwrap();
    assert_eq!(total_after_clear.as_f64().unwrap(), 43.0);

    // Pivot: group Orders by Customers[Region] and compute Total.
    let group_by = serde_wasm_bindgen::to_value(&vec![GroupByDto {
        table: "Customers".into(),
        column: "Region".into(),
    }])
    .unwrap();
    let measures = serde_wasm_bindgen::to_value(&vec![PivotMeasureDto {
        name: "Total".into(),
        expression: "[Total]".into(),
    }])
    .unwrap();

    let pivot_js = model.pivot("Orders", group_by, measures, None).unwrap();
    let pivot: PivotResultDto = serde_wasm_bindgen::from_value(pivot_js).unwrap();

    assert_eq!(pivot.columns, vec!["Customers[Region]", "Total"]);
    assert_eq!(pivot.rows.len(), 2);

    assert_eq!(pivot.rows[0][0].as_str().unwrap(), "East");
    assert_eq!(pivot.rows[0][1].as_f64().unwrap(), 38.0);

    assert_eq!(pivot.rows[1][0].as_str().unwrap(), "West");
    assert_eq!(pivot.rows[1][1].as_f64().unwrap(), 5.0);
}

#[wasm_bindgen_test]
fn dax_model_pivot_crosstab_with_filter_borrows_filter_context() {
    let model = build_basic_model(true);

    let mut filter = DaxFilterContext::new();
    filter
        .set_column_equals("Customers", "Region", JsValue::from_str("East"))
        .unwrap();

    let row_fields = serde_wasm_bindgen::to_value(&vec![GroupByDto {
        table: "Customers".into(),
        column: "Region".into(),
    }])
    .unwrap();
    let column_fields = serde_wasm_bindgen::to_value::<Vec<GroupByDto>>(&vec![]).unwrap();
    let measures = serde_wasm_bindgen::to_value(&vec![PivotMeasureDto {
        name: "Total".into(),
        expression: "[Total]".into(),
    }])
    .unwrap();

    #[derive(Debug, serde::Deserialize)]
    struct PivotGridDto {
        data: Vec<Vec<serde_json::Value>>,
    }

    let grid_js = model
        .pivot_crosstab_with_filter("Orders", row_fields, column_fields, measures, &filter)
        .unwrap();
    let grid: PivotGridDto = serde_wasm_bindgen::from_value(grid_js).unwrap();
    assert_eq!(grid.data.len(), 2);
    assert_eq!(grid.data[0][0].as_str().unwrap(), "Customers[Region]");
    assert_eq!(grid.data[0][1].as_str().unwrap(), "Total");
    assert_eq!(grid.data[1][0].as_str().unwrap(), "East");
    assert_eq!(grid.data[1][1].as_f64().unwrap(), 38.0);
}

#[wasm_bindgen_test]
fn dax_filter_context_set_column_in_supports_multi_value_filters() {
    let model = build_basic_model(true);

    let mut filter = DaxFilterContext::new();
    filter
        .set_column_in(
            "Orders",
            "CustomerId",
            vec![JsValue::from_f64(1.0), JsValue::from_f64(3.0)],
        )
        .unwrap();

    let value = model
        .evaluate("COUNTROWS(Orders)", Some(filter))
        .unwrap()
        .as_f64()
        .unwrap();
    assert_eq!(value, 3.0);
}

#[wasm_bindgen_test]
fn dax_model_apply_calculate_filters_supports_boolean_filter_args() {
    let mut model = DaxModel::new();

    let customers_rows =
        serde_wasm_bindgen::to_value(&vec![vec![json!(1), json!("East")], vec![json!(2), json!("West")]])
            .unwrap();
    model
        .add_table(
            "Customers",
            vec!["CustomerId".into(), "Region".into()],
            customers_rows,
        )
        .unwrap();

    let orders_rows = serde_wasm_bindgen::to_value(&vec![
        vec![json!(100), json!(1), json!(10.0)],
        vec![json!(101), json!(2), json!(5.0)],
        // Unmatched foreign key (should appear under the virtual BLANK customer row).
        vec![json!(102), json!(999), json!(7.0)],
    ])
    .unwrap();
    model
        .add_table(
            "Orders",
            vec!["OrderId".into(), "CustomerId".into(), "Amount".into()],
            orders_rows,
        )
        .unwrap();

    let relationship = RelationshipDto {
        name: "Orders_Customers".into(),
        from_table: "Orders".into(),
        from_column: "CustomerId".into(),
        to_table: "Customers".into(),
        to_column: "CustomerId".into(),
        cardinality: "OneToMany".into(),
        cross_filter_direction: "Single".into(),
        is_active: true,
        enforce_referential_integrity: false,
    };
    let relationship_js = serde_wasm_bindgen::to_value(&relationship).unwrap();
    model.add_relationship(relationship_js).unwrap();

    model.add_measure("Total", "SUM(Orders[Amount])").unwrap();

    // Distinct values should include the relationship-generated BLANK member when unmatched fact
    // rows exist.
    let region_values = model
        .get_distinct_column_values("Customers", "Region", None)
        .unwrap();
    let region_values: Vec<serde_json::Value> = serde_wasm_bindgen::from_value(region_values).unwrap();
    assert_eq!(
        region_values,
        vec![json!("East"), json!("West"), serde_json::Value::Null]
    );

    // Multi-value filters should support selecting BLANK (null) so pivot field items can include
    // the relationship-generated "(blank)" member.
    let mut blank_filter = DaxFilterContext::new();
    blank_filter
        .set_column_in("Customers", "Region", vec![JsValue::NULL])
        .unwrap();
    let total_blank = model.evaluate_with_filter("Total", &blank_filter).unwrap();
    assert_eq!(total_blank.as_f64().unwrap(), 7.0);

    let group_by = serde_wasm_bindgen::to_value(&vec![GroupByDto {
        table: "Customers".into(),
        column: "Region".into(),
    }])
    .unwrap();
    let measures = serde_wasm_bindgen::to_value(&vec![PivotMeasureDto {
        name: "Total".into(),
        expression: "[Total]".into(),
    }])
    .unwrap();

    let pivot_js = model.pivot("Orders", group_by.clone(), measures.clone(), None).unwrap();
    let pivot: PivotResultDto = serde_wasm_bindgen::from_value(pivot_js).unwrap();
    assert_eq!(pivot.rows.len(), 3);
    assert_eq!(pivot.rows[0][0].as_str().unwrap(), "East");
    assert_eq!(pivot.rows[0][1].as_f64().unwrap(), 10.0);
    assert_eq!(pivot.rows[1][0].as_str().unwrap(), "West");
    assert_eq!(pivot.rows[1][1].as_f64().unwrap(), 5.0);
    assert!(pivot.rows[2][0].is_null());
    assert_eq!(pivot.rows[2][1].as_f64().unwrap(), 7.0);

    // Excluding BLANK removes the virtual blank group and the unmatched fact rows that contribute
    // to it (matching the behavior tested in `crates/formula-dax/tests/pivot_star_schema_tests.rs`).
    let non_blank_filter = model
        .apply_calculate_filters(None, vec!["Customers[Region] <> BLANK()".to_string()])
        .unwrap();
    let region_values = model
        .get_distinct_column_values_with_filter("Customers", "Region", &non_blank_filter)
        .unwrap();
    let region_values: Vec<serde_json::Value> = serde_wasm_bindgen::from_value(region_values).unwrap();
    assert_eq!(region_values, vec![json!("East"), json!("West")]);
    let pivot_js = model
        .pivot("Orders", group_by, measures, Some(non_blank_filter.clone_js()))
        .unwrap();
    let pivot: PivotResultDto = serde_wasm_bindgen::from_value(pivot_js).unwrap();
    assert_eq!(pivot.rows.len(), 2);
    assert_eq!(pivot.rows[0][0].as_str().unwrap(), "East");
    assert_eq!(pivot.rows[0][1].as_f64().unwrap(), 10.0);
    assert_eq!(pivot.rows[1][0].as_str().unwrap(), "West");
    assert_eq!(pivot.rows[1][1].as_f64().unwrap(), 5.0);
}
