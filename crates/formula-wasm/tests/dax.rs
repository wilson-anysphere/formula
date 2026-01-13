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

#[wasm_bindgen_test]
fn dax_model_evaluate_and_pivot() {
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
        enforce_referential_integrity: true,
    };
    let relationship_js = serde_wasm_bindgen::to_value(&relationship).unwrap();
    model.add_relationship(relationship_js).unwrap();

    model.add_measure("Total", "SUM(Orders[Amount])").unwrap();

    // Measure evaluation (no filter context).
    let total = model.evaluate("Total", None).unwrap();
    assert_eq!(total.as_f64().unwrap(), 43.0);

    // Measure evaluation with a filter context (Customers[Region] == "East" should filter Orders).
    let mut filter = DaxFilterContext::new();
    filter
        .set_column_equals("Customers", "Region", JsValue::from_str("East"))
        .unwrap();
    let total_east = model.evaluate("Total", Some(filter)).unwrap();
    assert_eq!(total_east.as_f64().unwrap(), 38.0);

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
