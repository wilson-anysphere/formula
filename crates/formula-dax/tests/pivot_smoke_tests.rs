use formula_dax::{pivot, DataModel, FilterContext, GroupByColumn, PivotMeasure, Table, Value};

#[test]
fn pivot_smoke_in_memory_group_by_measure() {
    let mut model = DataModel::new();

    let mut fact = Table::new("Fact", vec!["Category", "Amount"]);
    fact.push_row(vec![Value::from("A"), Value::from(10.0)])
        .unwrap();
    fact.push_row(vec![Value::from("A"), Value::from(5.0)])
        .unwrap();
    fact.push_row(vec![Value::from("B"), Value::from(7.0)])
        .unwrap();
    model.add_table(fact).unwrap();

    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

    let group_by = vec![GroupByColumn::new("Fact", "Category")];
    let measures = vec![PivotMeasure::new("Total", "[Total]").unwrap()];

    let result = pivot(&model, "Fact", &group_by, &measures, &FilterContext::empty()).unwrap();

    assert_eq!(result.columns, vec!["Fact[Category]".to_string(), "Total".to_string()]);
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), Value::from(15.0)],
            vec![Value::from("B"), Value::from(7.0)],
        ]
    );
}

