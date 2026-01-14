use formula_dax::{Cardinality, CrossFilterDirection, DataModel, FilterContext, Relationship, Table, Value};
use formula_storage::Storage;

#[test]
fn data_model_round_trip_one_to_one_relationship() {
    let tmp = tempfile::NamedTempFile::new().expect("tmpfile");
    let path = tmp.path();

    let storage1 = Storage::open_path(path).expect("open storage");
    let workbook = storage1
        .create_workbook("Book", None)
        .expect("create workbook");

    let mut model = DataModel::new();

    let mut people = Table::new("People", vec!["PersonId", "Name"]);
    people
        .push_row(vec![Value::from(1.0), Value::from("Alice")])
        .unwrap();
    people
        .push_row(vec![Value::from(2.0), Value::from("Bob")])
        .unwrap();
    model.add_table(people).unwrap();

    let mut details = Table::new("Details", vec!["PersonId", "Email"]);
    details
        .push_row(vec![Value::from(1.0), Value::from("alice@example.com")])
        .unwrap();
    details
        .push_row(vec![Value::from(2.0), Value::from("bob@example.com")])
        .unwrap();
    model.add_table(details).unwrap();

    // Persisted models can specify cardinality = one_to_one. Ensure formula-dax accepts and indexes
    // the relationship so it survives a save/load cycle.
    model
        .add_relationship(Relationship {
            name: "Details_People".to_string(),
            from_table: "Details".to_string(),
            from_column: "PersonId".to_string(),
            to_table: "People".to_string(),
            to_column: "PersonId".to_string(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_measure("DetailCount", "COUNTROWS(Details)")
        .unwrap();

    let total = model
        .evaluate_measure("DetailCount", &FilterContext::empty())
        .unwrap();
    assert_eq!(total, Value::from(2.0));

    let alice_count = model
        .evaluate_measure(
            "DetailCount",
            &FilterContext::empty().with_column_equals("People", "Name", Value::from("Alice")),
        )
        .unwrap();
    assert_eq!(alice_count, Value::from(1.0));

    storage1.save_data_model(workbook.id, &model).unwrap();
    drop(storage1);

    let storage2 = Storage::open_path(path).expect("reopen storage");
    let loaded = storage2.load_data_model(workbook.id).expect("load data model");

    let alice_count2 = loaded
        .evaluate_measure(
            "DetailCount",
            &FilterContext::empty().with_column_equals("People", "Name", Value::from("Alice")),
        )
        .unwrap();
    assert_eq!(alice_count2, Value::from(1.0));
}

