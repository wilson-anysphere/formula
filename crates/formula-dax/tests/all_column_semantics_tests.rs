use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship, RowContext,
    Table,
};
use pretty_assertions::assert_eq;

#[test]
fn all_column_respects_other_filters_on_same_table() {
    let mut model = DataModel::new();
    let mut t = Table::new("T", vec!["Region", "Name"]);
    t.push_row(vec!["East".into(), "Alice".into()]).unwrap();
    t.push_row(vec!["East".into(), "Bob".into()]).unwrap();
    t.push_row(vec!["West".into(), "Carol".into()]).unwrap();
    t.push_row(vec!["West".into(), "Dan".into()]).unwrap();
    model.add_table(t).unwrap();

    let filter = FilterContext::empty().with_column_equals("T", "Region", "East".into());
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "COUNTROWS(ALL(T[Name]))",
            &filter,
            &RowContext::default(),
        )
        .unwrap();

    // Only the East-region names should be returned. Previously this incorrectly ignored the
    // Region filter and returned names across all regions.
    assert_eq!(value, 2.into());
}

#[test]
fn all_column_clears_filter_on_target_column_only() {
    let mut model = DataModel::new();
    let mut t = Table::new("T", vec!["Region", "Name"]);
    t.push_row(vec!["East".into(), "Alice".into()]).unwrap();
    t.push_row(vec!["East".into(), "Bob".into()]).unwrap();
    t.push_row(vec!["West".into(), "Carol".into()]).unwrap();
    t.push_row(vec!["West".into(), "Dan".into()]).unwrap();
    model.add_table(t).unwrap();

    let filter = FilterContext::empty().with_column_equals("T", "Region", "East".into());
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "COUNTROWS(ALL(T[Region]))",
            &filter,
            &RowContext::default(),
        )
        .unwrap();

    // The Region filter should be removed by ALL(T[Region]), so both regions are visible.
    assert_eq!(value, 2.into());
}

#[test]
fn all_column_includes_virtual_blank_row_when_unknown_member_exists() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Id"]);
    dim.push_row(vec![1.into()]).unwrap();
    dim.push_row(vec![2.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id"]);
    fact.push_row(vec![1.into()]).unwrap();
    fact.push_row(vec![999.into()]).unwrap(); // unmatched key -> unknown member in Dim
    model.add_table(fact).unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Id".into(),
            to_table: "Dim".into(),
            to_column: "Id".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "COUNTROWS(ALL(Dim[Id]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    // Dim has 2 physical values plus the relationship-generated "unknown" member.
    assert_eq!(value, 3.into());
}

#[test]
fn all_column_table_in_iterator_transitions_only_target_column_for_measures() {
    // In DAX, `ALL(Table[Column])` returns a single-column table. When iterating that table and
    // evaluating a measure, implicit context transition should only apply the column filter,
    // *not* all columns from an arbitrary representative physical row.
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Name", "Region"]);
    customers
        .push_row(vec![1.into(), "Alice".into(), "East".into()])
        .unwrap();
    customers
        .push_row(vec![2.into(), "Bob".into(), "West".into()])
        .unwrap();
    customers
        .push_row(vec![3.into(), "Carol".into(), "East".into()])
        .unwrap();
    model.add_table(customers).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "CustomerId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 1.into(), 20.0.into()])
        .unwrap();
    orders
        .push_row(vec![102.into(), 2.into(), 5.0.into()])
        .unwrap();
    orders
        .push_row(vec![103.into(), 3.into(), 8.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders_Customers".into(),
            from_table: "Orders".into(),
            from_column: "CustomerId".into(),
            to_table: "Customers".into(),
            to_column: "CustomerId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "SUMX(ALL(Customers[Region]), [Total Sales])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 43.0.into());
}
