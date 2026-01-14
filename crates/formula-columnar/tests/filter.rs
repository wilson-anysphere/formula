use formula_columnar::{
    BitVec, CmpOp, ColumnSchema, ColumnType, ColumnarTable, ColumnarTableBuilder, FilterExpr,
    FilterValue, PageCacheConfig, TableOptions, Value,
};
use std::sync::Arc;

fn options() -> TableOptions {
    TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 4 },
    }
}

fn build_table() -> ColumnarTable {
    let schema = vec![
        ColumnSchema {
            name: "n".to_owned(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "b".to_owned(),
            column_type: ColumnType::Boolean,
        },
        ColumnSchema {
            name: "s".to_owned(),
            column_type: ColumnType::String,
        },
    ];

    let mut builder = ColumnarTableBuilder::new(schema, options());
    let rows = vec![
        vec![
            Value::Number(1.0),
            Value::Boolean(true),
            Value::String(Arc::<str>::from("A")),
        ],
        vec![
            Value::Number(2.0),
            Value::Boolean(false),
            Value::String(Arc::<str>::from("B")),
        ],
        vec![Value::Number(3.0), Value::Null, Value::Null],
        vec![
            Value::Null,
            Value::Boolean(true),
            Value::String(Arc::<str>::from("A")),
        ],
        vec![
            Value::Number(2.0),
            Value::Boolean(false),
            Value::String(Arc::<str>::from("C")),
        ],
        vec![
            Value::Number(4.0),
            Value::Boolean(true),
            Value::String(Arc::<str>::from("a")),
        ],
    ];
    for row in rows {
        builder.append_row(&row);
    }
    builder.finalize()
}

fn mask_to_bools(mask: &BitVec) -> Vec<bool> {
    (0..mask.len()).map(|i| mask.get(i)).collect()
}

#[test]
fn bitvec_iter_ones_yields_increasing_indices() {
    let mut mask = BitVec::new();
    mask.push(false);
    mask.push(true);
    mask.push(false);
    mask.push(true);
    mask.push(true);
    mask.push(false);

    let ones: Vec<usize> = mask.iter_ones().collect();
    assert_eq!(ones, vec![1, 3, 4]);
}

#[test]
fn filter_numeric_comparisons() {
    let table = build_table();

    let eq = table
        .filter_mask(&FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Eq,
            value: FilterValue::Number(2.0),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&eq),
        vec![false, true, false, false, true, false]
    );

    let ne = table
        .filter_mask(&FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Ne,
            value: FilterValue::Number(2.0),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&ne),
        vec![true, false, true, false, false, true]
    );

    let lt = table
        .filter_mask(&FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Lt,
            value: FilterValue::Number(3.0),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&lt),
        vec![true, true, false, false, true, false]
    );

    let gte = table
        .filter_mask(&FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Gte,
            value: FilterValue::Number(2.0),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&gte),
        vec![false, true, true, false, true, true]
    );
}

#[test]
fn filter_boolean_without_nulls() {
    let schema = vec![ColumnSchema {
        name: "b".to_owned(),
        column_type: ColumnType::Boolean,
    }];
    let mut builder = ColumnarTableBuilder::new(schema, options());
    let rows = [
        Value::Boolean(true),
        Value::Boolean(false),
        Value::Boolean(false),
        Value::Boolean(true),
        Value::Boolean(true),
        Value::Boolean(false),
        Value::Boolean(true),
        Value::Boolean(false),
        Value::Boolean(false),
    ];
    for v in rows {
        builder.append_row(&[v]);
    }
    let table = builder.finalize();

    let eq_true = table
        .filter_mask(&FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Eq,
            value: FilterValue::Boolean(true),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&eq_true),
        vec![true, false, false, true, true, false, true, false, false]
    );

    let ne_true = table
        .filter_mask(&FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Ne,
            value: FilterValue::Boolean(true),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&ne_true),
        vec![false, true, true, false, false, true, false, true, true]
    );
}

#[test]
fn filter_or_and_not_combinators() {
    let table = build_table();

    let expr_or = FilterExpr::Or(
        Box::new(FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Eq,
            value: FilterValue::Number(1.0),
        }),
        Box::new(FilterExpr::Cmp {
            col: 2,
            op: CmpOp::Eq,
            value: FilterValue::String(Arc::<str>::from("C")),
        }),
    );
    let mask_or = table.filter_mask(&expr_or).unwrap();
    assert_eq!(
        mask_to_bools(&mask_or),
        vec![true, false, false, false, true, false]
    );

    let not_is_null = FilterExpr::Not(Box::new(FilterExpr::IsNull { col: 2 }));
    let mask_not = table.filter_mask(&not_is_null).unwrap();
    let mask_is_not_null = table.filter_mask(&FilterExpr::IsNotNull { col: 2 }).unwrap();
    assert_eq!(mask_to_bools(&mask_not), mask_to_bools(&mask_is_not_null));
    assert_eq!(
        mask_to_bools(&mask_is_not_null),
        vec![true, true, false, true, true, true]
    );
}

#[test]
fn filter_number_canonicalizes_negative_zero_and_nans_for_equality() {
    let schema = vec![ColumnSchema {
        name: "n".to_owned(),
        column_type: ColumnType::Number,
    }];
    let mut builder = ColumnarTableBuilder::new(schema, options());
    let rows = [
        Value::Number(0.0),
        Value::Number(-0.0),
        Value::Number(f64::NAN),
        Value::Null,
        Value::Number(1.0),
        Value::Number(f64::NAN),
    ];
    for v in rows {
        builder.append_row(&[v]);
    }
    let table = builder.finalize();

    let eq0 = table
        .filter_mask(&FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Eq,
            value: FilterValue::Number(0.0),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&eq0),
        vec![true, true, false, false, false, false]
    );

    let eq_nan = table
        .filter_mask(&FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Eq,
            value: FilterValue::Number(f64::NAN),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&eq_nan),
        vec![false, false, true, false, false, true]
    );

    let ne_nan = table
        .filter_mask(&FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Ne,
            value: FilterValue::Number(f64::NAN),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&ne_nan),
        vec![true, true, false, false, true, false]
    );
}

#[test]
fn filter_number_range_comparisons_against_nan_are_false() {
    let schema = vec![ColumnSchema {
        name: "n".to_owned(),
        column_type: ColumnType::Number,
    }];
    let mut builder = ColumnarTableBuilder::new(schema, options());
    let rows = [Value::Number(0.0), Value::Number(1.0), Value::Number(f64::NAN), Value::Null];
    for v in rows {
        builder.append_row(&[v]);
    }
    let table = builder.finalize();

    for op in [CmpOp::Lt, CmpOp::Lte, CmpOp::Gt, CmpOp::Gte] {
        let mask = table
            .filter_mask(&FilterExpr::Cmp {
                col: 0,
                op,
                value: FilterValue::Number(f64::NAN),
            })
            .unwrap();
        assert_eq!(mask_to_bools(&mask), vec![false, false, false, false]);
    }
}

#[test]
fn filter_number_ne_outside_range_matches_is_not_null() {
    let table = build_table();
    let ne = table
        .filter_mask(&FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Ne,
            value: FilterValue::Number(999.0),
        })
        .unwrap();
    // All non-null numeric values are != 999.0.
    assert_eq!(
        mask_to_bools(&ne),
        vec![true, true, true, false, true, true]
    );
}

#[test]
fn filter_boolean_and_nulls() {
    let table = build_table();

    let eq_true = table
        .filter_mask(&FilterExpr::Cmp {
            col: 1,
            op: CmpOp::Eq,
            value: FilterValue::Boolean(true),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&eq_true),
        vec![true, false, false, true, false, true]
    );

    let eq_false = table
        .filter_mask(&FilterExpr::Cmp {
            col: 1,
            op: CmpOp::Eq,
            value: FilterValue::Boolean(false),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&eq_false),
        vec![false, true, false, false, true, false]
    );
}

#[test]
fn filter_string_dictionary_and_case_sensitivity() {
    let table = build_table();

    let eq_a = table
        .filter_mask(&FilterExpr::Cmp {
            col: 2,
            op: CmpOp::Eq,
            value: FilterValue::String(Arc::<str>::from("A")),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&eq_a),
        vec![true, false, false, true, false, false]
    );

    let eq_lower = table
        .filter_mask(&FilterExpr::Cmp {
            col: 2,
            op: CmpOp::Eq,
            value: FilterValue::String(Arc::<str>::from("a")),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&eq_lower),
        vec![false, false, false, false, false, true]
    );

    let ne_a = table
        .filter_mask(&FilterExpr::Cmp {
            col: 2,
            op: CmpOp::Ne,
            value: FilterValue::String(Arc::<str>::from("A")),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&ne_a),
        vec![false, true, false, false, true, true]
    );

    // String missing from the dictionary.
    let eq_missing = table
        .filter_mask(&FilterExpr::Cmp {
            col: 2,
            op: CmpOp::Eq,
            value: FilterValue::String(Arc::<str>::from("Z")),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&eq_missing),
        vec![false, false, false, false, false, false]
    );

    let ne_missing = table
        .filter_mask(&FilterExpr::Cmp {
            col: 2,
            op: CmpOp::Ne,
            value: FilterValue::String(Arc::<str>::from("Z")),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&ne_missing),
        vec![true, true, false, true, true, true]
    );

    // Values outside the observed lexicographic min/max range can be rejected by stats.
    let eq_outside = table
        .filter_mask(&FilterExpr::Cmp {
            col: 2,
            op: CmpOp::Eq,
            value: FilterValue::String(Arc::<str>::from("!")),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&eq_outside),
        vec![false, false, false, false, false, false]
    );

    let ne_outside = table
        .filter_mask(&FilterExpr::Cmp {
            col: 2,
            op: CmpOp::Ne,
            value: FilterValue::String(Arc::<str>::from("!")),
        })
        .unwrap();
    assert_eq!(
        mask_to_bools(&ne_outside),
        vec![true, true, false, true, true, true]
    );
}

#[test]
fn filter_is_null_and_materialize_table() {
    let table = build_table();

    let is_null_n = table.filter_mask(&FilterExpr::IsNull { col: 0 }).unwrap();
    assert_eq!(
        mask_to_bools(&is_null_n),
        vec![false, false, false, true, false, false]
    );

    let expr = FilterExpr::And(
        Box::new(FilterExpr::Cmp {
            col: 0,
            op: CmpOp::Gte,
            value: FilterValue::Number(2.0),
        }),
        Box::new(FilterExpr::Cmp {
            col: 2,
            op: CmpOp::Ne,
            value: FilterValue::String(Arc::<str>::from("B")),
        }),
    );
    let mask = table.filter_mask(&expr).unwrap();
    assert_eq!(
        mask_to_bools(&mask),
        vec![false, false, false, false, true, true]
    );

    let filtered = table.filter_table(&mask).unwrap();
    assert_eq!(filtered.row_count(), 2);
    assert_eq!(filtered.column_count(), 3);

    assert_eq!(filtered.get_cell(0, 0), Value::Number(2.0));
    assert_eq!(filtered.get_cell(0, 1), Value::Boolean(false));
    assert_eq!(filtered.get_cell(0, 2), Value::String(Arc::<str>::from("C")));

    assert_eq!(filtered.get_cell(1, 0), Value::Number(4.0));
    assert_eq!(filtered.get_cell(1, 1), Value::Boolean(true));
    assert_eq!(filtered.get_cell(1, 2), Value::String(Arc::<str>::from("a")));
}

#[test]
fn filter_string_column_all_nulls_comparisons_are_false() {
    let schema = vec![ColumnSchema {
        name: "s".to_owned(),
        column_type: ColumnType::String,
    }];
    let mut builder = ColumnarTableBuilder::new(schema, options());
    for _ in 0..5 {
        builder.append_row(&[Value::Null]);
    }
    let table = builder.finalize();

    for op in [CmpOp::Eq, CmpOp::Ne] {
        let mask = table
            .filter_mask(&FilterExpr::Cmp {
                col: 0,
                op,
                value: FilterValue::String(Arc::<str>::from("A")),
            })
            .unwrap();
        assert_eq!(mask_to_bools(&mask), vec![false, false, false, false, false]);
    }
}
