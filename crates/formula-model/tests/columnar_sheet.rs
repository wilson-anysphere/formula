use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
};
use formula_model::{CellRef, CellValue, Range, RangeBatchBuffer, Worksheet, EXCEL_MAX_ROWS};
use std::sync::Arc;

fn build_datetime_string_table(row_count: usize) -> Arc<formula_columnar::ColumnarTable> {
    let schema = vec![
        ColumnSchema {
            name: "id".to_owned(),
            column_type: ColumnType::DateTime,
        },
        ColumnSchema {
            name: "category".to_owned(),
            column_type: ColumnType::String,
        },
    ];

    let mut builder = ColumnarTableBuilder::new(
        schema,
        TableOptions {
            page_size_rows: 64,
            cache: PageCacheConfig { max_entries: 4 },
        },
    );

    let categories = [Arc::<str>::from("A"), Arc::<str>::from("B")];
    for i in 0..row_count {
        builder.append_row(&[
            Value::DateTime(i as i64),
            Value::String(categories[i % categories.len()].clone()),
        ]);
    }
    Arc::new(builder.finalize())
}

#[test]
fn worksheet_can_be_backed_by_columnar_table_and_override_with_sparse_cells() {
    let table = build_datetime_string_table(10_000);
    let mut sheet = Worksheet::new(1, "Data");
    sheet.set_columnar_table(CellRef::new(0, 0), table);

    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::Number(0.0),
        "datetime values are exposed as numbers in the core cell model"
    );
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("A".to_string())
    );

    // Sparse overlay should override the backing store.
    sheet.set_value(
        CellRef::new(0, 1),
        CellValue::String("Override".to_string()),
    );
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("Override".to_string())
    );

    let range = sheet.get_range_batch(Range::new(CellRef::new(100, 0), CellRef::new(109, 1)));
    assert_eq!(range.columns.len(), 2);
    assert_eq!(range.columns[0].len(), 10);
    assert_eq!(range.columns[1].len(), 10);
}

#[test]
fn range_batch_into_reuses_buffer_memory_across_calls() {
    let table = build_datetime_string_table(64);
    let mut sheet = Worksheet::new(1, "Data");
    sheet.set_columnar_table(CellRef::new(0, 0), table);

    let mut buffer = RangeBatchBuffer::default();

    let range1 = Range::new(CellRef::new(0, 0), CellRef::new(9, 1));
    let batch1 = sheet.get_range_batch_into(range1, &mut buffer);
    assert_eq!(batch1.columns.len(), 2);
    assert_eq!(batch1.columns[0].len(), 10);

    let outer_ptr = buffer.columns.as_ptr();
    let inner_ptrs: Vec<*const CellValue> = buffer.columns.iter().map(|c| c.as_ptr()).collect();
    let inner_caps: Vec<usize> = buffer.columns.iter().map(|c| c.capacity()).collect();

    // Same shape, different origin.
    let range2 = Range::new(CellRef::new(10, 0), CellRef::new(19, 1));
    let batch2 = sheet.get_range_batch_into(range2, &mut buffer);
    assert_eq!(batch2.columns[0][0], CellValue::Number(10.0));

    assert_eq!(
        buffer.columns.as_ptr(),
        outer_ptr,
        "outer Vec should not reallocate when reusing the same viewport shape"
    );
    for (idx, col) in buffer.columns.iter().enumerate() {
        assert_eq!(
            col.as_ptr(),
            inner_ptrs[idx],
            "inner Vec for column {idx} should reuse its allocation"
        );
        assert_eq!(
            col.capacity(),
            inner_caps[idx],
            "inner Vec capacity for column {idx} should remain stable"
        );
    }
}

#[test]
fn range_batch_into_overlays_sparse_values_over_columnar() {
    let table = build_datetime_string_table(8);
    let mut sheet = Worksheet::new(1, "Data");
    sheet.set_columnar_table(CellRef::new(0, 0), table);

    sheet.set_value(
        CellRef::new(1, 1),
        CellValue::String("Override".to_string()),
    );

    let mut buffer = RangeBatchBuffer::default();
    let range = Range::new(CellRef::new(0, 0), CellRef::new(2, 1));
    let batch = sheet.get_range_batch_into(range, &mut buffer);

    assert_eq!(batch.columns[0][2], CellValue::Number(2.0));
    assert_eq!(
        batch.columns[1][0],
        CellValue::String("A".to_string()),
        "columnar values should be visible when not overridden"
    );
    assert_eq!(
        batch.columns[1][1],
        CellValue::String("Override".to_string()),
        "sparse overlay should take precedence over columnar backing"
    );
}

#[test]
fn effective_used_range_sparse_only_matches_used_range() {
    let mut sheet = Worksheet::new(1, "Data");
    sheet.set_value(CellRef::new(3, 4), CellValue::Number(1.0));

    assert_eq!(
        sheet.used_range(),
        Some(Range::new(CellRef::new(3, 4), CellRef::new(3, 4)))
    );
    assert_eq!(sheet.effective_used_range(), sheet.used_range());
}

#[test]
fn effective_used_range_columnar_only_matches_columnar_range() {
    let table = build_datetime_string_table(4);
    let mut sheet = Worksheet::new(1, "Data");
    sheet.set_columnar_table(CellRef::new(10, 20), table);

    assert_eq!(sheet.used_range(), None, "no sparse cells stored");
    assert_eq!(
        sheet.columnar_range(),
        Some(Range::new(CellRef::new(10, 20), CellRef::new(13, 21)))
    );
    assert_eq!(sheet.effective_used_range(), sheet.columnar_range());
}

#[test]
fn effective_used_range_sparse_overlaps_columnar() {
    let table = build_datetime_string_table(10);
    let mut sheet = Worksheet::new(1, "Data");
    sheet.set_columnar_table(CellRef::new(0, 0), table);

    // Sparse edit inside the columnar range should not expand the effective extent.
    sheet.set_value(CellRef::new(5, 1), CellValue::String("Edit".to_string()));

    assert_eq!(
        sheet.columnar_range(),
        Some(Range::new(CellRef::new(0, 0), CellRef::new(9, 1)))
    );
    assert_eq!(
        sheet.effective_used_range(),
        Some(Range::new(CellRef::new(0, 0), CellRef::new(9, 1)))
    );
}

#[test]
fn effective_used_range_sparse_outside_columnar_unions_extents() {
    let table = build_datetime_string_table(10);
    let mut sheet = Worksheet::new(1, "Data");
    sheet.set_columnar_table(CellRef::new(0, 0), table);

    // Sparse edit outside the table should expand the effective range.
    sheet.set_value(CellRef::new(20, 5), CellValue::Number(1.0));

    assert_eq!(
        sheet.effective_used_range(),
        Some(Range::new(CellRef::new(0, 0), CellRef::new(20, 5)))
    );
}

#[test]
fn columnar_backing_allows_reading_beyond_excel_row_limit() {
    let schema = vec![ColumnSchema {
        name: "id".to_owned(),
        column_type: ColumnType::DateTime,
    }];

    let mut builder = ColumnarTableBuilder::new(
        schema,
        TableOptions {
            page_size_rows: 64,
            cache: PageCacheConfig { max_entries: 4 },
        },
    );

    for i in 0..128 {
        builder.append_row(&[Value::DateTime(i)]);
    }

    let table = Arc::new(builder.finalize());
    let mut sheet = Worksheet::new(1, "Data");

    // Attach the table starting *after* Excel's maximum row, proving access doesn't
    // depend on sparse cell-key encoding.
    sheet.set_columnar_table(CellRef::new(EXCEL_MAX_ROWS, 0), table);

    assert_eq!(
        sheet.columnar_range(),
        Some(Range::new(
            CellRef::new(EXCEL_MAX_ROWS, 0),
            CellRef::new(EXCEL_MAX_ROWS + 127, 0)
        )),
        "columnar_range should work for tables beyond Excel's row limit"
    );

    assert_eq!(
        sheet.value(CellRef::new(EXCEL_MAX_ROWS + 5, 0)),
        CellValue::Number(5.0)
    );
}
