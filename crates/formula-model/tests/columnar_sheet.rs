use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
};
use formula_model::{CellRef, CellValue, Range, Worksheet, EXCEL_MAX_ROWS};
use std::sync::Arc;

#[test]
fn worksheet_can_be_backed_by_columnar_table_and_override_with_sparse_cells() {
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
            page_size_rows: 1024,
            cache: PageCacheConfig { max_entries: 8 },
        },
    );

    let categories = [Arc::<str>::from("A"), Arc::<str>::from("B")];
    for i in 0..10_000 {
        builder.append_row(&[
            Value::DateTime(i as i64),
            Value::String(categories[i % categories.len()].clone()),
        ]);
    }
    let table = Arc::new(builder.finalize());

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
    sheet.set_value(CellRef::new(0, 1), CellValue::String("Override".to_string()));
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
        sheet.value(CellRef::new(EXCEL_MAX_ROWS + 5, 0)),
        CellValue::Number(5.0)
    );
}

