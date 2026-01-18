use formula_engine::sort_filter::{
    apply_autofilter, AutoFilter, CellValue, ColumnFilter, FilterCriterion, FilterJoin,
    FilterValue, RangeData, RangeRef,
};
use std::collections::BTreeMap;

fn range(rows: Vec<Vec<CellValue>>) -> RangeData {
    let range = RangeRef {
        start_row: 0,
        start_col: 0,
        end_row: rows.len() - 1,
        end_col: rows[0].len() - 1,
    };
    RangeData::new(range, rows).unwrap()
}

fn render_visible_rows(data: &RangeData, visible_rows: &[bool]) -> Vec<Vec<CellValue>> {
    data.rows
        .iter()
        .zip(visible_rows.iter())
        .filter_map(|(row, visible)| visible.then(|| row.clone()))
        .collect()
}

#[test]
fn filter_hides_rows_from_render_and_can_be_cleared() {
    let data = range(vec![
        vec![CellValue::Text("Name".into())],
        vec![CellValue::Text("Alice".into())],
        vec![CellValue::Text("Bob".into())],
    ]);

    let filter = AutoFilter {
        range: data.range,
        columns: BTreeMap::from([(
            0,
            ColumnFilter {
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Equals(FilterValue::Text("Alice".into()))],
            },
        )]),
    };

    let result = apply_autofilter(&data, &filter).expect("filter should succeed");
    let rendered = render_visible_rows(&data, &result.visible_rows);
    assert_eq!(rendered.len(), 2);
    assert_eq!(rendered[1][0], CellValue::Text("Alice".into()));

    // Clearing filter restores all rows.
    let empty_filter = AutoFilter {
        range: data.range,
        columns: BTreeMap::new(),
    };
    let cleared = apply_autofilter(&data, &empty_filter).expect("filter should succeed");
    let rendered = render_visible_rows(&data, &cleared.visible_rows);
    assert_eq!(rendered.len(), 3);
}
