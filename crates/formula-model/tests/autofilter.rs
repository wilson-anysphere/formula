use formula_model::{
    CellRef, FilterColumn, FilterCriterion, FilterJoin, FilterValue, NumberComparison, Range,
    SheetAutoFilter, SortCondition, SortState, TextMatch, TextMatchKind,
};
use serde_json::json;

#[test]
fn sheet_autofilter_is_serde_roundtrippable() {
    let filter = SheetAutoFilter {
        range: Range::new(CellRef::new(0, 0), CellRef::new(10, 3)),
        filter_columns: vec![
            FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![
                    FilterCriterion::Blanks,
                    FilterCriterion::Equals(FilterValue::Text("Alice".into())),
                ],
                values: Vec::new(),
                raw_xml: vec![r#"<colorFilter dxfId="3"/>"#.to_string()],
            },
            FilterColumn {
                col_id: 1,
                join: FilterJoin::All,
                criteria: vec![FilterCriterion::Number(NumberComparison::Between {
                    min: 2.0,
                    max: 5.0,
                })],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 2,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::TextMatch(TextMatch {
                    kind: TextMatchKind::Contains,
                    pattern: "foo".into(),
                    case_sensitive: false,
                })],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
        ],
        sort_state: Some(SortState {
            conditions: vec![SortCondition {
                range: Range::new(CellRef::new(0, 0), CellRef::new(10, 0)),
                descending: true,
            }],
        }),
        raw_xml: vec![
            r#"<extLst><ext uri="{00000000-0000-0000-0000-000000000000}"/></extLst>"#.to_string(),
        ],
    };

    let json = serde_json::to_string(&filter).unwrap();
    let reparsed: SheetAutoFilter = serde_json::from_str(&json).unwrap();
    assert_eq!(reparsed, filter);
}

#[test]
fn legacy_table_filter_column_values_still_deserializes() {
    let payload = json!({
        "col_id": 0,
        "values": ["Apple", "Cherry"]
    });

    let col: FilterColumn = serde_json::from_value(payload).unwrap();
    assert_eq!(col.col_id, 0);
    assert_eq!(col.values, vec!["Apple", "Cherry"]);
    assert_eq!(col.join, FilterJoin::Any);
    assert_eq!(col.criteria, Vec::<FilterCriterion>::new());
}

#[test]
fn sheet_autofilter_defaults_missing_optional_fields() {
    let payload = json!({
        "range": {
            "start": { "row": 0, "col": 0 },
            "end": { "row": 1, "col": 0 }
        }
    });

    let filter: SheetAutoFilter = serde_json::from_value(payload).unwrap();
    assert!(filter.filter_columns.is_empty());
    assert!(filter.sort_state.is_none());
    assert!(filter.raw_xml.is_empty());
}
