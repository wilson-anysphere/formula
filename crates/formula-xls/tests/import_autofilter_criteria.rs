use formula_model::autofilter::{
    FilterColumn, FilterCriterion, FilterJoin, FilterValue, NumberComparison,
};
use formula_model::Range;

mod common;

use common::xls_fixture_builder;

#[test]
fn import_autofilter_criteria() {
    let bytes = xls_fixture_builder::build_autofilter_criteria_fixture_xls();
    let result = formula_xls::import_xls_bytes(&bytes).expect("import xls bytes");

    let sheet = result
        .workbook
        .sheet_by_name("FilterCriteria")
        .expect("FilterCriteria missing");
    let af = sheet.auto_filter.as_ref().expect("auto_filter missing");

    assert_eq!(af.range, Range::from_a1("A1:C5").unwrap());
    assert_eq!(
        af.filter_columns,
        vec![
            FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Equals(FilterValue::Text("Alice".to_string()))],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 1,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Number(NumberComparison::GreaterThan(1.0))],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 2,
                join: FilterJoin::All,
                criteria: vec![
                    FilterCriterion::Number(NumberComparison::GreaterThan(10.0)),
                    FilterCriterion::Number(NumberComparison::LessThan(20.0)),
                ],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
        ],
        "unexpected filter columns; warnings={:?}",
        result.warnings
    );
    assert!(af.sort_state.is_none());

    assert!(
        !result
            .warnings
            .iter()
            .any(|w| w.message.contains("failed to fully import `.xls` autofilter criteria")),
        "unexpected `.xls` autofilter criteria warning; warnings={:?}",
        result.warnings
    );
}

#[test]
fn import_autofilter_criteria_absolute_entry_index() {
    let bytes = xls_fixture_builder::build_autofilter_criteria_absolute_entry_fixture_xls();
    let result = formula_xls::import_xls_bytes(&bytes).expect("import xls bytes");

    let sheet = result
        .workbook
        .sheet_by_name("FilterCriteriaAbsEntry")
        .expect("FilterCriteriaAbsEntry missing");
    let af = sheet.auto_filter.as_ref().expect("auto_filter missing");
    assert_eq!(af.range, Range::from_a1("D1:F5").unwrap());

    assert_eq!(
        af.filter_columns,
        vec![
            FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Equals(FilterValue::Text("Alice".to_string()))],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 2,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Number(NumberComparison::GreaterThan(1.0))],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
        ],
        "unexpected filter columns; warnings={:?}",
        result.warnings
    );
    assert!(af.sort_state.is_none());

    assert!(
        !result
            .warnings
            .iter()
            .any(|w| w.message.contains("failed to fully import `.xls` autofilter criteria")),
        "unexpected `.xls` autofilter criteria warning; warnings={:?}",
        result.warnings
    );
}
