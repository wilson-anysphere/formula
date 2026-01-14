use std::io::Write;

use formula_model::autofilter::{
    FilterColumn, FilterCriterion, FilterJoin, FilterValue, NumberComparison, OpaqueCustomFilter,
};
use formula_model::Range;

mod common;

use common::xls_fixture_builder;

#[test]
fn import_autofilter_criteria() {
    let bytes = xls_fixture_builder::build_autofilter_criteria_fixture_xls();
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes).expect("write xls bytes");
    let result = formula_xls::import_xls_path(tmp.path()).expect("import xls");

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
                join: FilterJoin::All,
                criteria: vec![
                    FilterCriterion::Number(NumberComparison::GreaterThanOrEqual(10.0)),
                    FilterCriterion::Number(NumberComparison::LessThanOrEqual(20.0)),
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
fn import_autofilter_criteria_join_all() {
    let bytes = xls_fixture_builder::build_autofilter_criteria_join_all_fixture_xls();
    let result = formula_xls::import_xls_bytes(&bytes).expect("import xls bytes");

    let sheet = result
        .workbook
        .sheet_by_name("FilterCriteriaJoinAll")
        .expect("FilterCriteriaJoinAll missing");
    let af = sheet.auto_filter.as_ref().expect("auto_filter missing");

    assert_eq!(af.range, Range::from_a1("A1:A5").unwrap());
    assert_eq!(
        af.filter_columns,
        vec![FilterColumn {
            col_id: 0,
            join: FilterJoin::All,
            criteria: vec![
                FilterCriterion::Number(NumberComparison::GreaterThan(10.0)),
                FilterCriterion::Number(NumberComparison::LessThan(20.0)),
            ],
            values: Vec::new(),
            raw_xml: Vec::new(),
        }],
        "unexpected filter columns; warnings={:?}",
        result.warnings
    );

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
fn import_autofilter_criteria_operator_byte1_fallback() {
    let bytes = xls_fixture_builder::build_autofilter_criteria_operator_byte1_fixture_xls();
    let result = formula_xls::import_xls_bytes(&bytes).expect("import xls bytes");

    let sheet = result
        .workbook
        .sheet_by_name("FilterCriteriaOpByte1")
        .expect("FilterCriteriaOpByte1 missing");
    let af = sheet.auto_filter.as_ref().expect("auto_filter missing");

    assert_eq!(af.range, Range::from_a1("A1:A5").unwrap());
    assert_eq!(
        af.filter_columns,
        vec![FilterColumn {
            col_id: 0,
            join: FilterJoin::Any,
            criteria: vec![FilterCriterion::Equals(FilterValue::Text("Alice".to_string()))],
            values: Vec::new(),
            raw_xml: Vec::new(),
        }],
        "unexpected filter columns; warnings={:?}",
        result.warnings
    );

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
fn import_autofilter_criteria_between_operator_codes() {
    let bytes = xls_fixture_builder::build_autofilter_criteria_between_fixture_xls();
    let result = formula_xls::import_xls_bytes(&bytes).expect("import xls bytes");

    let sheet = result
        .workbook
        .sheet_by_name("FilterCriteriaBetween")
        .expect("FilterCriteriaBetween missing");
    let af = sheet.auto_filter.as_ref().expect("auto_filter missing");

    assert_eq!(af.range, Range::from_a1("A1:B5").unwrap());
    assert_eq!(
        af.filter_columns,
        vec![
            FilterColumn {
                col_id: 0,
                join: FilterJoin::All,
                criteria: vec![
                    FilterCriterion::Number(NumberComparison::GreaterThanOrEqual(10.0)),
                    FilterCriterion::Number(NumberComparison::LessThanOrEqual(20.0)),
                ],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 1,
                join: FilterJoin::Any,
                criteria: vec![
                    FilterCriterion::Number(NumberComparison::LessThan(10.0)),
                    FilterCriterion::Number(NumberComparison::GreaterThan(20.0)),
                ],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
        ],
        "unexpected filter columns; warnings={:?}",
        result.warnings
    );

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
fn import_autofilter_criteria_blanks_and_nonblanks() {
    let bytes = xls_fixture_builder::build_autofilter_criteria_blanks_fixture_xls();
    let result = formula_xls::import_xls_bytes(&bytes).expect("import xls bytes");

    let sheet = result
        .workbook
        .sheet_by_name("FilterCriteriaBlanks")
        .expect("FilterCriteriaBlanks missing");
    let af = sheet.auto_filter.as_ref().expect("auto_filter missing");

    assert_eq!(af.range, Range::from_a1("A1:D5").unwrap());
    assert_eq!(
        af.filter_columns,
        vec![
            FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Blanks],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 1,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::NonBlanks],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 2,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Blanks],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 3,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::NonBlanks],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
        ],
        "unexpected filter columns; warnings={:?}",
        result.warnings
    );

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
fn import_autofilter_criteria_text_operators_are_preserved_as_opaque_custom() {
    let bytes = xls_fixture_builder::build_autofilter_criteria_text_ops_fixture_xls();
    let result = formula_xls::import_xls_bytes(&bytes).expect("import xls bytes");

    let sheet = result
        .workbook
        .sheet_by_name("FilterCriteriaTextOps")
        .expect("FilterCriteriaTextOps missing");
    let af = sheet.auto_filter.as_ref().expect("auto_filter missing");

    assert_eq!(af.range, Range::from_a1("A1:C5").unwrap());
    assert_eq!(
        af.filter_columns,
        vec![
            FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                    operator: "contains".to_string(),
                    value: Some("Al".to_string()),
                })],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 1,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                    operator: "beginsWith".to_string(),
                    value: Some("B".to_string()),
                })],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 2,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                    operator: "endsWith".to_string(),
                    value: Some("z".to_string()),
                })],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
        ],
        "unexpected filter columns; warnings={:?}",
        result.warnings
    );

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
fn import_autofilter_criteria_negative_text_operators_are_preserved_as_opaque_custom() {
    let bytes = xls_fixture_builder::build_autofilter_criteria_text_ops_negative_fixture_xls();
    let result = formula_xls::import_xls_bytes(&bytes).expect("import xls bytes");

    let sheet = result
        .workbook
        .sheet_by_name("FilterCriteriaTextOpsNeg")
        .expect("FilterCriteriaTextOpsNeg missing");
    let af = sheet.auto_filter.as_ref().expect("auto_filter missing");

    assert_eq!(af.range, Range::from_a1("A1:C5").unwrap());
    assert_eq!(
        af.filter_columns,
        vec![
            FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                    operator: "doesNotContain".to_string(),
                    value: Some("Al".to_string()),
                })],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 1,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                    operator: "doesNotBeginWith".to_string(),
                    value: Some("B".to_string()),
                })],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 2,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                    operator: "doesNotEndWith".to_string(),
                    value: Some("z".to_string()),
                })],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
        ],
        "unexpected filter columns; warnings={:?}",
        result.warnings
    );

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
fn import_autofilter_criteria_bool_values() {
    let bytes = xls_fixture_builder::build_autofilter_criteria_bool_fixture_xls();
    let result = formula_xls::import_xls_bytes(&bytes).expect("import xls bytes");

    let sheet = result
        .workbook
        .sheet_by_name("FilterCriteriaBool")
        .expect("FilterCriteriaBool missing");
    let af = sheet.auto_filter.as_ref().expect("auto_filter missing");

    assert_eq!(af.range, Range::from_a1("A1:B5").unwrap());
    assert_eq!(
        af.filter_columns,
        vec![
            FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Equals(FilterValue::Bool(true))],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 1,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Equals(FilterValue::Bool(false))],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
        ],
        "unexpected filter columns; warnings={:?}",
        result.warnings
    );

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
fn import_autofilter_criteria_top10_is_preserved_as_raw_xml() {
    let bytes = xls_fixture_builder::build_autofilter_criteria_top10_fixture_xls();
    let result = formula_xls::import_xls_bytes(&bytes).expect("import xls bytes");

    let sheet = result
        .workbook
        .sheet_by_name("FilterCriteriaTop10")
        .expect("FilterCriteriaTop10 missing");
    let af = sheet.auto_filter.as_ref().expect("auto_filter missing");

    assert_eq!(af.range, Range::from_a1("A1:A5").unwrap());
    assert_eq!(
        af.filter_columns,
        vec![FilterColumn {
            col_id: 0,
            join: FilterJoin::Any,
            criteria: Vec::new(),
            values: Vec::new(),
            raw_xml: vec!["<top10 top=\"1\" percent=\"1\" val=\"5\"/>".to_string()],
        }],
        "unexpected filter columns; warnings={:?}",
        result.warnings
    );

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
