use formula_engine::sort_filter::{
    parse_a1_range, AutoFilter, ColumnFilter, DateComparison, FilterCriterion, FilterJoin,
    FilterValue, NumberComparison, RangeRef, TextMatch, TextMatchKind,
};
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::BTreeMap;
use thiserror::Error;

#[derive(Debug, Error)]
pub enum AutoFilterParseError {
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("XML attribute error: {0}")]
    Attr(#[from] quick_xml::events::attributes::AttrError),
    #[error("missing autoFilter ref attribute")]
    MissingRef,
    #[error("invalid ref: {0}")]
    InvalidRef(#[from] formula_engine::sort_filter::A1ParseError),
}

pub fn parse_autofilter(xml: &str) -> Result<AutoFilter, AutoFilterParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut range: Option<RangeRef> = None;
    let mut columns: BTreeMap<usize, ColumnFilter> = BTreeMap::new();

    let mut current_col: Option<usize> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"autoFilter" => {
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"ref" {
                            let val = a.unescape_value()?.to_string();
                            range = Some(parse_a1_range(&val)?);
                        }
                    }
                }
                b"filterColumn" => {
                    let mut col_id: Option<usize> = None;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"colId" {
                            col_id = Some(a.unescape_value()?.parse().unwrap_or(0));
                        }
                    }
                    current_col = col_id;
                }
                b"filters" => {
                    let Some(col) = current_col else {
                        continue;
                    };
                    let mut criteria: Vec<FilterCriterion> = Vec::new();
                    let mut include_blanks = false;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"blank" {
                            include_blanks = a.unescape_value()?.as_ref() == "1";
                        }
                    }
                    if include_blanks {
                        criteria.push(FilterCriterion::Blanks);
                    }
                    columns.insert(
                        col,
                        ColumnFilter {
                            join: FilterJoin::Any,
                            criteria,
                        },
                    );
                }
                b"filter" => {
                    let Some(col) = current_col else {
                        continue;
                    };
                    let Some(col_filter) = columns.get_mut(&col) else {
                        continue;
                    };

                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"val" {
                            let v = a.unescape_value()?.to_string();
                            if let Ok(n) = v.parse::<f64>() {
                                col_filter
                                    .criteria
                                    .push(FilterCriterion::Equals(FilterValue::Number(n)));
                            } else {
                                col_filter
                                    .criteria
                                    .push(FilterCriterion::Equals(FilterValue::Text(v)));
                            }
                        }
                    }
                }
                b"customFilters" => {
                    let Some(col) = current_col else {
                        continue;
                    };
                    let mut join = FilterJoin::Any;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"and" && a.unescape_value()?.as_ref() == "1" {
                            join = FilterJoin::All;
                        }
                    }
                    columns.insert(
                        col,
                        ColumnFilter {
                            join,
                            criteria: Vec::new(),
                        },
                    );
                }
                b"customFilter" => {
                    let Some(col) = current_col else {
                        continue;
                    };
                    let Some(col_filter) = columns.get_mut(&col) else {
                        continue;
                    };
                    let mut operator: Option<String> = None;
                    let mut val: Option<String> = None;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        match a.key.as_ref() {
                            b"operator" => operator = Some(a.unescape_value()?.to_string()),
                            b"val" => val = Some(a.unescape_value()?.to_string()),
                            _ => {}
                        }
                    }

                    let op = operator.unwrap_or_else(|| "equal".to_string());
                    let val = val.unwrap_or_default();
                    if let Some(c) = operator_to_criterion(&op, &val) {
                        col_filter.criteria.push(c);
                    }
                }
                b"dynamicFilter" => {
                    let Some(col) = current_col else {
                        continue;
                    };
                    let col_filter = columns.entry(col).or_insert_with(|| ColumnFilter {
                        join: FilterJoin::Any,
                        criteria: Vec::new(),
                    });
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"type" {
                            match a.unescape_value()?.as_ref() {
                                "today" => col_filter.criteria.push(FilterCriterion::Date(DateComparison::Today)),
                                "yesterday" => col_filter.criteria.push(FilterCriterion::Date(DateComparison::Yesterday)),
                                "tomorrow" => col_filter.criteria.push(FilterCriterion::Date(DateComparison::Tomorrow)),
                                _ => {}
                            }
                        }
                    }
                }
                _ => {}
            },
            Event::Empty(e) => match e.local_name().as_ref() {
                b"filter" => {
                    // `<filter val="..."/>`
                    let Some(col) = current_col else {
                        continue;
                    };
                    let Some(col_filter) = columns.get_mut(&col) else {
                        continue;
                    };

                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"val" {
                            let v = a.unescape_value()?.to_string();
                            if let Ok(n) = v.parse::<f64>() {
                                col_filter
                                    .criteria
                                    .push(FilterCriterion::Equals(FilterValue::Number(n)));
                            } else {
                                col_filter
                                    .criteria
                                    .push(FilterCriterion::Equals(FilterValue::Text(v)));
                            }
                        }
                    }
                }
                b"customFilter" => {
                    // `<customFilter operator="..." val="..."/>`
                    let Some(col) = current_col else {
                        continue;
                    };
                    let Some(col_filter) = columns.get_mut(&col) else {
                        continue;
                    };
                    let mut operator: Option<String> = None;
                    let mut val: Option<String> = None;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        match a.key.as_ref() {
                            b"operator" => operator = Some(a.unescape_value()?.to_string()),
                            b"val" => val = Some(a.unescape_value()?.to_string()),
                            _ => {}
                        }
                    }

                    let op = operator.unwrap_or_else(|| "equal".to_string());
                    let val = val.unwrap_or_default();
                    if let Some(c) = operator_to_criterion(&op, &val) {
                        col_filter.criteria.push(c);
                    }
                }
                b"dynamicFilter" => {
                    // `<dynamicFilter type="today"/>`
                    let Some(col) = current_col else {
                        continue;
                    };
                    let col_filter = columns.entry(col).or_insert_with(|| ColumnFilter {
                        join: FilterJoin::Any,
                        criteria: Vec::new(),
                    });
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"type" {
                            match a.unescape_value()?.as_ref() {
                                "today" => col_filter
                                    .criteria
                                    .push(FilterCriterion::Date(DateComparison::Today)),
                                "yesterday" => col_filter
                                    .criteria
                                    .push(FilterCriterion::Date(DateComparison::Yesterday)),
                                "tomorrow" => col_filter
                                    .criteria
                                    .push(FilterCriterion::Date(DateComparison::Tomorrow)),
                                _ => {}
                            }
                        }
                    }
                }
                _ => {}
            },
            Event::End(e) => match e.local_name().as_ref() {
                b"filterColumn" => current_col = None,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(AutoFilter {
        range: range.ok_or(AutoFilterParseError::MissingRef)?,
        columns,
    })
}

fn operator_to_criterion(operator: &str, val: &str) -> Option<FilterCriterion> {
    let as_number = val.parse::<f64>().ok();
    Some(match operator {
        "equal" => match as_number {
            Some(n) => FilterCriterion::Equals(FilterValue::Number(n)),
            None => FilterCriterion::Equals(FilterValue::Text(val.to_string())),
        },
        "notEqual" => {
            if val.is_empty() {
                FilterCriterion::NonBlanks
            } else {
                FilterCriterion::Number(NumberComparison::NotEqual(as_number?))
            }
        }
        "greaterThan" => FilterCriterion::Number(NumberComparison::GreaterThan(as_number?)),
        "greaterThanOrEqual" => FilterCriterion::Number(NumberComparison::GreaterThanOrEqual(as_number?)),
        "lessThan" => FilterCriterion::Number(NumberComparison::LessThan(as_number?)),
        "lessThanOrEqual" => FilterCriterion::Number(NumberComparison::LessThanOrEqual(as_number?)),
        "contains" => FilterCriterion::TextMatch(TextMatch {
            kind: TextMatchKind::Contains,
            pattern: val.to_string(),
            case_sensitive: false,
        }),
        "beginsWith" => FilterCriterion::TextMatch(TextMatch {
            kind: TextMatchKind::BeginsWith,
            pattern: val.to_string(),
            case_sensitive: false,
        }),
        "endsWith" => FilterCriterion::TextMatch(TextMatch {
            kind: TextMatchKind::EndsWith,
            pattern: val.to_string(),
            case_sensitive: false,
        }),
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn parse_filters_list() {
        let xml = r#"<autoFilter ref="A1:A3"><filterColumn colId="0"><filters blank="1"><filter val="Alice"/></filters></filterColumn></autoFilter>"#;
        let filter = parse_autofilter(xml).unwrap();
        assert_eq!(filter.range.start_row, 0);
        assert_eq!(filter.columns.len(), 1);
        let col = filter.columns.get(&0).unwrap();
        assert_eq!(col.join, FilterJoin::Any);
        assert_eq!(
            col.criteria,
            vec![
                FilterCriterion::Blanks,
                FilterCriterion::Equals(FilterValue::Text("Alice".into()))
            ]
        );
    }

    #[test]
    fn parse_custom_filters() {
        let xml = r#"<autoFilter ref="A1:A4"><filterColumn colId="0"><customFilters and="1"><customFilter operator="greaterThan" val="5"/><customFilter operator="lessThan" val="10"/></customFilters></filterColumn></autoFilter>"#;
        let filter = parse_autofilter(xml).unwrap();
        let col = filter.columns.get(&0).unwrap();
        assert_eq!(col.join, FilterJoin::All);
        assert_eq!(
            col.criteria,
            vec![
                FilterCriterion::Number(NumberComparison::GreaterThan(5.0)),
                FilterCriterion::Number(NumberComparison::LessThan(10.0))
            ]
        );
    }
}
