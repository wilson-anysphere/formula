use formula_model::autofilter::{
    DateComparison, FilterCriterion, FilterJoin, FilterValue, NumberComparison, OpaqueCustomFilter,
    OpaqueDynamicFilter, SheetAutoFilter, SortCondition, SortState, TextMatch, TextMatchKind,
};
use formula_model::{Range, RangeParseError};
use quick_xml::events::Event;
use quick_xml::Reader;
use quick_xml::Writer;
use std::collections::BTreeMap;
use std::io::Cursor;
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
    InvalidRef(#[from] RangeParseError),
}

fn parse_xml_bool(val: &str) -> bool {
    let trimmed = val.trim();
    trimmed == "1" || trimmed.eq_ignore_ascii_case("true")
}

pub fn parse_autofilter(xml: &str) -> Result<SheetAutoFilter, AutoFilterParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut range: Option<Range> = None;
    let mut filter_columns: BTreeMap<u32, formula_model::autofilter::FilterColumn> = BTreeMap::new();
    let mut sort_state: Option<SortState> = None;
    let mut raw_xml: Vec<String> = Vec::new();

    let mut current_col: Option<u32> = None;
    let mut in_filters = false;
    let mut in_custom_filters = false;
    let mut in_sort_state = false;
    let mut in_autofilter = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"autoFilter" => {
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"ref" {
                            let val = a.unescape_value()?.to_string();
                            range = Some(Range::from_a1(&val)?);
                        }
                    }
                    in_autofilter = true;
                }
                b"filterColumn" => {
                    if !in_autofilter {
                        continue;
                    }
                    let mut col_id: Option<u32> = None;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"colId" {
                            col_id = Some(a.unescape_value()?.trim().parse().unwrap_or(0));
                        }
                    }
                    current_col = col_id;
                    if let Some(col_id) = current_col {
                        filter_columns.entry(col_id).or_insert_with(|| formula_model::autofilter::FilterColumn {
                            col_id,
                            join: FilterJoin::Any,
                            criteria: Vec::new(),
                            values: Vec::new(),
                            raw_xml: Vec::new(),
                        });
                    }
                }
                b"filters" => {
                    if !in_autofilter {
                        continue;
                    }
                    let Some(col) = current_col else { continue };
                    in_filters = true;
                    let Some(col_filter) = filter_columns.get_mut(&col) else { continue };
                    col_filter.join = FilterJoin::Any;
                    col_filter.criteria.clear();
                    col_filter.values.clear();
                    let mut include_blanks = false;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"blank" {
                            include_blanks = parse_xml_bool(a.unescape_value()?.as_ref());
                        }
                    }
                    if include_blanks {
                        col_filter.criteria.push(FilterCriterion::Blanks);
                    }
                }
                b"filter" => {
                    if !in_filters {
                        continue;
                    }
                    if !in_autofilter {
                        continue;
                    }
                    let Some(col) = current_col else { continue };
                    let Some(col_filter) = filter_columns.get_mut(&col) else { continue };

                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"val" {
                            let v = a.unescape_value()?.to_string();
                            col_filter.values.push(v.clone());
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
                    if !in_autofilter {
                        continue;
                    }
                    let Some(col) = current_col else { continue };
                    in_custom_filters = true;
                    let Some(col_filter) = filter_columns.get_mut(&col) else { continue };
                    col_filter.criteria.clear();
                    col_filter.values.clear();
                    col_filter.join = FilterJoin::Any;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"and"
                            && parse_xml_bool(a.unescape_value()?.as_ref())
                        {
                            col_filter.join = FilterJoin::All;
                        }
                    }
                }
                b"customFilter" => {
                    if !in_custom_filters {
                        continue;
                    }
                    if !in_autofilter {
                        continue;
                    }
                    let Some(col) = current_col else { continue };
                    let Some(col_filter) = filter_columns.get_mut(&col) else { continue };
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
                    let c = operator_to_criterion(&op, val.as_deref());
                    col_filter.criteria.push(c);
                }
                b"dynamicFilter" => {
                    if !in_autofilter {
                        continue;
                    }
                    let Some(col) = current_col else { continue };
                    let Some(col_filter) = filter_columns.get_mut(&col) else { continue };
                    col_filter.join = FilterJoin::Any;
                    col_filter.criteria.clear();
                    col_filter.values.clear();
                    let mut filter_type: Option<String> = None;
                    let mut value: Option<String> = None;
                    let mut max_value: Option<String> = None;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        match a.key.as_ref() {
                            b"type" => filter_type = Some(a.unescape_value()?.to_string()),
                            b"val" => value = Some(a.unescape_value()?.to_string()),
                            b"maxVal" => max_value = Some(a.unescape_value()?.to_string()),
                            _ => {}
                        }
                    }
                    if let Some(filter_type) = filter_type {
                        match filter_type.as_str() {
                            "today" => col_filter.criteria.push(FilterCriterion::Date(DateComparison::Today)),
                            "yesterday" => col_filter
                                .criteria
                                .push(FilterCriterion::Date(DateComparison::Yesterday)),
                            "tomorrow" => col_filter
                                .criteria
                                .push(FilterCriterion::Date(DateComparison::Tomorrow)),
                            _ => col_filter.criteria.push(FilterCriterion::OpaqueDynamic(
                                OpaqueDynamicFilter {
                                    filter_type,
                                    value,
                                    max_value,
                                },
                            )),
                        }
                    }
                }
                b"sortState" => {
                    if !in_autofilter {
                        continue;
                    }
                    in_sort_state = true;
                    sort_state = Some(SortState { conditions: Vec::new() });
                }
                b"sortCondition" => {
                    if !in_sort_state {
                        continue;
                    }
                    if !in_autofilter {
                        continue;
                    }
                    let Some(sort_state) = sort_state.as_mut() else { continue };
                    let mut reference: Option<String> = None;
                    let mut descending = false;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        match a.key.as_ref() {
                            b"ref" => reference = Some(a.unescape_value()?.to_string()),
                            b"descending" => {
                                descending = parse_xml_bool(a.unescape_value()?.as_ref())
                            }
                            _ => {}
                        }
                    }
                    if let Some(reference) = reference {
                        let range = Range::from_a1(&reference)?;
                        sort_state.conditions.push(SortCondition { range, descending });
                    }
                }
                _ => {
                    if !in_autofilter {
                        continue;
                    }
                    let xml = capture_element_xml(&mut reader, Event::Start(e.into_owned()), &mut buf)?;
                    if let Some(col) = current_col {
                        if let Some(col_filter) = filter_columns.get_mut(&col) {
                            col_filter.raw_xml.push(xml);
                        }
                    } else {
                        raw_xml.push(xml);
                    }
                }
            },
            Event::Empty(e) => match e.local_name().as_ref() {
                b"autoFilter" => {
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"ref" {
                            let val = a.unescape_value()?.to_string();
                            range = Some(Range::from_a1(&val)?);
                        }
                    }
                    // Empty `<autoFilter/>` has no children.
                    break;
                }
                b"filterColumn" => {
                    if !in_autofilter {
                        continue;
                    }
                    let mut col_id: Option<u32> = None;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"colId" {
                            col_id = Some(a.unescape_value()?.trim().parse().unwrap_or(0));
                        }
                    }
                    if let Some(col_id) = col_id {
                        filter_columns.entry(col_id).or_insert_with(|| formula_model::autofilter::FilterColumn {
                            col_id,
                            join: FilterJoin::Any,
                            criteria: Vec::new(),
                            values: Vec::new(),
                            raw_xml: Vec::new(),
                        });
                    }
                }
                b"filter" => {
                    // `<filter val="..."/>`
                    if !in_filters {
                        continue;
                    }
                    let Some(col) = current_col else { continue };
                    let Some(col_filter) = filter_columns.get_mut(&col) else { continue };

                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        if a.key.as_ref() == b"val" {
                            let v = a.unescape_value()?.to_string();
                            col_filter.values.push(v.clone());
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
                    if !in_custom_filters {
                        continue;
                    }
                    let Some(col) = current_col else { continue };
                    let Some(col_filter) = filter_columns.get_mut(&col) else { continue };
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
                    col_filter
                        .criteria
                        .push(operator_to_criterion(&op, val.as_deref()));
                }
                b"dynamicFilter" => {
                    // `<dynamicFilter type="today"/>`
                    if !in_autofilter {
                        continue;
                    }
                    let Some(col) = current_col else { continue };
                    let Some(col_filter) = filter_columns.get_mut(&col) else { continue };
                    col_filter.join = FilterJoin::Any;
                    col_filter.criteria.clear();
                    col_filter.values.clear();
                    let mut filter_type: Option<String> = None;
                    let mut value: Option<String> = None;
                    let mut max_value: Option<String> = None;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        match a.key.as_ref() {
                            b"type" => filter_type = Some(a.unescape_value()?.to_string()),
                            b"val" => value = Some(a.unescape_value()?.to_string()),
                            b"maxVal" => max_value = Some(a.unescape_value()?.to_string()),
                            _ => {}
                        }
                    }
                    if let Some(filter_type) = filter_type {
                        match filter_type.as_str() {
                            "today" => col_filter.criteria.push(FilterCriterion::Date(DateComparison::Today)),
                            "yesterday" => col_filter
                                .criteria
                                .push(FilterCriterion::Date(DateComparison::Yesterday)),
                            "tomorrow" => col_filter
                                .criteria
                                .push(FilterCriterion::Date(DateComparison::Tomorrow)),
                            _ => col_filter.criteria.push(FilterCriterion::OpaqueDynamic(
                                OpaqueDynamicFilter {
                                    filter_type,
                                    value,
                                    max_value,
                                },
                            )),
                        }
                    }
                }
                b"sortState" => {
                    if !in_autofilter {
                        continue;
                    }
                    sort_state = Some(SortState { conditions: Vec::new() });
                }
                b"sortCondition" => {
                    let Some(sort_state) = sort_state.as_mut() else { continue };
                    let mut reference: Option<String> = None;
                    let mut descending = false;
                    for a in e.attributes().with_checks(false) {
                        let a = a?;
                        match a.key.as_ref() {
                            b"ref" => reference = Some(a.unescape_value()?.to_string()),
                            b"descending" => {
                                descending = parse_xml_bool(a.unescape_value()?.as_ref())
                            }
                            _ => {}
                        }
                    }
                    if let Some(reference) = reference {
                        let range = Range::from_a1(&reference)?;
                        sort_state.conditions.push(SortCondition { range, descending });
                    }
                }
                _ => {
                    if !in_autofilter {
                        continue;
                    }
                    let xml = capture_element_xml(&mut reader, Event::Empty(e.into_owned()), &mut buf)?;
                    if let Some(col) = current_col {
                        if let Some(col_filter) = filter_columns.get_mut(&col) {
                            col_filter.raw_xml.push(xml);
                        }
                    } else {
                        raw_xml.push(xml);
                    }
                }
            },
            Event::End(e) => match e.local_name().as_ref() {
                b"filterColumn" => current_col = None,
                b"filters" => in_filters = false,
                b"customFilters" => in_custom_filters = false,
                b"sortState" => in_sort_state = false,
                b"autoFilter" => break,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    let mut filter_columns = filter_columns.into_values().collect::<Vec<_>>();
    filter_columns.sort_by_key(|c| c.col_id);
    for col in &mut filter_columns {
        col.raw_xml.sort();
    }
    raw_xml.sort();

    Ok(SheetAutoFilter {
        range: range.ok_or(AutoFilterParseError::MissingRef)?,
        filter_columns,
        sort_state,
        raw_xml,
    })
}

fn capture_element_xml<B: std::io::BufRead>(
    reader: &mut Reader<B>,
    first: Event<'static>,
    buf: &mut Vec<u8>,
) -> Result<String, quick_xml::Error> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    match first {
        Event::Empty(e) => {
            writer.write_event(Event::Empty(e))?;
        }
        Event::Start(e) => {
            writer.write_event(Event::Start(e))?;
            let mut depth: usize = 0;
            loop {
                match reader.read_event_into(buf)? {
                    Event::Start(e) => {
                        depth += 1;
                        writer.write_event(Event::Start(e.into_owned()))?;
                    }
                    Event::Empty(e) => {
                        writer.write_event(Event::Empty(e.into_owned()))?;
                    }
                    Event::End(e) => {
                        writer.write_event(Event::End(e.into_owned()))?;
                        if depth == 0 {
                            break;
                        }
                        depth = depth.saturating_sub(1);
                    }
                    Event::Eof => break,
                    ev => {
                        writer.write_event(ev.into_owned())?;
                    }
                }
                buf.clear();
            }
        }
        _ => {}
    }

    let bytes = writer.into_inner().into_inner();
    Ok(String::from_utf8_lossy(&bytes).to_string())
}

fn operator_to_criterion(operator: &str, val: Option<&str>) -> FilterCriterion {
    let val = val.unwrap_or_default();
    let as_number = val.parse::<f64>().ok();
    match operator {
        "equal" => match as_number {
            Some(n) => FilterCriterion::Equals(FilterValue::Number(n)),
            None => FilterCriterion::Equals(FilterValue::Text(val.to_string())),
        },
        "notEqual" => {
            if val.is_empty() {
                FilterCriterion::NonBlanks
            } else if let Some(as_number) = as_number {
                FilterCriterion::Number(NumberComparison::NotEqual(as_number))
            } else {
                FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                    operator: operator.to_string(),
                    value: Some(val.to_string()),
                })
            }
        }
        "greaterThan" => as_number
            .map(|n| FilterCriterion::Number(NumberComparison::GreaterThan(n)))
            .unwrap_or_else(|| FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                operator: operator.to_string(),
                value: Some(val.to_string()),
            })),
        "greaterThanOrEqual" => as_number
            .map(|n| FilterCriterion::Number(NumberComparison::GreaterThanOrEqual(n)))
            .unwrap_or_else(|| FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                operator: operator.to_string(),
                value: Some(val.to_string()),
            })),
        "lessThan" => as_number
            .map(|n| FilterCriterion::Number(NumberComparison::LessThan(n)))
            .unwrap_or_else(|| FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                operator: operator.to_string(),
                value: Some(val.to_string()),
            })),
        "lessThanOrEqual" => as_number
            .map(|n| FilterCriterion::Number(NumberComparison::LessThanOrEqual(n)))
            .unwrap_or_else(|| FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                operator: operator.to_string(),
                value: Some(val.to_string()),
            })),
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
        _ => FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
            operator: operator.to_string(),
            value: Some(val.to_string()),
        }),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn parse_filters_list() {
        let xml = r#"<autoFilter ref="A1:D3"><filterColumn colId=" 2 "><filters blank=" true "><filter val="Alice"/></filters></filterColumn></autoFilter>"#;
        let filter = parse_autofilter(xml).unwrap();
        assert_eq!(filter.range.start, formula_model::CellRef::new(0, 0));
        assert_eq!(filter.filter_columns.len(), 1);
        let col = &filter.filter_columns[0];
        assert_eq!(col.col_id, 2);
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
        let xml = r#"<autoFilter ref="A1:D4"><filterColumn colId=" 1 "><customFilters and=" true "><customFilter operator="greaterThan" val="5"/><customFilter operator="lessThan" val="10"/></customFilters></filterColumn></autoFilter>"#;
        let filter = parse_autofilter(xml).unwrap();
        let col = &filter.filter_columns[0];
        assert_eq!(col.col_id, 1);
        assert_eq!(col.join, FilterJoin::All);
        assert_eq!(
            col.criteria,
            vec![
                FilterCriterion::Number(NumberComparison::GreaterThan(5.0)),
                FilterCriterion::Number(NumberComparison::LessThan(10.0))
            ]
        );
    }

    #[test]
    fn parse_sort_condition_descending_bool() {
        let xml = r#"<autoFilter ref="A1:D4"><sortState><sortCondition ref="C1:C4" descending=" true "/></sortState></autoFilter>"#;
        let filter = parse_autofilter(xml).unwrap();
        let sort_state = filter.sort_state.expect("sort state");
        assert_eq!(sort_state.conditions.len(), 1);
        assert_eq!(sort_state.conditions[0].descending, true);
    }
}
