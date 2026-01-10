use formula_engine::sort_filter::{
    to_a1_range, AutoFilter, ColumnFilter, DateComparison, FilterCriterion, FilterJoin, FilterValue,
    NumberComparison, TextMatchKind,
};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::Writer;
use std::io::Cursor;

pub fn write_autofilter(filter: &AutoFilter) -> Result<String, quick_xml::Error> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));

    let mut auto_filter = BytesStart::new("autoFilter");
    auto_filter.push_attribute(("ref", to_a1_range(filter.range).as_str()));
    writer.write_event(Event::Start(auto_filter))?;

    for (col_id, col_filter) in &filter.columns {
        write_filter_column(&mut writer, *col_id, col_filter)?;
    }

    writer.write_event(Event::End(BytesEnd::new("autoFilter")))?;
    let bytes = writer.into_inner().into_inner();
    Ok(String::from_utf8_lossy(&bytes).to_string())
}

fn write_filter_column<W: std::io::Write>(
    writer: &mut Writer<W>,
    col_id: usize,
    col_filter: &ColumnFilter,
) -> Result<(), quick_xml::Error> {
    let mut fc = BytesStart::new("filterColumn");
    fc.push_attribute(("colId", col_id.to_string().as_str()));
    writer.write_event(Event::Start(fc))?;

    if write_dynamic_filter(writer, col_filter)? {
        // Written.
    } else if can_write_as_filters(col_filter) {
        write_filters(writer, col_filter)?;
    } else {
        write_custom_filters(writer, col_filter)?;
    }

    writer.write_event(Event::End(BytesEnd::new("filterColumn")))?;
    Ok(())
}

fn write_dynamic_filter<W: std::io::Write>(
    writer: &mut Writer<W>,
    filter: &ColumnFilter,
) -> Result<bool, quick_xml::Error> {
    if filter.join != FilterJoin::Any || filter.criteria.len() != 1 {
        return Ok(false);
    }

    let filter_type = match &filter.criteria[0] {
        FilterCriterion::Date(DateComparison::Today) => Some("today"),
        FilterCriterion::Date(DateComparison::Yesterday) => Some("yesterday"),
        FilterCriterion::Date(DateComparison::Tomorrow) => Some("tomorrow"),
        _ => None,
    };

    let Some(filter_type) = filter_type else {
        return Ok(false);
    };

    let mut dyn_filter = BytesStart::new("dynamicFilter");
    dyn_filter.push_attribute(("type", filter_type));
    writer.write_event(Event::Empty(dyn_filter))?;
    Ok(true)
}

fn can_write_as_filters(filter: &ColumnFilter) -> bool {
    if filter.join != FilterJoin::Any {
        return false;
    }
    filter.criteria.iter().all(|c| matches!(c, FilterCriterion::Equals(_) | FilterCriterion::Blanks))
}

fn write_filters<W: std::io::Write>(
    writer: &mut Writer<W>,
    filter: &ColumnFilter,
) -> Result<(), quick_xml::Error> {
    let mut filters = BytesStart::new("filters");
    if filter.criteria.iter().any(|c| matches!(c, FilterCriterion::Blanks)) {
        filters.push_attribute(("blank", "1"));
    }
    writer.write_event(Event::Start(filters))?;

    for criterion in &filter.criteria {
        if let FilterCriterion::Equals(value) = criterion {
            let mut f = BytesStart::new("filter");
            f.push_attribute(("val", value_to_string(value).as_str()));
            writer.write_event(Event::Empty(f))?;
        }
    }

    writer.write_event(Event::End(BytesEnd::new("filters")))?;
    Ok(())
}

fn write_custom_filters<W: std::io::Write>(
    writer: &mut Writer<W>,
    filter: &ColumnFilter,
) -> Result<(), quick_xml::Error> {
    let mut entries: Vec<(String, Option<String>)> = Vec::new();
    let mut requires_and = filter.join == FilterJoin::All;

    for criterion in &filter.criteria {
        match criterion {
            FilterCriterion::Number(NumberComparison::Between { min, max }) => {
                requires_and = true;
                entries.push(("greaterThanOrEqual".into(), Some(min.to_string())));
                entries.push(("lessThanOrEqual".into(), Some(max.to_string())));
            }
            FilterCriterion::Date(DateComparison::Between { start, end }) => {
                requires_and = true;
                entries.push(("greaterThanOrEqual".into(), Some(start.to_string())));
                entries.push(("lessThanOrEqual".into(), Some(end.to_string())));
            }
            _ => entries.extend(criterion_to_custom_filters(criterion)),
        }
    }

    let mut custom = BytesStart::new("customFilters");
    if requires_and {
        custom.push_attribute(("and", "1"));
    }
    writer.write_event(Event::Start(custom))?;

    for (op, val) in entries {
        let mut cf = BytesStart::new("customFilter");
        cf.push_attribute(("operator", op.as_str()));
        if let Some(val) = val {
            cf.push_attribute(("val", val.as_str()));
        }
        writer.write_event(Event::Empty(cf))?;
    }

    writer.write_event(Event::End(BytesEnd::new("customFilters")))?;
    Ok(())
}

fn criterion_to_custom_filters(criterion: &FilterCriterion) -> Vec<(String, Option<String>)> {
    match criterion {
        FilterCriterion::Equals(v) => vec![("equal".into(), Some(value_to_string(v)))],
        FilterCriterion::TextMatch(m) => {
            let op = match m.kind {
                TextMatchKind::Contains => "contains",
                TextMatchKind::BeginsWith => "beginsWith",
                TextMatchKind::EndsWith => "endsWith",
            };
            vec![(op.into(), Some(m.pattern.clone()))]
        }
        FilterCriterion::Number(cmp) => match cmp {
            NumberComparison::GreaterThan(v) => vec![("greaterThan".into(), Some(v.to_string()))],
            NumberComparison::GreaterThanOrEqual(v) => {
                vec![("greaterThanOrEqual".into(), Some(v.to_string()))]
            }
            NumberComparison::LessThan(v) => vec![("lessThan".into(), Some(v.to_string()))],
            NumberComparison::LessThanOrEqual(v) => {
                vec![("lessThanOrEqual".into(), Some(v.to_string()))]
            }
            NumberComparison::Between { min, max } => vec![
                ("greaterThanOrEqual".into(), Some(min.to_string())),
                ("lessThanOrEqual".into(), Some(max.to_string())),
            ],
            NumberComparison::NotEqual(v) => vec![("notEqual".into(), Some(v.to_string()))],
        },
        FilterCriterion::Date(cmp) => match cmp {
            DateComparison::Today | DateComparison::Yesterday | DateComparison::Tomorrow => Vec::new(),
            DateComparison::OnDate(d) => vec![("equal".into(), Some(d.to_string()))],
            DateComparison::After(dt) => vec![("greaterThan".into(), Some(dt.to_string()))],
            DateComparison::Before(dt) => vec![("lessThan".into(), Some(dt.to_string()))],
            DateComparison::Between { start, end } => vec![
                ("greaterThanOrEqual".into(), Some(start.to_string())),
                ("lessThanOrEqual".into(), Some(end.to_string())),
            ],
        },
        FilterCriterion::Blanks => Vec::new(),
        FilterCriterion::NonBlanks => vec![("notEqual".into(), Some(String::new()))],
    }
}

fn value_to_string(value: &FilterValue) -> String {
    match value {
        FilterValue::Text(s) => s.clone(),
        FilterValue::Number(n) => n.to_string(),
        FilterValue::Bool(b) => b.to_string(),
        FilterValue::DateTime(dt) => dt.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_engine::sort_filter::RangeRef;
    use pretty_assertions::assert_eq;
    use std::collections::BTreeMap;

    #[test]
    fn write_and_parse_roundtrip_filters() {
        let filter = AutoFilter {
            range: RangeRef {
                start_row: 0,
                start_col: 0,
                end_row: 2,
                end_col: 0,
            },
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![
                        FilterCriterion::Blanks,
                        FilterCriterion::Equals(FilterValue::Text("Alice".into())),
                    ],
                },
            )]),
        };

        let xml = write_autofilter(&filter).unwrap();
        let parsed = crate::autofilter::parse_autofilter(&xml).unwrap();
        assert_eq!(parsed.range, filter.range);
        assert_eq!(parsed.columns.get(&0).unwrap().criteria.len(), 2);
    }

    #[test]
    fn write_dynamic_filter_today_roundtrip() {
        let filter = AutoFilter {
            range: RangeRef {
                start_row: 0,
                start_col: 0,
                end_row: 10,
                end_col: 0,
            },
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Date(DateComparison::Today)],
                },
            )]),
        };

        let xml = write_autofilter(&filter).unwrap();
        assert!(xml.contains("dynamicFilter"));
        let parsed = crate::autofilter::parse_autofilter(&xml).unwrap();
        assert_eq!(
            parsed.columns.get(&0).unwrap().criteria,
            vec![FilterCriterion::Date(DateComparison::Today)]
        );
    }

    #[test]
    fn write_between_expands_to_two_custom_filters() {
        let filter = AutoFilter {
            range: RangeRef {
                start_row: 0,
                start_col: 0,
                end_row: 10,
                end_col: 0,
            },
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Number(NumberComparison::Between {
                        min: 2.0,
                        max: 5.0,
                    })],
                },
            )]),
        };

        let xml = write_autofilter(&filter).unwrap();
        assert!(xml.contains("and=\"1\""));
        let parsed = crate::autofilter::parse_autofilter(&xml).unwrap();
        let col = parsed.columns.get(&0).unwrap();
        assert_eq!(col.join, FilterJoin::All);
        assert_eq!(
            col.criteria,
            vec![
                FilterCriterion::Number(NumberComparison::GreaterThanOrEqual(2.0)),
                FilterCriterion::Number(NumberComparison::LessThanOrEqual(5.0))
            ]
        );
    }
}
