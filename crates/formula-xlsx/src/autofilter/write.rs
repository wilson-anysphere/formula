use formula_model::autofilter::{
    DateComparison, FilterCriterion, FilterJoin, FilterValue, NumberComparison, OpaqueDynamicFilter,
    SheetAutoFilter, TextMatchKind,
};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::Writer;
use std::io::Cursor;
use std::sync::Arc;

pub fn write_autofilter(filter: &SheetAutoFilter) -> Result<String, quick_xml::Error> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    write_autofilter_to(&mut writer, filter, None)?;
    let bytes = writer.into_inner().into_inner();
    Ok(String::from_utf8_lossy(&bytes).to_string())
}

#[derive(Debug, Clone)]
struct AutoFilterTags {
    auto_filter: String,
    filter_column: String,
    dynamic_filter: String,
    filters: String,
    filter: String,
    custom_filters: String,
    custom_filter: String,
    sort_state: String,
    sort_condition: String,
}

impl AutoFilterTags {
    fn new(prefix: Option<&str>) -> Self {
        Self {
            auto_filter: crate::xml::prefixed_tag(prefix, "autoFilter"),
            filter_column: crate::xml::prefixed_tag(prefix, "filterColumn"),
            dynamic_filter: crate::xml::prefixed_tag(prefix, "dynamicFilter"),
            filters: crate::xml::prefixed_tag(prefix, "filters"),
            filter: crate::xml::prefixed_tag(prefix, "filter"),
            custom_filters: crate::xml::prefixed_tag(prefix, "customFilters"),
            custom_filter: crate::xml::prefixed_tag(prefix, "customFilter"),
            sort_state: crate::xml::prefixed_tag(prefix, "sortState"),
            sort_condition: crate::xml::prefixed_tag(prefix, "sortCondition"),
        }
    }
}

pub(crate) fn write_autofilter_to<W: std::io::Write>(
    writer: &mut Writer<W>,
    filter: &SheetAutoFilter,
    prefix: Option<&str>,
) -> Result<(), quick_xml::Error> {
    let tags = AutoFilterTags::new(prefix);

    let mut auto_filter = BytesStart::new(tags.auto_filter.as_str());
    auto_filter.push_attribute(("ref", filter.range.to_string().as_str()));
    writer.write_event(Event::Start(auto_filter))?;

    for col in &filter.filter_columns {
        write_filter_column(writer, col, &tags)?;
    }

    if let Some(sort_state) = &filter.sort_state {
        write_sort_state(writer, sort_state, &tags)?;
    }

    for raw in &filter.raw_xml {
        writer
            .get_mut()
            .write_all(raw.as_bytes())
            .map_err(|e| quick_xml::Error::Io(Arc::new(e)))?;
    }

    writer.write_event(Event::End(BytesEnd::new(tags.auto_filter.as_str())))?;
    Ok(())
}

fn write_filter_column<W: std::io::Write>(
    writer: &mut Writer<W>,
    col: &formula_model::autofilter::FilterColumn,
    tags: &AutoFilterTags,
) -> Result<(), quick_xml::Error> {
    let mut fc = BytesStart::new(tags.filter_column.as_str());
    fc.push_attribute(("colId", col.col_id.to_string().as_str()));
    writer.write_event(Event::Start(fc))?;

    if write_dynamic_filter(writer, col, tags)? {
        // Written.
    } else if can_write_as_filters(col) {
        write_filters(writer, col, tags)?;
    } else {
        write_custom_filters(writer, col, tags)?;
    }

    for raw in &col.raw_xml {
        writer
            .get_mut()
            .write_all(raw.as_bytes())
            .map_err(|e| quick_xml::Error::Io(Arc::new(e)))?;
    }

    writer.write_event(Event::End(BytesEnd::new(tags.filter_column.as_str())))?;
    Ok(())
}

fn write_dynamic_filter<W: std::io::Write>(
    writer: &mut Writer<W>,
    col: &formula_model::autofilter::FilterColumn,
    tags: &AutoFilterTags,
) -> Result<bool, quick_xml::Error> {
    let criteria = effective_criteria(col);
    if col.join != FilterJoin::Any || criteria.len() != 1 {
        return Ok(false);
    }

    match &criteria[0] {
        FilterCriterion::Date(DateComparison::Today) => {
            write_dynamic_filter_element(writer, tags, "today", None, None)?;
            Ok(true)
        }
        FilterCriterion::Date(DateComparison::Yesterday) => {
            write_dynamic_filter_element(writer, tags, "yesterday", None, None)?;
            Ok(true)
        }
        FilterCriterion::Date(DateComparison::Tomorrow) => {
            write_dynamic_filter_element(writer, tags, "tomorrow", None, None)?;
            Ok(true)
        }
        FilterCriterion::OpaqueDynamic(OpaqueDynamicFilter {
            filter_type,
            value,
            max_value,
        }) => {
            write_dynamic_filter_element(
                writer,
                tags,
                filter_type,
                value.as_deref(),
                max_value.as_deref(),
            )?;
            Ok(true)
        }
        _ => Ok(false),
    }
}

fn write_dynamic_filter_element<W: std::io::Write>(
    writer: &mut Writer<W>,
    tags: &AutoFilterTags,
    filter_type: &str,
    value: Option<&str>,
    max_value: Option<&str>,
) -> Result<(), quick_xml::Error> {
    let mut dyn_filter = BytesStart::new(tags.dynamic_filter.as_str());
    dyn_filter.push_attribute(("type", filter_type));
    if let Some(value) = value {
        dyn_filter.push_attribute(("val", value));
    }
    if let Some(max_value) = max_value {
        dyn_filter.push_attribute(("maxVal", max_value));
    }
    writer.write_event(Event::Empty(dyn_filter))?;
    Ok(())
}

fn can_write_as_filters(col: &formula_model::autofilter::FilterColumn) -> bool {
    if col.join != FilterJoin::Any {
        return false;
    }
    effective_criteria(col)
        .iter()
        .all(|c| matches!(c, FilterCriterion::Equals(_) | FilterCriterion::Blanks))
}

fn write_filters<W: std::io::Write>(
    writer: &mut Writer<W>,
    col: &formula_model::autofilter::FilterColumn,
    tags: &AutoFilterTags,
) -> Result<(), quick_xml::Error> {
    let criteria = effective_criteria(col);
    let mut filters = BytesStart::new(tags.filters.as_str());
    if criteria.iter().any(|c| matches!(c, FilterCriterion::Blanks)) {
        filters.push_attribute(("blank", "1"));
    }
    writer.write_event(Event::Start(filters))?;

    for criterion in &criteria {
        if let FilterCriterion::Equals(value) = criterion {
            let mut f = BytesStart::new(tags.filter.as_str());
            f.push_attribute(("val", value_to_string(value).as_str()));
            writer.write_event(Event::Empty(f))?;
        }
    }

    writer.write_event(Event::End(BytesEnd::new(tags.filters.as_str())))?;
    Ok(())
}

fn write_custom_filters<W: std::io::Write>(
    writer: &mut Writer<W>,
    col: &formula_model::autofilter::FilterColumn,
    tags: &AutoFilterTags,
) -> Result<(), quick_xml::Error> {
    let criteria = effective_criteria(col);
    let mut entries: Vec<(String, Option<String>)> = Vec::new();
    let mut requires_and = col.join == FilterJoin::All;

    for criterion in &criteria {
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

    let mut custom = BytesStart::new(tags.custom_filters.as_str());
    if requires_and {
        custom.push_attribute(("and", "1"));
    }
    writer.write_event(Event::Start(custom))?;

    for (op, val) in entries {
        let mut cf = BytesStart::new(tags.custom_filter.as_str());
        cf.push_attribute(("operator", op.as_str()));
        if let Some(val) = val {
            cf.push_attribute(("val", val.as_str()));
        }
        writer.write_event(Event::Empty(cf))?;
    }

    writer.write_event(Event::End(BytesEnd::new(tags.custom_filters.as_str())))?;
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
        FilterCriterion::OpaqueCustom(c) => vec![(c.operator.clone(), c.value.clone())],
        FilterCriterion::OpaqueDynamic(_) => Vec::new(),
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

fn effective_criteria(col: &formula_model::autofilter::FilterColumn) -> Vec<FilterCriterion> {
    if !col.criteria.is_empty() {
        return col.criteria.clone();
    }
    col.values
        .iter()
        .map(|v| FilterCriterion::Equals(FilterValue::Text(v.clone())))
        .collect()
}

fn write_sort_state<W: std::io::Write>(
    writer: &mut Writer<W>,
    sort_state: &formula_model::autofilter::SortState,
    tags: &AutoFilterTags,
) -> Result<(), quick_xml::Error> {
    let sort = BytesStart::new(tags.sort_state.as_str());
    writer.write_event(Event::Start(sort))?;
    for cond in &sort_state.conditions {
        let mut sc = BytesStart::new(tags.sort_condition.as_str());
        sc.push_attribute(("ref", cond.range.to_string().as_str()));
        if cond.descending {
            sc.push_attribute(("descending", "1"));
        }
        writer.write_event(Event::Empty(sc))?;
    }
    writer.write_event(Event::End(BytesEnd::new(tags.sort_state.as_str())))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;
    use formula_model::autofilter::FilterColumn;
    use formula_model::{CellRef, Range};

    #[test]
    fn write_and_parse_roundtrip_filters() {
        let filter = SheetAutoFilter {
            range: Range::new(CellRef::new(0, 0), CellRef::new(2, 0)),
            filter_columns: vec![FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![
                    FilterCriterion::Blanks,
                    FilterCriterion::Equals(FilterValue::Text("Alice".into())),
                ],
                values: Vec::new(),
                raw_xml: Vec::new(),
            }],
            sort_state: None,
            raw_xml: Vec::new(),
        };

        let xml = write_autofilter(&filter).unwrap();
        let parsed = crate::autofilter::parse_autofilter(&xml).unwrap();
        assert_eq!(parsed.range, filter.range);
        assert_eq!(parsed.filter_columns[0].criteria.len(), 2);
    }

    #[test]
    fn write_dynamic_filter_today_roundtrip() {
        let filter = SheetAutoFilter {
            range: Range::new(CellRef::new(0, 0), CellRef::new(10, 0)),
            filter_columns: vec![FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Date(DateComparison::Today)],
                values: Vec::new(),
                raw_xml: Vec::new(),
            }],
            sort_state: None,
            raw_xml: Vec::new(),
        };

        let xml = write_autofilter(&filter).unwrap();
        assert!(xml.contains("dynamicFilter"));
        let parsed = crate::autofilter::parse_autofilter(&xml).unwrap();
        assert_eq!(
            parsed.filter_columns[0].criteria,
            vec![FilterCriterion::Date(DateComparison::Today)]
        );
    }

    #[test]
    fn write_between_expands_to_two_custom_filters() {
        let filter = SheetAutoFilter {
            range: Range::new(CellRef::new(0, 0), CellRef::new(10, 0)),
            filter_columns: vec![FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Number(NumberComparison::Between {
                    min: 2.0,
                    max: 5.0,
                })],
                values: Vec::new(),
                raw_xml: Vec::new(),
            }],
            sort_state: None,
            raw_xml: Vec::new(),
        };

        let xml = write_autofilter(&filter).unwrap();
        assert!(xml.contains("and=\"1\""));
        let parsed = crate::autofilter::parse_autofilter(&xml).unwrap();
        let col = &parsed.filter_columns[0];
        assert_eq!(col.join, FilterJoin::All);
        assert_eq!(
            col.criteria,
            vec![
                FilterCriterion::Number(NumberComparison::GreaterThanOrEqual(2.0)),
                FilterCriterion::Number(NumberComparison::LessThanOrEqual(5.0))
            ]
        );
    }

    #[test]
    fn roundtrip_preserves_unknown_xml_payloads() {
        let xml = r#"<autoFilter ref="A1:A3"><filterColumn colId="0"><filters><filter val="Alice"/></filters><colorFilter dxfId="3"/></filterColumn><extLst><ext uri="x"/></extLst></autoFilter>"#;
        let model = crate::autofilter::parse_autofilter(xml).unwrap();
        assert_eq!(model.filter_columns.len(), 1);
        assert_eq!(model.filter_columns[0].raw_xml.len(), 1);
        assert_eq!(model.raw_xml.len(), 1);

        let out = write_autofilter(&model).unwrap();
        assert!(out.contains("colorFilter"));
        assert!(out.contains("extLst"));

        let reparsed = crate::autofilter::parse_autofilter(&out).unwrap();
        assert_eq!(reparsed, model);
    }

    #[test]
    fn sort_state_roundtrips() {
        let filter = SheetAutoFilter {
            range: Range::new(CellRef::new(0, 0), CellRef::new(10, 1)),
            filter_columns: Vec::new(),
            sort_state: Some(formula_model::autofilter::SortState {
                conditions: vec![formula_model::autofilter::SortCondition {
                    range: Range::new(CellRef::new(0, 1), CellRef::new(10, 1)),
                    descending: true,
                }],
            }),
            raw_xml: Vec::new(),
        };

        let xml = write_autofilter(&filter).unwrap();
        assert!(xml.contains("sortState"));
        let parsed = crate::autofilter::parse_autofilter(&xml).unwrap();
        assert_eq!(parsed, filter);
    }
}
