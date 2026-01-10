use crate::autofilter::parse::{parse_autofilter, AutoFilterParseError};
use crate::autofilter::write::write_autofilter_to;
use crate::XlsxError;
use formula_engine::sort_filter::AutoFilter;
use quick_xml::events::Event;
use quick_xml::{Reader, Writer};

pub fn parse_worksheet_autofilter(xml: &str) -> Result<Option<AutoFilter>, AutoFilterParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut has_autofilter = false;
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"autoFilter" => {
                has_autofilter = true;
                break;
            }
            _ => {}
        }
        buf.clear();
    }

    if !has_autofilter {
        return Ok(None);
    }

    Ok(Some(parse_autofilter(xml)?))
}

pub fn write_worksheet_autofilter(
    worksheet_xml: &str,
    filter: Option<&AutoFilter>,
) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(worksheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut wrote_autofilter = false;
    let mut skip_depth: usize = 0;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if skip_depth == 0 && e.local_name().as_ref() == b"autoFilter" => {
                if let Some(filter) = filter {
                    write_autofilter_to(&mut writer, filter)?;
                    wrote_autofilter = true;
                }
                skip_depth = 1;
            }
            Event::Empty(ref e) if skip_depth == 0 && e.local_name().as_ref() == b"autoFilter" => {
                if let Some(filter) = filter {
                    write_autofilter_to(&mut writer, filter)?;
                    wrote_autofilter = true;
                }
            }
            Event::Start(_) if skip_depth > 0 => {
                skip_depth += 1;
            }
            Event::End(ref e) if skip_depth > 0 => {
                skip_depth = skip_depth.saturating_sub(1);
            }
            Event::End(ref e) if e.local_name().as_ref() == b"worksheet" => {
                if !wrote_autofilter {
                    if let Some(filter) = filter {
                        write_autofilter_to(&mut writer, filter)?;
                        wrote_autofilter = true;
                    }
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            _ => {
                if skip_depth == 0 {
                    writer.write_event(event.to_owned())?;
                }
            }
        }
        buf.clear();
    }

    Ok(String::from_utf8(writer.into_inner())?)
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_engine::sort_filter::{ColumnFilter, FilterCriterion, FilterJoin, FilterValue, RangeRef};
    use pretty_assertions::assert_eq;
    use std::collections::BTreeMap;

    #[test]
    fn inserts_autofilter_when_missing() {
        let worksheet_xml = r#"<worksheet><sheetData/></worksheet>"#;
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
                    criteria: vec![FilterCriterion::Equals(FilterValue::Text("Alice".into()))],
                },
            )]),
        };

        let written = write_worksheet_autofilter(worksheet_xml, Some(&filter)).unwrap();
        assert!(written.contains("autoFilter"));
        let parsed = parse_worksheet_autofilter(&written).unwrap().unwrap();
        assert_eq!(parsed.range, filter.range);
    }

    #[test]
    fn removes_autofilter_when_cleared() {
        let worksheet_xml =
            r#"<worksheet><sheetData/><autoFilter ref="A1:A3"/></worksheet>"#;
        let written = write_worksheet_autofilter(worksheet_xml, None).unwrap();
        assert!(!written.contains("autoFilter"));
        assert_eq!(parse_worksheet_autofilter(&written).unwrap(), None);
    }

    #[test]
    fn replaces_existing_autofilter() {
        let worksheet_xml =
            r#"<worksheet><sheetData/><autoFilter ref="A1:A3"><filterColumn colId="0"><filters><filter val="Bob"/></filters></filterColumn></autoFilter></worksheet>"#;
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
                    criteria: vec![FilterCriterion::Equals(FilterValue::Text("Alice".into()))],
                },
            )]),
        };

        let written = write_worksheet_autofilter(worksheet_xml, Some(&filter)).unwrap();
        let parsed = parse_worksheet_autofilter(&written).unwrap().unwrap();
        assert_eq!(
            parsed.columns.get(&0).unwrap().criteria,
            vec![FilterCriterion::Equals(FilterValue::Text("Alice".into()))]
        );
    }
}

