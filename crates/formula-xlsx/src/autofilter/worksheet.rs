use crate::autofilter::parse::{parse_autofilter, AutoFilterParseError};
use crate::autofilter::write::write_autofilter_to;
use crate::XlsxError;
use formula_model::SheetAutoFilter;
use quick_xml::events::Event;
use quick_xml::{Reader, Writer};

pub fn parse_worksheet_autofilter(
    xml: &str,
) -> Result<Option<SheetAutoFilter>, AutoFilterParseError> {
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
    filter: Option<&SheetAutoFilter>,
) -> Result<String, XlsxError> {
    let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(worksheet_xml)?;
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
                    write_autofilter_to(&mut writer, filter, worksheet_prefix.as_deref())?;
                    wrote_autofilter = true;
                }
                skip_depth = 1;
            }
            Event::Empty(ref e) if skip_depth == 0 && e.local_name().as_ref() == b"autoFilter" => {
                if let Some(filter) = filter {
                    write_autofilter_to(&mut writer, filter, worksheet_prefix.as_deref())?;
                    wrote_autofilter = true;
                }
            }
            Event::Start(_) if skip_depth > 0 => {
                skip_depth += 1;
            }
            Event::End(ref e) if skip_depth > 0 => {
                skip_depth = skip_depth.saturating_sub(1);
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if skip_depth == 0
                    && !wrote_autofilter
                    && matches!(
                        e.local_name().as_ref(),
                        b"mergeCells" | b"tableParts" | b"extLst"
                    ) =>
            {
                if let Some(filter) = filter {
                    write_autofilter_to(&mut writer, filter, worksheet_prefix.as_deref())?;
                    wrote_autofilter = true;
                }
                writer.write_event(event.to_owned())?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"worksheet" => {
                if !wrote_autofilter {
                    if let Some(filter) = filter {
                        write_autofilter_to(&mut writer, filter, worksheet_prefix.as_deref())?;
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
    use formula_model::autofilter::{FilterColumn, FilterCriterion, FilterJoin, FilterValue};
    use formula_model::{CellRef, Range};
    use pretty_assertions::assert_eq;

    #[test]
    fn inserts_autofilter_when_missing() {
        let worksheet_xml = r#"<worksheet><sheetData/></worksheet>"#;
        let filter = SheetAutoFilter {
            range: Range::new(CellRef::new(0, 0), CellRef::new(2, 0)),
            filter_columns: vec![FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Equals(FilterValue::Text("Alice".into()))],
                values: Vec::new(),
                raw_xml: Vec::new(),
            }],
            sort_state: None,
            raw_xml: Vec::new(),
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
        let filter = SheetAutoFilter {
            range: Range::new(CellRef::new(0, 0), CellRef::new(2, 0)),
            filter_columns: vec![FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Equals(FilterValue::Text("Alice".into()))],
                values: Vec::new(),
                raw_xml: Vec::new(),
            }],
            sort_state: None,
            raw_xml: Vec::new(),
        };

        let written = write_worksheet_autofilter(worksheet_xml, Some(&filter)).unwrap();
        let parsed = parse_worksheet_autofilter(&written).unwrap().unwrap();
        assert_eq!(
            parsed.filter_columns[0].criteria,
            vec![FilterCriterion::Equals(FilterValue::Text("Alice".into()))]
        );
    }

    #[test]
    fn inserts_autofilter_before_table_parts_when_missing() {
        let worksheet_xml = r#"<worksheet><sheetData/><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>"#;
        let filter = SheetAutoFilter {
            range: Range::new(CellRef::new(0, 0), CellRef::new(2, 0)),
            filter_columns: vec![FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Equals(FilterValue::Text("Alice".into()))],
                values: Vec::new(),
                raw_xml: Vec::new(),
            }],
            sort_state: None,
            raw_xml: Vec::new(),
        };

        let written = write_worksheet_autofilter(worksheet_xml, Some(&filter)).unwrap();
        let auto_pos = written.find("<autoFilter").expect("autofilter inserted");
        let table_pos = written.find("<tableParts").expect("tableParts exists");
        assert!(auto_pos < table_pos, "expected autofilter before tableParts");
    }
}
