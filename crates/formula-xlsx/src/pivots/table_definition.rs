use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml::local_name;
use crate::XlsxError;

/// Metadata extracted from an `xl/pivotTables/pivotTable*.xml` part.
///
/// This struct intentionally captures only the subset of the pivot table definition that we need
/// for sheet rendering and recomputation, while leaving the original XML untouched for round-trip
/// fidelity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTableDefinition {
    /// OPC part path, e.g. `xl/pivotTables/pivotTable1.xml`.
    pub path: String,
    pub name: Option<String>,
    pub cache_id: Option<u32>,
    /// Output range on the destination worksheet (A1-style range).
    pub location_ref: Option<String>,
    pub first_header_row: Option<u32>,
    pub first_data_row: Option<u32>,
    pub first_data_col: Option<u32>,
    /// `pivotTableDefinition@dataOnRows` (defaults to `false`).
    pub data_on_rows: bool,
    /// `pivotTableDefinition@rowGrandTotals` (defaults to `true`).
    pub row_grand_totals: bool,
    /// `pivotTableDefinition@colGrandTotals` (defaults to `true`).
    pub col_grand_totals: bool,
    /// `pivotTableDefinition@outline` (if present).
    pub outline: Option<bool>,
    /// `pivotTableDefinition@compact` (if present).
    pub compact: Option<bool>,
    /// `pivotTableDefinition@compactData` (if present).
    pub compact_data: Option<bool>,
}

impl PivotTableDefinition {
    pub fn parse(path: &str, xml: &[u8]) -> Result<Self, XlsxError> {
        let mut reader = Reader::from_reader(Cursor::new(xml));
        reader.config_mut().trim_text(true);

        let mut def = PivotTableDefinition {
            path: path.to_string(),
            name: None,
            cache_id: None,
            location_ref: None,
            first_header_row: None,
            first_data_row: None,
            first_data_col: None,
            data_on_rows: false,
            row_grand_totals: true,
            col_grand_totals: true,
            outline: None,
            compact: None,
            compact_data: None,
        };

        let mut buf = Vec::new();
        let mut parsed_root = false;

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(start) => {
                    parse_start_element(&mut def, &start, &mut parsed_root)?;
                }
                Event::Empty(start) => {
                    parse_start_element(&mut def, &start, &mut parsed_root)?;
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(def)
    }
}

fn parse_start_element(
    def: &mut PivotTableDefinition,
    start: &quick_xml::events::BytesStart<'_>,
    parsed_root: &mut bool,
) -> Result<(), XlsxError> {
    let name = start.name();
    let tag = local_name(name.as_ref());
    if !*parsed_root && tag.eq_ignore_ascii_case(b"pivotTableDefinition") {
        *parsed_root = true;
        for attr in start.attributes().with_checks(false) {
            let attr = attr?;
            let key = local_name(attr.key.as_ref());
            let value = attr.unescape_value()?.into_owned();

            if key.eq_ignore_ascii_case(b"name") {
                def.name = Some(value);
            } else if key.eq_ignore_ascii_case(b"cacheId") {
                def.cache_id = value.parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"dataOnRows") {
                if let Some(v) = parse_bool(&value) {
                    def.data_on_rows = v;
                }
            } else if key.eq_ignore_ascii_case(b"rowGrandTotals") {
                if let Some(v) = parse_bool(&value) {
                    def.row_grand_totals = v;
                }
            } else if key.eq_ignore_ascii_case(b"colGrandTotals") {
                if let Some(v) = parse_bool(&value) {
                    def.col_grand_totals = v;
                }
            } else if key.eq_ignore_ascii_case(b"outline") {
                def.outline = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"compact") {
                def.compact = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"compactData") {
                def.compact_data = parse_bool(&value);
            }
        }
        return Ok(());
    }

    if tag.eq_ignore_ascii_case(b"location") {
        for attr in start.attributes().with_checks(false) {
            let attr = attr?;
            let key = local_name(attr.key.as_ref());
            let value = attr.unescape_value()?.into_owned();

            if key.eq_ignore_ascii_case(b"ref") {
                def.location_ref = Some(value);
            } else if key.eq_ignore_ascii_case(b"firstHeaderRow") {
                def.first_header_row = value.parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"firstDataRow") {
                def.first_data_row = value.parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"firstDataCol") {
                def.first_data_col = value.parse::<u32>().ok();
            }
        }
    }

    Ok(())
}

fn parse_bool(value: &str) -> Option<bool> {
    match value.trim() {
        "1" | "true" | "TRUE" | "True" => Some(true),
        "0" | "false" | "FALSE" | "False" => Some(false),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use pretty_assertions::assert_eq;

    #[test]
    fn parses_location_and_layout_flags() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:pivotTableDefinition xmlns:p="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1"
  cacheId="7"
  dataOnRows="1"
  rowGrandTotals="0"
  colGrandTotals="1"
  outline="1"
  compact="0"
  compactData="1">
  <p:location ref="B3:F20" firstHeaderRow="2" firstDataRow="3" firstDataCol="2"/>
</p:pivotTableDefinition>"#;

        let parsed = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml)
            .expect("parse pivotTableDefinition");

        assert_eq!(parsed.name.as_deref(), Some("PivotTable1"));
        assert_eq!(parsed.cache_id, Some(7));
        assert_eq!(parsed.location_ref.as_deref(), Some("B3:F20"));
        assert_eq!(parsed.first_header_row, Some(2));
        assert_eq!(parsed.first_data_row, Some(3));
        assert_eq!(parsed.first_data_col, Some(2));
        assert_eq!(parsed.data_on_rows, true);
        assert_eq!(parsed.row_grand_totals, false);
        assert_eq!(parsed.col_grand_totals, true);
        assert_eq!(parsed.outline, Some(true));
        assert_eq!(parsed.compact, Some(false));
        assert_eq!(parsed.compact_data, Some(true));
    }
}
