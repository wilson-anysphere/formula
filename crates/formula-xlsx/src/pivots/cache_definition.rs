use std::collections::BTreeSet;
use std::io::Cursor;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::{XlsxError, XlsxPackage};

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PivotCacheDefinition {
    pub record_count: Option<u64>,
    pub refresh_on_load: Option<bool>,
    pub created_version: Option<u32>,
    pub refreshed_version: Option<u32>,
    pub cache_source_type: PivotCacheSourceType,
    pub worksheet_source_sheet: Option<String>,
    pub worksheet_source_ref: Option<String>,
    pub cache_fields: Vec<PivotCacheField>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PivotCacheSourceType {
    Worksheet,
    External,
    Consolidation,
    Scenario,
    Unknown(String),
}

impl Default for PivotCacheSourceType {
    fn default() -> Self {
        Self::Unknown(String::new())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PivotCacheField {
    pub name: String,
    pub num_fmt_id: Option<u32>,
    pub database_field: Option<bool>,
    pub formula: Option<String>,
    pub sql_type: Option<i32>,
    pub hierarchy: Option<u32>,
}

impl XlsxPackage {
    /// Parse every pivot cache definition part in the package.
    ///
    /// Returns a sorted list of `(part_name, parsed_definition)` pairs.
    pub fn pivot_cache_definitions(&self) -> Result<Vec<(String, PivotCacheDefinition)>, XlsxError> {
        let mut paths: BTreeSet<String> = BTreeSet::new();
        for name in self.part_names() {
            if name.starts_with("xl/pivotCache/")
                && name.contains("pivotCacheDefinition")
                && name.ends_with(".xml")
            {
                paths.insert(name.to_string());
            }
        }

        let mut out = Vec::new();
        for path in paths {
            let Some(bytes) = self.part(&path) else {
                continue;
            };
            out.push((path, parse_pivot_cache_definition(bytes)?));
        }
        Ok(out)
    }

    /// Parse a single pivot cache definition part.
    pub fn pivot_cache_definition(
        &self,
        part_name: &str,
    ) -> Result<Option<PivotCacheDefinition>, XlsxError> {
        let Some(bytes) = self.part(part_name) else {
            return Ok(None);
        };
        Ok(Some(parse_pivot_cache_definition(bytes)?))
    }
}

fn parse_pivot_cache_definition(xml: &[u8]) -> Result<PivotCacheDefinition, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut def = PivotCacheDefinition::default();

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) | Event::Empty(e) => {
                handle_element(&mut def, &e)?;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(def)
}

fn handle_element(def: &mut PivotCacheDefinition, e: &BytesStart<'_>) -> Result<(), XlsxError> {
    match e.local_name().as_ref() {
        b"pivotCacheDefinition" => {
            for attr in e.attributes().with_checks(false) {
                let attr = attr.map_err(quick_xml::Error::from)?;
                let value = attr.unescape_value()?;
                match attr.key.local_name().as_ref() {
                    b"recordCount" => def.record_count = value.parse::<u64>().ok(),
                    b"refreshOnLoad" => def.refresh_on_load = parse_bool(&value),
                    b"createdVersion" => def.created_version = value.parse::<u32>().ok(),
                    b"refreshedVersion" => def.refreshed_version = value.parse::<u32>().ok(),
                    _ => {}
                }
            }
        }
        b"cacheSource" => {
            for attr in e.attributes().with_checks(false) {
                let attr = attr.map_err(quick_xml::Error::from)?;
                if attr.key.local_name().as_ref() != b"type" {
                    continue;
                }
                let raw_value = attr.unescape_value()?.to_string();
                let value = raw_value.to_ascii_lowercase();
                def.cache_source_type = match value.as_str() {
                    "worksheet" => PivotCacheSourceType::Worksheet,
                    "external" => PivotCacheSourceType::External,
                    "consolidation" => PivotCacheSourceType::Consolidation,
                    "scenario" => PivotCacheSourceType::Scenario,
                    _ => PivotCacheSourceType::Unknown(raw_value),
                };
                break;
            }
        }
        b"worksheetSource" => {
            let mut sheet: Option<String> = None;
            let mut reference: Option<String> = None;
            let mut name: Option<String> = None;
            for attr in e.attributes().with_checks(false) {
                let attr = attr.map_err(quick_xml::Error::from)?;
                let value = attr.unescape_value()?.to_string();
                match attr.key.local_name().as_ref() {
                    b"sheet" => sheet = Some(value),
                    b"ref" => reference = Some(value),
                    b"name" => name = Some(value),
                    _ => {}
                }
            }
            def.worksheet_source_sheet = sheet;
            def.worksheet_source_ref = reference.or(name);
        }
        b"cacheField" => {
            let mut field = PivotCacheField::default();
            for attr in e.attributes().with_checks(false) {
                let attr = attr.map_err(quick_xml::Error::from)?;
                let value = attr.unescape_value()?;
                match attr.key.local_name().as_ref() {
                    b"name" => field.name = value.to_string(),
                    b"numFmtId" => field.num_fmt_id = value.parse::<u32>().ok(),
                    b"databaseField" => field.database_field = parse_bool(&value),
                    b"formula" => field.formula = Some(value.to_string()),
                    b"sqlType" => field.sql_type = value.parse::<i32>().ok(),
                    b"hierarchy" => field.hierarchy = value.parse::<u32>().ok(),
                    _ => {}
                }
            }
            def.cache_fields.push(field);
        }
        _ => {}
    }
    Ok(())
}

fn parse_bool(value: &str) -> Option<bool> {
    match value {
        "1" => Some(true),
        "0" => Some(false),
        _ if value.eq_ignore_ascii_case("true") => Some(true),
        _ if value.eq_ignore_ascii_case("false") => Some(false),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use pretty_assertions::assert_eq;

    #[test]
    fn parses_named_source_when_ref_missing() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet">
    <worksheetSource name="MyNamedRange"/>
  </cacheSource>
</pivotCacheDefinition>"#;

        let def = parse_pivot_cache_definition(xml).expect("parse");
        assert_eq!(def.cache_source_type, PivotCacheSourceType::Worksheet);
        assert_eq!(def.worksheet_source_sheet, None);
        assert_eq!(def.worksheet_source_ref.as_deref(), Some("MyNamedRange"));
        assert!(def.cache_fields.is_empty());
    }

    #[test]
    fn handles_missing_cache_fields() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet">
    <worksheetSource sheet="Sheet1" ref="A1:B2"/>
  </cacheSource>
</pivotCacheDefinition>"#;

        let def = parse_pivot_cache_definition(xml).expect("parse");
        assert_eq!(def.cache_source_type, PivotCacheSourceType::Worksheet);
        assert_eq!(def.worksheet_source_sheet.as_deref(), Some("Sheet1"));
        assert_eq!(def.worksheet_source_ref.as_deref(), Some("A1:B2"));
        assert!(def.cache_fields.is_empty());
    }

    #[test]
    fn parses_cache_source_type_case_insensitively() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="Worksheet"/>
</pivotCacheDefinition>"#;

        let def = parse_pivot_cache_definition(xml).expect("parse");
        assert_eq!(def.cache_source_type, PivotCacheSourceType::Worksheet);
    }

    #[test]
    fn preserves_unknown_cache_source_type() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="WeIrD"/>
</pivotCacheDefinition>"#;

        let def = parse_pivot_cache_definition(xml).expect("parse");
        assert_eq!(
            def.cache_source_type,
            PivotCacheSourceType::Unknown("WeIrD".to_string())
        );
    }

    #[test]
    fn tolerates_namespaced_elements_and_unknown_tags() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:pivotCacheDefinition xmlns:p="http://schemas.openxmlformats.org/spreadsheetml/2006/main" p:recordCount="4">
  <p:cacheSource p:type="worksheet">
    <p:worksheetSource p:sheet="Sheet1" p:ref="A1:B2"/>
  </p:cacheSource>
  <p:cacheFields p:count="1">
    <p:cacheField p:name="Field1" p:numFmtId="0"/>
  </p:cacheFields>
  <p:unknownTag foo="bar"/>
</p:pivotCacheDefinition>"#;

        let def = parse_pivot_cache_definition(xml).expect("parse");
        assert_eq!(def.record_count, Some(4));
        assert_eq!(def.cache_source_type, PivotCacheSourceType::Worksheet);
        assert_eq!(def.worksheet_source_sheet.as_deref(), Some("Sheet1"));
        assert_eq!(def.worksheet_source_ref.as_deref(), Some("A1:B2"));
        assert_eq!(def.cache_fields.len(), 1);
        assert_eq!(def.cache_fields[0].name, "Field1");
    }
}
