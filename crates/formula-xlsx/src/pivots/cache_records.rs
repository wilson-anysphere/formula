use chrono::NaiveDate;
use std::collections::BTreeMap;

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::Reader;

use formula_engine::date::{serial_to_ymd, ExcelDateSystem};

use crate::{XlsxDocument, XlsxError};

/// A typed value found in a `<r>` record inside `pivotCacheRecords*.xml`.
#[derive(Debug, Clone, PartialEq)]
pub enum PivotCacheValue {
    /// `<m/>`
    Missing,
    /// `<n v="..."/>`
    Number(f64),
    /// `<s v="..."/>`
    String(String),
    /// `<b v="0|1"/>`
    Bool(bool),
    /// `<e v="..."/>`
    Error(String),
    /// `<d v="..."/>` (often ISO-8601 / RFC3339).
    ///
    /// Excel stores real date/time values as strings in pivot caches, so we keep
    /// the raw attribute value (instead of coercing to an Excel serial number).
    DateTime(String),
    /// `<x v="..."/>` (shared item index).
    ///
    /// Excel can store record values as indices into a per-field `<sharedItems>` table in the
    /// pivot cache definition. Use [`crate::pivots::PivotCacheDefinition::resolve_record_value`]
    /// (or [`crate::pivots::PivotCacheDefinition::resolve_record_values`] for whole records) to
    /// turn this into the corresponding typed value.
    Index(u32),
}

impl XlsxDocument {
    /// Create a streaming reader for a pivot cache records part
    /// (e.g. `xl/pivotCache/pivotCacheRecords1.xml`).
    pub fn pivot_cache_records<'a>(
        &'a self,
        part_name: &str,
    ) -> Result<PivotCacheRecordsReader<'a>, XlsxError> {
        let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
        let bytes = self
            .parts()
            .get(part_name)
            .ok_or_else(|| XlsxError::MissingPart(part_name.to_string()))?;
        Ok(PivotCacheRecordsReader::new(bytes))
    }

    /// Parse all `pivotCacheRecords*.xml` parts in the document into memory.
    ///
    /// Prefer [`Self::pivot_cache_records`] for large caches.
    pub fn pivot_cache_records_all(&self) -> BTreeMap<String, Vec<Vec<PivotCacheValue>>> {
        let mut out = BTreeMap::new();
        for (name, bytes) in self.parts() {
            if name.starts_with("xl/pivotCache/")
                && name.contains("pivotCacheRecords")
                && name.ends_with(".xml")
            {
                let mut reader = PivotCacheRecordsReader::new(bytes);
                out.insert(name.to_string(), reader.parse_all_records());
            }
        }
        out
    }
}

/// Streaming reader for `xl/pivotCache/pivotCacheRecords*.xml`.
///
/// This parser is namespace-insensitive (uses local names) and is designed to
/// avoid loading the entire DOM into memory.
pub struct PivotCacheRecordsReader<'a> {
    reader: Reader<&'a [u8]>,
    buf: Vec<u8>,
    skip_buf: Vec<u8>,
    done: bool,
}

impl<'a> PivotCacheRecordsReader<'a> {
    pub fn new(bytes: &'a [u8]) -> Self {
        let mut reader = Reader::from_reader(bytes);
        reader.config_mut().trim_text(false);
        Self {
            reader,
            buf: Vec::new(),
            skip_buf: Vec::new(),
            done: false,
        }
    }

    /// Return the next `<r>` record, if present.
    ///
    /// The record is returned as a list of typed values in the order they appeared in XML.
    /// Unknown tags are ignored.
    ///
    /// Note: record values may include [`PivotCacheValue::Index`] entries, which need to be
    /// resolved against the corresponding cache definition's `<sharedItems>` table using
    /// [`crate::pivots::PivotCacheDefinition::resolve_record_value`] (or
    /// [`crate::pivots::PivotCacheDefinition::resolve_record_values`]).
    pub fn next_record(&mut self) -> Option<Vec<PivotCacheValue>> {
        if self.done {
            return None;
        }

        loop {
            let event = match self.reader.read_event_into(&mut self.buf) {
                Ok(ev) => ev,
                Err(_) => {
                    self.done = true;
                    return None;
                }
            };

            match event {
                Event::Start(e) if e.local_name().as_ref() == b"r" => {
                    drop(e);
                    self.buf.clear();
                    let record = self.read_record();
                    return Some(record);
                }
                Event::Empty(e) if e.local_name().as_ref() == b"r" => {
                    drop(e);
                    self.buf.clear();
                    return Some(Vec::new());
                }
                Event::Eof => {
                    self.done = true;
                    self.buf.clear();
                    return None;
                }
                _ => {}
            }

            self.buf.clear();
        }
    }

    /// Convenience helper for small fixtures: parse all records into memory.
    pub fn parse_all_records(&mut self) -> Vec<Vec<PivotCacheValue>> {
        let mut out = Vec::new();
        while let Some(record) = self.next_record() {
            out.push(record);
        }
        out
    }

    fn read_record(&mut self) -> Vec<PivotCacheValue> {
        let mut values = Vec::new();

        loop {
            let event = match self.reader.read_event_into(&mut self.buf) {
                Ok(ev) => ev,
                Err(_) => {
                    self.done = true;
                    break;
                }
            };

            match event {
                Event::Empty(e) => {
                    if let Some(value) = parse_value_empty(&e) {
                        values.push(value);
                    }
                }
                Event::Start(e) => {
                    let e = e.into_owned();
                    self.buf.clear();
                    if let Some(value) = self.parse_value_start(&e) {
                        values.push(value);
                    }
                }
                Event::End(e) if e.local_name().as_ref() == b"r" => break,
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }

            self.buf.clear();
        }

        values
    }

    fn parse_value_start(&mut self, e: &BytesStart<'_>) -> Option<PivotCacheValue> {
        let local_name = e.local_name();
        let local_name = local_name.as_ref();

        // Most records use self-closing tags (`Event::Empty`), but some producers
        // emit `<n><v>...</v></n>` instead of `<n v="..."/>`.
        let attr_v = attr_value_local(e, b"v");

        match local_name {
            b"m" => {
                self.skip_to_end(e.name());
                Some(PivotCacheValue::Missing)
            }
            b"n" => {
                let v = match attr_v {
                    Some(v) => {
                        self.skip_to_end(e.name());
                        Some(v)
                    }
                    None => self.read_value_text_from_element(e.name()),
                };
                Some(parse_number(v))
            }
            b"d" => {
                let v = match attr_v {
                    Some(v) => {
                        self.skip_to_end(e.name());
                        Some(v)
                    }
                    None => self.read_value_text_from_element(e.name()),
                };
                Some(parse_datetime(v))
            }
            b"x" => {
                let v = match attr_v {
                    Some(v) => {
                        self.skip_to_end(e.name());
                        Some(v)
                    }
                    None => self.read_value_text_from_element(e.name()),
                };
                Some(parse_index(v))
            }
            b"s" => {
                let v = match attr_v {
                    Some(v) => {
                        self.skip_to_end(e.name());
                        Some(v)
                    }
                    None => self.read_value_text_from_element(e.name()),
                };
                Some(parse_string(v))
            }
            b"e" => {
                let v = match attr_v {
                    Some(v) => {
                        self.skip_to_end(e.name());
                        Some(v)
                    }
                    None => self.read_value_text_from_element(e.name()),
                };
                Some(parse_error(v))
            }
            b"b" => {
                let v = match attr_v {
                    Some(v) => {
                        self.skip_to_end(e.name());
                        Some(v)
                    }
                    None => self.read_value_text_from_element(e.name()),
                };
                Some(parse_bool(v))
            }
            _ => {
                // Unknown tags should be ignored, but we must still advance the reader.
                self.skip_to_end(e.name());
                None
            }
        }
    }

    fn read_value_text_from_element(&mut self, outer_end: QName<'_>) -> Option<String> {
        let mut value: Option<String> = None;

        loop {
            let event = match self.reader.read_event_into(&mut self.buf) {
                Ok(ev) => ev,
                Err(_) => {
                    self.done = true;
                    break;
                }
            };

            match event {
                Event::Start(e) => {
                    let e = e.into_owned();
                    self.buf.clear();

                    if e.local_name().as_ref() == b"v" {
                        let v = self.read_text_to_end(e.name());
                        if value.is_none() {
                            value = Some(v);
                        }
                    } else {
                        // Skip unknown nested elements inside the value wrapper.
                        self.skip_to_end(e.name());
                    }
                }
                Event::Empty(e) if e.local_name().as_ref() == b"v" => {
                    value.get_or_insert_with(String::new);
                }
                Event::Text(e) if value.is_none() => {
                    if let Ok(text) = e.unescape() {
                        let text = text.into_owned();
                        if !text.trim().is_empty() {
                            value = Some(text);
                        }
                    }
                }
                Event::CData(e) if value.is_none() => {
                    if let Ok(text) = std::str::from_utf8(e.as_ref()) {
                        if !text.trim().is_empty() {
                            value = Some(text.to_string());
                        }
                    }
                }
                Event::End(e) if e.name() == outer_end => break,
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }

            self.buf.clear();
        }

        value
    }

    fn read_text_to_end(&mut self, end: QName<'_>) -> String {
        let mut text = String::new();

        loop {
            let event = match self.reader.read_event_into(&mut self.buf) {
                Ok(ev) => ev,
                Err(_) => {
                    self.done = true;
                    break;
                }
            };

            match event {
                Event::Text(e) => {
                    if let Ok(t) = e.unescape() {
                        text.push_str(&t);
                    }
                }
                Event::CData(e) => {
                    if let Ok(t) = std::str::from_utf8(e.as_ref()) {
                        text.push_str(t);
                    }
                }
                Event::Start(e) => {
                    let e = e.into_owned();
                    self.buf.clear();
                    // Unlikely, but keep the parser resilient.
                    self.skip_to_end(e.name());
                }
                Event::End(e) if e.name() == end => break,
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }

            self.buf.clear();
        }

        text
    }

    fn skip_to_end(&mut self, end: QName<'_>) {
        self.skip_buf.clear();
        let _ = self.reader.read_to_end_into(end, &mut self.skip_buf);
    }
}

fn parse_value_empty(e: &BytesStart<'_>) -> Option<PivotCacheValue> {
    let local_name = e.local_name();
    let local_name = local_name.as_ref();

    match local_name {
        b"m" => Some(PivotCacheValue::Missing),
        b"n" => Some(parse_number(attr_value_local(e, b"v"))),
        b"s" => Some(parse_string(attr_value_local(e, b"v"))),
        b"b" => Some(parse_bool(attr_value_local(e, b"v"))),
        b"e" => Some(parse_error(attr_value_local(e, b"v"))),
        b"d" => Some(parse_datetime(attr_value_local(e, b"v"))),
        b"x" => Some(parse_index(attr_value_local(e, b"v"))),
        _ => None,
    }
}

fn attr_value_local(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for attr in e.attributes().with_checks(false) {
        let Ok(attr) = attr else {
            continue;
        };
        if attr.key.local_name().as_ref() != key {
            continue;
        }
        if let Ok(v) = attr.unescape_value() {
            return Some(v.into_owned());
        }
    }
    None
}

fn parse_number(v: Option<String>) -> PivotCacheValue {
    let Some(v) = v else {
        return PivotCacheValue::Missing;
    };
    let Ok(n) = v.trim().parse::<f64>() else {
        return PivotCacheValue::Missing;
    };
    PivotCacheValue::Number(n)
}

fn parse_datetime(v: Option<String>) -> PivotCacheValue {
    let Some(v) = v else {
        return PivotCacheValue::Missing;
    };
    PivotCacheValue::DateTime(v)
}

fn parse_index(v: Option<String>) -> PivotCacheValue {
    let Some(v) = v else {
        return PivotCacheValue::Missing;
    };
    let Ok(idx) = v.trim().parse::<u32>() else {
        return PivotCacheValue::Missing;
    };
    PivotCacheValue::Index(idx)
}

fn parse_string(v: Option<String>) -> PivotCacheValue {
    let Some(v) = v else {
        return PivotCacheValue::Missing;
    };
    PivotCacheValue::String(v)
}

fn parse_error(v: Option<String>) -> PivotCacheValue {
    let Some(v) = v else {
        return PivotCacheValue::Missing;
    };
    PivotCacheValue::Error(v)
}

fn parse_bool(v: Option<String>) -> PivotCacheValue {
    let Some(v) = v else {
        return PivotCacheValue::Missing;
    };

    let v = v.trim();
    PivotCacheValue::Bool(v == "1" || v.eq_ignore_ascii_case("true"))
}

/// Best-effort conversion of a pivot cache `<d v="..."/>` value into a `NaiveDate`.
///
/// Pivot caches commonly store ISO-8601 / RFC3339 strings (e.g. `2024-01-15T00:00:00Z`).
/// Timelines typically operate on the date component.
pub fn pivot_cache_datetime_to_naive_date(v: &str) -> Option<NaiveDate> {
    let v = v.trim();
    if v.is_empty() {
        return None;
    }

    // Common case: RFC3339/ISO8601 strings such as `2024-01-15T00:00:00Z`.
    let date_part = v.split(['T', ' ']).next().unwrap_or(v);
    if date_part.len() >= 10 {
        let mut parts = date_part.split('-');
        if let (Some(year), Some(month), Some(day)) = (parts.next(), parts.next(), parts.next()) {
            if let (Ok(year), Ok(month), Ok(day)) =
                (year.parse::<i32>(), month.parse::<u32>(), day.parse::<u32>())
            {
                if let Some(date) = NaiveDate::from_ymd_opt(year, month, day) {
                    return Some(date);
                }
            }
        }
    }

    // Some producers emit `YYYYMMDD`-style compact dates.
    if date_part.len() == 8 && date_part.as_bytes().iter().all(|b| b.is_ascii_digit()) {
        let year = date_part[..4].parse::<i32>().ok()?;
        let month = date_part[4..6].parse::<u32>().ok()?;
        let day = date_part[6..8].parse::<u32>().ok()?;
        if let Some(date) = NaiveDate::from_ymd_opt(year, month, day) {
            return Some(date);
        }
    }

    // Fallback: interpret numeric `d` values as Excel serial dates. Pivot caches typically
    // store dates using the 1900 date system.
    if let Ok(serial) = date_part.parse::<f64>() {
        let serial = serial.trunc() as i32;
        if let Ok(excel_date) = serial_to_ymd(serial, ExcelDateSystem::EXCEL_1900) {
            return NaiveDate::from_ymd_opt(
                excel_date.year,
                excel_date.month as u32,
                excel_date.day as u32,
            );
        }
    }

    None
}

#[cfg(test)]
mod tests {
    use super::*;

    use pretty_assertions::assert_eq;

    #[test]
    fn parses_record_item_tag_variants() {
        let xml = r##"
            <pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <r>
                <m/>
                <b v="1"/>
                <b v="0"/>
                <e v="#DIV/0!"/>
                <n v="42"/>
                <n>42</n>
                <s v="Hello"/>
                <s>Hello</s>
                <x v="3"/>
              </r>
            </pivotCacheRecords>
        "##;

        let mut reader = PivotCacheRecordsReader::new(xml.as_bytes());
        let records = reader.parse_all_records();

        assert_eq!(
            records,
            vec![vec![
                PivotCacheValue::Missing,
                PivotCacheValue::Bool(true),
                PivotCacheValue::Bool(false),
                PivotCacheValue::Error("#DIV/0!".to_string()),
                PivotCacheValue::Number(42.0),
                PivotCacheValue::Number(42.0),
                PivotCacheValue::String("Hello".to_string()),
                PivotCacheValue::String("Hello".to_string()),
                PivotCacheValue::Index(3),
            ]]
        );
    }

    #[test]
    fn parses_namespace_prefixed_value_tags() {
        let xml = r#"
            <pc:pivotCacheRecords xmlns:pc="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <pc:r>
                <pc:n v="1"/>
                <pc:s>World</pc:s>
                <pc:m/>
              </pc:r>
            </pc:pivotCacheRecords>
        "#;

        let mut reader = PivotCacheRecordsReader::new(xml.as_bytes());
        let records = reader.parse_all_records();

        assert_eq!(
            records,
            vec![vec![
                PivotCacheValue::Number(1.0),
                PivotCacheValue::String("World".to_string()),
                PivotCacheValue::Missing,
            ]]
        );
    }

    #[test]
    fn parses_text_node_and_wrapped_value_forms() {
        let xml = r##"
            <pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <r>
                <m></m>
                <b>1</b>
                <b>false</b>
                <e>#DIV/0!</e>
                <n><v>42</v></n>
                <s><v>Hello</v></s>
                <x><v>3</v></x>
              </r>
            </pivotCacheRecords>
        "##;

        let mut reader = PivotCacheRecordsReader::new(xml.as_bytes());
        let records = reader.parse_all_records();

        assert_eq!(
            records,
            vec![vec![
                PivotCacheValue::Missing,
                PivotCacheValue::Bool(true),
                PivotCacheValue::Bool(false),
                PivotCacheValue::Error("#DIV/0!".to_string()),
                PivotCacheValue::Number(42.0),
                PivotCacheValue::String("Hello".to_string()),
                PivotCacheValue::Index(3),
            ]]
        );
    }
}
