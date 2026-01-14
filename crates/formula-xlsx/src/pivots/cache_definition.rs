use std::collections::BTreeSet;
use std::io::Cursor;

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::Reader;

use chrono::NaiveDate;
use formula_engine::date::{serial_to_ymd, ExcelDateSystem};
use formula_model::pivots::ScalarValue;

use super::cache_records::pivot_cache_datetime_to_naive_date;
use super::PivotCacheValue;
use crate::openxml::resolve_relationship_target;
use crate::{XlsxDocument, XlsxError, XlsxPackage};
#[derive(Debug, Clone, PartialEq, Default)]
pub struct PivotCacheDefinition {
    pub record_count: Option<u64>,
    pub refresh_on_load: Option<bool>,
    pub created_version: Option<u32>,
    pub refreshed_version: Option<u32>,
    pub cache_source_type: PivotCacheSourceType,
    pub cache_source_connection_id: Option<u32>,
    pub worksheet_source_sheet: Option<String>,
    pub worksheet_source_ref: Option<String>,
    pub cache_fields: Vec<PivotCacheField>,
}

impl PivotCacheDefinition {
    pub fn calculated_fields(&self) -> Vec<formula_model::pivots::CalculatedField> {
        self.cache_fields
            .iter()
            .filter_map(|field| {
                let formula = field.formula.as_deref()?;
                let formula = formula.strip_prefix('=').unwrap_or(formula);
                Some(formula_model::pivots::CalculatedField {
                    name: field.name.clone(),
                    formula: formula.to_string(),
                })
            })
            .collect()
    }
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

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PivotCacheField {
    pub name: String,
    pub caption: Option<String>,
    pub property_name: Option<String>,
    pub num_fmt_id: Option<u32>,
    pub database_field: Option<bool>,
    pub server_field: Option<bool>,
    pub unique_list: Option<bool>,
    pub formula: Option<String>,
    pub sql_type: Option<i32>,
    pub hierarchy: Option<u32>,
    pub level: Option<u32>,
    pub mapping_count: Option<u32>,
    /// Shared item table for this cache field, as found in the cache definition's
    /// `<sharedItems>` element.
    ///
    /// When present, pivot cache records can encode values as `<x v="..."/>`
    /// indices into this list.
    pub shared_items: Option<Vec<PivotCacheValue>>,
}

impl PivotCacheDefinition {
    /// Resolve a pivot cache record value against this cache definition.
    ///
    /// Excel can store record values as an `<x v="..."/>` shared item index instead of an inline
    /// typed value (`<s>`, `<n>`, etc.). In that case this helper looks up the corresponding item
    /// from `cache_fields[field_idx].shared_items` and returns a clone of it.
    ///
    /// Resolution rules:
    /// - If `value` is not [`PivotCacheValue::Index`], it is returned unchanged (no allocation).
    /// - If the index is out of range, the cache field is missing, or the cache field has no
    ///   `shared_items`, this returns [`PivotCacheValue::Missing`].
    ///
    /// If you need to resolve an entire record at once (based on value position), see
    /// [`Self::resolve_record_values`].
    #[inline]
    pub fn resolve_record_value(
        &self,
        field_idx: usize,
        value: PivotCacheValue,
    ) -> PivotCacheValue {
        let PivotCacheValue::Index(shared_idx) = value else {
            return value;
        };

        let Some(field) = self.cache_fields.get(field_idx) else {
            return PivotCacheValue::Missing;
        };
        let Some(shared_items) = field.shared_items.as_ref() else {
            return PivotCacheValue::Missing;
        };
        let Ok(shared_idx) = usize::try_from(shared_idx) else {
            return PivotCacheValue::Missing;
        };
        shared_items
            .get(shared_idx)
            .cloned()
            .unwrap_or(PivotCacheValue::Missing)
    }

    /// Resolve a record's values against this cache definition.
    ///
    /// This is a convenience wrapper around [`Self::resolve_record_value`] that uses the value's
    /// position in the record as `field_idx` (0-based).
    ///
    /// This helper returns an iterator so callers can resolve values while streaming records
    /// without allocating a new record buffer.
    #[inline]
    pub fn resolve_record_values<'a, I>(
        &'a self,
        values: I,
    ) -> impl Iterator<Item = PivotCacheValue> + 'a
    where
        I: IntoIterator<Item = PivotCacheValue>,
        I::IntoIter: 'a,
    {
        values
            .into_iter()
            .enumerate()
            .map(|(field_idx, value)| self.resolve_record_value(field_idx, value))
    }

    /// Resolve a shared item index for a given cache field into a typed [`ScalarValue`].
    ///
    /// Excel encodes some pivot cache values (and slicer selections) as indices into a per-field
    /// `<sharedItems>` table. This helper turns those indices back into typed values so slicer and
    /// timeline selections match the pivot cache's key semantics.
    pub fn resolve_shared_item(
        &self,
        field_idx: usize,
        shared_item_index: u32,
    ) -> Option<ScalarValue> {
        let field = self.cache_fields.get(field_idx)?;
        let item = field
            .shared_items
            .as_ref()?
            .get(shared_item_index as usize)?;

        Some(match item {
            PivotCacheValue::Missing => ScalarValue::Blank,
            PivotCacheValue::String(s) => ScalarValue::Text(s.clone()),
            PivotCacheValue::Number(n) => ScalarValue::from(*n),
            PivotCacheValue::Bool(b) => ScalarValue::Bool(*b),
            PivotCacheValue::Error(_) => ScalarValue::Blank,
            PivotCacheValue::DateTime(s) => {
                let trimmed = s.trim();
                if trimmed.is_empty() {
                    ScalarValue::Blank
                } else if let Some(date) = pivot_cache_datetime_to_naive_date(trimmed) {
                    ScalarValue::Date(date)
                } else if let Ok(serial) = trimmed.parse::<f64>() {
                    let serial = serial.trunc() as i32;
                    if let Ok(excel_date) = serial_to_ymd(serial, ExcelDateSystem::EXCEL_1900) {
                        NaiveDate::from_ymd_opt(
                            excel_date.year,
                            excel_date.month as u32,
                            excel_date.day as u32,
                        )
                        .map(ScalarValue::Date)
                        .unwrap_or_else(|| ScalarValue::Text(s.clone()))
                    } else {
                        ScalarValue::Text(s.clone())
                    }
                } else {
                    ScalarValue::Text(s.clone())
                }
            }
            PivotCacheValue::Index(idx) => ScalarValue::Text(idx.to_string()),
        })
    }
}

impl XlsxPackage {
    /// Parse every pivot cache definition part in the package.
    ///
    /// Returns a sorted list of `(part_name, parsed_definition)` pairs.
    pub fn pivot_cache_definitions(
        &self,
    ) -> Result<Vec<(String, PivotCacheDefinition)>, XlsxError> {
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

    /// Resolve and parse the pivot cache definition for a given `cacheId`.
    ///
    /// Excel typically stores cache definitions as `xl/pivotCache/pivotCacheDefinitionN.xml`, but
    /// the `N` in the filename does not always match the `cacheId`. When present, the authoritative
    /// mapping is the workbook-level `<pivotCaches>` list and `xl/_rels/workbook.xml.rels`.
    pub fn pivot_cache_definition_for_cache_id(
        &self,
        cache_id: u32,
    ) -> Result<Option<(String, PivotCacheDefinition)>, XlsxError> {
        let workbook_xml = match self.part("xl/workbook.xml") {
            Some(bytes) => bytes,
            None => return Ok(None),
        };

        // Prefer the workbook-level pivotCaches mapping over filename guessing. In practice the
        // numeric suffix in `pivotCacheDefinitionN.xml` does not always line up with `cacheId`.
        //
        // Best-effort: tolerate malformed `xl/workbook.xml` by falling back to filename guessing
        // instead of returning an error.
        let rel_id = match workbook_pivot_cache_rel_id(workbook_xml, cache_id) {
            Ok(rel_id) => rel_id,
            Err(_) => None,
        };
        if let Some(rel_id) = rel_id {
            // Best-effort: if the relationships part is malformed (or otherwise unparsable), treat
            // it as "mapping cannot be resolved" and fall back to filename guessing.
            if let Ok(Some(part_name)) =
                resolve_relationship_target(self, "xl/workbook.xml", &rel_id)
            {
                if let Some(bytes) = self.part(&part_name) {
                    return Ok(Some((part_name, parse_pivot_cache_definition(bytes)?)));
                }
            }
        }

        // Fall back to the historical filename guess only when the workbook mapping is missing
        // or cannot be resolved.
        let guess = format!("xl/pivotCache/pivotCacheDefinition{cache_id}.xml");
        let Some(bytes) = self.part(&guess) else {
            return Ok(None);
        };
        Ok(Some((guess, parse_pivot_cache_definition(bytes)?)))
    }

    /// Parse a single pivot cache definition part.
    pub fn pivot_cache_definition(
        &self,
        part_name: &str,
    ) -> Result<Option<PivotCacheDefinition>, XlsxError> {
        let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
        let Some(bytes) = self.part(part_name) else {
            return Ok(None);
        };
        Ok(Some(parse_pivot_cache_definition(bytes)?))
    }
}

impl XlsxDocument {
    /// Parse every pivot cache definition part preserved in the document.
    ///
    /// Returns a sorted list of `(part_name, parsed_definition)` pairs.
    pub fn pivot_cache_definitions(
        &self,
    ) -> Result<Vec<(String, PivotCacheDefinition)>, XlsxError> {
        let mut paths: BTreeSet<String> = BTreeSet::new();
        for name in self.parts().keys() {
            if name.starts_with("xl/pivotCache/")
                && name.contains("pivotCacheDefinition")
                && name.ends_with(".xml")
            {
                paths.insert(name.to_string());
            }
        }

        let mut out = Vec::new();
        for path in paths {
            let Some(bytes) = self.parts().get(&path) else {
                continue;
            };
            out.push((path, parse_pivot_cache_definition(bytes)?));
        }
        Ok(out)
    }

    /// Parse a single pivot cache definition part preserved in the document.
    pub fn pivot_cache_definition(
        &self,
        part_name: &str,
    ) -> Result<Option<PivotCacheDefinition>, XlsxError> {
        let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
        let Some(bytes) = self.parts().get(part_name) else {
            return Ok(None);
        };
        Ok(Some(parse_pivot_cache_definition(bytes)?))
    }
}

fn parse_pivot_cache_definition(xml: &[u8]) -> Result<PivotCacheDefinition, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut nested_buf = Vec::new();
    let mut skip_buf = Vec::new();
    let mut def = PivotCacheDefinition::default();
    let mut current_field_idx: Option<usize> = None;
    let mut in_shared_items = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) => {
                let tag = e.local_name();
                let tag = tag.as_ref();

                if in_shared_items {
                    if let Some(field_idx) = current_field_idx {
                        if let Some(item) = parse_shared_item_start(
                            &mut reader,
                            &e,
                            &mut nested_buf,
                            &mut skip_buf,
                        )? {
                            if let Some(field) = def.cache_fields.get_mut(field_idx) {
                                field.shared_items.get_or_insert_with(Vec::new).push(item);
                            }
                        }
                    } else {
                        skip_to_end(&mut reader, e.name(), &mut skip_buf);
                    }
                } else if tag.eq_ignore_ascii_case(b"cacheField") {
                    let field = parse_cache_field(&e)?;
                    def.cache_fields.push(field);
                    current_field_idx = def.cache_fields.len().checked_sub(1);
                    in_shared_items = false;
                } else if tag.eq_ignore_ascii_case(b"sharedItems") {
                    // `sharedItems` appears as a child of `cacheField`. Record that we should treat
                    // the upcoming elements as shared item values until we hit `</sharedItems>`.
                    if let Some(field_idx) = current_field_idx {
                        if let Some(field) = def.cache_fields.get_mut(field_idx) {
                            field.shared_items.get_or_insert_with(Vec::new);
                        }
                        in_shared_items = true;
                    } else {
                        in_shared_items = false;
                    }
                } else {
                    handle_element(&mut def, &e)?;
                }
            }
            Event::Empty(e) => {
                let tag = e.local_name();
                let tag = tag.as_ref();

                if in_shared_items {
                    if let Some(field_idx) = current_field_idx {
                        if let Some(item) = parse_shared_item_empty(&e) {
                            if let Some(field) = def.cache_fields.get_mut(field_idx) {
                                field.shared_items.get_or_insert_with(Vec::new).push(item);
                            }
                        }
                    }
                } else if tag.eq_ignore_ascii_case(b"cacheField") {
                    let field = parse_cache_field(&e)?;
                    def.cache_fields.push(field);
                } else if tag.eq_ignore_ascii_case(b"sharedItems") {
                    // Empty `<sharedItems/>` list.
                    if let Some(field_idx) = current_field_idx {
                        if let Some(field) = def.cache_fields.get_mut(field_idx) {
                            field.shared_items.get_or_insert_with(Vec::new);
                        }
                    }
                } else {
                    handle_element(&mut def, &e)?;
                }
            }
            Event::End(e) => {
                let tag = e.local_name();
                let tag = tag.as_ref();
                if tag.eq_ignore_ascii_case(b"cacheField") {
                    current_field_idx = None;
                    in_shared_items = false;
                } else if tag.eq_ignore_ascii_case(b"sharedItems") {
                    in_shared_items = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(def)
}

fn parse_shared_item_empty(e: &BytesStart<'_>) -> Option<PivotCacheValue> {
    let local_name = e.local_name();
    let local_name = local_name.as_ref();

    match local_name {
        b"m" => Some(PivotCacheValue::Missing),
        b"n" => Some(parse_shared_number(attr_value_local(e, b"v"))),
        b"s" => Some(parse_shared_string(attr_value_local(e, b"v"))),
        b"b" => Some(parse_shared_bool(attr_value_local(e, b"v"))),
        b"e" => Some(parse_shared_error(attr_value_local(e, b"v"))),
        b"d" => Some(parse_shared_datetime(attr_value_local(e, b"v"))),
        // `<x>` is record-level (shared item index) and should not be treated as a shared item.
        b"x" => None,
        // Preserve unknown tags as text so shared-items indices remain stable.
        _ => Some(PivotCacheValue::String(
            attr_value_local(e, b"v").unwrap_or_default(),
        )),
    }
}

fn parse_shared_item_start<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    e: &BytesStart<'_>,
    buf: &mut Vec<u8>,
    skip_buf: &mut Vec<u8>,
) -> Result<Option<PivotCacheValue>, XlsxError> {
    let local_name = e.local_name();
    let local_name = local_name.as_ref();

    // Most shared items are self-closing tags (`Event::Empty`), but some producers emit
    // `<s><v>...</v></s>` or `<n>42</n>`.
    let attr_v = attr_value_local(e, b"v");

    let mut value_text =
        |reader: &mut Reader<R>| read_value_text_from_element(reader, e.name(), buf, skip_buf);

    match local_name {
        b"m" => {
            skip_to_end(reader, e.name(), skip_buf);
            Ok(Some(PivotCacheValue::Missing))
        }
        b"n" => {
            let v = match attr_v {
                Some(v) => {
                    skip_to_end(reader, e.name(), skip_buf);
                    Some(v)
                }
                None => value_text(reader)?,
            };
            Ok(Some(parse_shared_number(v)))
        }
        b"d" => {
            let v = match attr_v {
                Some(v) => {
                    skip_to_end(reader, e.name(), skip_buf);
                    Some(v)
                }
                None => value_text(reader)?,
            };
            Ok(Some(parse_shared_datetime(v)))
        }
        // `<x>` is record-level (shared item index) and should not be treated as a shared item.
        b"x" => {
            // Still advance the reader so the outer parse loop stays in sync.
            skip_to_end(reader, e.name(), skip_buf);
            Ok(None)
        }
        b"s" => {
            let v = match attr_v {
                Some(v) => {
                    skip_to_end(reader, e.name(), skip_buf);
                    Some(v)
                }
                None => value_text(reader)?,
            };
            Ok(Some(parse_shared_string(v)))
        }
        b"e" => {
            let v = match attr_v {
                Some(v) => {
                    skip_to_end(reader, e.name(), skip_buf);
                    Some(v)
                }
                None => value_text(reader)?,
            };
            Ok(Some(parse_shared_error(v)))
        }
        b"b" => {
            let v = match attr_v {
                Some(v) => {
                    skip_to_end(reader, e.name(), skip_buf);
                    Some(v)
                }
                None => value_text(reader)?,
            };
            Ok(Some(parse_shared_bool(v)))
        }
        _ => {
            // Preserve unknown tags as text so shared-items indices remain stable.
            let v = match attr_v {
                Some(v) => {
                    skip_to_end(reader, e.name(), skip_buf);
                    v
                }
                None => value_text(reader)?.unwrap_or_default(),
            };
            Ok(Some(PivotCacheValue::String(v)))
        }
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

fn parse_shared_number(v: Option<String>) -> PivotCacheValue {
    let Some(v) = v else {
        return PivotCacheValue::Missing;
    };
    let Ok(n) = v.trim().parse::<f64>() else {
        return PivotCacheValue::Missing;
    };
    PivotCacheValue::Number(n)
}

fn parse_shared_datetime(v: Option<String>) -> PivotCacheValue {
    let Some(v) = v else {
        return PivotCacheValue::Missing;
    };
    PivotCacheValue::DateTime(v)
}

fn parse_shared_string(v: Option<String>) -> PivotCacheValue {
    let Some(v) = v else {
        return PivotCacheValue::Missing;
    };
    PivotCacheValue::String(v)
}

fn parse_shared_error(v: Option<String>) -> PivotCacheValue {
    let Some(v) = v else {
        return PivotCacheValue::Missing;
    };
    PivotCacheValue::Error(v)
}

fn parse_shared_bool(v: Option<String>) -> PivotCacheValue {
    let Some(v) = v else {
        return PivotCacheValue::Missing;
    };
    let v = v.trim();
    PivotCacheValue::Bool(v == "1" || v.eq_ignore_ascii_case("true"))
}

fn skip_to_end<R: std::io::BufRead>(reader: &mut Reader<R>, end: QName<'_>, buf: &mut Vec<u8>) {
    buf.clear();
    let _ = reader.read_to_end_into(end, buf);
}

fn read_value_text_from_element<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    outer_end: QName<'_>,
    buf: &mut Vec<u8>,
    skip_buf: &mut Vec<u8>,
) -> Result<Option<String>, XlsxError> {
    let mut value: Option<String> = None;

    loop {
        let event = reader.read_event_into(buf)?;

        match event {
            Event::Start(e) => {
                let e = e.into_owned();
                buf.clear();

                if e.local_name().as_ref() == b"v" {
                    let v = read_text_to_end(reader, e.name(), buf, skip_buf)?;
                    if value.is_none() {
                        value = Some(v);
                    }
                } else {
                    // Skip unknown nested elements inside the value wrapper.
                    skip_to_end(reader, e.name(), skip_buf);
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
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(value)
}

fn read_text_to_end<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    end: QName<'_>,
    buf: &mut Vec<u8>,
    skip_buf: &mut Vec<u8>,
) -> Result<String, XlsxError> {
    let mut text = String::new();

    loop {
        let event = reader.read_event_into(buf)?;

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
                buf.clear();
                // Keep the parser resilient by skipping nested elements.
                skip_to_end(reader, e.name(), skip_buf);
            }
            Event::End(e) if e.name() == end => break,
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(text)
}

fn workbook_pivot_cache_rel_id(xml: &[u8], cache_id: u32) -> Result<Option<String>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) | Event::Empty(e) => {
                if e.local_name().as_ref().eq_ignore_ascii_case(b"pivotCache") {
                    let mut found_cache_id = None;
                    let mut rel_id = None;

                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(quick_xml::Error::from)?;
                        let key = attr.key.local_name();
                        let key = key.as_ref();
                        let value = attr.unescape_value()?.into_owned();

                        if key.eq_ignore_ascii_case(b"cacheId") {
                            found_cache_id = value.parse::<u32>().ok();
                        } else if key.eq_ignore_ascii_case(b"id") {
                            rel_id = Some(value);
                        }
                    }

                    if found_cache_id == Some(cache_id) {
                        return Ok(rel_id);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(None)
}

fn handle_element(def: &mut PivotCacheDefinition, e: &BytesStart<'_>) -> Result<(), XlsxError> {
    let tag = e.local_name();
    let tag = tag.as_ref();

    if tag.eq_ignore_ascii_case(b"pivotCacheDefinition") {
        for attr in e.attributes().with_checks(false) {
            let attr = attr.map_err(quick_xml::Error::from)?;
            let key = attr.key.local_name();
            let key = key.as_ref();
            let value = attr.unescape_value()?;

            if key.eq_ignore_ascii_case(b"recordCount") {
                def.record_count = value.parse::<u64>().ok();
            } else if key.eq_ignore_ascii_case(b"refreshOnLoad") {
                def.refresh_on_load = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"createdVersion") {
                def.created_version = value.parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"refreshedVersion") {
                def.refreshed_version = value.parse::<u32>().ok();
            }
        }
    } else if tag.eq_ignore_ascii_case(b"cacheSource") {
        for attr in e.attributes().with_checks(false) {
            let attr = attr.map_err(quick_xml::Error::from)?;
            let key = attr.key.local_name();
            let key = key.as_ref();
            if key.eq_ignore_ascii_case(b"type") {
                let raw_value = attr.unescape_value()?.to_string();
                let value = raw_value.to_ascii_lowercase();
                def.cache_source_type = match value.as_str() {
                    "worksheet" => PivotCacheSourceType::Worksheet,
                    "external" => PivotCacheSourceType::External,
                    "consolidation" => PivotCacheSourceType::Consolidation,
                    "scenario" => PivotCacheSourceType::Scenario,
                    _ => PivotCacheSourceType::Unknown(raw_value),
                };
            } else if key.eq_ignore_ascii_case(b"connectionId") {
                def.cache_source_connection_id = attr.unescape_value()?.parse::<u32>().ok();
            }
        }
    } else if tag.eq_ignore_ascii_case(b"worksheetSource") {
        let mut sheet: Option<String> = None;
        let mut reference: Option<String> = None;
        let mut name: Option<String> = None;
        for attr in e.attributes().with_checks(false) {
            let attr = attr.map_err(quick_xml::Error::from)?;
            let key = attr.key.local_name();
            let key = key.as_ref();
            let value = attr.unescape_value()?.to_string();
            if key.eq_ignore_ascii_case(b"sheet") {
                sheet = Some(value);
            } else if key.eq_ignore_ascii_case(b"ref") {
                reference = Some(value);
            } else if key.eq_ignore_ascii_case(b"name") {
                name = Some(value);
            }
        }

        // Some non-standard producers encode the sheet in the ref (e.g. `Sheet1!A1:C5`)
        // instead of using the `sheet="..."` attribute.
        if sheet.is_none() {
            if let Some(ref_value) = reference.as_deref() {
                if let Some((parsed_sheet, parsed_ref)) = split_sheet_ref(ref_value) {
                    sheet = Some(parsed_sheet);
                    reference = Some(parsed_ref);
                }
            }
        }

        def.worksheet_source_sheet = sheet;
        def.worksheet_source_ref = reference.or(name);
    }
    Ok(())
}

fn parse_cache_field(e: &BytesStart<'_>) -> Result<PivotCacheField, XlsxError> {
    let mut field = PivotCacheField::default();
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(quick_xml::Error::from)?;
        let key = attr.key.local_name();
        let key = key.as_ref();
        let value = attr.unescape_value()?;
        if key.eq_ignore_ascii_case(b"name") {
            field.name = value.to_string();
        } else if key.eq_ignore_ascii_case(b"caption") {
            field.caption = Some(value.to_string());
        } else if key.eq_ignore_ascii_case(b"propertyName") {
            field.property_name = Some(value.to_string());
        } else if key.eq_ignore_ascii_case(b"numFmtId") {
            field.num_fmt_id = value.parse::<u32>().ok();
        } else if key.eq_ignore_ascii_case(b"databaseField") {
            field.database_field = parse_bool(&value);
        } else if key.eq_ignore_ascii_case(b"serverField") {
            field.server_field = parse_bool(&value);
        } else if key.eq_ignore_ascii_case(b"uniqueList") {
            field.unique_list = parse_bool(&value);
        } else if key.eq_ignore_ascii_case(b"formula") {
            field.formula = Some(value.to_string());
        } else if key.eq_ignore_ascii_case(b"sqlType") {
            field.sql_type = value.parse::<i32>().ok();
        } else if key.eq_ignore_ascii_case(b"hierarchy") {
            field.hierarchy = value.parse::<u32>().ok();
        } else if key.eq_ignore_ascii_case(b"level") {
            field.level = value.parse::<u32>().ok();
        } else if key.eq_ignore_ascii_case(b"mappingCount") {
            field.mapping_count = value.parse::<u32>().ok();
        }
    }
    Ok(field)
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

pub(crate) fn split_sheet_ref(reference: &str) -> Option<(String, String)> {
    let (sheet_part, ref_part) = reference.rsplit_once('!')?;
    if ref_part.is_empty() {
        return None;
    }

    let sheet_part = sheet_part.trim();
    let sheet_part = if let Some(stripped) = sheet_part
        .strip_prefix('\'')
        .and_then(|s| s.strip_suffix('\''))
    {
        stripped.replace("''", "'")
    } else {
        sheet_part.to_string()
    };

    Some((sheet_part, ref_part.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    use pretty_assertions::assert_eq;
    use formula_engine::pivot::PivotFieldRef;

    #[test]
    fn parses_shared_items_and_resolves_indices() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheFields count="1">
    <cacheField name="Field1" numFmtId="0">
      <sharedItems count="6">
        <m/>
        <n v="42"/>
        <b v="1"/>
        <s v="Hello"/>
        <d v="2024-01-15T00:00:00Z"/>
        <weird v="??"/>
      </sharedItems>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>"#;

        let def = parse_pivot_cache_definition(xml).expect("parse");
        assert_eq!(def.cache_fields.len(), 1);
        let field = &def.cache_fields[0];
        assert_eq!(
            field.shared_items,
            Some(vec![
                PivotCacheValue::Missing,
                PivotCacheValue::Number(42.0),
                PivotCacheValue::Bool(true),
                PivotCacheValue::String("Hello".to_string()),
                PivotCacheValue::DateTime("2024-01-15T00:00:00Z".to_string()),
                PivotCacheValue::String("??".to_string()),
            ])
        );

        assert_eq!(def.resolve_shared_item(0, 0), Some(ScalarValue::Blank));
        assert_eq!(def.resolve_shared_item(0, 1), Some(ScalarValue::from(42.0)));
        assert_eq!(def.resolve_shared_item(0, 2), Some(ScalarValue::Bool(true)));
        assert_eq!(
            def.resolve_shared_item(0, 3),
            Some(ScalarValue::from("Hello"))
        );
        assert_eq!(
            def.resolve_shared_item(0, 4),
            Some(ScalarValue::Date(
                NaiveDate::from_ymd_opt(2024, 1, 15).unwrap()
            ))
        );
        assert_eq!(def.resolve_shared_item(0, 5), Some(ScalarValue::from("??")));
    }

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
        assert_eq!(def.cache_source_connection_id, None);
        assert_eq!(def.worksheet_source_sheet.as_deref(), Some("Sheet1"));
        assert_eq!(def.worksheet_source_ref.as_deref(), Some("A1:B2"));
        assert_eq!(def.cache_fields.len(), 1);
        assert_eq!(def.cache_fields[0].name, "Field1");
    }

    #[test]
    fn parses_cache_field_common_attributes() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheFields count="1">
    <cacheField name="Field1" caption="Caption" propertyName="Prop" numFmtId="5" databaseField="1" serverField="0" uniqueList="1" formula="=A1" sqlType="4" hierarchy="2" level="3" mappingCount="7"/>
  </cacheFields>
</pivotCacheDefinition>"#;

        let def = parse_pivot_cache_definition(xml).expect("parse");
        assert_eq!(def.cache_fields.len(), 1);
        let field = &def.cache_fields[0];
        assert_eq!(field.name, "Field1");
        assert_eq!(field.caption.as_deref(), Some("Caption"));
        assert_eq!(field.property_name.as_deref(), Some("Prop"));
        assert_eq!(field.num_fmt_id, Some(5));
        assert_eq!(field.database_field, Some(true));
        assert_eq!(field.server_field, Some(false));
        assert_eq!(field.unique_list, Some(true));
        assert_eq!(field.formula.as_deref(), Some("=A1"));
        assert_eq!(field.sql_type, Some(4));
        assert_eq!(field.hierarchy, Some(2));
        assert_eq!(field.level, Some(3));
        assert_eq!(field.mapping_count, Some(7));
    }

    #[test]
    fn parses_cache_field_shared_items() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheFields count="1">
    <cacheField name="Field1">
      <sharedItems count="6">
        <m/>
        <n v="42"/>
        <n>43</n>
        <s v="Hello"/>
        <s>World</s>
        <b v="1"/>
      </sharedItems>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>"#;

        let def = parse_pivot_cache_definition(xml).expect("parse");
        assert_eq!(def.cache_fields.len(), 1);
        assert_eq!(
            def.cache_fields[0].shared_items,
            Some(vec![
                PivotCacheValue::Missing,
                PivotCacheValue::Number(42.0),
                PivotCacheValue::Number(43.0),
                PivotCacheValue::String("Hello".to_string()),
                PivotCacheValue::String("World".to_string()),
                PivotCacheValue::Bool(true),
            ])
        );
    }

    #[test]
    fn parses_cache_source_connection_id() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="external" connectionId="42"/>
</pivotCacheDefinition>"#;

        let def = parse_pivot_cache_definition(xml).expect("parse");
        assert_eq!(def.cache_source_type, PivotCacheSourceType::External);
        assert_eq!(def.cache_source_connection_id, Some(42));
    }

    #[test]
    fn parses_sheet_from_ref_when_sheet_attr_missing() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet">
    <worksheetSource ref="'Sheet 1'!A1:C5"/>
  </cacheSource>
</pivotCacheDefinition>"#;

        let def = parse_pivot_cache_definition(xml).expect("parse");
        assert_eq!(def.cache_source_type, PivotCacheSourceType::Worksheet);
        assert_eq!(def.worksheet_source_sheet.as_deref(), Some("Sheet 1"));
        assert_eq!(def.worksheet_source_ref.as_deref(), Some("A1:C5"));
    }

    #[test]
    fn resolves_record_value_valid_shared_item_index() {
        let def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Field1".to_string(),
                shared_items: Some(vec![
                    PivotCacheValue::String("A".to_string()),
                    PivotCacheValue::Number(42.0),
                ]),
                ..PivotCacheField::default()
            }],
            ..PivotCacheDefinition::default()
        };

        assert_eq!(
            def.resolve_record_value(0, PivotCacheValue::Index(1)),
            PivotCacheValue::Number(42.0)
        );
    }

    #[test]
    fn resolves_record_values_iterator_uses_field_position() {
        let def = PivotCacheDefinition {
            cache_fields: vec![
                PivotCacheField {
                    name: "Field1".to_string(),
                    shared_items: Some(vec![PivotCacheValue::String("A".to_string())]),
                    ..PivotCacheField::default()
                },
                PivotCacheField {
                    name: "Field2".to_string(),
                    shared_items: Some(vec![
                        PivotCacheValue::Number(42.0),
                        PivotCacheValue::Number(43.0),
                    ]),
                    ..PivotCacheField::default()
                },
            ],
            ..PivotCacheDefinition::default()
        };

        let record = vec![PivotCacheValue::Index(0), PivotCacheValue::Index(1)];
        let resolved = def.resolve_record_values(record).collect::<Vec<_>>();
        assert_eq!(
            resolved,
            vec![PivotCacheValue::String("A".to_string()), PivotCacheValue::Number(43.0)]
        );
    }

    #[test]
    fn resolves_record_value_non_index_passthrough() {
        let def = PivotCacheDefinition::default();
        let value = PivotCacheValue::Number(1.234);
        assert_eq!(
            def.resolve_record_value(0, value.clone()),
            value
        );
    }

    #[test]
    fn resolves_record_value_when_cache_field_is_missing() {
        let def = PivotCacheDefinition::default();
        assert_eq!(
            def.resolve_record_value(0, PivotCacheValue::Index(0)),
            PivotCacheValue::Missing
        );
    }

    #[test]
    fn resolves_record_value_out_of_range_shared_item_index() {
        let def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Field1".to_string(),
                shared_items: Some(vec![PivotCacheValue::String("A".to_string())]),
                ..PivotCacheField::default()
            }],
            ..PivotCacheDefinition::default()
        };

        assert_eq!(
            def.resolve_record_value(0, PivotCacheValue::Index(5)),
            PivotCacheValue::Missing
        );
    }

    #[test]
    fn resolves_record_value_when_cache_field_has_no_shared_items() {
        let def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Field1".to_string(),
                shared_items: None,
                ..PivotCacheField::default()
            }],
            ..PivotCacheDefinition::default()
        };

        assert_eq!(
            def.resolve_record_value(0, PivotCacheValue::Index(0)),
            PivotCacheValue::Missing
        );
    }

    #[test]
    fn imports_calculated_fields_into_engine_config() {
        let cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheFields count="3">
    <cacheField name="Sales"/>
    <cacheField name="Cost"/>
    <cacheField name="Profit" formula="Sales-Cost"/>
  </cacheFields>
</pivotCacheDefinition>"#;

        let cache_def = parse_pivot_cache_definition(cache_xml).expect("parse cache definition");
        assert_eq!(
            cache_def.calculated_fields(),
            vec![formula_model::pivots::CalculatedField {
                name: "Profit".to_string(),
                formula: "Sales-Cost".to_string(),
            }]
        );

        let table_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataFields count="1">
    <dataField fld="2" subtotal="sum" name="Sum of Profit"/>
  </dataFields>
</pivotTableDefinition>"#;

        let table =
            crate::pivots::PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", table_xml)
                .expect("parse pivot table definition");

        let cfg = crate::pivots::engine_bridge::pivot_table_to_engine_config(&table, &cache_def);
        assert_eq!(
            cfg.calculated_fields,
            vec![formula_engine::pivot::CalculatedField {
                name: "Profit".to_string(),
                formula: "Sales-Cost".to_string(),
            }]
        );
        assert_eq!(cfg.value_fields.len(), 1);
        assert_eq!(
            cfg.value_fields[0].source_field,
            PivotFieldRef::CacheFieldName("Profit".to_string())
        );
    }

    #[test]
    fn calculated_fields_strip_leading_equals() {
        let cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheFields count="1">
    <cacheField name="Profit" formula="=Sales-Cost"/>
  </cacheFields>
</pivotCacheDefinition>"#;

        let cache_def = parse_pivot_cache_definition(cache_xml).expect("parse cache definition");
        assert_eq!(
            cache_def.calculated_fields(),
            vec![formula_model::pivots::CalculatedField {
                name: "Profit".to_string(),
                formula: "Sales-Cost".to_string(),
            }]
        );
    }
}
