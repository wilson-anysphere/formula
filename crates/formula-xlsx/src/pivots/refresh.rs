use std::collections::{BTreeMap, HashMap};
use std::io::Cursor;

use formula_model::{CellRef, Range};
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml::{local_name, parse_relationships, rels_part_name, resolve_target};
use crate::package::{XlsxError, XlsxPackage};
use crate::pivots::cache_records::PivotCacheValue;
use crate::shared_strings::parse_shared_strings_xml;
use crate::xml::{QName, XmlElement, XmlNode};

const WORKBOOK_PART: &str = "xl/workbook.xml";
const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

#[derive(Debug, Clone, PartialEq, Eq)]
struct WorksheetSource {
    sheet: String,
    reference: String,
}

fn pivot_cache_value_is_blank(value: &PivotCacheValue) -> bool {
    match value {
        PivotCacheValue::Missing => true,
        PivotCacheValue::String(s) => s.is_empty(),
        PivotCacheValue::Number(_)
        | PivotCacheValue::Bool(_)
        | PivotCacheValue::Error(_)
        | PivotCacheValue::DateTime(_)
        | PivotCacheValue::Index(_) => false,
    }
}

fn pivot_cache_value_header_text(value: &PivotCacheValue) -> String {
    match value {
        PivotCacheValue::String(s) => s.clone(),
        PivotCacheValue::Number(n) => format_pivot_cache_number(*n),
        PivotCacheValue::Bool(true) => "TRUE".to_string(),
        PivotCacheValue::Bool(false) => "FALSE".to_string(),
        PivotCacheValue::Error(e) => e.clone(),
        PivotCacheValue::DateTime(dt) => dt.clone(),
        PivotCacheValue::Index(idx) => idx.to_string(),
        PivotCacheValue::Missing => String::new(),
    }
}

fn format_pivot_cache_number(value: f64) -> String {
    // Use Excel-like "General" formatting for stable `<n v="..."/>` output.
    //
    // `f64::to_string()` prefers scientific notation for some magnitudes; Excel's pivot cache
    // records generally use a more human-friendly fixed-point form when possible.
    if !value.is_finite() {
        return value.to_string();
    }
    formula_format::format_value(
        formula_format::Value::Number(value),
        Some("General"),
        &formula_format::FormatOptions::default(),
    )
    .text
}

impl XlsxPackage {
    /// Refresh a pivot cache's cache definition/records from the referenced worksheet source range.
    ///
    /// This helper is intentionally conservative: it updates `recordCount`, `cacheFields`, and
    /// rewrites the cache records payload, while preserving unrelated XML nodes/attributes where
    /// possible.
    pub fn refresh_pivot_cache_from_worksheet(
        &mut self,
        cache_definition_part: &str,
    ) -> Result<(), XlsxError> {
        let cache_definition_bytes = self
            .part(cache_definition_part)
            .ok_or_else(|| XlsxError::MissingPart(cache_definition_part.to_string()))?
            .to_vec();

        let worksheet_source =
            parse_worksheet_source_from_cache_definition(&cache_definition_bytes)?;
        let range = Range::from_a1(&worksheet_source.reference).map_err(|e| {
            XlsxError::Invalid(format!(
                "invalid worksheetSource ref {ref_:?}: {e}",
                ref_ = worksheet_source.reference
            ))
        })?;

        let worksheet_part = resolve_worksheet_part(self, &worksheet_source.sheet)?;
        let worksheet_xml = self
            .part(&worksheet_part)
            .ok_or_else(|| XlsxError::MissingPart(worksheet_part.clone()))?;

        let shared_strings = load_shared_strings(self)?;
        let cells = parse_worksheet_cells_in_range(worksheet_xml, range, &shared_strings)?;
        let (field_names, records) = build_cache_fields_and_records(range, &cells);

        let record_count = records.len() as u64;

        let cache_records_part = resolve_cache_records_part(self, cache_definition_part)?;
        let cache_records_bytes = self
            .part(&cache_records_part)
            .ok_or_else(|| XlsxError::MissingPart(cache_records_part.clone()))?
            .to_vec();

        let updated_cache_definition =
            refresh_cache_definition_xml(&cache_definition_bytes, &field_names, record_count)?;
        let updated_cache_records =
            refresh_cache_records_xml(&cache_records_bytes, &records, record_count)?;

        self.set_part(cache_definition_part.to_string(), updated_cache_definition);
        self.set_part(cache_records_part, updated_cache_records);
        Ok(())
    }
}

fn parse_worksheet_source_from_cache_definition(xml: &[u8]) -> Result<WorksheetSource, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut sheet = None;
    let mut reference = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) => {
                if local_name(e.name().as_ref()) == b"worksheetSource" {
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr?;
                        match local_name(attr.key.as_ref()) {
                            b"sheet" => sheet = Some(attr.unescape_value()?.into_owned()),
                            b"ref" => reference = Some(attr.unescape_value()?.into_owned()),
                            _ => {}
                        }
                    }
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    let sheet = sheet.ok_or(XlsxError::MissingAttr("worksheetSource@sheet"))?;
    let reference = reference.ok_or(XlsxError::MissingAttr("worksheetSource@ref"))?;
    Ok(WorksheetSource { sheet, reference })
}

fn resolve_worksheet_part(package: &XlsxPackage, sheet_name: &str) -> Result<String, XlsxError> {
    let sheets = package.workbook_sheets()?;
    let sheet = sheets
        .iter()
        .find(|s| s.name == sheet_name)
        .ok_or_else(|| XlsxError::Invalid(format!("sheet {sheet_name:?} not found in workbook")))?;

    let rels_part = rels_part_name("xl/workbook.xml");
    let rels_bytes = package
        .part(&rels_part)
        .ok_or_else(|| XlsxError::MissingPart(rels_part.clone()))?;
    let rels = parse_relationships(rels_bytes)?;
    let rel = rels.iter().find(|r| r.id == sheet.rel_id).ok_or_else(|| {
        XlsxError::Invalid(format!(
            "missing workbook relationship {rid:?} for sheet {sheet_name:?}",
            rid = sheet.rel_id
        ))
    })?;

    Ok(resolve_target("xl/workbook.xml", &rel.target))
}

fn resolve_cache_records_part(
    package: &XlsxPackage,
    cache_definition_part: &str,
) -> Result<String, XlsxError> {
    let rels_part = rels_part_name(cache_definition_part);
    let rels_bytes = package
        .part(&rels_part)
        .ok_or_else(|| XlsxError::MissingPart(rels_part.clone()))?;
    let rels = parse_relationships(rels_bytes)?;

    let Some(rel) = rels
        .into_iter()
        .find(|r| r.type_uri.ends_with("/pivotCacheRecords"))
    else {
        return Err(XlsxError::Invalid(format!(
            "pivot cache definition {cache_definition_part:?} does not reference a pivotCacheRecords relationship"
        )));
    };

    Ok(resolve_target(cache_definition_part, &rel.target))
}

fn load_shared_strings(package: &XlsxPackage) -> Result<Vec<String>, XlsxError> {
    let shared_strings_part = resolve_shared_strings_part_name(package)?.or_else(|| {
        package
            .part("xl/sharedStrings.xml")
            .map(|_| "xl/sharedStrings.xml".to_string())
    });
    let Some(shared_strings_part) = shared_strings_part else {
        return Ok(Vec::new());
    };
    let Some(bytes) = package.part(&shared_strings_part) else {
        return Ok(Vec::new());
    };

    let xml = std::str::from_utf8(bytes).map_err(|e| {
        XlsxError::Invalid(format!("{shared_strings_part:?} is not valid UTF-8: {e}"))
    })?;
    let parsed = parse_shared_strings_xml(xml)
        .map_err(|e| XlsxError::Invalid(format!("failed to parse {shared_strings_part:?}: {e}")))?;

    Ok(parsed.items.into_iter().map(|rt| rt.text).collect())
}

fn resolve_shared_strings_part_name(package: &XlsxPackage) -> Result<Option<String>, XlsxError> {
    let rels_part = rels_part_name(WORKBOOK_PART);
    let rels_bytes = match package.part(&rels_part) {
        Some(bytes) => bytes,
        None => return Ok(None),
    };
    let rels = parse_relationships(rels_bytes)?;
    Ok(rels
        .into_iter()
        .find(|rel| {
            rel.type_uri == REL_TYPE_SHARED_STRINGS
                && !rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        })
        .map(|rel| resolve_target(WORKBOOK_PART, &rel.target)))
}

fn parse_worksheet_cells_in_range(
    worksheet_xml: &[u8],
    range: Range,
    shared_strings: &[String],
) -> Result<HashMap<CellRef, PivotCacheValue>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(worksheet_xml));
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut in_sheet_data = false;

    let mut current_ref: Option<CellRef> = None;
    let mut current_t: Option<String> = None;
    let mut current_value_text: Option<String> = None;
    let mut current_inline_text: Option<String> = None;
    let mut in_v = false;

    let mut cells = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = true
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"sheetData" => in_sheet_data = false,
            Event::Empty(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = false;
                drop(e);
            }

            Event::Start(e) if in_sheet_data && local_name(e.name().as_ref()) == b"c" => {
                current_ref = None;
                current_t = None;
                current_value_text = None;
                current_inline_text = None;
                in_v = false;

                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    match local_name(attr.key.as_ref()) {
                        b"r" => {
                            let a1 = attr.unescape_value()?.into_owned();
                            let parsed = CellRef::from_a1(&a1).map_err(|e| {
                                XlsxError::Invalid(format!("invalid cell reference {a1:?}: {e}"))
                            })?;
                            current_ref = Some(parsed);
                        }
                        b"t" => current_t = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
            }
            Event::Empty(e) if in_sheet_data && local_name(e.name().as_ref()) == b"c" => {
                // We only care about values; empty `<c/>` entries represent blanks (possibly styled).
                drop(e);
            }

            Event::End(e) if in_sheet_data && local_name(e.name().as_ref()) == b"c" => {
                if let Some(cell_ref) = current_ref.take() {
                    if range.contains(cell_ref) {
                        let value = interpret_worksheet_cell_value(
                            current_t.as_deref(),
                            current_value_text.as_deref(),
                            current_inline_text.as_deref(),
                            shared_strings,
                        );
                        if !matches!(value, PivotCacheValue::Missing) {
                            cells.insert(cell_ref, value);
                        }
                    }
                }

                current_t = None;
                current_value_text = None;
                current_inline_text = None;
                in_v = false;
            }

            Event::Start(e)
                if in_sheet_data
                    && current_ref.is_some()
                    && local_name(e.name().as_ref()) == b"v" =>
            {
                in_v = true;
            }
            Event::End(e) if in_sheet_data && local_name(e.name().as_ref()) == b"v" => in_v = false,
            Event::Text(e) if in_sheet_data && in_v => {
                current_value_text = Some(e.unescape()?.into_owned());
            }

            Event::Start(e)
                if in_sheet_data
                    && current_ref.is_some()
                    && current_t.as_deref() == Some("inlineStr")
                    && local_name(e.name().as_ref()) == b"is" =>
            {
                current_inline_text = Some(parse_inline_is_text(&mut reader)?);
            }
            Event::Empty(e)
                if in_sheet_data
                    && current_ref.is_some()
                    && current_t.as_deref() == Some("inlineStr")
                    && local_name(e.name().as_ref()) == b"is" =>
            {
                current_inline_text = Some(String::new());
            }

            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(cells)
}

fn parse_inline_is_text<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<String, XlsxError> {
    let mut buf = Vec::new();
    let mut out = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"t" => {
                out.push_str(&read_text(reader, b"t")?);
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"r" => {
                out.push_str(&parse_inline_r_text(reader)?);
            }
            Event::Start(e) => {
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"is" => break,
            Event::Eof => {
                return Err(XlsxError::Invalid(
                    "unexpected EOF while parsing inline string <is>".to_string(),
                ))
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn parse_inline_r_text<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<String, XlsxError> {
    let mut buf = Vec::new();
    let mut out = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"t" => {
                out.push_str(&read_text(reader, b"t")?);
            }
            Event::Start(e) => {
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"r" => break,
            Event::Eof => {
                return Err(XlsxError::Invalid(
                    "unexpected EOF while parsing inline string <r>".to_string(),
                ))
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn read_text<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    end_local: &[u8],
) -> Result<String, XlsxError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => text.push_str(&e.unescape()?.into_owned()),
            Event::CData(e) => text.push_str(std::str::from_utf8(e.as_ref()).map_err(|err| {
                XlsxError::Invalid(format!("inline string <t> contains invalid utf-8: {err}"))
            })?),
            Event::End(e) if local_name(e.name().as_ref()) == end_local => break,
            Event::Eof => {
                return Err(XlsxError::Invalid(
                    "unexpected EOF while parsing inline string <t>".to_string(),
                ))
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

fn interpret_worksheet_cell_value(
    t: Option<&str>,
    v_text: Option<&str>,
    inline_text: Option<&str>,
    shared_strings: &[String],
) -> PivotCacheValue {
    match t {
        Some("s") => {
            let raw = v_text.unwrap_or_default();
            let idx = raw.parse::<usize>().unwrap_or(0);
            let text = shared_strings.get(idx).cloned().unwrap_or_default();
            PivotCacheValue::String(text)
        }
        Some("b") => {
            let raw = v_text.unwrap_or_default();
            PivotCacheValue::Bool(raw == "1" || raw.eq_ignore_ascii_case("true"))
        }
        Some("str") => PivotCacheValue::String(v_text.unwrap_or_default().to_string()),
        Some("inlineStr") => PivotCacheValue::String(inline_text.unwrap_or_default().to_string()),
        Some(_) | None => {
            let Some(raw) = v_text else {
                return PivotCacheValue::Missing;
            };

            match raw.parse::<f64>() {
                Ok(n) => PivotCacheValue::Number(n),
                Err(_) => PivotCacheValue::String(raw.to_string()),
            }
        }
    }
}

fn build_cache_fields_and_records(
    range: Range,
    cells: &HashMap<CellRef, PivotCacheValue>,
) -> (Vec<String>, Vec<Vec<PivotCacheValue>>) {
    let mut fields = Vec::new();
    for col in range.start.col..=range.end.col {
        let cell = CellRef::new(range.start.row, col);
        let value = cells
            .get(&cell)
            .cloned()
            .unwrap_or(PivotCacheValue::Missing);
        fields.push(pivot_cache_value_header_text(&value));
    }

    let mut records = Vec::new();
    if range.start.row < range.end.row {
        for row in (range.start.row + 1)..=range.end.row {
            let mut record = Vec::new();
            let mut all_blank = true;
            for col in range.start.col..=range.end.col {
                let cell = CellRef::new(row, col);
                let value = cells
                    .get(&cell)
                    .cloned()
                    .unwrap_or(PivotCacheValue::Missing);
                if !pivot_cache_value_is_blank(&value) {
                    all_blank = false;
                }
                record.push(value);
            }
            if !all_blank {
                records.push(record);
            }
        }
    }

    (fields, records)
}

fn refresh_cache_definition_xml(
    existing: &[u8],
    field_names: &[String],
    record_count: u64,
) -> Result<Vec<u8>, XlsxError> {
    let mut root = XmlElement::parse(existing).map_err(|e| {
        XlsxError::Invalid(format!("failed to parse pivot cache definition xml: {e}"))
    })?;

    root.set_attr("recordCount", record_count.to_string());

    if let Some(cache_fields) = root.child_mut("cacheFields") {
        cache_fields.set_attr("count", field_names.len().to_string());

        let old_children = std::mem::take(&mut cache_fields.children);
        let mut new_children = Vec::new();
        let mut field_idx = 0usize;
        for child in old_children {
            match child {
                XmlNode::Element(mut el) if el.name.local == "cacheField" => {
                    if field_idx >= field_names.len() {
                        continue;
                    }
                    el.set_attr("name", field_names[field_idx].clone());
                    field_idx += 1;
                    new_children.push(XmlNode::Element(el));
                }
                other => new_children.push(other),
            }
        }

        while field_idx < field_names.len() {
            new_children.push(XmlNode::Element(build_cache_field_element(
                cache_fields.name.ns.clone(),
                &field_names[field_idx],
            )));
            field_idx += 1;
        }

        cache_fields.children = new_children;
    } else {
        root.children
            .push(XmlNode::Element(build_cache_fields_element(
                root.name.ns.clone(),
                field_names,
            )));
    }

    Ok(root.to_xml_string().into_bytes())
}

fn build_cache_fields_element(ns: Option<String>, field_names: &[String]) -> XmlElement {
    let mut attrs = BTreeMap::new();
    attrs.insert(
        QName {
            ns: None,
            local: "count".to_string(),
        },
        field_names.len().to_string(),
    );

    let mut children = Vec::new();
    for name in field_names {
        children.push(XmlNode::Element(build_cache_field_element(
            ns.clone(),
            name,
        )));
    }

    XmlElement {
        name: QName {
            ns,
            local: "cacheFields".to_string(),
        },
        attrs,
        children,
    }
}

fn build_cache_field_element(ns: Option<String>, name: &str) -> XmlElement {
    let mut attrs = BTreeMap::new();
    attrs.insert(
        QName {
            ns: None,
            local: "name".to_string(),
        },
        name.to_string(),
    );
    attrs.insert(
        QName {
            ns: None,
            local: "numFmtId".to_string(),
        },
        "0".to_string(),
    );

    XmlElement {
        name: QName {
            ns,
            local: "cacheField".to_string(),
        },
        attrs,
        children: Vec::new(),
    }
}

fn refresh_cache_records_xml(
    existing: &[u8],
    records: &[Vec<PivotCacheValue>],
    record_count: u64,
) -> Result<Vec<u8>, XlsxError> {
    let mut root = XmlElement::parse(existing)
        .map_err(|e| XlsxError::Invalid(format!("failed to parse pivot cache records xml: {e}")))?;
    root.set_attr("count", record_count.to_string());

    let ns = root.name.ns.clone();
    let mut new_r_nodes: Vec<XmlNode> = records
        .iter()
        .map(|record| XmlNode::Element(build_record_element(ns.clone(), record)))
        .collect();

    let old_children = std::mem::take(&mut root.children);
    let mut children = Vec::new();
    let mut inserted = false;

    for child in old_children {
        let is_r = matches!(child, XmlNode::Element(ref el) if el.name.local == "r");
        if is_r {
            if !inserted {
                children.append(&mut new_r_nodes);
                inserted = true;
            }
            continue;
        }
        children.push(child);
    }

    if !inserted {
        children.append(&mut new_r_nodes);
    }
    root.children = children;

    Ok(root.to_xml_string().into_bytes())
}

fn build_record_element(ns: Option<String>, record: &[PivotCacheValue]) -> XmlElement {
    let mut out = XmlElement {
        name: QName {
            ns: ns.clone(),
            local: "r".to_string(),
        },
        attrs: BTreeMap::new(),
        children: Vec::new(),
    };

    for value in record {
        out.children.push(XmlNode::Element(match value {
            PivotCacheValue::String(s) => build_value_element(ns.clone(), "s", Some(s.as_str())),
            PivotCacheValue::Number(n) => {
                let formatted = format_pivot_cache_number(*n);
                build_value_element(ns.clone(), "n", Some(formatted.as_str()))
            }
            PivotCacheValue::Bool(b) => {
                build_value_element(ns.clone(), "b", Some(if *b { "1" } else { "0" }))
            }
            PivotCacheValue::Error(e) => build_value_element(ns.clone(), "e", Some(e.as_str())),
            PivotCacheValue::DateTime(dt) => build_value_element(ns.clone(), "d", Some(dt.as_str())),
            PivotCacheValue::Index(idx) => {
                let idx_str = idx.to_string();
                build_value_element(ns.clone(), "x", Some(idx_str.as_str()))
            }
            PivotCacheValue::Missing => build_value_element(ns.clone(), "m", None),
        }));
    }

    out
}

fn build_value_element(ns: Option<String>, tag: &str, value: Option<&str>) -> XmlElement {
    let mut attrs = BTreeMap::new();
    if let Some(value) = value {
        attrs.insert(
            QName {
                ns: None,
                local: "v".to_string(),
            },
            value.to_string(),
        );
    }

    XmlElement {
        name: QName {
            ns,
            local: tag.to_string(),
        },
        attrs,
        children: Vec::new(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::XlsxPackage;
    use std::io::{Cursor, Write};

    #[test]
    fn parse_inline_string_ignores_phonetic_text() {
        let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <t>Base</t>
          <phoneticPr fontId="0" type="noConversion"/>
          <rPh sb="0" eb="4"><t>PHO</t></rPh>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let range = Range::from_a1("A1:A1").expect("range");
        let cells = parse_worksheet_cells_in_range(worksheet_xml, range, &[]).expect("parse cells");
        let cell_ref = CellRef::from_a1("A1").expect("A1");
        assert_eq!(
            cells.get(&cell_ref),
            Some(&PivotCacheValue::String("Base".to_string()))
        );
    }

    #[test]
    fn resolve_shared_strings_part_name_ignores_external_relationship() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="https://example.com/sharedStrings.xml" TargetMode="External"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData/>
</worksheet>"#;

        let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/sharedStrings.xml", options).unwrap();
        zip.write_all(shared_strings_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        assert_eq!(
            super::resolve_shared_strings_part_name(&pkg).expect("resolve shared strings"),
            Some("xl/sharedStrings.xml".to_string())
        );
    }
}
